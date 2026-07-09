"""Spike harness — the media **long-tail**: do the `MediaFormat` playback setters
round-trip, and does the poster-frame API marshal?

The v1.7 media cut shipped insert + reads (`Muted`/`Volume`/`StartPoint`/`EndPoint`
were *read* by `media_video_spike.py`). This spike pins the **writes** the long-tail
needs, all on a temp audio clip (a synthesized WAV, no dependency):

1. **Mute** — `MediaFormat.Muted = msoTrue/msoFalse` round-trips.
2. **Volume** — `MediaFormat.Volume = 0.0..1.0` round-trips (what does out-of-range do?).
3. **Trim** — `MediaFormat.StartPoint` / `EndPoint` (milliseconds) settable + read back;
   note clamping vs. the clip `Length`, and whether EndPoint<StartPoint is rejected.
4. **Poster frame** — `MediaFormat.SetDisplayPicture(path)` (a video concept). We can't
   synthesize a real video without a dependency, so this only probes whether the method
   **marshals** on an audio shape (and how it fails) — enough to decide build-vs-defer.

Net-zero: the temp slide + audio are removed in `finally`; nothing on the user's deck
is mutated.

Run:  uv run python scripts/media_longtail_spike.py
"""

from __future__ import annotations

import base64
import json
import math
import os
import struct
import tempfile
from typing import Any

import pptlive as pl
from pptlive import _selection

_MSO_TRUE = -1
_MSO_FALSE = 0

# A 1x1 transparent PNG (base64) — a stand-in poster image, no dependency.
_PNG_1x1 = base64.b64decode(
    b"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _make_wav(path: str, *, seconds: float = 2.4, freq: float = 440.0, rate: int = 8000) -> None:
    n = int(seconds * rate)
    frames = b"".join(
        struct.pack("<h", int(0.3 * 32767 * math.sin(2 * math.pi * freq * i / rate)))
        for i in range(n)
    )
    data_size = len(frames)
    with open(path, "wb") as fh:
        fh.write(b"RIFF")
        fh.write(struct.pack("<I", 36 + data_size))
        fh.write(b"WAVE")
        fh.write(b"fmt ")
        fh.write(struct.pack("<IHHIIHH", 16, 1, 1, rate, rate * 2, 2, 16))
        fh.write(b"data")
        fh.write(struct.pack("<I", data_size))
        fh.write(frames)


def _read_mf(mf: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for name in ("Length", "Muted", "Volume", "StartPoint", "EndPoint"):
        try:
            out[name] = getattr(mf, name)
        except Exception as exc:
            out[name] = _err(exc)
    return out


def probe(slide_com: Any, wav_path: str, png_path: str) -> dict[str, Any]:
    out: dict[str, Any] = {}
    shape = slide_com.Shapes.AddMediaObject2(wav_path, False, True, 40.0, 40.0, 60.0, 60.0)
    mf = shape.MediaFormat
    out["initial"] = _read_mf(mf)

    # 1. Mute round-trip.
    try:
        mf.Muted = _MSO_TRUE
        muted_on = int(mf.Muted)
        mf.Muted = _MSO_FALSE
        muted_off = int(mf.Muted)
        out["mute"] = {"set_true_read": muted_on, "set_false_read": muted_off,
                       "round_trips": muted_on == _MSO_TRUE and muted_off == _MSO_FALSE}
    except Exception as exc:
        out["mute"] = {"error": _err(exc)}

    # 2. Volume round-trip + out-of-range behavior.
    try:
        mf.Volume = 0.5
        v_mid = float(mf.Volume)
        mf.Volume = 0.0
        v_lo = float(mf.Volume)
        mf.Volume = 1.0
        v_hi = float(mf.Volume)
        out["volume"] = {"set_0.5": v_mid, "set_0.0": v_lo, "set_1.0": v_hi}
        try:
            mf.Volume = 2.0
            out["volume"]["set_2.0_read"] = float(mf.Volume)
        except Exception as exc:
            out["volume"]["set_2.0_error"] = _err(exc)
        mf.Volume = 1.0
    except Exception as exc:
        out["volume"] = {"error": _err(exc)}

    # 3. Trim: StartPoint / EndPoint (ms).
    try:
        length = float(mf.Length)
        mf.StartPoint = 500.0
        mf.EndPoint = length - 500.0 if length > 1000 else length
        out["trim"] = {
            "length_ms": length,
            "start_read": float(mf.StartPoint),
            "end_read": float(mf.EndPoint),
        }
        # Over-length EndPoint — clamp or raise?
        try:
            mf.EndPoint = length + 5000.0
            out["trim"]["end_over_length_read"] = float(mf.EndPoint)
        except Exception as exc:
            out["trim"]["end_over_length_error"] = _err(exc)
        # EndPoint < StartPoint — rejected?
        try:
            mf.StartPoint = 1000.0
            mf.EndPoint = 200.0
            out["trim"]["end_lt_start_read"] = {"start": float(mf.StartPoint),
                                                "end": float(mf.EndPoint)}
        except Exception as exc:
            out["trim"]["end_lt_start_error"] = _err(exc)
        # Reset.
        mf.StartPoint = 0.0
        mf.EndPoint = length
    except Exception as exc:
        out["trim"] = {"error": _err(exc)}

    # 4. Poster frame — does SetDisplayPicture marshal (on an audio shape)?
    try:
        mf.SetDisplayPicture(png_path)
        out["poster"] = {"marshalled": True, "note": "accepted on audio shape"}
    except Exception as exc:
        out["poster"] = {"marshalled": False, "error": _err(exc)}

    return out


def main() -> int:
    findings: dict[str, Any] = {}
    wav_path: str | None = None
    png_path: str | None = None
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)

        snap = _selection.snapshot(ppt)

        fd, wav_path = tempfile.mkstemp(suffix=".wav", prefix="pptlive_lt_")
        os.close(fd)
        _make_wav(wav_path)
        fd2, png_path = tempfile.mkstemp(suffix=".png", prefix="pptlive_lt_")
        with os.fdopen(fd2, "wb") as fh:
            fh.write(_PNG_1x1)

        temp_ids: list[int] = []
        try:
            with deck.edit("media long-tail spike"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide = deck.slides[temp.index]
                findings["probe"] = probe(slide.com, wav_path, png_path)
        finally:
            try:
                with deck.edit("media long-tail spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            if deck.slides[idx].id in temp_ids:
                                deck.slides[idx].delete()
                        except Exception:
                            continue
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    for p in (wav_path, png_path):
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

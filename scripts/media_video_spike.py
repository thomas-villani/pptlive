"""Spike harness — the 'gold idea': insert media (audio narration) and export the
deck to **video** against real PowerPoint.

Tests the full LLM-authoring-a-narrated-video path, end to end, on a temp slide:

1. **Insert audio** — `Shapes.AddMediaObject2(FileName, LinkToFile,
   SaveWithDocument, Left, Top, Width, Height)` with a synthesized WAV (no
   dependency). Read back `Shape.MediaType` (ppMediaTypeSound=2), and
   `Shape.MediaFormat.Length` (the clip duration — needed to time the slide).
2. **Auto-play + timing** — `Shape.AnimationSettings.PlaySettings.PlayOnEntry`
   (= msoTrue, so the narration plays on slide entry), and set the slide to
   advance after the clip length (`SlideShowTransition.AdvanceOnTime` +
   `.AdvanceTime`) so the video paces itself to the narration.
3. **Export video** — `Presentation.CreateVideo(FileName,
   UseTimingsAndNarrations, DefaultSlideDuration, VertResolution,
   FramesPerSecond, Quality)`. THE marshalling question (PDF's
   `ExportAsFixedFormat` would NOT marshal late-bound — does `CreateVideo`?).
   It is **async**: poll `Presentation.CreateVideoStatus` (PpMediaTaskStatus:
   None=0, InProgress=1, Queued=2, Done=3, Failed=4) to completion, then confirm
   the .mp4 exists and is non-empty.

CreateVideo exports the WHOLE deck (incl. the user's real slide) — that's a
read, it does not mutate the deck — to a throwaway temp .mp4 that is deleted.
Net-zero on the deck: the temp slide + audio are removed in `finally`.

Run:  uv run python scripts/media_video_spike.py
"""

from __future__ import annotations

import json
import math
import os
import struct
import tempfile
import time
from typing import Any

import pptlive as pl
from pptlive import _selection

# PpMediaTaskStatus
_STATUS = {0: "None", 1: "InProgress", 2: "Queued", 3: "Done", 4: "Failed"}
# PpSaveAsFileType alternates (not run here, just documented)
_PP_SAVE_AS_MP4 = 39
_PP_SAVE_AS_WMV = 37


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _make_wav(path: str, *, seconds: float = 1.2, freq: float = 440.0, rate: int = 8000) -> None:
    """Write a tiny mono 16-bit PCM sine-tone WAV — a 'real' audible narration stand-in."""
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


def probe_insert_audio(slide_com: Any, wav_path: str) -> tuple[dict[str, Any], Any]:
    out: dict[str, Any] = {}
    audio_shape = None
    try:
        shapes_com = slide_com.Shapes
        try:
            audio_shape = shapes_com.AddMediaObject2(wav_path, False, True, 40.0, 40.0, 60.0, 60.0)
            out["api"] = "AddMediaObject2"
        except Exception as exc:
            out["addmediaobject2_error"] = _err(exc)
            # Legacy fallback
            audio_shape = shapes_com.AddMediaObject(wav_path, 40.0, 40.0, 60.0, 60.0)
            out["api"] = "AddMediaObject(legacy)"
        out["shape_name"] = str(audio_shape.Name)
        for name in ("MediaType",):
            try:
                out[name] = int(getattr(audio_shape, name))
            except Exception as exc:
                out[name] = _err(exc)
        # MediaFormat — clip length (ms) etc.
        try:
            mf = audio_shape.MediaFormat
            for name in ("Length", "Muted", "Volume", "StartPoint", "EndPoint"):
                try:
                    out[f"MediaFormat.{name}"] = getattr(mf, name)
                except Exception as exc:
                    out[f"MediaFormat.{name}"] = _err(exc)
        except Exception as exc:
            out["mediaformat_error"] = _err(exc)
        # Auto-play on entry (legacy AnimationSettings.PlaySettings).
        try:
            ps = audio_shape.AnimationSettings.PlaySettings
            ps.PlayOnEntry = -1  # msoTrue
            ps.HideWhileNotPlaying = -1
            out["play_on_entry_set"] = True
            out["play_on_entry_read"] = int(ps.PlayOnEntry)
        except Exception as exc:
            out["playsettings_error"] = _err(exc)
    except Exception as exc:
        out["error"] = _err(exc)
    return out, audio_shape


def probe_video_export(deck_com: Any, mp4_path: str, *, timeout_s: float = 120.0) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        # CreateVideo(FileName, UseTimingsAndNarrations, DefaultSlideDuration,
        #             VertResolution, FramesPerSecond, Quality)
        deck_com.CreateVideo(mp4_path, True, 3, 480, 24, 85)
        out["createvideo_marshalled"] = True
        statuses: list[str] = []
        t0 = time.time()
        last = None
        while time.time() - t0 < timeout_s:
            try:
                st = int(deck_com.CreateVideoStatus)
            except Exception as exc:
                out["status_read_error"] = _err(exc)
                break
            label = _STATUS.get(st, f"?{st}")
            if label != last:
                statuses.append(f"{round(time.time()-t0,1)}s:{label}")
                last = label
            if st in (3, 4):  # Done / Failed
                break
            time.sleep(1.5)
        out["status_trace"] = statuses
        out["final_status"] = last
        # Encoder writes the file after status flips to Done; give it a beat.
        for _ in range(20):
            if os.path.exists(mp4_path) and os.path.getsize(mp4_path) > 0:
                break
            time.sleep(0.5)
        out["file_exists"] = os.path.exists(mp4_path)
        out["file_size"] = os.path.getsize(mp4_path) if os.path.exists(mp4_path) else 0
        out["ok"] = out["final_status"] == "Done" and out["file_size"] > 0
    except Exception as exc:
        out["error"] = _err(exc)
        out["createvideo_marshalled"] = False
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    wav_path: str | None = None
    mp4_path: str | None = None
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        findings["slide_count_before"] = count_before

        snap = _selection.snapshot(ppt)
        findings["viewed_slide"] = snap.slide_index

        fd, wav_path = tempfile.mkstemp(suffix=".wav", prefix="pptlive_narr_")
        os.close(fd)
        _make_wav(wav_path)
        findings["wav_bytes"] = os.path.getsize(wav_path)
        fd2, mp4_path = tempfile.mkstemp(suffix=".mp4", prefix="pptlive_video_")
        os.close(fd2)
        os.remove(mp4_path)  # let PowerPoint create it

        temp_ids: list[int] = []
        try:
            with deck.edit("media spike: insert audio + auto-play + advance timing"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide = deck.slides[temp.index]
                slide_com = slide.com
                audio_info, audio_shape = probe_insert_audio(slide_com, wav_path)
                findings["insert_audio"] = audio_info
                # Pace the slide to the clip: AdvanceOnTime + AdvanceTime (seconds).
                try:
                    clip_ms = float(audio_shape.MediaFormat.Length) if audio_shape else 0.0
                    secs = max(1.0, clip_ms / 1000.0)
                    tr = slide_com.SlideShowTransition
                    tr.AdvanceOnTime = -1
                    tr.AdvanceTime = secs
                    findings["advance_timing"] = {"clip_ms": clip_ms, "advance_time_s": secs, "ok": True}
                except Exception as exc:
                    findings["advance_timing"] = {"error": _err(exc)}

            # Export OUTSIDE the edit fence (a read); temp slide still present.
            findings["video_export"] = probe_video_export(deck.com, mp4_path)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("media spike: cleanup"):
                    for _ in range(len(temp_ids) + 2):
                        victim = None
                        for idx in range(len(deck.slides), 0, -1):
                            try:
                                sid = deck.slides[idx].id
                            except Exception:
                                continue
                            if sid in temp_ids and sid not in deleted:
                                victim = (idx, sid)
                                break
                        if victim is None:
                            break
                        deck.slides[victim[0]].delete()
                        deleted.append(victim[1])
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        count_after = len(deck.slides)
        findings["slide_count_after"] = count_after
        findings["net_zero_ok"] = count_after == count_before
        findings["save_as_video_constants"] = {"mp4": _PP_SAVE_AS_MP4, "wmv": _PP_SAVE_AS_WMV}

    for p in (wav_path, mp4_path):
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except OSError:
                pass

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

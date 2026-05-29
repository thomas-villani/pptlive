"""Spike — verify the v0.7 picture wrappers against real PowerPoint.

Exercises alt text + per-shape image export end to end on a *running* PowerPoint
with a deck open:

    uv run python scripts/picture_spike.py

Three probes, all net-zero and polite:

1. **Alt text round-trip.** Capture the first shape's `AlternativeText`, set a new
   value via `set_alt_text`, read it back (wrapper + listing), then restore the
   original. Confirms `Shape.AlternativeText` is the right COM handle.
2. **`add_picture` + alt text.** Add a tiny temp PNG (embedded) with `alt_text=`,
   read its alt text + geometry, then delete it. Confirms the create-path sets
   alt text and the shape lands.
3. **Per-shape export — the one genuine unknown (RESOLVED 2026-05-28).**
   `Shape.Export(PathName, Filter, ...)` where `Filter` is the `PpShapeFormat`
   *int* enum (not `Slide.Export`'s string). Export the first shape at native size
   (the shipped wrapper path) and, raw via `.com`, at an explicit 400x300, then
   read each PNG's IHDR back. **Finding:** native export is reliable (a 720 pt
   shape on a 960 pt slide -> 960 px wide). But ScaleWidth/ScaleHeight do **not**
   map to output pixels the way `Slide.Export`'s do — requesting 400x300 gave
   399x241 (width roughly tracked, height didn't, aspect wasn't preserved). So
   `Shape.export_image` ships **native-only**; this probe documents why.

Prints one JSON findings object. Adds/deletes nothing net (the temp picture is
removed; the alt text is restored), and the viewed slide + Selection are restored.
"""

from __future__ import annotations

import base64
import json
import os
import tempfile
from typing import Any

import pptlive as pl
from pptlive import Shape, _selection

# A 2x2 red PNG (real bytes, so PowerPoint's AddPicture accepts it).
_TINY_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAEklEQVR4nGP8z8Dwn4EIwDiqEF0R"
    "AAhKBQHWB1l3AAAAAElFTkSuQmCC"
)


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _png_dims(path: str) -> tuple[int, int] | None:
    """Read (width, height) from a PNG's IHDR, or None if it isn't a PNG."""
    try:
        with open(path, "rb") as fh:
            data = fh.read(24)
        if not data.startswith(b"\x89PNG\r\n\x1a\n"):
            return None
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    except Exception:
        return None


def _first_shape(deck: pl.Presentation, slide_index: int) -> Shape | None:
    for sh in deck.slides[slide_index].shapes:
        if isinstance(sh, Shape):
            return sh
    return None


def probe_alt_text(deck: pl.Presentation) -> dict[str, Any]:
    sh = _first_shape(deck, 1)
    if sh is None:
        return {"skipped": "no shape on slide 1"}
    original = sh.alt_text
    out: dict[str, Any] = {"anchor": sh.anchor_id, "original": original}
    try:
        with deck.edit("picture spike: set alt"):
            sh.set_alt_text("pptlive-spike-alt")
        out["read_back"] = sh.alt_text
        out["in_listing"] = deck.slides[1].shapes.list()[sh.index - 1]["alt_text"]
    finally:
        with deck.edit("picture spike: restore alt"):
            sh.set_alt_text(original)
    out["round_trip_ok"] = out.get("read_back") == "pptlive-spike-alt"
    return out


def probe_add_picture(deck: pl.Presentation) -> dict[str, Any]:
    fd, img = tempfile.mkstemp(prefix="pptlive_spike_", suffix=".png")
    os.close(fd)
    with open(img, "wb") as fh:
        fh.write(_TINY_PNG)
    out: dict[str, Any] = {}
    try:
        with deck.edit("picture spike: add picture"):
            pic = deck.slides[1].shapes.add_picture(
                img, left=36.0, top=36.0, alt_text="pptlive-spike-pic"
            )
        out["anchor"] = pic.anchor_id
        out["type"] = pic.shape_type
        out["alt_text"] = pic.alt_text
        out["geometry"] = pic.geometry()
        with deck.edit("picture spike: delete picture"):
            pic.delete()
        out["deleted"] = True
    except Exception as exc:
        out["error"] = _err(exc)
    finally:
        os.remove(img)
    return out


def probe_export(deck: pl.Presentation) -> dict[str, Any]:
    sh = _first_shape(deck, 1)
    if sh is None:
        return {"skipped": "no shape on slide 1"}
    out: dict[str, Any] = {"anchor": sh.anchor_id, "geometry_pt": sh.geometry()}
    tmps: list[str] = []
    try:
        # The shipped API: native-size export via the wrapper (the reliable path).
        native = str(sh.export_image(fmt="png"))
        tmps.append(native)
        out["native_px"] = _png_dims(native)

        # The finding that made us drop a size override: Shape.Export's
        # ScaleWidth/ScaleHeight do NOT map to output pixels like Slide.Export's.
        # Run it raw via the .com escape hatch and read the result back.
        fd, sized = tempfile.mkstemp(prefix="pptlive_spike_", suffix=".png")
        os.close(fd)
        os.remove(sized)
        tmps.append(sized)
        sh.com.Export(os.path.abspath(sized), 2, 400, 300)  # filter=PpShapeFormat.PNG
        out["raw_requested_400x300"] = _png_dims(sized)
        out["scale_args_mean_pixels"] = out["raw_requested_400x300"] == (400, 300)
    except Exception as exc:
        out["error"] = _err(exc)
    finally:
        for p in tmps:
            try:
                os.remove(p)
            except OSError:
                pass
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        shapes_before = len(deck.slides[1].shapes)
        snap = _selection.snapshot(ppt)

        try:
            findings["alt_text"] = probe_alt_text(deck)
            findings["add_picture"] = probe_add_picture(deck)
            findings["export"] = probe_export(deck)
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            try:
                _selection.restore(ppt, snap)
            except Exception as exc:
                findings["restore_error"] = _err(exc)

        findings["net_zero_ok"] = (
            len(deck.slides) == count_before and len(deck.slides[1].shapes) == shapes_before
        )

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

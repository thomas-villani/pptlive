"""Pin PowerPoint's picture-crop COM semantics before designing a `crop` verb.

Asked for by the 2026-08-03 Claude Code authoring review
(`docs/reviews/claude-code-python-review-03Aug2026.md`):

> A crop verb. I had a 2.08:1 tissue panorama and a [16:9 box]; the only way to
> get a full-bleed title panel was to oversize the [image] and let [it] hang off
> the slide edge. It works, but geometry_report then flags it as a defect forever,
> so I lose the signal.

That last clause is the real cost: the workaround permanently poisons the
`off_slide` signal that `geometry_report` exists to provide, so the cheapest
pre-render lint stops being trustworthy.

WHAT THIS PROBE MUST ANSWER (nothing is assumed — see the CLAUDE.md lesson that a
round-trip spike cannot validate a value, only its own echo):

  Q1  Do `PictureFormat.CropLeft/Right/Top/Bottom` exist and round-trip?
  Q2  **Which edge does each one actually crop?** A name is not evidence. Each is
      probed independently against a picture with four distinctly-coloured edge
      stripes: crop one edge by half, export the shape, and sample where that
      edge's stripe used to be. A wrong mapping shows a *different colour*, so
      "CropTop secretly crops the bottom" cannot pass.
  Q3  What are the units — points? And measured against the ORIGINAL picture or
      the displayed shape?
  Q4  Does cropping shrink `Shape.Width/Height`, or keep the box and crop within?
      (This decides whether a `crop` verb must restore geometry afterwards.)
  Q5  Is the modern `PictureFormat.Crop` object reachable under our late-bound
      dispatch, and does it expose `PictureWidth`/`PictureHeight`? That is what a
      real **cover-fit** ("fill this box, centre-crop the overflow") needs, and
      `ExportAsFixedFormat` has already taught us that a nominally-correct COM
      member can simply refuse to marshal here.

Net-zero: adds one temp slide, deletes it, restores the viewed slide. Run with
    uv run python scripts/crop_spike.py

FINDINGS (2026-08-03, live PowerPoint — all five answered, every check passed):

  A1  `CropLeft/CropRight/CropTop/CropBottom` all exist, read `0.0` on a fresh
      picture, and round-trip exactly (set 75.0 -> read 75.0).
  A2  **Each property crops the edge its name claims.** Pixel evidence, not an
      echo: the green left stripe survived `CropTop`/`CropBottom`/`CropRight` and
      vanished only under `CropLeft`, and likewise for all four. A transposed
      mapping could not have passed.
  A3  The unit is **points, measured against the ORIGINAL picture** — not a
      fraction, and not points-of-the-current-shape. A 400x400 px image inserted
      at its native 300x300 pt, given `CropLeft = 75.0`, narrowed by exactly
      75.0 pt.
  A4  **Cropping SHRINKS `Shape.Width`/`Height`** (300.0 -> 225.0); it does not
      keep the box and letterbox inside it. So a `crop` verb that means to hold a
      layout must re-apply the target geometry *after* setting the crops — the
      same shape of gotcha as `set_picture`'s locked aspect ratio.
  A5  The modern `PictureFormat.Crop` object **does marshal** under our late-bound
      dispatch (unlike `ExportAsFixedFormat`) and exposes `PictureWidth`,
      `PictureHeight`, `PictureOffsetX/Y`, `ShapeWidth`, `ShapeHeight` — all
      readable. **Cover-fit is therefore computable entirely from COM**, with no
      need to re-open and measure the source file.

DESIGN IMPLICATIONS for the verb:

  * `Shape.crop(*, left=, right=, top=, bottom=)` is the raw primitive — points off
    each edge of the original, matching COM 1:1. Must restore `width`/`height`
    afterwards if the caller's box is to be preserved (A4).
  * `Shape.crop_to_fit(*, left, top, width, height)` ("cover") is the verb the
    review actually wants: read `Crop.PictureWidth`/`PictureHeight` for the source
    aspect (A5), compute the symmetric centre-crop that makes the source aspect
    match the target box, apply it, then set the box exactly. That turns
    "full-bleed this panorama into a 16:9 panel" into one call **without** pushing
    anything off-slide, so `geometry_report`'s `off_slide` signal stays clean.
  * Validation before COM: a crop must not exceed the picture (left+right <
    PictureWidth, top+bottom < PictureHeight) — otherwise PowerPoint is being
    asked for a zero/negative-width shape.
  * Non-picture shape -> `ValueError`, mirroring `set_picture`.
"""

from __future__ import annotations

import os
import struct
import sys
import tempfile
import zlib
from typing import Any

import pptlive as pl

# The test image: four distinct edge stripes around a white core, so cropping any
# one edge by half removes exactly one colour and nothing else.
IMG = 400  # px, square
BAND = IMG // 4
RED = (220, 30, 30)  # top stripe
BLUE = (40, 60, 220)  # bottom stripe
GREEN = (30, 200, 60)  # left stripe
YELLOW = (230, 210, 40)  # right stripe
WHITE = (255, 255, 255)  # core

#: edge -> (the COM property, the stripe colour that edge carries, where to sample
#: the exported PNG as (x, y) fractions once that edge is cropped away).
EDGES: dict[str, tuple[str, tuple[int, int, int], tuple[float, float]]] = {
    "top": ("CropTop", RED, (0.5, 0.04)),
    "bottom": ("CropBottom", BLUE, (0.5, 0.96)),
    "left": ("CropLeft", GREEN, (0.04, 0.5)),
    "right": ("CropRight", YELLOW, (0.96, 0.5)),
}


def _chunk(tag: bytes, data: bytes) -> bytes:
    return (
        struct.pack(">I", len(data))
        + tag
        + data
        + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    )


def _striped_png() -> bytes:
    """A 400x400 PNG: red top band, blue bottom, green left, yellow right, white core."""
    rows = bytearray()
    for y in range(IMG):
        rows.append(0)  # filter type 0
        for x in range(IMG):
            if y < BAND:
                px = RED
            elif y >= IMG - BAND:
                px = BLUE
            elif x < BAND:
                px = GREEN
            elif x >= IMG - BAND:
                px = YELLOW
            else:
                px = WHITE
            rows.extend(px)
    ihdr = struct.pack(">IIBBBBB", IMG, IMG, 8, 2, 0, 0, 0)  # 8-bit truecolour
    return (
        b"\x89PNG\r\n\x1a\n"
        + _chunk(b"IHDR", ihdr)
        + _chunk(b"IDAT", zlib.compress(bytes(rows), 9))
        + _chunk(b"IEND", b"")
    )


def _decode(png_path: str) -> tuple[int, int, int, bytearray, bytes, int] | None:
    """`(width, height, stride, pixels, palette, mode)` for a PNG, or None.

    `mode` is the colour type. Handles PowerPoint's `Shape.Export` output, which
    can come back palettised (the lesson from `set_picture_spike.py`).
    """
    try:
        with open(png_path, "rb") as fh:
            data = fh.read()
    except OSError:
        return None
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        return None
    pos, width, height, depth, ctype, idat, plte = 8, 0, 0, 0, 0, b"", b""
    while pos + 8 <= len(data):
        (length,) = struct.unpack(">I", data[pos : pos + 4])
        tag = data[pos + 4 : pos + 8]
        chunk = data[pos + 8 : pos + 8 + length]
        if tag == b"IHDR":
            width, height, depth, ctype = struct.unpack(">IIBB", chunk[:10])
        elif tag == b"PLTE":
            plte = chunk
        elif tag == b"IDAT":
            idat += chunk
        elif tag == b"IEND":
            break
        pos += 12 + length
    if ctype == 3:
        if depth not in (1, 2, 4, 8):
            return None
        stride, bpp = (IMG * 0 + (width * depth + 7) // 8), 1
    elif ctype in (2, 6) and depth == 8:
        channels = 3 if ctype == 2 else 4
        stride, bpp = width * channels, channels
    else:
        return None
    try:
        raw = zlib.decompress(idat)
    except zlib.error:
        return None
    out, prev, p = bytearray(), bytearray(stride), 0
    for _y in range(height):
        if p >= len(raw):
            break
        ft = raw[p]
        p += 1
        line = bytearray(raw[p : p + stride])
        p += stride
        for i in range(stride):
            a = line[i - bpp] if i >= bpp else 0
            b = prev[i]
            c = prev[i - bpp] if i >= bpp else 0
            if ft == 1:
                line[i] = (line[i] + a) & 0xFF
            elif ft == 2:
                line[i] = (line[i] + b) & 0xFF
            elif ft == 3:
                line[i] = (line[i] + ((a + b) >> 1)) & 0xFF
            elif ft == 4:
                pp = a + b - c
                pa, pb, pc = abs(pp - a), abs(pp - b), abs(pp - c)
                line[i] = (
                    line[i] + (a if (pa <= pb and pa <= pc) else (b if pb <= pc else c))
                ) & 0xFF
        out.extend(line)
        prev = line
    return width, height, stride, out, plte, ctype


def pixel_at(png_path: str, fx: float, fy: float) -> tuple[int, int, int] | None:
    """The pixel at fractional position (fx, fy) of a PNG, as RGB."""
    dec = _decode(png_path)
    if dec is None:
        return None
    width, height, stride, out, plte, ctype = dec
    x = min(width - 1, max(0, int(width * fx)))
    y = min(height - 1, max(0, int(height * fy)))
    if ctype == 3:
        depth = 8 if stride >= width else (8 * stride) // width
        if depth == 8:
            idx = out[y * stride + x]
        else:
            byte = out[y * stride + (x * depth) // 8]
            idx = (byte >> (8 - depth - (x * depth) % 8)) & ((1 << depth) - 1)
        po = idx * 3
        return (plte[po], plte[po + 1], plte[po + 2]) if po + 2 < len(plte) else None
    channels = 3 if ctype == 2 else 4
    base = y * stride + x * channels
    return (out[base], out[base + 1], out[base + 2])


def nearest(px: tuple[int, int, int] | None) -> str:
    """Name the closest of our five known colours (so output reads as evidence)."""
    if px is None:
        return "unreadable"
    named = {"red": RED, "blue": BLUE, "green": GREEN, "yellow": YELLOW, "white": WHITE}
    best, dist = "?", 1 << 30
    for name, want in named.items():
        d = sum((a - b) ** 2 for a, b in zip(px, want, strict=False))
        if d < dist:
            best, dist = name, d
    return best if dist <= 3 * 60**2 else f"other{px}"


def main() -> int:
    fd, img_path = tempfile.mkstemp(prefix="pptlive_crop_", suffix=".png")
    os.write(fd, _striped_png())
    os.close(fd)
    try:
        with pl.connect() as ppt:
            if not len(ppt.presentations):
                ppt.com.Presentations.Add()
            return run(ppt, img_path)
    finally:
        os.unlink(img_path)


def run(ppt: pl.PowerPoint, img_path: str) -> int:  # noqa: C901 - a probe, read top-down
    deck = ppt.presentations.active
    if not len(deck.slides):
        with deck.edit("crop spike: seed"):
            deck.slides.add(layout="blank")
    before = len(deck.slides)
    viewed = ppt.viewed_slide_index()
    findings: dict[str, Any] = {}
    exports: list[str] = []

    slide = deck.slides.add(layout="blank")
    try:
        # ---- Q1/Q3/Q4: existence, units, and effect on the shape box ----------
        print("Q1/Q3/Q4  crop properties: existence, units, effect on Shape.Width")
        with deck.edit("crop spike: baseline picture"):
            pic = slide.shapes.add_picture(img_path, left=40.0, top=40.0)
        com = pic.com
        native_w, native_h = float(com.Width), float(com.Height)
        print(f"  native inserted size: {native_w:.1f} x {native_h:.1f} pt (image {IMG}x{IMG} px)")
        findings["native_size_pt"] = [round(native_w, 2), round(native_h, 2)]

        try:
            pf = com.PictureFormat
            readback = {}
            for edge, (prop, _c, _s) in EDGES.items():
                readback[edge] = float(getattr(pf, prop))
            print(f"  initial crops: {readback}")
            findings["crops_exist"] = True
            findings["initial_crops"] = readback
        except Exception as exc:  # noqa: BLE001 - the answer to Q1 is "no"
            print(f"  [FAIL] PictureFormat.Crop* unreadable: {exc}")
            findings["crops_exist"] = False
            return _finish(deck, slide, before, viewed, findings, exports, ok=False)

        # Crop the left by a quarter of the NATIVE WIDTH IN POINTS. If the unit is
        # points-of-the-original, the shape should narrow by exactly that much.
        quarter = native_w / 4.0
        with deck.edit("crop spike: quarter off the left"):
            com.PictureFormat.CropLeft = quarter
        after_w = float(com.Width)
        shrank_by = native_w - after_w
        print(f"  CropLeft = {quarter:.1f}pt -> Shape.Width {native_w:.1f} -> {after_w:.1f}")
        print(
            f"    shrank by {shrank_by:.1f}pt ({'matches' if abs(shrank_by - quarter) < 1 else 'DOES NOT match'} the crop)"
        )
        findings["unit_is_points_of_original"] = abs(shrank_by - quarter) < 1.0
        findings["crop_shrinks_shape"] = shrank_by > 1.0
        # Read the crop back — does it echo what we set?
        echoed = float(com.PictureFormat.CropLeft)
        print(
            f"    CropLeft reads back as {echoed:.1f}pt ({'round-trips' if abs(echoed - quarter) < 1 else 'DRIFTED'})"
        )
        findings["crop_round_trips"] = abs(echoed - quarter) < 1.0
        with deck.edit("crop spike: reset"):
            com.PictureFormat.CropLeft = 0.0
            com.Width, com.Height = native_w, native_h

        # ---- Q2: which edge does each property actually crop? -----------------
        print("\nQ2  which edge does each Crop* property really cut?")
        print("    (crop one edge by half; that edge's stripe must be GONE)")
        ok = True
        for edge, (prop, colour, (fx, fy)) in EDGES.items():
            with deck.edit(f"crop spike: {prop}"):
                probe = slide.shapes.add_picture(img_path, left=40.0, top=40.0)
            pcom = probe.com
            # Sample BEFORE, to prove the stripe was there to begin with.
            before_png = _export(pcom, exports)
            was = nearest(pixel_at(before_png, fx, fy)) if before_png else "unreadable"
            with deck.edit(f"crop spike: apply {prop}"):
                setattr(
                    pcom.PictureFormat,
                    prop,
                    float(native_w if edge in ("left", "right") else native_h) / 2.0,
                )
            after_png = _export(pcom, exports)
            now = nearest(pixel_at(after_png, fx, fy)) if after_png else "unreadable"
            want_before = nearest(colour)
            good = was == want_before and now != want_before
            ok = ok and good
            print(
                f"  [{'ok' if good else 'FAIL'}] {prop:<11} sample({fx:.2f},{fy:.2f}): "
                f"{was} -> {now}   (expected {want_before} -> not {want_before})"
            )
            findings[f"{prop}_removes_{edge}_stripe"] = good
            with deck.edit("crop spike: drop probe"):
                probe.delete()

        # ---- Q5: the modern Crop object (what cover-fit would want) -----------
        print("\nQ5  is the modern PictureFormat.Crop object reachable + useful?")
        crop_obj: dict[str, Any] = {}
        try:
            c = com.PictureFormat.Crop
            for name in (
                "PictureWidth",
                "PictureHeight",
                "PictureOffsetX",
                "PictureOffsetY",
                "ShapeWidth",
                "ShapeHeight",
            ):
                try:
                    crop_obj[name] = round(float(getattr(c, name)), 2)
                except Exception as exc:  # noqa: BLE001
                    crop_obj[name] = f"unreadable: {type(exc).__name__}"
            print(f"  Crop object reachable: {crop_obj}")
            findings["crop_object"] = crop_obj
            usable = isinstance(crop_obj.get("PictureWidth"), float)
            print(
                f"  [{'ok' if usable else 'FAIL'}] PictureWidth/Height readable "
                f"-> cover-fit can be computed {'without' if usable else 'WITHOUT'} "
                "re-measuring the source file"
            )
            findings["cover_fit_computable_from_com"] = usable
        except Exception as exc:  # noqa: BLE001 - the ExportAsFixedFormat problem again?
            print(f"  [note] PictureFormat.Crop did not marshal: {type(exc).__name__}: {exc}")
            findings["crop_object"] = None
            findings["cover_fit_computable_from_com"] = False
            print("  -> cover-fit must derive the source aspect from the inserted native size")

        return _finish(deck, slide, before, viewed, findings, exports, ok=ok)
    except Exception:
        _finish(deck, slide, before, viewed, findings, exports, ok=False)
        raise


def _export(com_shape: Any, exports: list[str]) -> str | None:
    """Export one shape to a temp PNG and return the path (or None)."""
    fd, path = tempfile.mkstemp(prefix="pptlive_cropx_", suffix=".png")
    os.close(fd)
    exports.append(path)
    try:
        com_shape.Export(path, 2)  # 2 = ppShapeFormatPNG
        return path
    except Exception:  # noqa: BLE001
        return None


def _finish(
    deck: Any,
    slide: Any,
    before: int,
    viewed: int | None,
    findings: dict[str, Any],
    exports: list[str],
    *,
    ok: bool,
) -> int:
    with deck.edit("crop spike: clean up"):
        slide.delete()
    if viewed is not None and viewed <= len(deck.slides):
        deck.go_to(deck.slides[viewed])
    for path in exports:
        try:
            os.unlink(path)
        except OSError:
            pass
    print(f"\nnet-zero: {len(deck.slides) == before} ({before} slides before and after)")
    print("\nfindings:")
    for key, value in findings.items():
        print(f"  {key}: {value}")
    print("\n" + ("all checks passed" if ok else "SOME CHECKS FAILED — read the Q2 table above"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

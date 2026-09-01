"""Pin the COM semantics behind GitHub issues #49, #50, #53 before fixing them.

The 2026-08-31 live-testing round (narrated slide decks driven by an agent) filed
six issues; three were already fixed on main (#48/#51/#52 — the #44/#45 round).
The remaining three each hinge on a live-PowerPoint behavior nothing in the fake
COM can answer, so (per the standing lesson) each gets a probe designed so a
wrong assumption produces a *different* observable, not an echo.

WHAT THIS PROBE MUST ANSWER:

  #49 — empty content placeholder captures `add_picture`
  A1  Does `Shapes.AddPicture` on a slide with an empty content placeholder
      return the placeholder (capture), and does `Shapes.Count` even grow?
      (If Count doesn't grow, our `_added()` — "last shape by Count" — hands the
      caller an arbitrary WRONG shape, worse than the issue reports.)
  A2  Is the requested left/top/width honored, or does the placeholder frame win?
  A3  Do `PictureFormat.CropLeft` (and the modern `Crop` object) work on a
      picture-filled placeholder? (Decides whether `crop`/`crop_to_fit` can
      simply accept captured pictures instead of refusing them.)
  A4  Is capture avoided when every content placeholder already has text?
  A5  Post-capture, do `Left/Top/Width/Height` writes stick on the placeholder?
      (Decides whether honoring the caller's geometry after capture is viable.)

  #50 — `media add` silently re-centers on negative --left/--top
  B1  Where does `AddMediaObject2(left=-90, top=-90)` actually put the icon?
  B2  Does assigning `Shape.Left = -90` AFTER the insert stick?
  B3  Are positive on-slide coords honored (sanity)?
  B4  Are positive OFF-slide coords honored, or is the clamp negative-only?
  B5  Does `AddPicture` honor negative coords (is this media-specific)?

  #53 — autofit overrides explicit sizes from `set_paragraphs`
  C1  With a content placeholder's default autofit, do explicit `size=` values
      written by `set_paragraphs` read back shrunk (real `Font.Size` rewrite) or
      nominal (display-only scale)?
  C2  Disable autofit FIRST (`set_text_frame(autosize="none")`), then the same
      `set_paragraphs` — do the sizes stick? (The fix ordering.)
  C3  Disable autofit AFTER the shrink — do sizes revert on their own, or is the
      damage permanent (the known "autofit does not re-fit retroactively"
      caveat, which is why the reporter needed a third re-apply call)?

Net-zero: adds temp slides, deletes them, restores the viewed slide. Run with
    uv run python scripts/gh_feedback_spike.py
"""

from __future__ import annotations

import os
import struct
import sys
import tempfile
import zlib
from typing import Any

import pptlive as pl

MSO_PLACEHOLDER = 14  # MsoShapeType.msoPlaceholder
MSO_PICTURE = 13


def _tiny_png() -> bytes:
    """A 200x100 solid-teal PNG (2:1, so a placeholder frame mismatch is visible)."""
    w, h = 200, 100
    rows = bytearray()
    for _y in range(h):
        rows.append(0)
        rows.extend(bytes((20, 160, 160)) * w)
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(bytes(rows), 9))
        + chunk(b"IEND", b"")
    )


def _tiny_wav() -> bytes:
    """0.2 s of 16-bit mono silence at 8 kHz — enough for AddMediaObject2."""
    n = 1600
    data = b"\x00\x00" * n
    return (
        b"RIFF"
        + struct.pack("<I", 36 + len(data))
        + b"WAVEfmt "
        + struct.pack("<IHHIIHH", 16, 1, 1, 8000, 16000, 2, 16)
        + b"data"
        + struct.pack("<I", len(data))
        + data
    )


def _tmpfile(suffix: str, payload: bytes) -> str:
    fd, path = tempfile.mkstemp(prefix="pptlive_ghspike_", suffix=suffix)
    os.write(fd, payload)
    os.close(fd)
    return path


def _box(com_shape: Any) -> dict[str, float]:
    return {
        "left": round(float(com_shape.Left), 1),
        "top": round(float(com_shape.Top), 1),
        "width": round(float(com_shape.Width), 1),
        "height": round(float(com_shape.Height), 1),
    }


def probe_capture(deck: Any, img_path: str, findings: dict[str, Any]) -> list[Any]:
    """#49 — placeholder capture of AddPicture. Returns temp slides to delete."""
    print("\n== #49  empty content placeholder vs AddPicture ==")
    slide = deck.slides.add(layout="two_content", title="capture spike")
    shapes_com = slide.com.Shapes
    n_before = int(shapes_com.Count)

    ret = shapes_com.AddPicture(img_path, 0, -1, 500.0, 120.0, 200.0, -1.0)
    n_after = int(shapes_com.Count)
    ret_type = int(ret.Type)
    captured = ret_type == MSO_PLACEHOLDER
    print(
        f"  A1 Shapes.Count {n_before} -> {n_after}; returned Type={ret_type} "
        f"({'PLACEHOLDER — captured' if captured else 'free picture'})"
    )
    findings["a1_captured"] = captured
    findings["a1_count_grew"] = n_after > n_before
    if not findings["a1_count_grew"]:
        last = shapes_com(n_after)
        wrong = int(last.Id) != int(ret.Id)
        print(
            f"     last-by-Count Id={int(last.Id)} vs returned Id={int(ret.Id)} "
            f"-> _added() would hand back the {'WRONG shape' if wrong else 'same shape'}"
        )
        findings["a1_added_wrong_shape"] = wrong

    got = _box(ret)
    print(f"  A2 requested left=500 top=120 width=200 -> got {got}")
    findings["a2_geometry_honored"] = (
        abs(got["left"] - 500.0) < 1
        and abs(got["top"] - 120.0) < 1
        and abs(got["width"] - 200.0) < 1
    )

    try:
        ret.PictureFormat.CropLeft = 10.0
        echo = float(ret.PictureFormat.CropLeft)
        ret.PictureFormat.CropLeft = 0.0
        crop_ok = abs(echo - 10.0) < 0.5
    except Exception as exc:  # noqa: BLE001 - "no" is a finding
        crop_ok = False
        print(f"     CropLeft raised: {type(exc).__name__}: {exc}")
    try:
        pw = float(ret.PictureFormat.Crop.PictureWidth)
        crop_obj_ok = pw > 0
    except Exception as exc:  # noqa: BLE001
        crop_obj_ok = False
        pw = None
        print(f"     Crop object raised: {type(exc).__name__}: {exc}")
    print(
        f"  A3 CropLeft round-trips on captured shape: {crop_ok}; "
        f"Crop.PictureWidth readable: {crop_obj_ok} ({pw})"
    )
    findings["a3_crop_works_on_placeholder"] = crop_ok and crop_obj_ok

    slide_full = deck.slides.add(
        layout="two_content",
        title="capture spike (filled)",
        content={"body:1": "left text", "body:2": "right text"},
    )
    sc2 = slide_full.com.Shapes
    n2_before = int(sc2.Count)
    ret2 = sc2.AddPicture(img_path, 0, -1, 500.0, 120.0, 200.0, -1.0)
    free = int(ret2.Type) != MSO_PLACEHOLDER and int(sc2.Count) > n2_before
    print(
        f"  A4 with both bodies filled: returned Type={int(ret2.Type)}, "
        f"Count {n2_before} -> {int(sc2.Count)} ({'free picture' if free else 'STILL captured'})"
    )
    findings["a4_filled_placeholders_avoid_capture"] = free

    ret.Left, ret.Top, ret.Width = 500.0, 120.0, 200.0
    after = _box(ret)
    stuck = (
        abs(after["left"] - 500.0) < 1
        and abs(after["top"] - 120.0) < 1
        and abs(after["width"] - 200.0) < 1
    )
    print(f"  A5 post-capture geometry writes: -> {after} ({'stick' if stuck else 'DO NOT stick'})")
    findings["a5_post_capture_geometry_writable"] = stuck
    return [slide, slide_full]


def probe_media_coords(
    deck: Any, wav_path: str, img_path: str, findings: dict[str, Any]
) -> list[Any]:
    """#50 — AddMediaObject2 vs negative / off-slide coordinates."""
    print("\n== #50  AddMediaObject2 vs negative --left/--top ==")
    slide = deck.slides.add(layout="blank")
    shapes_com = slide.com.Shapes
    slide_w = float(deck.com.PageSetup.SlideWidth)

    m1 = shapes_com.AddMediaObject2(wav_path, 0, -1, -90.0, -90.0, 60.0, 60.0)
    b1 = _box(m1)
    print(f"  B1 requested left=-90 top=-90 -> got {b1}")
    findings["b1_negative_honored_at_insert"] = (
        abs(b1["left"] + 90.0) < 1 and abs(b1["top"] + 90.0) < 1
    )

    m1.Left, m1.Top = -90.0, -90.0
    b2 = _box(m1)
    print(f"  B2 post-insert Left/Top = -90 -> got {b2}")
    findings["b2_negative_sticks_after_insert"] = (
        abs(b2["left"] + 90.0) < 1 and abs(b2["top"] + 90.0) < 1
    )

    m3 = shapes_com.AddMediaObject2(wav_path, 0, -1, 100.0, 100.0, 60.0, 60.0)
    b3 = _box(m3)
    print(f"  B3 requested left=100 top=100 -> got {b3}")
    findings["b3_positive_honored"] = abs(b3["left"] - 100.0) < 1 and abs(b3["top"] - 100.0) < 1

    off = slide_w + 50.0
    m4 = shapes_com.AddMediaObject2(wav_path, 0, -1, off, 100.0, 60.0, 60.0)
    b4 = _box(m4)
    print(f"  B4 requested left={off:.0f} (off-slide right) -> got {b4}")
    findings["b4_positive_offslide_honored"] = abs(b4["left"] - off) < 1

    p = shapes_com.AddPicture(img_path, 0, -1, -90.0, -90.0, 60.0, -1.0)
    b5 = _box(p)
    print(f"  B5 AddPicture at left=-90 top=-90 -> got {b5}")
    findings["b5_addpicture_negative_honored"] = (
        abs(b5["left"] + 90.0) < 1 and abs(b5["top"] + 90.0) < 1
    )
    return [slide]


BULLETS = [
    {
        "text": "Lead finding with a deliberately long line of prose so the frame overflows",
        "size": 20,
    },
    {
        "text": "Supporting detail line one, also padded out to take real horizontal space",
        "size": 16,
    },
    {
        "text": "Supporting detail line two, also padded out to take real horizontal space",
        "size": 16,
    },
    {"text": "Lead finding number two, again long enough that wrapping is guaranteed", "size": 20},
    {"text": "Supporting detail line three, more prose padding for genuine overflow", "size": 16},
    {"text": "Supporting detail line four, more prose padding for genuine overflow", "size": 16},
    {"text": "Lead finding number three, the frame is by now far past its height", "size": 20},
    {"text": "Supporting detail line five, and still more padding to be certain", "size": 16},
    {"text": "Supporting detail line six, and still more padding to be certain", "size": 16},
    {"text": "Closing line, the tenth paragraph, well beyond any content box", "size": 16},
]
WANT = [float(item["size"]) for item in BULLETS]


def _read_sizes(ph: Any) -> list[float]:
    tr = ph.com.TextFrame.TextRange
    return [
        round(float(tr.Paragraphs(i).Font.Size), 1)
        for i in range(1, int(tr.Paragraphs().Count) + 1)
    ]


def probe_autofit(deck: Any, findings: dict[str, Any]) -> list[Any]:
    """#53 — does autofit rewrite Font.Size under set_paragraphs?"""
    print("\n== #53  autofit vs explicit set_paragraphs sizes ==")
    slide1 = deck.slides.add(layout="title_and_content", title="autofit spike C1")
    ph1 = deck.anchor_by_id(f"ph:{slide1.index}:body")
    auto_before = int(ph1.com.TextFrame2.AutoSize)
    ph1.set_paragraphs(BULLETS)
    sizes1 = _read_sizes(ph1)
    print(f"  C1 default autofit (TextFrame2.AutoSize={auto_before}): wrote {WANT}")
    print(f"     read back {sizes1}")
    findings["c1_sizes_survive_default_autofit"] = sizes1 == WANT
    findings["c1_autosize_mode"] = auto_before

    slide2 = deck.slides.add(layout="title_and_content", title="autofit spike C2")
    ph2 = deck.anchor_by_id(f"ph:{slide2.index}:body")
    ph2.set_text_frame(autosize="none")
    ph2.set_paragraphs(BULLETS)
    sizes2 = _read_sizes(ph2)
    print(f"  C2 autosize='none' FIRST, then set_paragraphs: read back {sizes2}")
    findings["c2_sizes_stick_with_autofit_off_first"] = sizes2 == WANT

    ph1.set_text_frame(autosize="none")
    sizes3 = _read_sizes(ph1)
    print(f"  C3 autosize='none' AFTER the C1 shrink: read back {sizes3}")
    findings["c3_late_autosize_off_restores_sizes"] = sizes3 == WANT
    return [slide1, slide2]


def main() -> int:
    img_path = _tmpfile(".png", _tiny_png())
    wav_path = _tmpfile(".wav", _tiny_wav())
    try:
        with pl.connect() as ppt:
            if not len(ppt.presentations):
                ppt.com.Presentations.Add()
            deck = ppt.presentations.active
            if not len(deck.slides):
                with deck.edit("gh spike: seed"):
                    deck.slides.add(layout="blank")
            before = len(deck.slides)
            viewed = ppt.viewed_slide_index()
            findings: dict[str, Any] = {}
            temps: list[Any] = []
            try:
                with deck.edit("gh spike: #49 capture"):
                    temps += probe_capture(deck, img_path, findings)
                with deck.edit("gh spike: #50 media coords"):
                    temps += probe_media_coords(deck, wav_path, img_path, findings)
                with deck.edit("gh spike: #53 autofit"):
                    temps += probe_autofit(deck, findings)
            finally:
                with deck.edit("gh spike: clean up"):
                    for slide in reversed(temps):
                        slide.delete()
                if viewed is not None and viewed <= len(deck.slides):
                    deck.go_to(deck.slides[viewed])
                print(
                    f"\nnet-zero: {len(deck.slides) == before} ({before} slides before and after)"
                )
                print("\nfindings:")
                for key, value in findings.items():
                    print(f"  {key}: {value}")
            return 0
    finally:
        for path in (img_path, wav_path):
            try:
                os.unlink(path)
            except OSError:
                pass


if __name__ == "__main__":
    sys.exit(main())

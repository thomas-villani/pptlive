"""Verify `Shape.crop` / `Shape.crop_to_fit` against live PowerPoint.

The follow-up to `scripts/crop_spike.py`, which pinned the four `PictureFormat.
Crop*` properties but left one question open — and that question decides whether
`crop_to_fit` is correct or merely plausible:

  S1  **What is a crop point a point OF?** The first spike cropped a picture
      sitting at its *native* size, where "points of the original picture" and
      "points of the displayed shape" are numerically identical — so it could not
      tell them apart. This one inserts a picture at native size, **resizes it to
      half**, and then crops: under the display reading the shape narrows by the
      crop value, under the picture reading it narrows by half of it. One probe,
      two distinguishable outcomes.
  S2  Is `Crop.PictureWidth` in that same unit? The over-crop guard in
      `Shape.crop` compares the requested crop against it, so a mismatch would
      reject valid crops (or wave through a shape-destroying one).
  S3  Does zeroing a crop restore the uncropped shape size? `crop_to_fit` clears
      first and then measures the source aspect off the box, so if zeroing left
      the box small every fit would be computed against the wrong ratio.
  S4  Does `crop_to_fit(fit="cover")` land the box exactly **and keep the right
      content**? Pixel evidence, not an echo: a 400x400 image with four
      distinctly-coloured edge stripes, cover-fitted into a 2:1 box, must lose
      the top and bottom stripes and keep the left and right ones. A transposed
      axis shows the opposite pair, so it cannot pass by accident.
  S5  Does `fit="contain"` crop nothing, fit inside the box, and centre?

`_shapes._crop_axis_scale` deliberately *measures* the S1 ratio at runtime rather
than assuming an answer, so `crop_to_fit` is correct either way — but S1 still
has to be recorded, because the `crop()` primitive is documented as 1:1 with COM
and its guard depends on S2.

Net-zero: adds one temp slide, deletes it, restores the viewed slide. Run with
    uv run python scripts/crop_fit_spike.py

FINDINGS (2026-08-03, live PowerPoint — 12/12 checks pass, after one real fix):

  A1  **A crop point is a point of the ORIGINAL picture.** On a picture inserted
      at its native 300 pt and then resized to 150 pt, `CropLeft = 75` removed
      **37.5** display points, not 75. So crop values are independent of how the
      picture is scaled on the slide, and any fit has to convert between the two.
  A2  **`Crop.PictureWidth` is a DIFFERENT unit — display points.** On that same
      half-scale picture it read **150.00**, the uncropped *displayed* extent, not
      the 300 pt original. This spike's first run asserted the opposite and
      failed, which caught a genuine bug: `Shape.crop`'s over-crop guard compared
      the requested crop (original points) against `PictureWidth` (display
      points), so it would have rejected valid crops on any shrunk picture and
      waved through a shape-destroying one on any enlarged picture. Fixed by
      dividing through the measured scale — `_picture_extent` now returns
      (300.00, 300.00) crop points for that picture, and the probe it uses puts
      the caller's existing crop back rather than zeroing the edge it borrowed.
  A3  Zeroing a crop **does** restore the uncropped shape size (150.00 pt back
      from 112.50), so `crop_to_fit`'s "clear first, then measure the source
      aspect" is sound.
  A4  `crop_to_fit(fit="cover")` lands the box **exactly** (400.00 x 200.00 at
      (60.00, 80.00)) and keeps the right content: the square four-stripe source
      into a 2:1 box lost the red top and blue bottom stripes and kept the green
      left and yellow right ones, with a symmetric 75/75 vertical crop. A
      transposed axis would have shown the opposite pair — pixel evidence, not an
      echo.
  A5  `fit="contain"` cropped nothing, fitted to 200.00 x 200.00 inside the
      400x200 box (height-limited), centred the 200 pt shortfall horizontally
      (left 60 -> 160), and still showed all four stripes.

THE LESSON, again: the unit question was invisible to `crop_spike.py` because it
probed a picture at native size, where both readings give the same number — the
same shape of blind spot as validating a constant against its own echo. Design
the probe so a wrong answer produces a *different* observable value: here, resize
the picture first.
"""

from __future__ import annotations

import os
import sys
import tempfile
from typing import Any

from crop_spike import BLUE, GREEN, IMG, RED, YELLOW, _striped_png, nearest, pixel_at

import pptlive as pl

NATIVE = 300.0  # the 400x400 px image inserts at 300x300 pt (96 dpi -> 72 pt/in)
TOL = 0.75  # pt — PowerPoint stores points as floats; ~0.01 pt of noise is normal


def main() -> int:
    fd, img_path = tempfile.mkstemp(prefix="pptlive_cropfit_", suffix=".png")
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
        with deck.edit("crop-fit spike: seed"):
            deck.slides.add(layout="blank")
    before = len(deck.slides)
    viewed = ppt.viewed_slide_index()
    exports: list[str] = []
    checks: list[tuple[str, bool, str]] = []

    def check(label: str, ok: bool, detail: str) -> None:
        checks.append((label, ok, detail))
        print(f"  [{'ok' if ok else 'FAIL'}] {label}: {detail}")

    slide = deck.slides.add(layout="blank")
    try:
        # ---- S1: what is a crop point a point OF? -----------------------------
        print("S1  crop unit, measured on a picture that is NOT at native size")
        with deck.edit("crop-fit spike: half-scale picture"):
            pic = slide.shapes.add_picture(img_path, left=40.0, top=40.0)
            pic.resize(width=NATIVE / 2, height=NATIVE / 2)
        com = pic.com
        half_w = float(com.Width)
        print(f"  picture inserted native ~{NATIVE:.0f}pt, resized to {half_w:.1f}pt")
        with deck.edit("crop-fit spike: crop a quarter"):
            com.PictureFormat.CropLeft = NATIVE / 4.0  # 75.0
        removed = half_w - float(com.Width)
        display_reading = abs(removed - NATIVE / 4.0) < TOL  # 75 pt off a 150 pt shape
        picture_reading = abs(removed - NATIVE / 8.0) < TOL  # 37.5 pt off a 150 pt shape
        unit = (
            "points of the DISPLAYED SHAPE"
            if display_reading
            else "points of the ORIGINAL PICTURE"
            if picture_reading
            else f"NEITHER (removed {removed:.2f}pt)"
        )
        check(
            "crop unit is points of the ORIGINAL picture",
            picture_reading,
            f"CropLeft=75 removed {removed:.2f}pt -> {unit}",
        )
        print(f"  >>> S1 ANSWER: a crop point is {unit}")

        # ---- S2: Crop.PictureWidth is a DIFFERENT unit ------------------------
        # The first run of this spike assumed PictureWidth was in crop units and
        # failed, which is the finding: it is the uncropped picture in *display*
        # points. `_shapes._picture_extent` divides it by the measured scale to
        # get back to crop units; both readings are pinned here so a regression
        # in either shows up.
        print("\nS2  Crop.PictureWidth — display points, NOT crop points")
        crop_obj = com.PictureFormat.Crop
        pw, ph = float(crop_obj.PictureWidth), float(crop_obj.PictureHeight)
        check(
            "PictureWidth is the uncropped DISPLAY extent",
            abs(pw - half_w) < TOL,
            f"PictureWidth={pw:.2f} PictureHeight={ph:.2f}; "
            f"display={half_w:.2f}, crop-unit original={NATIVE:.2f}",
        )
        from pptlive._shapes import _picture_extent  # noqa: PLC0415 - probe-local

        extent = _picture_extent(com)
        check(
            "_picture_extent converts back to crop units",
            extent is not None and abs(extent[0] - NATIVE) < TOL,
            f"_picture_extent -> {extent} (the original is {NATIVE:.2f} crop points)",
        )
        check(
            "the extent probe left the caller's crop alone",
            abs(float(com.PictureFormat.CropLeft) - NATIVE / 4.0) < 0.01,
            f"CropLeft still {float(com.PictureFormat.CropLeft):.2f}",
        )

        # ---- S3: does zeroing a crop restore the box? -------------------------
        print("\nS3  zeroing a crop restores the uncropped size")
        with deck.edit("crop-fit spike: un-crop"):
            com.PictureFormat.CropLeft = 0.0
        check(
            "zero restores",
            abs(float(com.Width) - half_w) < TOL,
            f"width back to {float(com.Width):.2f}pt (was {half_w:.2f} before the crop)",
        )
        with deck.edit("crop-fit spike: drop probe"):
            pic.delete()

        # ---- S4: crop_to_fit cover — exact box + the RIGHT content kept -------
        print("\nS4  crop_to_fit(fit='cover') into a 2:1 box")
        print("    (square source -> must lose the RED top and BLUE bottom stripes,")
        print("     and KEEP the GREEN left and YELLOW right ones)")
        with deck.edit("crop-fit spike: cover"):
            cover = slide.shapes.add_picture(img_path, left=20.0, top=20.0)
            result = cover.crop_to_fit(left=60.0, top=80.0, width=400.0, height=200.0)
        geo = result["geometry"]
        check(
            "cover box exact",
            all(
                abs(geo[k] - want) < TOL
                for k, want in (("left", 60.0), ("top", 80.0), ("width", 400.0), ("height", 200.0))
            ),
            f"{geo['width']:.2f}x{geo['height']:.2f} at ({geo['left']:.2f}, {geo['top']:.2f})",
        )
        crop = result["crop"]
        check(
            "cover cropped the vertical axis, symmetrically",
            crop["top"] > 0 and abs(crop["top"] - crop["bottom"]) < 0.01 and crop["left"] == 0,
            f"crop={crop}",
        )
        png = _export(cover.com, exports)
        seen = {
            "top (was red)": nearest(pixel_at(png, 0.5, 0.04)) if png else "unreadable",
            "bottom (was blue)": nearest(pixel_at(png, 0.5, 0.96)) if png else "unreadable",
            "left (green)": nearest(pixel_at(png, 0.04, 0.5)) if png else "unreadable",
            "right (yellow)": nearest(pixel_at(png, 0.96, 0.5)) if png else "unreadable",
        }
        check(
            "cover kept the right content",
            seen["top (was red)"] != nearest(RED)
            and seen["bottom (was blue)"] != nearest(BLUE)
            and seen["left (green)"] == nearest(GREEN)
            and seen["right (yellow)"] == nearest(YELLOW),
            str(seen),
        )
        with deck.edit("crop-fit spike: drop cover"):
            cover.delete()

        # ---- S5: crop_to_fit contain — nothing cropped, fitted + centred ------
        print("\nS5  crop_to_fit(fit='contain') into the same 2:1 box")
        with deck.edit("crop-fit spike: contain"):
            fitted = slide.shapes.add_picture(img_path, left=20.0, top=20.0)
            result = fitted.crop_to_fit(
                left=60.0, top=80.0, width=400.0, height=200.0, fit="contain"
            )
        geo, crop = result["geometry"], result["crop"]
        check("contain crops nothing", all(v == 0.0 for v in crop.values()), f"crop={crop}")
        check(
            "contain fits inside the box, aspect kept",
            abs(geo["height"] - 200.0) < TOL and abs(geo["width"] - 200.0) < TOL,
            f"{geo['width']:.2f}x{geo['height']:.2f} (square source, height-limited by the 200pt box)",
        )
        check(
            "contain centres the shortfall",
            abs(geo["left"] - (60.0 + (400.0 - geo["width"]) / 2.0)) < TOL
            and abs(geo["top"] - 80.0) < TOL,
            f"at ({geo['left']:.2f}, {geo['top']:.2f})",
        )
        png = _export(fitted.com, exports)
        seen = {
            edge: nearest(pixel_at(png, fx, fy)) if png else "unreadable"
            for edge, (fx, fy) in (
                ("top", (0.5, 0.04)),
                ("bottom", (0.5, 0.96)),
                ("left", (0.04, 0.5)),
                ("right", (0.96, 0.5)),
            )
        }
        check(
            "contain still shows all four stripes",
            seen
            == {
                "top": nearest(RED),
                "bottom": nearest(BLUE),
                "left": nearest(GREEN),
                "right": nearest(YELLOW),
            },
            str(seen),
        )
        with deck.edit("crop-fit spike: drop contain"):
            fitted.delete()

        return _finish(deck, slide, before, viewed, checks, exports)
    except Exception:
        _finish(deck, slide, before, viewed, checks, exports)
        raise


def _export(com_shape: Any, exports: list[str]) -> str | None:
    """Export one shape to a temp PNG and return the path (or None)."""
    fd, path = tempfile.mkstemp(prefix="pptlive_cropfitx_", suffix=".png")
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
    checks: list[tuple[str, bool, str]],
    exports: list[str],
) -> int:
    with deck.edit("crop-fit spike: clean up"):
        slide.delete()
    if viewed is not None and viewed <= len(deck.slides):
        deck.go_to(deck.slides[viewed])
    for path in exports:
        try:
            os.unlink(path)
        except OSError:
            pass
    print(f"\nnet-zero: {len(deck.slides) == before} ({before} slides before and after)")
    failed = [label for label, ok, _ in checks if not ok]
    print(f"\n{len(checks) - len(failed)}/{len(checks)} checks passed")
    if failed:
        print("FAILED: " + ", ".join(failed))
    return 1 if failed else 0


if __name__ == "__main__":
    print(f"(source image: {IMG}x{IMG} px, four coloured edge stripes)")
    sys.exit(main())

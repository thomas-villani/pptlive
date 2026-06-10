"""Snapshots — render slides to PNG so a vision model can *see* the deck.

The PowerPoint analog of wordlive's `_snapshot.py`, but *shorter*: where Word has
no pixel-faithful in-memory raster and wordlive routes through
`ExportAsFixedFormat` → PDF → PyMuPDF, PowerPoint's `Slide.Export(FileName,
FilterName, ScaleWidth, ScaleHeight)` already renders a slide to a sized PNG on
disk (verified in v0.4 — the reason pptlive exposes a size override for slides
but not shapes). So a snapshot is just: per slide, export to a temp file at a
capped size, read the bytes back, and clean up. No extra dependency.

The token lever is **`max_dim`** — a long-edge pixel cap. A vision model is
billed on an image's pixel **area**, not its DPI, so capping the long edge gives
a *predictable per-slide token budget*. And because every slide in a deck shares
one geometry (`PageSetup` is deck-wide), one `max_dim` yields a *uniform*
per-slide cost across the whole deck — the right knob for a cheap "did my styling
land across all slides" read. `max_dim=None` renders at the slide's native size.

A snapshot is a **read**: `Slide.Export` doesn't move the view or change the
Selection, so (unlike the edit verbs) it needs no `deck.edit(...)` fence.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING

from . import _com
from .constants import image_filter_for
from .exceptions import SlideNotFoundError

if TYPE_CHECKING:
    from ._presentation import Presentation
    from ._slides import Slide

#: PowerPoint exports a slide at 96 DPI natively (1 in = 72 pt), so the native
#: pixel size of a slide edge is its length in points × 96/72. We use this to
#: keep `max_dim` a *cap* — it only ever lowers resolution, never upscales.
_NATIVE_PX_PER_PT = 96.0 / 72.0


@dataclass(frozen=True)
class Snapshot:
    """One rendered slide of a deck.

    `slide` is the 1-based slide index; `image` is the encoded image bytes in the
    chosen `fmt` (PNG by default, JPEG when `fmt="jpg"`) — feed it straight to a
    vision model, or write it yourself. `path` is where the image was written when
    a `snapshot(out=...)` call saved it to disk, otherwise `None`.
    """

    slide: int
    image: bytes
    path: Path | None = None


def _capped_dims(
    slide_w_pt: float, slide_h_pt: float, max_dim: int | None
) -> tuple[int, int] | None:
    """Resolve `(width_px, height_px)` for a slide capped to `max_dim` on its long edge.

    The cap *only ever lowers* resolution: the target long edge is
    `min(max_dim, native_long_px)` where `native_long_px` is the slide's 96-DPI
    native pixel size, so a generous `max_dim` never upscales past native. Both
    dimensions scale together, preserving the slide's aspect ratio. Returns
    `None` when `max_dim` is `None` (render at native size — pass no size override
    to `Slide.Export`). Pure / COM-free, so it unit-tests without PowerPoint.
    """
    if max_dim is None:
        return None
    long_pt = max(slide_w_pt, slide_h_pt)
    if long_pt <= 0:
        return None
    native_long_px = long_pt * _NATIVE_PX_PER_PT
    target_long_px = min(float(max_dim), native_long_px)
    scale = target_long_px / long_pt
    return max(1, round(slide_w_pt * scale)), max(1, round(slide_h_pt * scale))


def _resolve_slides(deck: Presentation, slides: int | tuple[int, int] | None) -> list[Slide]:
    """Resolve a `slides` selector to the list of `Slide`s to render.

    `None` → every slide; an `int` → that single 1-based slide; a `(start, end)`
    tuple → the inclusive 1-based span. An out-of-range index/span raises
    `SlideNotFoundError` (exit 2).
    """
    all_slides = list(deck.slides)
    n = len(all_slides)
    if slides is None:
        return all_slides
    if isinstance(slides, bool):  # bool is an int subclass — reject it explicitly
        raise TypeError(f"slides must be int, (start, end), or None, got {type(slides).__name__}")
    if isinstance(slides, int):
        if not (1 <= slides <= n):
            raise SlideNotFoundError(slides)
        return [all_slides[slides - 1]]
    if isinstance(slides, tuple) and len(slides) == 2:
        start, end = int(slides[0]), int(slides[1])
        if not (1 <= start <= end <= n):
            raise SlideNotFoundError(start if not (1 <= start <= n) else end)
        return all_slides[start - 1 : end]
    raise TypeError(f"slides must be int, (start, end), or None, got {type(slides).__name__}")


def render(
    deck: Presentation,
    *,
    slides: int | tuple[int, int] | None = None,
    fmt: str = "png",
    max_dim: int | None = None,
) -> list[tuple[int, bytes]]:
    """Render the selected slides to `(slide_index, png_bytes)` pairs.

    Each slide is exported to a temp file at the `max_dim`-capped size via
    `Slide.export_image`, read back to bytes, and the temp file removed. The
    deck's `PageSetup` (slide geometry, deck-wide) is read once to size the cap.
    Isolated from file placement (`build_snapshots`) so the rendering is testable
    on its own.
    """
    image_filter_for(fmt)  # validate the format up front (ValueError before any COM)
    targets = _resolve_slides(deck, slides)
    with _com.translate_com_errors():
        ps = deck.com.PageSetup
        w_pt, h_pt = float(ps.SlideWidth), float(ps.SlideHeight)
    dims = _capped_dims(w_pt, h_pt, max_dim)
    out: list[tuple[int, bytes]] = []
    for slide in targets:
        if dims is None:
            path = slide.export_image(None, fmt=fmt)
        else:
            path = slide.export_image(None, width=dims[0], height=dims[1], fmt=fmt)
        try:
            png = Path(path).read_bytes()
        finally:
            try:
                os.unlink(path)
            except OSError:
                pass
        out.append((slide.index, png))
    return out


def build_snapshots(
    rendered: list[tuple[int, bytes]], out: str | os.PathLike[str] | None
) -> list[Snapshot]:
    """Wrap `(slide, png)` pairs as `Snapshot`s, writing files when `out` is given.

    A single slide writes straight to `out`. Multiple slides can't share one
    path, so each is written next to `out` as `<stem>-s<N><suffix>` (N = slide
    index).
    """
    if out is None:
        return [Snapshot(slide=i, image=img, path=None) for i, img in rendered]
    out_path = Path(os.fspath(out))
    single = len(rendered) == 1
    snaps: list[Snapshot] = []
    for i, img in rendered:
        dest = out_path if single else out_path.with_name(f"{out_path.stem}-s{i}{out_path.suffix}")
        dest.write_bytes(img)
        snaps.append(Snapshot(slide=i, image=img, path=dest))
    return snaps


def snapshot(
    deck: Presentation,
    out: str | os.PathLike[str] | None = None,
    *,
    slides: int | tuple[int, int] | None = None,
    fmt: str = "png",
    max_dim: int | None = None,
) -> list[Snapshot]:
    """Render slides to PNG and (optionally) write them — see `Presentation.snapshot`."""
    return build_snapshots(render(deck, slides=slides, fmt=fmt, max_dim=max_dim), out)

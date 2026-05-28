"""Viewed-slide + Selection snapshot/restore — the politeness primitives.

The PowerPoint analog of wordlive's `_selection.py`, redesigned for the 2-D
object model. Where Word snapshots a single cursor offset, here we snapshot:

  - the **viewed slide** (`ActiveWindow.View.Slide.SlideIndex`) — the full-screen
    thing the user is looking at; moving it is the jarring "stomp", and
    restoring it is the whole point of `edit()`;
  - the **Selection** — its `PpSelectionType` plus, for a shape selection, the
    selected shapes' names so we can re-select them.

Everything is best-effort and swallows COM failures: politeness must never turn
a successful edit into an error. `ActiveWindow` may be absent (no open window, or
a slide show running), in which case there's nothing to preserve.

Spike item (IMPLEMENTATION.md): confirm a shape-range Selection round-trips
cleanly; if not, we already fall back to restoring just the viewed slide and
unselecting.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

from .constants import PpSelectionType

if TYPE_CHECKING:
    from ._app import PowerPoint


@dataclass(frozen=True)
class SelectionSnapshot:
    """A point-in-time capture of what the user is looking at and has selected."""

    slide_index: int | None
    """The viewed slide (1-based), or None if no Normal/Slide-view window exists."""
    selection_type: int = int(PpSelectionType.NONE)
    """The `PpSelectionType` at snapshot time."""
    shape_names: tuple[str, ...] = ()
    """Names of the selected shapes, when `selection_type` is SHAPES."""


def _active_window(ppt: PowerPoint) -> Any | None:
    try:
        return ppt.com.ActiveWindow
    except Exception:
        return None


def snapshot(ppt: PowerPoint) -> SelectionSnapshot:
    """Capture the viewed slide and current Selection. Never raises."""
    win = _active_window(ppt)
    if win is None:
        return SelectionSnapshot(slide_index=None)

    slide_index: int | None = None
    try:
        slide_index = int(win.View.Slide.SlideIndex)
    except Exception:
        slide_index = None

    selection_type = int(PpSelectionType.NONE)
    shape_names: tuple[str, ...] = ()
    try:
        sel = win.Selection
        selection_type = int(sel.Type)
        if selection_type == int(PpSelectionType.SHAPES):
            shape_names = tuple(str(sh.Name) for sh in sel.ShapeRange)
    except Exception:
        selection_type = int(PpSelectionType.NONE)
        shape_names = ()

    return SelectionSnapshot(
        slide_index=slide_index,
        selection_type=selection_type,
        shape_names=shape_names,
    )


def restore(ppt: PowerPoint, snap: SelectionSnapshot) -> None:
    """Best-effort restore of the viewed slide and Selection. Never raises."""
    win = _active_window(ppt)
    if win is None:
        return

    # 1. Put the user back on the slide they were looking at.
    if snap.slide_index is not None:
        try:
            win.View.GotoSlide(snap.slide_index)
        except Exception:
            pass

    # 2. Restore the Selection. A shape selection is re-selected by name on the
    #    restored slide; anything else (None/Slides, or a text selection we
    #    can't faithfully round-trip) collapses to "no selection".
    try:
        if snap.selection_type == int(PpSelectionType.SHAPES) and snap.shape_names:
            try:
                slide = win.View.Slide
                slide.Shapes.Range(list(snap.shape_names)).Select()
            except Exception:
                _unselect(win)
        else:
            _unselect(win)
    except Exception:
        pass


def _unselect(win: Any) -> None:
    try:
        win.Selection.Unselect()
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Reading the live Selection (v0.4) — the cursor/`here` marker, polite by being
# a pure read. This is the *query* side; `Presentation.anchor_by_id("here:")` is
# the opt-in *target* side (both resolve through the same Type -> anchor map).
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SelectionInfo:
    """What the user currently has selected, resolved to pptlive anchors.

    `type` is `"none"` / `"slides"` / `"shapes"` / `"text"`. For a shape
    selection, `shapes` lists each selected shape (`anchor_id`/`name`/`id`/
    `index`); for a text selection, `paragraph` and `text` describe the caret's
    paragraph. `anchor_id` is the single targetable anchor (`here:` resolves to
    it) — the first selected shape, or the text paragraph — or None when nothing
    is targetable (empty/slide selection).
    """

    type: str
    slide: int | None
    shapes: tuple[dict[str, Any], ...] = ()
    shape_index: int | None = None
    paragraph: int | None = None
    text: str | None = None

    @property
    def anchor_id(self) -> str | None:
        if (
            self.type == "text"
            and self.slide is not None
            and self.shape_index is not None
            and self.paragraph is not None
        ):
            return f"para:{self.slide}:{self.shape_index}:{self.paragraph}"
        if self.type == "shapes" and self.slide is not None and self.shape_index is not None:
            return f"shape:{self.slide}:{self.shape_index}"
        return None

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": self.type,
            "slide": self.slide,
            "anchor_id": self.anchor_id,
            "shapes": [dict(s) for s in self.shapes],
            "paragraph": self.paragraph,
            "text": self.text,
        }


def _zorder_index(slide_com: Any, shape_id: int) -> int | None:
    """Map a stable `Shape.Id` back to its 1-based z-order index on the slide."""
    try:
        for idx, sh in enumerate(slide_com.Shapes, start=1):
            if int(sh.Id) == int(shape_id):
                return idx
    except Exception:
        return None
    return None


def _shape_entry(slide_index: int | None, com_shape: Any, z: int | None) -> dict[str, Any]:
    return {
        "anchor_id": f"shape:{slide_index}:{z}" if z is not None else None,
        "name": str(com_shape.Name),
        "id": int(com_shape.Id),
        "index": z,
    }


def read_selection(ppt: PowerPoint) -> SelectionInfo:
    """Resolve the active window's current Selection to pptlive anchors.

    Read-only and best-effort: it never raises and never perturbs the selection.
    Returns `type="none"` when nothing is selected or no window is open. A SLIDES
    selection (sorter/thumbnail pane) reports the slide but has no text/shape
    anchor; SHAPES and TEXT resolve to `shape:S:N` / `para:S:N:P`.
    """
    win = _active_window(ppt)
    if win is None:
        return SelectionInfo("none", None)

    slide_com: Any | None
    slide_index: int | None
    try:
        slide_com = win.View.Slide
        slide_index = int(slide_com.SlideIndex)
    except Exception:
        slide_com, slide_index = None, None

    try:
        sel = win.Selection
        stype = int(sel.Type)
    except Exception:
        return SelectionInfo("none", slide_index)

    if stype == int(PpSelectionType.SHAPES) and slide_com is not None:
        shapes: list[dict[str, Any]] = []
        try:
            for sh in sel.ShapeRange:
                shapes.append(_shape_entry(slide_index, sh, _zorder_index(slide_com, int(sh.Id))))
        except Exception:
            shapes = []
        first_z = shapes[0]["index"] if shapes else None
        return SelectionInfo("shapes", slide_index, tuple(shapes), shape_index=first_z)

    if stype == int(PpSelectionType.TEXT) and slide_com is not None:
        try:
            host = sel.ShapeRange(1)
            z = _zorder_index(slide_com, int(host.Id))
            full = str(host.TextFrame.TextRange.Text)
            start = int(sel.TextRange.Start)  # 1-based char offset of the caret
            paragraph = full[: max(start - 1, 0)].count("\r") + 1
            return SelectionInfo(
                "text",
                slide_index,
                (_shape_entry(slide_index, host, z),),
                shape_index=z,
                paragraph=paragraph,
                text=str(sel.TextRange.Text),
            )
        except Exception:
            return SelectionInfo("text", slide_index)

    if stype == int(PpSelectionType.SLIDES):
        si = slide_index
        try:
            si = int(sel.SlideRange(1).SlideIndex)
        except Exception:
            pass
        return SelectionInfo("slides", si)

    return SelectionInfo("none", slide_index)

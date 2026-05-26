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

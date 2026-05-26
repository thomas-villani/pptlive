"""EditScope: a view/Selection-preservation scope — NOT an atomic-undo scope.

This is the headline PowerPoint diff from wordlive. Word's `EditScope` opens an
`Application.UndoRecord` (`StartCustomRecord`/`EndCustomRecord`) so a whole block
of mutations collapses into a single Ctrl-Z. **PowerPoint has no equivalent** —
there is no API to group programmatic mutations into one undo step. So pptlive's
`edit()` only snapshots and restores the user's *viewed slide and Selection*;
each mutation inside the block remains its own separate Ctrl-Z entry.

Say this loudly and never imply parity: a 3-op `edit()` block is 3 undo steps,
not 1. The user keeps full Ctrl-Z — just N presses, not one.

(Spike item: confirm no undo-grouping primitive snuck into a recent build before
finalizing. If one ever appears, this is where it would be wired in.)
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _selection
from ._selection import SelectionSnapshot

if TYPE_CHECKING:
    from ._app import PowerPoint


class EditScope:
    """Snapshots the viewed slide + Selection on enter; restores them on exit.

    ```
    with deck.edit("Revise agenda slide"):
        deck.anchor_by_id("ph:2:title").set_text("Agenda")
        deck.anchor_by_id("ph:2:body").set_text("Intro\\nDemo\\nQ&A")
    ```

    On clean exit the user is returned to the slide they were looking at, with
    their shape selection re-selected — unless code inside the scope called
    `allow_view_move()` (the analog of wordlive's `allow_cursor_move()`), which
    opts out so a deliberate `go_to`/jump survives. If the block raises, the view
    is left wherever the failing op put it, so the user can see what happened.

    **No UndoRecord** — see the module docstring. This scope does not make the
    enclosed mutations atomically undoable.
    """

    def __init__(self, ppt: PowerPoint, label: str) -> None:
        self._ppt = ppt
        self._label = label
        self._snapshot: SelectionSnapshot | None = None
        self._move_allowed = False

    @property
    def ppt(self) -> PowerPoint:
        return self._ppt

    @property
    def label(self) -> str:
        return self._label

    def allow_view_move(self) -> None:
        """Opt out of restoring the viewed slide + Selection on scope exit."""
        self._move_allowed = True

    def __enter__(self) -> EditScope:
        self._snapshot = _selection.snapshot(self._ppt)
        return self

    def __exit__(
        self, exc_type: type[BaseException] | None, exc: BaseException | None, tb: Any
    ) -> None:
        if exc_type is None and not self._move_allowed and self._snapshot is not None:
            try:
                _selection.restore(self._ppt, self._snapshot)
            except Exception:
                pass

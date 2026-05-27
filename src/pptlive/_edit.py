"""EditScope: an atomic-undo + view/Selection-preservation scope.

This is the headline PowerPoint diff from wordlive — and the 2026-05-26 spike
overturned the original pessimistic assumption. Word's `EditScope` brackets a
block with `Application.UndoRecord` (`StartCustomRecord`/`EndCustomRecord`) so
the whole block collapses into one Ctrl-Z. PowerPoint has **no `UndoRecord`** and
no explicit begin/end bracket — *but it turns out not to need one*: PowerPoint
**groups consecutive COM edits made within one automation session into a single
undo entry by default**, and `Application.StartNewUndoEntry()` is a boundary
primitive that starts a fresh entry. So `edit()` calls `StartNewUndoEntry()` on
entry to fence the block, and every mutation inside it collapses into a single
Ctrl-Z step. We can offer near-parity with wordlive's atomic undo after all.

The mechanism differs from Word's (a *start* fence + default grouping, not an
explicit start/end bracket), so two honest caveats remain:
  - there is no explicit "end" — the block is closed by the next `edit()` (which
    re-fences) or by the user's next manual action (which self-fences);
  - mutations made *outside* an `edit()` block aren't fenced and can merge into
    an adjacent entry, so always wrap mutations in `deck.edit(...)`.

Verified interactively via `scripts/undo_test.py`: 2 in-session edits → 1 Ctrl-Z
(with *and* without `StartNewUndoEntry`, so default grouping is real); and
`[e1, e2] · StartNewUndoEntry() · [e3]` → the 1st Ctrl-Z reverts only e3, the 2nd
reverts e1+e2 together (so `StartNewUndoEntry` is a verified boundary).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _selection
from ._selection import SelectionSnapshot

if TYPE_CHECKING:
    from ._app import PowerPoint


def _start_new_undo_entry(ppt: PowerPoint) -> None:
    """Fence the start of an edit block as a fresh undo entry (best-effort).

    PowerPoint then groups the subsequent in-session COM edits into this entry,
    so the whole block becomes one Ctrl-Z. Never raises: if the call is missing
    on some build, edits still group by default — we just lose the clean
    boundary. Operates on the raw COM object via the `.com` escape hatch, the
    same way `_selection` drives `ActiveWindow`.
    """
    try:
        ppt.com.StartNewUndoEntry()
    except Exception:
        pass


class EditScope:
    """Snapshots the viewed slide + Selection on enter; restores them on exit.

    ```
    with deck.edit("Revise agenda slide"):
        deck.anchor_by_id("ph:2:title").set_text("Agenda")
        deck.anchor_by_id("ph:2:body").set_text("Intro\\nDemo\\nQ&A")
    ```

    On enter it fences a fresh undo entry (`StartNewUndoEntry`), so the whole
    block is a **single Ctrl-Z** (see the module docstring). On clean exit the
    user is returned to the slide they were looking at, with their shape
    selection re-selected — unless code inside the scope called
    `allow_view_move()` (the analog of wordlive's `allow_cursor_move()`), which
    opts out so a deliberate `go_to`/jump survives. If the block raises, the view
    is left wherever the failing op put it, so the user can see what happened.
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
        _start_new_undo_entry(self._ppt)
        return self

    def __exit__(
        self, exc_type: type[BaseException] | None, exc: BaseException | None, tb: Any
    ) -> None:
        if exc_type is None and not self._move_allowed and self._snapshot is not None:
            try:
                _selection.restore(self._ppt, self._snapshot)
            except Exception:
                pass

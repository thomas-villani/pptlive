"""Politeness: viewed-slide + Selection snapshot/restore, and the edit() scope.

These prove the *view*-preservation behaviour and the atomic-undo fence: the
2026-05-26 spike showed PowerPoint groups in-session edits into one undo entry,
so edit() calls `StartNewUndoEntry()` on entry to fence the block (see EditScope).
"""

from __future__ import annotations

from pptlive import _selection


def test_snapshot_captures_viewed_slide_and_shape_selection(ppt, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._selection_type = 2  # PpSelectionType.SHAPES
    fake_powerpoint._selected_names = ("Content Placeholder 2",)

    snap = _selection.snapshot(ppt)
    assert snap.slide_index == 2
    assert snap.selection_type == 2
    assert snap.shape_names == ("Content Placeholder 2",)


def test_edit_restores_viewed_slide(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    with deck.edit("touch the deck"):
        # Simulate an operation that moved the view (e.g. an internal jump).
        fake_powerpoint.ActiveWindow.View.GotoSlide(3)
        assert fake_powerpoint._viewed == 3
    # On clean exit the user is put back on slide 1.
    assert fake_powerpoint._viewed == 1


def test_edit_fences_a_single_undo_entry(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # edit() calls StartNewUndoEntry() once on entry, so the whole block — no
    # matter how many mutations — collapses to one Ctrl-Z (PowerPoint groups
    # the rest of the in-session edits into that entry).
    assert fake_powerpoint._undo_entries == 0
    with deck.edit("multi-op block"):
        deck.anchor_by_id("ph:2:title").set_text("Agenda")
        deck.anchor_by_id("ph:2:body").set_text("Intro\nDemo\nQ&A")
    assert fake_powerpoint._undo_entries == 1


def test_each_edit_block_fences_its_own_entry(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # Consecutive blocks each re-fence, so they don't bleed into one another.
    with deck.edit("first"):
        deck.anchor_by_id("ph:2:title").set_text("A")
    with deck.edit("second"):
        deck.anchor_by_id("ph:2:title").set_text("B")
    assert fake_powerpoint._undo_entries == 2


def test_allow_view_move_opts_out_of_restore(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    with deck.edit("deliberate jump") as scope:
        scope.allow_view_move()
        fake_powerpoint.ActiveWindow.View.GotoSlide(3)
    assert fake_powerpoint._viewed == 3  # not snapped back


def test_edit_does_not_restore_on_exception(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    try:
        with deck.edit("boom"):
            fake_powerpoint.ActiveWindow.View.GotoSlide(3)
            raise RuntimeError("op failed")
    except RuntimeError:
        pass
    # Left where the failing op put it, so the user can see what happened.
    assert fake_powerpoint._viewed == 3


def test_restore_reselects_shapes(ppt, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    snap = _selection.SelectionSnapshot(slide_index=2, selection_type=2, shape_names=("Title 1",))
    fake_powerpoint._viewed = 1  # something moved the view away
    _selection.restore(ppt, snap)
    assert fake_powerpoint._viewed == 2  # slide restored
    assert fake_powerpoint._selected_names == ("Title 1",)  # selection restored


def test_go_to_moves_view_and_selects(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    deck.go_to(deck.anchor_by_id("shape:3:1"))
    assert fake_powerpoint._viewed == 3
    # The targeted shape got selected.
    target = fake_powerpoint.ActivePresentation.Slides(3).Shapes(1)
    assert target.selected is True

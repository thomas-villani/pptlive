"""Politeness: viewed-slide + Selection snapshot/restore, and the edit() scope.

These prove the *view*-preservation behaviour. There is deliberately no
atomic-undo assertion: PowerPoint has no UndoRecord, so edit() never groups
mutations into one Ctrl-Z (see EditScope).
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

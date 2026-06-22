"""Slide lifecycle (v0.1): add / delete / duplicate / move_to / set_layout, plus
layout-name resolution. Runs against the fake COM deck (3 slides, the standard
Office CustomLayouts). Indices are positional in the fake, so add/delete/move
shift them exactly as real PowerPoint does while `SlideID` stays stable.
"""

from __future__ import annotations

import pytest

from pptlive.exceptions import LayoutNotFoundError, PowerPointBusyError, SlideNotFoundError

# -- add --------------------------------------------------------------------


def test_add_appends_with_default_layout(deck) -> None:  # type: ignore[no-untyped-def]
    new = deck.slides.add()
    assert len(deck.slides) == 4
    assert new.index == 4
    assert new.layout_name == "Title and Content"  # the default alias
    ids = [s.id for s in deck.slides]
    assert new.id not in {256, 257, 258}  # fresh SlideID
    assert len(set(ids)) == 4  # all unique


def test_add_at_index_shifts_following(deck) -> None:  # type: ignore[no-untyped-def]
    new = deck.slides.add(layout="blank", index=2)
    assert new.index == 2
    assert new.layout_name == "Blank"
    assert deck.slides[3].title == "Agenda"  # old slide 2 pushed to 3
    assert len(deck.slides) == 4


def test_add_resolves_friendly_layout_names(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides.add(layout="two_content").layout_name == "Two Content"
    assert deck.slides.add(layout="Section Header").layout_name == "Section Header"
    assert deck.slides.add(layout="title").layout_name == "Title Slide"  # alias
    assert deck.slides.add(layout="content").layout_name == "Title and Content"  # alias


def test_add_layout_by_index(deck) -> None:  # type: ignore[no-untyped-def]
    # A 1-based index into CustomLayouts; "Title Slide" is layout 1.
    assert deck.slides.add(layout=1).layout_name == "Title Slide"


def test_add_unknown_layout_raises_with_available(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(LayoutNotFoundError) as ei:
        deck.slides.add(layout="nonsense")
    assert "Two Content" in str(ei.value)  # message lists the real names
    assert "Title and Content" in ei.value.available


def test_add_out_of_range_index_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.slides.add(index=0)
    with pytest.raises(SlideNotFoundError):
        deck.slides.add(index=99)
    # count+1 (append position) is valid and does not raise.
    deck.slides.add(index=len(deck.slides) + 1)


# -- add(placeholders=...) — reposition layout placeholders in one op --------


def test_apply_placeholder_geometry_moves_and_resizes(deck) -> None:  # type: ignore[no-untyped-def]
    # Slide 2 has a body placeholder; shrink it to a left-half content area.
    deck.slides._apply_placeholder_geometry(2, {"body": {"left": 40, "top": 140, "width": 440}})
    geo = deck.anchor_by_id("ph:2:body").geometry()
    assert geo["left"] == 40
    assert geo["top"] == 140
    assert geo["width"] == 440


def test_add_rejects_bad_placeholder_geometry(deck) -> None:  # type: ignore[no-untyped-def]
    # Validation runs before the slide is created (an unknown key is a typo guard).
    before = len(deck.slides)
    with pytest.raises(ValueError, match="unknown key"):
        deck.slides.add(layout="blank", placeholders={"body": {"wide": 5}})
    with pytest.raises(ValueError, match="must be a number"):
        deck.slides.add(layout="blank", placeholders={"body": {"left": True}})
    with pytest.raises(ValueError, match="non-empty geometry dict"):
        deck.slides.add(layout="blank", placeholders={"body": {}})
    assert len(deck.slides) == before  # nothing was created on a rejected call


def test_apply_placeholder_geometry_is_atomic_on_bad_kind(deck) -> None:  # type: ignore[no-untyped-def]
    from pptlive.exceptions import AnchorNotFoundError

    # A later KIND that doesn't resolve must abort *before* any earlier
    # placeholder is moved — slide 2 has no subtitle, so the whole call fails
    # and the (valid) body placeholder keeps its original geometry.
    before = deck.anchor_by_id("ph:2:body").geometry()
    with pytest.raises(AnchorNotFoundError):
        deck.slides._apply_placeholder_geometry(
            2, {"body": {"left": 40, "width": 440}, "subtitle": {"left": 10}}
        )
    after = deck.anchor_by_id("ph:2:body").geometry()
    assert (after["left"], after["width"]) == (before["left"], before["width"])


# -- delete -----------------------------------------------------------------


def test_delete_shifts_indices(deck) -> None:  # type: ignore[no-untyped-def]
    deck.slides[2].delete()
    assert len(deck.slides) == 2
    assert deck.slides[2].layout_name == "Blank"  # old slide 3 now at 2
    assert [s.id for s in deck.slides] == [256, 258]


# -- duplicate --------------------------------------------------------------


def test_duplicate_inserts_after_with_new_id(deck) -> None:  # type: ignore[no-untyped-def]
    new = deck.slides[1].duplicate()
    assert len(deck.slides) == 4
    assert new.index == 2  # immediately after the original
    assert new.id != 256
    assert new.title == "Welcome"  # content copied
    assert deck.slides[1].id == 256  # original stays put


# -- move_to ----------------------------------------------------------------


def test_move_to_reorders(deck) -> None:  # type: ignore[no-untyped-def]
    moved = deck.slides[1].move_to(3)
    assert moved.index == 3
    assert moved.id == 256  # same slide, new position
    assert [s.id for s in deck.slides] == [257, 258, 256]


def test_move_to_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.slides[1].move_to(0)
    with pytest.raises(SlideNotFoundError):
        deck.slides[1].move_to(99)


# -- set_layout -------------------------------------------------------------


def test_set_layout_changes_name(deck) -> None:  # type: ignore[no-untyped-def]
    deck.slides[3].set_layout("two_content")  # slide 3 was Blank
    assert deck.slides[3].layout_name == "Two Content"


def test_set_layout_unknown_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(LayoutNotFoundError):
        deck.slides[1].set_layout("does-not-exist")


# -- layouts discovery ------------------------------------------------------


def test_layouts_lists_names(deck) -> None:  # type: ignore[no-untyped-def]
    rows = deck.layouts()
    assert rows[0] == {"index": 1, "name": "Title Slide"}
    names = [r["name"] for r in rows]
    assert "Two Content" in names
    assert len(rows) == 9


def test_custom_layouts_surfaces_busy_instead_of_empty(deck) -> None:  # type: ignore[no-untyped-def]
    # A transient busy reading the master's layouts must surface (exit 3), not be
    # masked as "no layouts" (which would silently fall back to legacy add).
    class _Boom:
        @property
        def CustomLayouts(self) -> object:
            raise PowerPointBusyError(hresult=0x80010001)

    deck.com.SlideMaster = _Boom()
    with pytest.raises(PowerPointBusyError):
        deck.layouts()

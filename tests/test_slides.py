"""Slide-level reads: list, read, outline, page setup, title, has_notes."""

from __future__ import annotations

import pytest

from pptlive.exceptions import SlideNotFoundError


def test_slides_list_shape(deck) -> None:  # type: ignore[no-untyped-def]
    rows = deck.slides.list()
    assert [r["index"] for r in rows] == [1, 2, 3]
    assert [r["id"] for r in rows] == [256, 257, 258]
    by_index = {r["index"]: r for r in rows}
    assert by_index[1]["title"] == "Welcome"
    assert by_index[1]["layout"] == "Title Slide"
    assert by_index[1]["has_notes"] is True
    assert by_index[2]["has_notes"] is False
    assert by_index[2]["shape_count"] == 3
    assert by_index[3]["title"] is None


def test_slide_indexing_is_one_based(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[1].index == 1
    assert len(deck.slides) == 3


def test_slide_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.slides[0]
    with pytest.raises(SlideNotFoundError):
        deck.slides[4]


def test_slide_read_payload(deck) -> None:  # type: ignore[no-untyped-def]
    grid = deck.slides[2].read()
    assert grid["index"] == 2
    assert grid["id"] == 257
    assert grid["layout"] == "Title and Content"
    assert grid["title"] == "Agenda"
    shapes = grid["shapes"]
    assert [s["anchor_id"] for s in shapes] == ["shape:2:1", "shape:2:2", "shape:2:3"]
    assert shapes[0]["placeholder"] == "title"
    assert shapes[1]["placeholder"] == "body"
    assert shapes[2]["type"] == "picture"
    assert shapes[2]["has_text_frame"] is False
    assert shapes[2]["text"] is None
    # geometry is in points
    assert shapes[2]["geometry"] == {
        "left": 400.0,
        "top": 120.0,
        "width": 300.0,
        "height": 200.0,
        "rotation": 0.0,
    }


def test_outline(deck) -> None:  # type: ignore[no-untyped-def]
    items = deck.outline()
    assert items[1] == {"slide": 2, "title": "Agenda", "bullets": ["Intro", "Demo", "Q&A"]}
    # Title slide has no body placeholder -> no bullets, but keeps its title.
    assert items[0]["title"] == "Welcome"
    assert items[0]["bullets"] == []
    assert items[2]["title"] is None


def test_page_setup_points(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.page_setup() == {"width": 960.0, "height": 540.0}


def test_iteration_yields_slides_in_order(deck) -> None:  # type: ignore[no-untyped-def]
    assert [s.index for s in deck.slides] == [1, 2, 3]

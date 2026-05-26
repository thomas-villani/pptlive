"""ShapeCollection + geometry verbs + drift-proof identity."""

from __future__ import annotations

import pytest

from pptlive.exceptions import AnchorNotFoundError


def test_shapes_list(deck) -> None:  # type: ignore[no-untyped-def]
    rows = deck.slides[2].shapes.list()
    assert len(rows) == 3
    assert rows[0]["name"] == "Title 1"
    assert rows[0]["id"] == 2
    assert rows[0]["type"] == "placeholder"


def test_shape_by_zorder_index(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[3]
    assert shape.anchor_id == "shape:2:3"
    assert shape.name == "Picture 3"
    assert shape.shape_id == 4


def test_shape_by_name(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes["Content Placeholder 2"]
    assert shape.anchor_id == "shape:2:2"
    assert shape.text == "Intro\rDemo\rQ&A"


def test_shape_by_unknown_name_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[2].shapes["Nonexistent"]


def test_shape_index_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[2].shapes[99]


def test_membership(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[2].shapes
    assert "Title 1" in shapes
    assert 1 in shapes
    assert 99 not in shapes
    assert "ghost" not in shapes


def test_geometry_read(deck) -> None:  # type: ignore[no-untyped-def]
    geo = deck.slides[2].shapes[3].geometry()
    assert geo["left"] == 400.0
    assert geo["height"] == 200.0


def test_move_and_resize(deck) -> None:  # type: ignore[no-untyped-def]
    picture = deck.slides[2].shapes[3]
    picture.move(top=140.0)
    picture.resize(width=320.0)
    geo = picture.geometry()
    assert geo["top"] == 140.0
    assert geo["width"] == 320.0
    # left/height unchanged
    assert geo["left"] == 400.0
    assert geo["height"] == 200.0


def test_move_with_no_args_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        deck.slides[2].shapes[3].move()


def test_shape_type_names(deck) -> None:  # type: ignore[no-untyped-def]
    types = [s["type"] for s in deck.slides[3].shapes.list()]
    assert types == ["textbox", "line"]

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


def test_shape_name_propagates_missing_shape(deck) -> None:  # type: ignore[no-untyped-def]
    # `.name` must NOT fabricate an anchor_id-shaped string for a vanished shape:
    # that would collide with the `shape:S:N` format and read as a real name. It
    # propagates the lookup failure, exactly like `shape_id` / `shape_type`.
    from pptlive import Shape

    ghost = Shape(deck.slides[2], 99)
    with pytest.raises(AnchorNotFoundError):
        _ = ghost.name


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


# -- creators (v0.2) --------------------------------------------------------


def test_add_textbox_appends_with_text(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[3].shapes
    before = len(shapes)
    box = shapes.add_textbox("Hello", left=50.0, top=60.0, width=200.0, height=40.0)
    assert len(shapes) == before + 1
    assert box.index == before + 1  # appended at the top of the z-order
    assert box.shape_type == "textbox"
    assert box.text == "Hello"
    geo = box.geometry()
    assert (geo["left"], geo["top"], geo["width"], geo["height"]) == (50.0, 60.0, 200.0, 40.0)


def test_add_textbox_defaults_geometry(deck) -> None:  # type: ignore[no-untyped-def]
    box = deck.slides[3].shapes.add_textbox()
    geo = box.geometry()
    assert geo["left"] == 72.0 and geo["top"] == 72.0
    assert geo["width"] == 288.0 and geo["height"] == 72.0
    assert box.text == ""


def test_add_shape_friendly_name(deck) -> None:  # type: ignore[no-untyped-def]
    rect = deck.slides[3].shapes.add_shape("rectangle", left=10.0, top=10.0)
    assert rect.shape_type == "auto_shape"
    assert rect.com.AutoShapeType == 1  # MsoAutoShapeType.RECTANGLE


def test_add_shape_alias_and_separators(deck) -> None:  # type: ignore[no-untyped-def]
    # "ellipse" -> oval (9); "Rounded Rectangle" normalizes to roundedrectangle (5).
    assert deck.slides[3].shapes.add_shape("ellipse").com.AutoShapeType == 9
    assert deck.slides[3].shapes.add_shape("Rounded Rectangle").com.AutoShapeType == 5


def test_add_shape_raw_int_passthrough(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[3].shapes.add_shape(33).com.AutoShapeType == 33  # right arrow


def test_add_shape_unknown_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown autoshape"):
        deck.slides[3].shapes.add_shape("nonsense")


def test_add_picture_embeds(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    img = tmp_path / "logo.png"
    img.write_bytes(b"\x89PNG\r\n\x1a\n")  # contents irrelevant to the fake
    pic = deck.slides[3].shapes.add_picture(img, left=30.0, top=40.0)
    assert pic.shape_type == "picture"
    assert pic.has_text_frame is False
    geo = pic.geometry()
    assert geo["left"] == 30.0 and geo["top"] == 40.0


def test_add_picture_missing_file_raises(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(FileNotFoundError):
        deck.slides[3].shapes.add_picture(tmp_path / "absent.png")


# -- delete (v0.2) ----------------------------------------------------------


def test_delete_shape_shifts_indices(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[2].shapes
    assert len(shapes) == 3
    deck.slides[2].shapes[1].delete()  # was "Title 1"
    assert len(shapes) == 2
    # The old shape 2 ("Content Placeholder 2") is now at z-order 1.
    assert deck.slides[2].shapes[1].name == "Content Placeholder 2"

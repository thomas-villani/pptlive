"""ShapeCollection + geometry verbs + drift-proof identity."""

from __future__ import annotations

import pytest

from pptlive.exceptions import AnchorNotFoundError, NoTextFrameError


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


def test_reset_to_layout_restores_placeholder(deck) -> None:  # type: ignore[no-untyped-def]
    body = deck.slides[2].shapes[2]  # the body placeholder, manually wrecked below
    body.move(left=5.0, top=5.0)
    body.resize(width=10.0, height=10.0)
    body.format_text(size=5.0)
    restored = body.reset_to_layout()
    # Geometry + default font come from the slide layout's body placeholder.
    assert restored == {
        "left": 66.0,
        "top": 143.75,
        "width": 828.0,
        "height": 342.625,
        "font_size": 28.0,
    }
    assert body.geometry()["width"] == 828.0
    assert float(body.com.TextFrame.TextRange.Font.Size) == 28.0


def test_reset_to_layout_on_non_placeholder_raises(deck) -> None:  # type: ignore[no-untyped-def]
    # Slide 3 shape 1 is a plain textbox, not a placeholder.
    with pytest.raises(ValueError, match="placeholder"):
        deck.slides[3].shapes[1].reset_to_layout()


def test_text_frame_status(deck) -> None:  # type: ignore[no-untyped-def]
    st = deck.slides[2].shapes[2].text_frame_status()
    assert st.autosize == "text_to_fit_shape"  # off TextFrame2, not the classic mixed
    assert st.word_wrap is True
    assert st.margins == {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}
    assert st.overflow_risk == "low"  # an autofit mode is active
    assert st.to_dict()["autosize"] == "text_to_fit_shape"


def test_text_frame_status_no_frame_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(NoTextFrameError):
        deck.slides[3].shapes[2].text_frame_status()  # "Line 2" has no text frame


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


# -- set_picture: re-source a picture in place (v-next) ----------------------


def _png(tmp_path, name: str = "new.png"):  # type: ignore[no-untyped-def]
    img = tmp_path / name
    img.write_bytes(b"\x89PNG\r\n\x1a\n")  # contents irrelevant to the fake
    return img


def test_set_picture_preserves_box_name_and_zorder(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    pic = deck.slides[2].shapes[3]  # "Picture 3" at z-order 3
    assert pic.shape_type == "picture"
    old_id = pic.shape_id
    old_geo = pic.geometry()
    new = pic.set_picture(_png(tmp_path))
    # A fresh, drift-proof handle to a new shape (the old Id is gone).
    assert new.anchor_id.startswith("shapeid:2:")
    assert new.shape_id != old_id
    assert new.shape_type == "picture"
    # Box, name, and z-order slot are preserved.
    assert new.geometry() == old_geo
    assert new.name == "Picture 3"
    assert new.index == 3  # restacked back to the old slot
    # Net shape count is unchanged (delete + re-insert).
    assert len(deck.slides[2].shapes) == 3


def test_set_picture_carries_alt_text(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    pic = deck.slides[2].shapes[3]
    pic.set_alt_text("company logo")
    new = pic.set_picture(_png(tmp_path))
    assert new.alt_text == "company logo"


def test_set_picture_alt_text_override(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    pic = deck.slides[2].shapes[3]
    pic.set_alt_text("old alt")
    new = pic.set_picture(_png(tmp_path), alt_text="new alt")
    assert new.alt_text == "new alt"


def test_set_picture_missing_file_raises(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(FileNotFoundError):
        deck.slides[2].shapes[3].set_picture(tmp_path / "absent.png")


def test_set_picture_on_non_picture_raises(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="needs a picture shape"):
        deck.slides[3].shapes[1].set_picture(_png(tmp_path))  # a textbox


def test_shapeid_index_follows_collection_not_zorderposition() -> None:
    # On a flat slide ZOrderPosition == Shapes-collection index, but with
    # grouped/placeholder orderings they diverge. ShapeById must report the
    # collection index (the basis shape:S:N resolves by), not ZOrderPosition —
    # else the emitted shape:S:N would point at a different shape.
    from types import SimpleNamespace

    from pptlive._shapes import ShapeById
    from pptlive._slides import Slide

    class _Coll:
        def __init__(self, shapes: list) -> None:  # type: ignore[type-arg]
            self._shapes = shapes

        @property
        def Count(self) -> int:
            return len(self._shapes)

        def __call__(self, idx: int) -> object:
            return self._shapes[idx - 1]

    # One shape at collection index 1 whose ZOrderPosition lies (reports 9).
    sh = SimpleNamespace(Id=42, ZOrderPosition=9)
    slide_com = SimpleNamespace(Shapes=_Coll([sh]), SlideIndex=3)
    handle = ShapeById(Slide(None, slide_com), 42)  # type: ignore[arg-type]
    assert handle.index == 1  # collection index, not ZOrderPosition (9)


# -- delete (v0.2) ----------------------------------------------------------


def test_delete_shape_shifts_indices(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[2].shapes
    assert len(shapes) == 3
    deck.slides[2].shapes[1].delete()  # was "Title 1"
    assert len(shapes) == 2
    # The old shape 2 ("Content Placeholder 2") is now at z-order 1.
    assert deck.slides[2].shapes[1].name == "Content Placeholder 2"

"""ShapeCollection + geometry verbs + drift-proof identity."""

from __future__ import annotations

import pytest

from pptlive import constants as K
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


def test_shape_by_duplicate_name_raises_ambiguous(deck) -> None:  # type: ignore[no-untyped-def]
    from pptlive.exceptions import AmbiguousMatchError

    # PowerPoint allows duplicate shape names; a name lookup that matches more
    # than one must surface the ambiguity (with shape:S:N candidates), not
    # silently pick the first — consistent with placeholder resolution.
    shapes = deck.slides[2].shapes
    shapes[1].com.Name = "Dup"
    shapes[2].com.Name = "Dup"
    with pytest.raises(AmbiguousMatchError) as exc:
        _ = shapes["Dup"]
    assert "shape:2:1" in str(exc.value) and "shape:2:2" in str(exc.value)
    assert "Dup" in shapes  # membership is still True (the name exists)


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


def test_text_frame_status_reports_vertical_anchor(deck) -> None:  # type: ignore[no-untyped-def]
    st = deck.slides[2].shapes[2].text_frame_status()
    assert st.vertical_anchor == "top"  # msoAnchorTop, the default
    assert st.to_dict()["vertical_anchor"] == "top"


def test_text_frame_status_no_frame_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(NoTextFrameError):
        deck.slides[3].shapes[2].text_frame_status()  # "Line 2" has no text frame


# -- set_text_frame (the setter half of the diagnostic) ---------------------


def test_set_text_frame_sets_every_knob(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[2]
    st = shape.set_text_frame(
        autosize="none", word_wrap=False, vertical_anchor="middle", margins=0.0
    )
    tf = shape.com.TextFrame
    assert int(shape.com.TextFrame2.AutoSize) == 0  # msoAutoSizeNone
    assert int(tf.WordWrap) == 0
    assert int(tf.VerticalAnchor) == 3  # msoAnchorMiddle
    assert (tf.MarginLeft, tf.MarginRight, tf.MarginTop, tf.MarginBottom) == (0.0, 0.0, 0.0, 0.0)
    # The return value is the resulting status, so a caller sees the effect.
    assert st.autosize == "none"
    assert st.word_wrap is False
    assert st.vertical_anchor == "middle"
    assert st.overflow_risk == "possible"  # autosize off — text can now clip


def test_set_text_frame_margins_scalar_then_per_edge_override(deck) -> None:  # type: ignore[no-untyped-def]
    # margins= sets all four; a specific margin_* then overrides that one edge.
    st = deck.slides[2].shapes[2].set_text_frame(margins=0.0, margin_left=12.0)
    assert st.margins == {"left": 12.0, "right": 0.0, "top": 0.0, "bottom": 0.0}


def test_set_text_frame_partial_leaves_others_alone(deck) -> None:  # type: ignore[no-untyped-def]
    st = deck.slides[2].shapes[2].set_text_frame(word_wrap=False)
    assert st.word_wrap is False
    assert st.autosize == "text_to_fit_shape"  # untouched
    assert st.margins["left"] == 7.2  # untouched


def test_set_text_frame_autosize_aliases(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[2]
    assert shape.set_text_frame(autosize="off").autosize == "none"
    assert shape.set_text_frame(autosize="grow").autosize == "shape_to_fit_text"
    assert shape.set_text_frame(autosize="shrink").autosize == "text_to_fit_shape"


def test_set_text_frame_accepts_its_own_canonical_names(deck) -> None:  # type: ignore[no-untyped-def]
    # Regression: the multi-word canonical names round-trip. They were initially
    # written into the lookup with underscores, but `_normalize_name` strips
    # non-alphanumerics — so every name the CHOICES list advertised was rejected
    # while the single-word aliases worked. The live spike caught it; the earlier
    # unit tests only exercised aliases, which is exactly why they missed it.
    shape = deck.slides[2].shapes[2]
    for name in K.AUTOSIZE_CHOICES:
        assert shape.set_text_frame(autosize=name).autosize == name
    for name in K.VERTICAL_ANCHOR_CHOICES:
        assert shape.set_text_frame(vertical_anchor=name).vertical_anchor == name


def test_every_advertised_choice_coerces() -> None:
    # The same trap, guarded at the constants layer for both new enums.
    for name in K.AUTOSIZE_CHOICES:
        assert K.autosize_name(K.autosize_for(name)) == name
    for name in K.VERTICAL_ANCHOR_CHOICES:
        assert K.vertical_anchor_name(K.vertical_anchor_for(name)) == name


def test_set_text_frame_anchor_aliases(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[2]
    assert shape.set_text_frame(vertical_anchor="center").vertical_anchor == "middle"
    assert shape.set_text_frame(vertical_anchor="Top Baseline").vertical_anchor == "top_baseline"


def test_set_text_frame_requires_an_argument(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="at least one"):
        deck.slides[2].shapes[2].set_text_frame()


def test_set_text_frame_rejects_bad_values_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[2]
    for kwargs in (
        {"autosize": "enormous"},
        {"vertical_anchor": "sideways"},
        {"margins": -1.0},
        {"margin_top": -0.5},
    ):
        with pytest.raises(ValueError):
            shape.set_text_frame(**kwargs)  # type: ignore[arg-type]
    # Nothing was written — the whole point of validating before COM.
    assert int(shape.com.TextFrame.WordWrap) == -1
    assert shape.com.TextFrame.MarginLeft == 7.2


def test_set_text_frame_no_frame_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(NoTextFrameError):
        deck.slides[3].shapes[2].set_text_frame(autosize="none")


def test_add_textbox_accepts_text_frame_kwargs(deck) -> None:  # type: ignore[no-untyped-def]
    # The reported gap: a new text box autosizes, so a passed height drifts.
    # autosize="none" + margins=0 is the precise-layout opener.
    box = deck.slides[3].shapes.add_textbox(
        "Tight", height=40.0, autosize="none", margins=0.0, vertical_anchor="middle"
    )
    st = box.text_frame_status()
    assert st.autosize == "none"
    assert st.vertical_anchor == "middle"
    assert st.margins == {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}


def test_add_shape_accepts_text_frame_kwargs(deck) -> None:  # type: ignore[no-untyped-def]
    rect = deck.slides[3].shapes.add_shape(
        "rectangle", text="Centered", vertical_anchor="middle", word_wrap=False
    )
    st = rect.text_frame_status()
    assert st.vertical_anchor == "middle"
    assert st.word_wrap is False


@pytest.mark.parametrize(
    ("verb", "args"),
    [("add_textbox", ("Pinned",)), ("add_shape", ("rectangle",))],
)
def test_creation_applies_text_frame_before_text(deck, monkeypatch, verb, args) -> None:  # type: ignore[no-untyped-def]
    # Regression, found by scripts/text_frame_setter_spike.py against live
    # PowerPoint: a new text box autofits, so text written FIRST resizes the box
    # and a later autosize="none" merely pins the already-wrong height (a 40pt box
    # arrived at 29.1pt). The frame must be configured before any text lands.
    #
    # Asserted by capturing the shape's text at the moment apply_text_frame runs:
    # if the order is right, there is nothing in the frame yet.
    from pptlive import _shapes

    seen: list[str] = []
    real = _shapes.apply_text_frame

    def spy(com_shape, **kwargs):  # type: ignore[no-untyped-def]
        seen.append(str(com_shape.TextFrame.TextRange.Text))
        return real(com_shape, **kwargs)

    monkeypatch.setattr(_shapes, "apply_text_frame", spy)
    shapes = deck.slides[3].shapes
    box = (
        getattr(shapes, verb)(*args, text="Pinned", autosize="none")
        if verb == "add_shape"
        else shapes.add_textbox(*args, autosize="none")
    )
    assert seen == [""]  # the frame was configured while the box was still empty
    assert box.text == "Pinned"
    assert box.text_frame_status().autosize == "none"


def test_add_textbox_rejects_bad_frame_kwargs_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    before = len(deck.slides[3].shapes)
    with pytest.raises(ValueError, match="unknown autosize"):
        deck.slides[3].shapes.add_textbox("x", autosize="nope")
    assert len(deck.slides[3].shapes) == before  # no half-created shape


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


def test_add_shape_writes_text(deck) -> None:  # type: ignore[no-untyped-def]
    # Parity with add_textbox: the CLI's `shape add --kind shape --text` used to
    # be the only surface that could do this in one call.
    rect = deck.slides[3].shapes.add_shape("rectangle", text="Label")
    assert rect.com.TextFrame.TextRange.Text == "Label"
    assert rect.text == "Label"


def test_add_shape_text_defaults_empty(deck) -> None:  # type: ignore[no-untyped-def]
    rect = deck.slides[3].shapes.add_shape("rectangle")
    assert rect.com.TextFrame.TextRange.Text == ""


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


# -- crop / crop_to_fit (v-next) ---------------------------------------------
#
# The fake picture ("Picture 3") displays 300x200 pt from a 600x400 picture-point
# image — a 0.5 display scale, deliberately NOT 1:1 (see `_FakePictureFormat`), so
# a wrapper that confused crop units with display units fails these outright.


def _pic(deck):  # type: ignore[no-untyped-def]
    return deck.slides[2].shapes[3]


def _visible_aspect(shape) -> float:  # type: ignore[no-untyped-def]
    """Aspect of the picture area still showing, from the fake's own ground truth.

    Reads the fake's native picture size directly (crop units) rather than going
    through any library helper, so this stays an independent check on the fit math.
    """
    pf = shape.com.PictureFormat
    width = pf._pic_w - pf.CropLeft - pf.CropRight
    height = pf._pic_h - pf.CropTop - pf.CropBottom
    return width / height


@pytest.mark.parametrize(
    ("source", "box", "expected"),
    [
        ((200.0, 100.0), (100.0, 100.0), (0.5, 0.0)),  # wide source -> trim the sides
        ((100.0, 200.0), (100.0, 100.0), (0.0, 0.5)),  # tall source -> trim top/bottom
        ((160.0, 90.0), (320.0, 180.0), (0.0, 0.0)),  # same aspect -> no crop at all
    ],
)
def test_cover_crop_fractions(source, box, expected) -> None:  # type: ignore[no-untyped-def]
    from pptlive._shapes import cover_crop_fractions

    got = cover_crop_fractions(source[0], source[1], box[0], box[1])
    assert got == pytest.approx(expected)


@pytest.mark.parametrize(
    ("source", "box", "expected"),
    [
        ((200.0, 100.0), (100.0, 100.0), (100.0, 50.0)),  # width-limited
        ((100.0, 200.0), (100.0, 100.0), (50.0, 100.0)),  # height-limited
        ((160.0, 90.0), (320.0, 180.0), (320.0, 180.0)),  # exact fit, scaled up
    ],
)
def test_contain_size(source, box, expected) -> None:  # type: ignore[no-untyped-def]
    from pptlive._shapes import contain_size

    got = contain_size(source[0], source[1], box[0], box[1])
    assert got == pytest.approx(expected)


def test_crop_trims_and_shrinks_the_box(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    result = pic.crop(left=60.0)
    assert result["crop"] == {"left": 60.0, "right": 0.0, "top": 0.0, "bottom": 0.0}
    # Crop spike A4: cropping SHRINKS the shape box rather than letterboxing.
    # 60 picture-points off a 600-wide picture at 0.5 scale = 30 display points.
    assert result["geometry"]["width"] == pytest.approx(270.0)
    assert result["geometry"]["height"] == pytest.approx(200.0)


def test_crop_only_touches_the_edges_passed(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    pic.crop(top=40.0)
    result = pic.crop(left=60.0)  # must not reset the top crop
    assert result["crop"] == {"left": 60.0, "right": 0.0, "top": 40.0, "bottom": 0.0}


def test_crop_zero_uncrops_an_edge(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    pic.crop(left=60.0)
    result = pic.crop(left=0.0)
    assert result["crop"]["left"] == 0.0
    assert result["geometry"]["width"] == pytest.approx(300.0)


def test_crop_requires_an_edge(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="at least one of"):
        _pic(deck).crop()


def test_crop_rejects_negative(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="must be >= 0"):
        _pic(deck).crop(left=-5.0)


def test_crop_rejects_consuming_the_whole_picture(deck) -> None:  # type: ignore[no-untyped-def]
    # 400 + 300 >= the picture's 600 picture-point width.
    with pytest.raises(ValueError, match="consumes the picture"):
        _pic(deck).crop(left=400.0, right=300.0)


def test_crop_guard_uses_crop_units_not_display_points(deck) -> None:  # type: ignore[no-untyped-def]
    # The bug `scripts/crop_fit_spike.py` caught: `Crop.PictureWidth` is the
    # picture in DISPLAY points (300 here) while `Crop*` takes points of the
    # ORIGINAL picture (600 here). A guard that compared the two directly would
    # reject this perfectly valid 350pt crop.
    assert deck.slides[2].shapes[3].com.PictureFormat.Crop.PictureWidth == 300.0
    result = _pic(deck).crop(left=350.0)
    assert result["crop"]["left"] == 350.0
    assert result["geometry"]["width"] == pytest.approx(125.0)  # (600-350) * 0.5


def test_crop_extent_probe_preserves_an_existing_crop(deck) -> None:  # type: ignore[no-untyped-def]
    # Measuring the crop scale may need a probe crop; it must put the caller's
    # value back, not merely zero the edge it borrowed.
    pic = _pic(deck)
    pic.crop(left=90.0)
    assert pic.crop(bottom=40.0)["crop"]["left"] == 90.0


def test_crop_on_non_picture_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="needs a picture shape"):
        deck.slides[3].shapes[1].crop(left=10.0)  # a textbox


def test_crop_to_fit_cover_fills_the_box_exactly(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    result = pic.crop_to_fit(left=10.0, top=20.0, width=300.0, height=300.0)
    assert result["fit"] == "cover"
    geo = result["geometry"]
    assert (geo["left"], geo["top"]) == pytest.approx((10.0, 20.0))
    assert (geo["width"], geo["height"]) == pytest.approx((300.0, 300.0))
    # The 1.5:1 source loses the sides, symmetrically, and nothing off the top.
    assert result["crop"]["left"] == pytest.approx(result["crop"]["right"])
    assert result["crop"]["top"] == 0.0 and result["crop"]["bottom"] == 0.0
    # The real assertion: what still shows matches the box's aspect. This is what
    # a units bug breaks — a naive 1:1 crop would leave 1.25:1 showing in a 1:1 box.
    assert _visible_aspect(pic) == pytest.approx(300.0 / 300.0)


def test_crop_to_fit_cover_trims_top_and_bottom_for_a_tall_box(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    result = pic.crop_to_fit(width=600.0, height=100.0)  # 6:1, wider than the 1.5:1 source
    assert result["crop"]["left"] == 0.0 and result["crop"]["right"] == 0.0
    assert result["crop"]["top"] == pytest.approx(result["crop"]["bottom"])
    assert result["crop"]["top"] > 0.0
    assert _visible_aspect(pic) == pytest.approx(6.0)


def test_crop_to_fit_cover_leaves_a_matching_aspect_uncropped(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    result = pic.crop_to_fit(width=450.0, height=300.0)  # 1.5:1, the source's own aspect
    assert result["crop"] == {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}
    assert (result["geometry"]["width"], result["geometry"]["height"]) == pytest.approx(
        (450.0, 300.0)
    )


def test_crop_to_fit_contain_letterboxes_and_centres(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    result = pic.crop_to_fit(left=0.0, top=0.0, width=300.0, height=300.0, fit="contain")
    assert result["fit"] == "contain"
    # Nothing is cropped — the whole picture shows.
    assert result["crop"] == {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}
    geo = result["geometry"]
    # Width-limited: 300 wide, 200 tall, centred vertically in the 300pt box.
    assert (geo["width"], geo["height"]) == pytest.approx((300.0, 200.0))
    assert (geo["left"], geo["top"]) == pytest.approx((0.0, 50.0))


def test_crop_to_fit_defaults_to_the_current_box(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    before = pic.geometry()
    result = pic.crop_to_fit(width=100.0)  # left/top/height all default
    assert (result["geometry"]["left"], result["geometry"]["top"]) == pytest.approx(
        (before["left"], before["top"])
    )
    assert result["geometry"]["height"] == pytest.approx(before["height"])
    assert result["geometry"]["width"] == pytest.approx(100.0)


def test_crop_to_fit_is_idempotent(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    once = pic.crop_to_fit(width=300.0, height=300.0)
    twice = pic.crop_to_fit(width=300.0, height=300.0)
    assert once["crop"] == pytest.approx(twice["crop"])
    assert once["geometry"] == pytest.approx(twice["geometry"])


def test_crop_to_fit_clears_a_previous_crop(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    pic.crop(left=120.0, top=80.0)
    result = pic.crop_to_fit(width=450.0, height=300.0)  # the source's own aspect
    # Starts from the whole picture, so a matching aspect needs no crop at all.
    assert result["crop"] == {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}


def test_crop_to_fit_rejects_unknown_fit(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown fit"):
        _pic(deck).crop_to_fit(width=100.0, fit="stretch")


def test_crop_to_fit_rejects_non_positive_box(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="width must be > 0"):
        _pic(deck).crop_to_fit(width=0.0)


def test_crop_to_fit_on_non_picture_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="needs a picture shape"):
        deck.slides[3].shapes[1].crop_to_fit(width=100.0)


def test_every_advertised_fit_mode_is_accepted(deck) -> None:  # type: ignore[no-untyped-def]
    # The `set_text_frame` lesson: a CHOICES tuple must have every member actually
    # round-trip, or the CLI advertises a name the library refuses.
    from pptlive._shapes import FIT_MODES

    for mode in FIT_MODES:
        assert _pic(deck).crop_to_fit(width=200.0, height=150.0, fit=mode)["fit"] == mode


def test_read_reports_crop_only_when_cropped(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _pic(deck)
    assert "crop" not in pic.to_dict()
    pic.crop(left=60.0)
    assert pic.to_dict()["crop"]["left"] == 60.0
    # Non-pictures never carry the key.
    assert "crop" not in deck.slides[3].shapes[1].to_dict()


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

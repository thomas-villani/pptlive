"""Shape fill/line color (PPTLIVE-007), z-order (PPTLIVE-008), and the
delete-proof `shapeid:S:ID` anchor (PPTLIVE-010).

The fake deck's slide 2 has three shapes: Title 1 (id 2, z1), Content
Placeholder 2 (id 3, z2), Picture 3 (id 4, z3). Slide 3 has a textbox + a line.
"""

from __future__ import annotations

import pytest

from pptlive.constants import color_hex_or_none, zorder_cmd_for
from pptlive.exceptions import AnchorNotFoundError

# -- constants helpers ------------------------------------------------------


def test_zorder_cmd_for_friendly_and_aliases() -> None:
    assert zorder_cmd_for("front") == 0
    assert zorder_cmd_for("back") == 1
    assert zorder_cmd_for("forward") == 2
    assert zorder_cmd_for("backward") == 3
    # case / separator insensitive + verbose Office spellings
    assert zorder_cmd_for("Bring To Front") == 0
    assert zorder_cmd_for("send_to_back") == 1
    # raw int passes through
    assert zorder_cmd_for(3) == 3


def test_zorder_cmd_for_unknown_raises() -> None:
    with pytest.raises(ValueError, match="unknown z-order command"):
        zorder_cmd_for("sideways")


def test_color_hex_or_none_guards_theme_sentinel() -> None:
    assert color_hex_or_none(0x0000FF) == "#FF0000"  # COM RGB is 0xBBGGRR
    assert color_hex_or_none(0x80000000) is None  # automatic sentinel
    assert color_hex_or_none(-1) is None
    assert color_hex_or_none(0x1000000) is None  # out of RGB range


# -- 007: fill / line color -------------------------------------------------


def test_set_fill_solid_color(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_fill(fill="#FF0000")
    d = shape.to_dict()
    assert d["fill"] == {"type": "solid", "color": "#FF0000", "visible": True}


def test_set_fill_none_makes_transparent(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_fill(fill="none")
    assert shape.to_dict()["fill"]["visible"] is False


def test_set_fill_line_color_and_width(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_fill(line=(0, 0, 255), line_width=4.5)
    line = shape.to_dict()["line"]
    assert line["color"] == "#0000FF"
    assert line["weight"] == 4.5
    assert line["visible"] is True


def test_set_fill_line_none_removes_border(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_fill(line="none")
    assert shape.to_dict()["line"]["visible"] is False


def test_set_fill_bad_color_raises_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with pytest.raises(ValueError):
        shape.set_fill(fill="#ZZZ")
    # the bad call must not have mutated the fill
    assert shape.com.Fill.Type != 1  # Solid() never ran


def test_set_fill_requires_an_arg(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="at least one"):
        deck.slides[2].shapes[1].set_fill()


def test_shape_to_dict_default_fill_is_theme_none(deck) -> None:  # type: ignore[no-untyped-def]
    # An untouched shape inherits a theme fill -> color reads back as None, not #000000.
    d = deck.slides[2].shapes[1].to_dict()
    assert d["fill"] == {"type": "background", "color": None, "visible": True}
    assert d["line"]["color"] is None


def test_add_shape_with_fill_and_line(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[3].shapes
    rect = shapes.add_shape("rectangle", fill="#102030", line="none", line_width=2.0)
    d = rect.to_dict()
    assert d["fill"] == {"type": "solid", "color": "#102030", "visible": True}
    assert d["line"]["visible"] is False
    assert d["line"]["weight"] == 2.0


def test_add_shape_bad_fill_raises_before_creating(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[3].shapes
    before = len(shapes)
    with pytest.raises(ValueError):
        shapes.add_shape("rectangle", fill="not-a-color")
    assert len(shapes) == before  # nothing was created


def test_add_textbox_with_fill(deck) -> None:  # type: ignore[no-untyped-def]
    box = deck.slides[3].shapes.add_textbox("hi", fill="#FFFFFF")
    assert box.to_dict()["fill"] == {"type": "solid", "color": "#FFFFFF", "visible": True}


# -- 008: z-order -----------------------------------------------------------


def test_reorder_send_to_back(deck) -> None:  # type: ignore[no-untyped-def]
    # Picture 3 is the top shape (z3); send it to the back -> z1.
    pic = deck.slides[2].shapes[3]
    new_index = pic.reorder("back")
    assert new_index == 1
    assert deck.slides[2].shapes[1].name == "Picture 3"


def test_reorder_bring_to_front(deck) -> None:  # type: ignore[no-untyped-def]
    title = deck.slides[2].shapes[1]  # z1
    new_index = title.reorder("front")
    assert new_index == 3
    assert deck.slides[2].shapes[3].name == "Title 1"


def test_reorder_forward_and_backward(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[2].shapes
    # Content Placeholder 2 starts at z2; forward -> z3, then backward -> z2.
    content = shapes[2]
    assert content.reorder("forward") == 3
    assert shapes[3].name == "Content Placeholder 2"
    assert shapes[3].reorder("backward") == 2
    assert shapes[2].name == "Content Placeholder 2"


def test_reorder_unknown_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown z-order command"):
        deck.slides[2].shapes[1].reorder("up")


# -- 010: shapeid:S:ID delete-proof anchor ----------------------------------


def test_by_id_resolves(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes.by_id(4)  # Picture 3
    assert shape.anchor_id == "shapeid:2:4"
    assert shape.name == "Picture 3"
    assert shape.index == 3  # current z-order position
    assert shape.target_id == 4


def test_by_id_missing_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[2].shapes.by_id(999)


def test_anchor_by_id_shapeid(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.anchor_by_id("shapeid:2:3")
    assert shape.name == "Content Placeholder 2"
    assert shape.anchor_id == "shapeid:2:3"


def test_anchor_by_id_shapeid_malformed_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id("shapeid:2")
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id("shapeid:2:notanint")


def test_shapeid_survives_delete_that_shifts_index(deck) -> None:  # type: ignore[no-untyped-def]
    # Picture 3 is id 4 at z-order 3. Capture it by stable id, then delete the
    # shape below it: the z-order index shifts, but the shapeid still resolves.
    pic = deck.slides[2].shapes.by_id(4)
    assert pic.index == 3
    deck.slides[2].shapes[1].delete()  # delete Title 1 -> everything shifts down
    # shape:2:3 is now out of range, but shapeid:2:4 still points at Picture 3.
    assert pic.name == "Picture 3"
    assert pic.index == 2
    with pytest.raises(AnchorNotFoundError):
        deck.slides[2].shapes[3].name  # the old positional handle is now invalid


def test_shapeid_to_dict_reports_current_index(deck) -> None:  # type: ignore[no-untyped-def]
    pic = deck.slides[2].shapes.by_id(4)
    d = pic.to_dict()
    assert d["anchor_id"] == "shape:2:3"  # the dict still emits the canonical z-order id
    assert d["id"] == 4
    assert d["index"] == 3

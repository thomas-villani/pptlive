"""Live selection read (v0.4): `deck.selection()` + the `here:` anchor.

The fake's `_select_shapes` / `_select_text` / `_select_slide` drive
`ActiveWindow.Selection`; `read_selection` maps each `Selection.Type` to our
anchor vocabulary, and `anchor_by_id("here:")` resolves it to a live Shape or
Paragraph (the explicit opt-in to target the user's selection).
"""

from __future__ import annotations

import pytest

from pptlive import Paragraph, Shape
from pptlive.exceptions import AnchorNotFoundError


def test_selection_none_by_default(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.selection()
    assert info.type == "none"
    assert info.anchor_id is None


def test_selection_shapes(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Content Placeholder 2")
    info = deck.selection()
    assert info.type == "shapes"
    assert info.slide == 2
    assert info.shapes[0]["anchor_id"] == "shape:2:2"
    assert info.shapes[0]["name"] == "Content Placeholder 2"
    assert info.anchor_id == "shape:2:2"


def test_selection_multiple_shapes_anchor_is_first(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Title 1", "Picture 3")
    info = deck.selection()
    assert [s["anchor_id"] for s in info.shapes] == ["shape:2:1", "shape:2:3"]
    assert info.anchor_id == "shape:2:1"


def test_selection_text_recovers_paragraph(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_text("Content Placeholder 2", 2)  # body = Intro/Demo/Q&A
    info = deck.selection()
    assert info.type == "text"
    assert info.slide == 2
    assert info.paragraph == 2
    assert info.text == "Demo\r"
    assert info.anchor_id == "para:2:2:2"


def test_selection_slides(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    fake_powerpoint._select_slide()
    info = deck.selection()
    assert info.type == "slides"
    assert info.slide == 1
    assert info.anchor_id is None  # a slide is a container, not a text/shape anchor


def test_selection_to_dict_shape(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Title 1")
    d = deck.selection().to_dict()
    assert d["type"] == "shapes"
    assert d["anchor_id"] == "shape:2:1"
    assert d["shapes"][0]["id"] == 2


def test_here_resolves_to_selected_shape(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Content Placeholder 2")
    anchor = deck.anchor_by_id("here:")
    assert isinstance(anchor, Shape)
    assert anchor.anchor_id == "shape:2:2"
    assert anchor.text == "Intro\rDemo\rQ&A"


def test_here_resolves_to_selected_paragraph(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_text("Content Placeholder 2", 2)
    anchor = deck.anchor_by_id("here:")
    assert isinstance(anchor, Paragraph)
    assert anchor.anchor_id == "para:2:2:2"
    assert anchor.text == "Demo"


def test_here_with_no_selection_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id("here:")

"""anchor_by_id resolution + error mapping."""

from __future__ import annotations

import pytest

from pptlive import Notes, PlaceholderShape, Shape
from pptlive.exceptions import AnchorNotFoundError, SlideNotFoundError


def test_resolve_shape(deck) -> None:  # type: ignore[no-untyped-def]
    anchor = deck.anchor_by_id("shape:2:1")
    assert isinstance(anchor, Shape)
    assert anchor.text == "Agenda"


def test_resolve_placeholder(deck) -> None:  # type: ignore[no-untyped-def]
    anchor = deck.anchor_by_id("ph:2:body")
    assert isinstance(anchor, PlaceholderShape)
    assert anchor.text == "Intro\rDemo\rQ&A"


def test_resolve_notes(deck) -> None:  # type: ignore[no-untyped-def]
    anchor = deck.anchor_by_id("notes:1")
    assert isinstance(anchor, Notes)
    assert anchor.text == "Lead with the vision."


@pytest.mark.parametrize(
    "anchor_id",
    [
        "slide:2",
        "garbage",
        "shape:2",
        "shape:two:1",
        "ph:2",
        "ph:2:banner",
        "notes:x",
        "para:1:1:1",
    ],
)
def test_bad_anchor_ids_raise_anchor_not_found(deck, anchor_id) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id(anchor_id)


def test_missing_slide_raises_slide_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.anchor_by_id("shape:9:1")


def test_missing_shape_raises_anchor_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id("shape:2:99")

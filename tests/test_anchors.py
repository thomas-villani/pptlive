"""Text anchors: shapes, placeholders, notes, and the no-text-frame guard."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from pptlive import PlaceholderShape, Shape
from pptlive._anchors import _font_color_hex, font_to_dict
from pptlive.constants import tristate_value
from pptlive.exceptions import AnchorNotFoundError, NoTextFrameError


def test_tristate_value_preserves_mixed() -> None:
    assert tristate_value(-1) is True
    assert tristate_value(0) is False
    assert tristate_value(-2) == "mixed"  # msoTriStateMixed — the lost signal


def test_font_color_hex_guards_theme_sentinel() -> None:
    # A literal RGB renders; the 0x80000000 "automatic"/theme sentinel -> None.
    assert _font_color_hex(SimpleNamespace(Color=SimpleNamespace(RGB=0x38A7D0))) == "#D0A738"
    assert _font_color_hex(SimpleNamespace(Color=SimpleNamespace(RGB=-2147483648))) is None


def test_font_to_dict_reports_effective_attributes() -> None:
    tr = SimpleNamespace(
        Font=SimpleNamespace(
            Bold=-1,
            Italic=0,
            Underline=0,
            Size=24.0,
            Name="Georgia",
            Color=SimpleNamespace(RGB=0x0000FF),  # PowerPoint R-low-byte long
        )
    )
    font = font_to_dict(tr)
    assert font["bold"] is True and font["italic"] is False
    assert font["size"] == 24.0 and font["font"] == "Georgia"
    assert font["color"] == "#FF0000"


def test_shape_text_read(deck) -> None:  # type: ignore[no-untyped-def]
    body = deck.slides[2].shapes[2]
    assert body.text == "Intro\rDemo\rQ&A"


def test_shape_set_text_roundtrips(deck) -> None:  # type: ignore[no-untyped-def]
    title = deck.slides[2].shapes[1]
    title.set_text("New Agenda")
    assert title.text == "New Agenda"


def test_set_text_on_frameless_shape_raises_no_text_frame(deck) -> None:  # type: ignore[no-untyped-def]
    picture = deck.slides[2].shapes[3]
    assert picture.has_text_frame is False
    with pytest.raises(NoTextFrameError):
        picture.set_text("nope")
    with pytest.raises(NoTextFrameError):
        picture.text  # noqa: B018 — property access triggers the raising path


def test_placeholder_by_kind_title(deck) -> None:  # type: ignore[no-untyped-def]
    ph = deck.slides[2].placeholder("title")
    assert isinstance(ph, PlaceholderShape)
    assert ph.anchor_id == "ph:2:title"
    assert ph.text == "Agenda"


def test_placeholder_title_falls_back_to_center_title(deck) -> None:  # type: ignore[no-untyped-def]
    # Slide 1 is a Title Slide: its title is a *center* title (type 3); the
    # friendly "title" kind should still resolve it.
    ph = deck.slides[1].placeholder("title")
    assert ph.text == "Welcome"
    assert ph.placeholder_kind == "title"


def test_placeholder_body(deck) -> None:  # type: ignore[no-untyped-def]
    ph = deck.slides[2].placeholder("body")
    assert ph.text == "Intro\rDemo\rQ&A"
    # `\n` is normalized to a real paragraph break (`\r`), so the body becomes two
    # separately-addressable paragraphs (PPTLIVE-001).
    ph.set_text("One\nTwo")
    assert deck.slides[2].shapes[2].text == "One\rTwo"


def test_set_text_soft_break_stays_one_paragraph(deck) -> None:  # type: ignore[no-untyped-def]
    # `\v` (SOFT_BREAK) is a within-paragraph line break, not a paragraph break.
    ph = deck.slides[2].placeholder("body")
    ph.set_text("One\vTwo")
    assert deck.slides[2].shapes[2].text == "One\vTwo"
    assert ph.paragraph_count() == 1


def test_placeholder_missing_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[3].placeholder("body")  # Blank slide has no body placeholder


def test_unknown_placeholder_kind_is_value_error(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        deck.slides[2].placeholder("banner")


def test_notes_read_and_write(deck) -> None:  # type: ignore[no-untyped-def]
    notes = deck.slides[1].notes
    assert notes.anchor_id == "notes:1"
    assert notes.text == "Lead with the vision."
    notes.set_text("Open with the demo instead.")
    assert deck.slides[1].notes.text == "Open with the demo instead."


def test_notes_missing_body_raises(deck) -> None:  # type: ignore[no-untyped-def]
    # Slide 3's notes page has no body placeholder in the fake.
    with pytest.raises(AnchorNotFoundError):
        deck.slides[3].notes.text  # noqa: B018


def test_shape_com_returns_raw_shape_not_text_range(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    assert isinstance(shape, Shape)
    assert shape.com.Name == "Title 1"  # the raw Shape, exposing .Name

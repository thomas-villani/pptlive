"""Paragraph anchors + text-structure verbs (v0.3) against the fake deck.

Slide 2 shape 2 is the body placeholder "Intro\\rDemo\\rQ&A" (3 paragraphs); the
fake's paragraph-aware text model splits/joins on \\r exactly like PowerPoint's
TextRange (verified live in scripts/text_spike.py), so these prove the wrapper
logic — addressing, break-preserving set, insert, formatting, bullets.
"""

from __future__ import annotations

import pytest

from pptlive import Paragraph
from pptlive.exceptions import AnchorNotFoundError, NoTextFrameError


def _body(deck):  # type: ignore[no-untyped-def]
    return deck.slides[2].shapes[2]  # "Intro\rDemo\rQ&A"


def _body_com(deck):  # type: ignore[no-untyped-def]
    return deck.slides[2].shapes[2].com.TextFrame.TextRange


# -- addressing & reads -----------------------------------------------------


def test_paragraph_count_and_read(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    assert len(body.paragraphs) == 3
    assert body.paragraphs[2].text == "Demo"  # break stripped
    assert isinstance(body.paragraph(1), Paragraph)
    assert body.paragraph(1).anchor_id == "para:2:2:1"


def test_paragraphs_list(deck) -> None:  # type: ignore[no-untyped-def]
    rows = _body(deck).paragraphs.list()
    assert [r["text"] for r in rows] == ["Intro", "Demo", "Q&A"]
    assert rows[0]["anchor_id"] == "para:2:2:1"
    assert rows[0]["indent_level"] == 1
    assert rows[0]["bullet"] == "none"


def test_paragraph_index_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        _body(deck).paragraphs[9]


def test_paragraphs_on_frameless_shape_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(NoTextFrameError):
        len(deck.slides[3].shapes[2].paragraphs)  # slide 3 shape 2 is a Line


# -- set / insert / delete --------------------------------------------------


def test_set_paragraph_preserves_siblings(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    body.paragraphs[2].set_text("DEMO")
    assert body.text == "Intro\rDEMO\rQ&A"
    assert len(body.paragraphs) == 3


def test_insert_paragraph_after_middle(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[2].insert_paragraph_after("New")
    assert _body(deck).text == "Intro\rDemo\rNew\rQ&A"


def test_insert_paragraph_after_last(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[3].insert_paragraph_after("New")
    assert _body(deck).text == "Intro\rDemo\rQ&A\rNew"


def test_insert_paragraph_before(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[1].insert_paragraph_before("New")
    assert _body(deck).text == "New\rIntro\rDemo\rQ&A"


def test_append_paragraph_via_shape_anchor(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).insert_paragraph_after("New")  # whole-shape range -> append
    assert _body(deck).text == "Intro\rDemo\rQ&A\rNew"


def test_delete_paragraph(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[2].delete()  # remove "Demo"
    assert _body(deck).text == "Intro\rQ&A"
    assert len(_body(deck).paragraphs) == 2


# -- font formatting (the apply_style reframe) ------------------------------


def test_format_text_bold_size_color(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[1].format_text(bold=True, size=24.0, color="#FF0000")
    com = _body_com(deck).Paragraphs(1, 1)
    assert int(com.Font.Bold) == -1
    assert float(com.Font.Size) == 24.0
    assert int(com.Font.Color.RGB) == 255  # red -> 0x0000FF


def test_format_text_whole_shape_applies_to_all_paragraphs(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).format_text(italic=True)
    com = _body_com(deck)
    assert all(int(com.Paragraphs(p, 1).Font.Italic) == -1 for p in (1, 2, 3))


# -- paragraph formatting ---------------------------------------------------


def test_format_paragraph(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[1].format_paragraph(
        alignment="center", space_before=12.0, line_spacing=1.5, indent_level=2
    )
    com = _body_com(deck).Paragraphs(1, 1)
    assert int(com.ParagraphFormat.Alignment) == 2
    assert float(com.ParagraphFormat.SpaceBefore) == 12.0
    assert float(com.ParagraphFormat.SpaceWithin) == 1.5
    assert int(com.IndentLevel) == 2


def test_format_paragraph_bad_indent_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        _body(deck).paragraphs[1].format_paragraph(indent_level=9)


def test_format_paragraph_bad_alignment_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        _body(deck).paragraphs[1].format_paragraph(alignment="sideways")


# -- bullets / lists --------------------------------------------------------


def test_apply_and_remove_list(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    body.apply_list("numbered")
    com = _body_com(deck)
    assert int(com.ParagraphFormat.Bullet.Visible) == -1
    assert int(com.ParagraphFormat.Bullet.Type) == 2  # numbered

    body.apply_list("bulleted", character="•")
    assert int(com.ParagraphFormat.Bullet.Type) == 1  # unnumbered
    assert int(com.ParagraphFormat.Bullet.Character) == ord("•")

    body.remove_list()
    assert int(com.ParagraphFormat.Bullet.Visible) == 0


def test_apply_list_unknown_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        _body(deck).apply_list("circular")

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


def test_paragraph_diagnostics_spacing_and_runs(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    body.paragraphs[1].format_paragraph(line_spacing_points=24.0, space_before=6.0)
    row = body.paragraphs.list()[0]
    # Spacing reads carry the unit mode read off the paired LineRule*.
    assert row["line_spacing"] == {"value": 24.0, "mode": "points"}
    assert row["space_before"] == {"value": 6.0, "mode": "points"}
    assert row["space_after"]["mode"] == "points"  # default (0 pt)
    # run_sizes lists the distinct per-run font sizes (uniform here).
    body.paragraphs[1].format_text(size=18.0)
    assert _body(deck).paragraphs.list()[0]["run_sizes"] == [18.0]


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


def test_format_text_bad_color_does_not_half_format(deck) -> None:  # type: ignore[no-untyped-def]
    # color is validated before any font property is written, so a bad color
    # raises ValueError without leaving bold applied (no partial mutation).
    para = _body(deck).paragraphs[1]
    before = int(_body_com(deck).Paragraphs(1, 1).Font.Bold)
    with pytest.raises(ValueError):
        para.format_text(bold=True, color="not-a-color")
    assert int(_body_com(deck).Paragraphs(1, 1).Font.Bold) == before


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


def test_line_spacing_multiple_sets_msotrue(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[1].format_paragraph(line_spacing=1.5)
    pf = _body_com(deck).Paragraphs(1, 1).ParagraphFormat
    assert float(pf.SpaceWithin) == 1.5
    assert int(pf.LineRuleWithin) == -1  # msoTrue -> the value is a multiple


def test_line_spacing_points_sets_msofalse(deck) -> None:  # type: ignore[no-untyped-def]
    # The reviewer's "24" footgun, done right: exact points, not 24x.
    _body(deck).paragraphs[1].format_paragraph(line_spacing_points=24.0)
    pf = _body_com(deck).Paragraphs(1, 1).ParagraphFormat
    assert float(pf.SpaceWithin) == 24.0
    assert int(pf.LineRuleWithin) == 0  # msoFalse -> the value is in points


def test_space_before_points_vs_lines(deck) -> None:  # type: ignore[no-untyped-def]
    para = _body(deck).paragraphs[1]
    para.format_paragraph(space_before=18.0)
    pf = _body_com(deck).Paragraphs(1, 1).ParagraphFormat
    assert float(pf.SpaceBefore) == 18.0
    assert int(pf.LineRuleBefore) == 0  # points
    para.format_paragraph(space_after_lines=0.5)
    assert float(pf.SpaceAfter) == 0.5
    assert int(pf.LineRuleAfter) == -1  # multiple


def test_line_spacing_guardrail_rejects_big_multiple(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="line_spacing_points"):
        _body(deck).paragraphs[1].format_paragraph(line_spacing=24.0)


def test_line_spacing_guardrail_force_allows(deck) -> None:  # type: ignore[no-untyped-def]
    _body(deck).paragraphs[1].format_paragraph(line_spacing=24.0, force=True)
    pf = _body_com(deck).Paragraphs(1, 1).ParagraphFormat
    assert float(pf.SpaceWithin) == 24.0
    assert int(pf.LineRuleWithin) == -1  # still a (forced) multiple


def test_line_spacing_both_forms_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="not both"):
        _body(deck).paragraphs[1].format_paragraph(line_spacing=1.5, line_spacing_points=24.0)


def test_format_paragraph_bad_indent_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        _body(deck).paragraphs[1].format_paragraph(indent_level=9)


def test_format_paragraph_bad_alignment_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        _body(deck).paragraphs[1].format_paragraph(alignment="sideways")


# -- set_paragraphs (safe structured writing) -------------------------------


def test_set_paragraphs_builds_distinct_paragraphs(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    ids = body.set_paragraphs(["First", "Second", "Third"])
    assert body.text == "First\rSecond\rThird"
    assert ids == ["para:2:2:1", "para:2:2:2", "para:2:2:3"]


def test_set_paragraphs_applies_per_item_formatting(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    body.set_paragraphs(
        [
            {"text": "Heading", "size": 28.0, "alignment": "center"},
            {"text": "Point one", "list_type": "bulleted", "indent_level": 2},
        ]
    )
    com = _body_com(deck)
    assert int(com.Paragraphs(1, 1).ParagraphFormat.Alignment) == 2
    assert float(com.Paragraphs(1, 1).Font.Size) == 28.0
    assert int(com.Paragraphs(2, 1).ParagraphFormat.Bullet.Visible) == -1
    assert int(com.Paragraphs(2, 1).IndentLevel) == 2


def test_set_paragraphs_folds_inner_newline_to_soft_break(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    # A newline *inside* an item must not split it — it stays one paragraph.
    body.set_paragraphs(["Line A\nstill A", "Item B"])
    assert len(body.paragraphs) == 2
    assert body.paragraphs[1].text == "Line A\vstill A"


def test_set_paragraphs_guardrail_propagates(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="line_spacing_points"):
        _body(deck).set_paragraphs([{"text": "x", "line_spacing": 24.0}])


def test_set_paragraphs_rejects_bad_item(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        _body(deck).set_paragraphs([{"no_text": "oops"}])


def test_set_paragraphs_on_single_paragraph_anchor_raises(deck) -> None:  # type: ignore[no-untyped-def]
    # A `para:` anchor is a single paragraph; set_paragraphs replaces a whole
    # frame's list and would corrupt it / silently drop formatting. Reject it,
    # leaving the surrounding paragraphs untouched.
    para = _body(deck).paragraph(2)
    with pytest.raises(ValueError, match="para:2:2:2"):
        para.set_paragraphs(["a", "b", "c"])
    assert _body(deck).text == "Intro\rDemo\rQ&A"  # nothing was written


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


# -- reset_format (recovery) ------------------------------------------------


def test_reset_format_normalizes_spacing(deck) -> None:  # type: ignore[no-untyped-def]
    body = _body(deck)
    # Push it into the reviewer's bad state, then recover.
    body.format_paragraph(line_spacing_points=240.0, space_before=99.0, indent_level=4)
    body.reset_format()
    pf = _body_com(deck).ParagraphFormat
    assert float(pf.SpaceWithin) == 1.0
    assert int(pf.LineRuleWithin) == -1  # back to a multiple (single)
    assert float(pf.SpaceBefore) == 0.0
    assert float(pf.SpaceAfter) == 0.0
    assert int(_body_com(deck).IndentLevel) == 1

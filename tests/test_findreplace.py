"""find / find_replace — the fuzzy-matching core, the deck traversal, and the CLI.

Runs entirely against the `fake_powerpoint` MagicMock graph (no PowerPoint), so
it proves the traversal, offset math, scoping, and exit-code contract without COM.
The live behaviours these rest on were pinned by `scripts/findreplace_spike.py`.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive import _findreplace
from pptlive.cli.main import main
from pptlive.exceptions import AmbiguousMatchError, AnchorNotFoundError


def _json(result):  # type: ignore[no-untyped-def]
    return json.loads(result.output)


# ---------------------------------------------------------------------------
# The pure matching core (_findreplace.find_matches) — OS-independent
# ---------------------------------------------------------------------------


def test_find_matches_locates_every_occurrence() -> None:
    matches = _findreplace.find_matches("alpha beta alpha gamma alpha", "alpha")
    assert [m.start for m in matches] == [0, 11, 23]
    assert all(m.text == "alpha" for m in matches)


def test_find_matches_no_match_and_empty_needle() -> None:
    assert _findreplace.find_matches("hello world", "zzz") == []
    assert _findreplace.find_matches("hello", "") == []


def test_find_matches_is_smart_quote_fuzzy_but_keeps_original_text() -> None:
    # The needle has a straight apostrophe; the haystack a curly one — they match,
    # and the returned text is the *original* (curly) substring.
    [m] = _findreplace.find_matches("don’t stop", "don't")
    assert m.text == "don’t"
    assert (m.start, m.end) == (0, 5)


def test_find_matches_collapses_whitespace_runs() -> None:
    [m] = _findreplace.find_matches("foo   \t bar", "foo bar")
    assert m.text == "foo   \t bar"  # original span, whitespace and all


def test_find_matches_is_case_sensitive() -> None:
    assert _findreplace.find_matches("Demo demo", "demo") == [
        _findreplace.Match(start=5, end=9, text="demo")
    ]


# ---------------------------------------------------------------------------
# deck.find — the traversal
# ---------------------------------------------------------------------------


def test_find_reaches_a_shape_and_reports_a_resolvable_anchor(deck) -> None:  # type: ignore[no-untyped-def]
    hits = deck.find("Welcome")
    assert len(hits) == 1
    hit = hits[0]
    assert hit["anchor_id"] == "shape:1:1"
    assert hit["start"] == 0
    assert hit["length"] == len("Welcome")
    assert hit["text"] == "Welcome"
    # the anchor it reports actually resolves
    assert deck.anchor_by_id(hit["anchor_id"]).text == "Welcome"


def test_find_reaches_body_bullets_with_in_frame_offset(deck) -> None:  # type: ignore[no-untyped-def]
    # body is "Intro\rDemo\rQ&A" — "Demo" starts after "Intro\r" (offset 6).
    [hit] = deck.find("Demo")
    assert hit["anchor_id"] == "shape:2:2"
    assert hit["start"] == 6
    assert hit["text"] == "Demo"


def test_find_reaches_speaker_notes(deck) -> None:  # type: ignore[no-untyped-def]
    [hit] = deck.find("vision")
    assert hit["anchor_id"] == "notes:1"


def test_find_is_whole_deck_in_document_order(deck) -> None:  # type: ignore[no-untyped-def]
    # "de" (lowercase) occurs twice, both inside slide 1's subtitle "A demo deck".
    hits = deck.find("de")
    assert [h["anchor_id"] for h in hits] == ["shape:1:2", "shape:1:2"]
    assert [h["start"] for h in hits] == [2, 7]


def test_find_scope_slide_restricts_the_search(deck) -> None:  # type: ignore[no-untyped-def]
    deck.anchor_by_id("shape:3:1").set_text("Welcome back")  # add a 2nd "Welcome"
    assert len(deck.find("Welcome")) == 2
    assert len(deck.find("Welcome", scope="slide:1")) == 1


def test_find_scope_anchor_restricts_to_one_frame(deck) -> None:  # type: ignore[no-untyped-def]
    hits = deck.find("de", scope="shape:1:2")
    assert len(hits) == 2
    assert deck.find("de", scope="ph:2:title") == []  # "Agenda" has no "de"


def test_find_no_match_returns_empty_list(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.find("nonexistent-zzz") == []


def test_find_includes_a_context_snippet(deck) -> None:  # type: ignore[no-untyped-def]
    [hit] = deck.find("Demo")
    assert "Demo" in hit["context"]
    # Paragraph breaks render as a visible glyph (not a raw `\r`) for legibility.
    assert "\r" not in hit["context"]
    assert "⏎" in hit["context"]  # the body "Intro\rDemo\rQ&A" surrounds the hit


def test_find_reaches_table_cells(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("add table"):
        shape = deck.slides[3].shapes.add_table(2, 2)
        shape.table.cell(1, 2).set_text("needle in a cell")
    n = shape.index
    [hit] = deck.find("needle")
    assert hit["anchor_id"] == f"cell:3:{n}:1:2"
    assert hit["text"] == "needle"


# ---------------------------------------------------------------------------
# deck.find_replace — the write path
# ---------------------------------------------------------------------------


def test_find_replace_single_auto_applies(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("fr"):
        applied = deck.find_replace("Welcome", "Hello")
    assert len(applied) == 1
    assert applied[0]["anchor_id"] == "shape:1:1"
    assert deck.anchor_by_id("shape:1:1").text == "Hello"


def test_find_replace_only_rewrites_the_span_keeping_other_paragraphs(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("fr"):
        deck.find_replace("Demo", "Show")
    assert deck.anchor_by_id("shape:2:2").text == "Intro\rShow\rQ&A"


def test_find_replace_all_in_one_frame_applies_in_reverse(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("fr"):
        applied = deck.find_replace("de", "XX", all=True)
    assert len(applied) == 2
    assert deck.anchor_by_id("shape:1:2").text == "A XXmo XXck"


def test_find_replace_occurrence_picks_one(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("fr"):
        deck.find_replace("de", "X", occurrence=2)
    assert deck.anchor_by_id("shape:1:2").text == "A demo Xck"


def test_find_replace_ambiguous_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AmbiguousMatchError) as exc:
        deck.find_replace("de", "X")
    assert len(exc.value.matches) == 2


def test_find_replace_zero_matches_raises_anchor_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.find_replace("nonexistent-zzz", "x")


def test_find_replace_out_of_range_occurrence_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.find_replace("de", "x", occurrence=5)


def test_find_replace_does_not_re_match_inside_replacement(deck) -> None:  # type: ignore[no-untyped-def]
    # The live drift hazard: a naive replace-until-empty loop on a replacement
    # that re-contains the search text spins forever. We match once up front, so
    # "alpha" -> "alpha_X" is applied exactly once.
    deck.anchor_by_id("shape:3:1").set_text("alpha")
    with deck.edit("fr"):
        applied = deck.find_replace("alpha", "alpha_X", all=True)
    assert len(applied) == 1
    assert deck.anchor_by_id("shape:3:1").text == "alpha_X"


def test_find_replace_all_spans_multiple_frames(deck) -> None:  # type: ignore[no-untyped-def]
    deck.anchor_by_id("shape:2:1").set_text("needle one")
    deck.anchor_by_id("shape:3:1").set_text("needle two")
    with deck.edit("fr"):
        applied = deck.find_replace("needle", "N", all=True)
    assert len(applied) == 2
    assert deck.anchor_by_id("shape:2:1").text == "N one"
    assert deck.anchor_by_id("shape:3:1").text == "N two"


# ---------------------------------------------------------------------------
# CLI — find / replace --find, payloads and exit codes
# ---------------------------------------------------------------------------


def test_cli_find_emits_match_array(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["find", "--text", "Welcome"])
    assert result.exit_code == 0
    rows = _json(result)
    assert isinstance(rows, list)
    assert rows[0]["anchor_id"] == "shape:1:1"


def test_cli_find_no_match_is_empty_array_exit_0(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["find", "--text", "nonexistent-zzz"])
    assert result.exit_code == 0
    assert _json(result) == []


def test_cli_find_scoped(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["find", "--text", "de", "--in", "shape:1:2"])
    assert result.exit_code == 0
    assert len(_json(result)) == 2


def test_cli_replace_find_single(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["replace", "--find", "Welcome", "--text", "Hi"])
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["ok"] is True
    assert payload["count"] == 1
    assert fake_powerpoint.ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.Text == "Hi"


def test_cli_replace_find_all(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["replace", "--find", "de", "--text", "XX", "--all"])
    assert result.exit_code == 0
    assert _json(result)["count"] == 2
    assert fake_powerpoint.ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange.Text == (
        "A XXmo XXck"
    )


def test_cli_replace_find_ambiguous_exit_5(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["replace", "--find", "de", "--text", "X"])
    assert result.exit_code == 5
    # Failures follow the CLI contract: no JSON object on stdout, the actionable
    # hint goes to stderr (same as every other failure path). The message names
    # both the count and the disambiguators so an LLM driver can retry.
    assert "2 matches" in result.output
    assert "occurrence" in result.output


def test_cli_replace_find_zero_matches_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["replace", "--find", "nonexistent-zzz", "--text", "x"])
    assert result.exit_code == 2


def test_cli_replace_occurrence(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["replace", "--find", "de", "--text", "X", "--occurrence", "2"]
    )
    assert result.exit_code == 0
    assert fake_powerpoint.ActivePresentation.Slides(1).Shapes(2).TextFrame.TextRange.Text == (
        "A demo Xck"
    )


def test_cli_replace_anchor_mode_still_works(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["replace", "--anchor-id", "shape:3:1", "--text", "Z"])
    assert result.exit_code == 0
    assert fake_powerpoint.ActivePresentation.Slides(3).Shapes(1).TextFrame.TextRange.Text == "Z"


def test_cli_replace_requires_exactly_one_mode(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    both = CliRunner().invoke(
        main, ["replace", "--anchor-id", "shape:3:1", "--find", "x", "--text", "y"]
    )
    assert both.exit_code != 0
    assert "exactly one" in both.output
    neither = CliRunner().invoke(main, ["replace", "--text", "y"])
    assert neither.exit_code != 0


def test_cli_replace_flags_rejected_in_anchor_mode(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["replace", "--anchor-id", "shape:3:1", "--text", "y", "--all"]
    )
    assert result.exit_code != 0
    assert "only valid with --find" in result.output

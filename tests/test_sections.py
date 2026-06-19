"""Deck sections (v0.6.0 batch-2) — `Presentation.SectionProperties` wrapper.

The 1-based section model proved against the fake `_FakeSectionProperties`
(mirrors the batch2 spike: AddBeforeSlide starts a span + auto-creates a leading
"Default Section"; AddSection appends an empty boundary; Delete keeps slides).
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, ReadOp, _edit_core, _read_core
from pptlive.cli.main import main
from pptlive.exceptions import AnchorNotFoundError, SlideNotFoundError


def _three_slide_deck(deck):  # type: ignore[no-untyped-def]
    """The default fake deck has 3 slides — enough for multi-section spans."""
    assert len(deck.slides) == 3
    return deck


# -- library: add / list ----------------------------------------------------


def test_sections_empty_by_default(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.sections.list() == []
    assert len(deck.sections) == 0


def test_add_before_slide_auto_creates_default_section(deck) -> None:  # type: ignore[no-untyped-def]
    _three_slide_deck(deck)
    with deck.edit("s"):
        row = deck.sections.add("Results", before_slide=2)
    # Adding the first section before slide 2 auto-spawns a leading Default Section.
    rows = deck.sections.list()
    assert [r["name"] for r in rows] == ["Default Section", "Results"]
    assert rows[0]["first_slide"] == 1 and rows[0]["slide_count"] == 1
    assert rows[1]["first_slide"] == 2 and rows[1]["slide_count"] == 2
    assert row["name"] == "Results" and row["index"] == 2


def test_add_trailing_empty_section(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("s"):
        deck.sections.add("Intro", before_slide=1)
        row = deck.sections.add("Appendix")  # no before_slide -> empty trailing
    assert row["first_slide"] is None and row["slide_count"] == 0


def test_add_before_slide_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.sections.add("Nope", before_slide=99)


# -- rename / delete / move -------------------------------------------------


def test_rename(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("s"):
        deck.sections.add("Old", before_slide=1)
        row = deck.sections.rename(1, "New")
    assert row["name"] == "New"
    assert deck.sections.list()[0]["name"] == "New"


def test_delete_keeps_slides(deck) -> None:  # type: ignore[no-untyped-def]
    n = len(deck.slides)
    with deck.edit("s"):
        deck.sections.add("A", before_slide=1)
        out = deck.sections.delete(1)
    assert out["deleted"] is True and out["slides_deleted"] is False
    assert deck.sections.list() == []
    assert len(deck.slides) == n  # boundary gone, slides intact


def test_move(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("s"):
        deck.sections.add("First", before_slide=1)
        deck.sections.add("Second", before_slide=3)
        row = deck.sections.move(2, 1)
    assert row["index"] == 1 and row["name"] == "Second"


def test_rename_bad_index_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.sections.rename(5, "x")


# -- batch ops --------------------------------------------------------------


def test_batch_section_add_and_read(deck, ppt) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("s"):
        out = _edit_core(deck, EditOp.SECTION_ADD, {"name": "Part 1", "before_slide": 1})
    assert out["section"]["name"] == "Part 1"
    read = _read_core(ppt, ReadOp.SECTIONS, {})
    assert read["sections"][0]["name"] == "Part 1"


def test_batch_section_add_requires_name(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("s"), pytest.raises(BatchOpError, match="requires `name`"):
        _edit_core(deck, EditOp.SECTION_ADD, {"before_slide": 1})


def test_batch_section_delete_and_move(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("s"):
        _edit_core(deck, EditOp.SECTION_ADD, {"name": "A", "before_slide": 1})
        _edit_core(deck, EditOp.SECTION_ADD, {"name": "B", "before_slide": 3})
        moved = _edit_core(deck, EditOp.SECTION_MOVE, {"section": 2, "to": 1})
        assert moved["section"]["name"] == "B"
        out = _edit_core(deck, EditOp.SECTION_DELETE, {"section": 1})
    assert out["name"] == "B"


# -- CLI --------------------------------------------------------------------


def test_cli_section_add_list(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    add = runner.invoke(
        main, ["--json", "section", "add", "--name", "Intro", "--before-slide", "1"]
    )
    assert add.exit_code == 0, add.output
    assert json.loads(add.output)["section"]["name"] == "Intro"
    listing = runner.invoke(main, ["--json", "section", "list"])
    assert listing.exit_code == 0
    assert json.loads(listing.output)[0]["name"] == "Intro"


def test_cli_section_rename_move_delete(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    runner.invoke(main, ["section", "add", "--name", "A", "--before-slide", "1"])
    runner.invoke(main, ["section", "add", "--name", "B", "--before-slide", "3"])
    ren = runner.invoke(main, ["--json", "section", "rename", "--section", "1", "--name", "AA"])
    assert json.loads(ren.output)["section"]["name"] == "AA"
    mov = runner.invoke(main, ["--json", "section", "move", "--section", "2", "--to", "1"])
    assert json.loads(mov.output)["section"]["name"] == "B"
    dele = runner.invoke(main, ["--json", "section", "delete", "--section", "1"])
    assert json.loads(dele.output)["deleted"] is True


def test_cli_section_delete_bad_index_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["section", "delete", "--section", "9"])
    assert res.exit_code == 2  # AnchorNotFoundError

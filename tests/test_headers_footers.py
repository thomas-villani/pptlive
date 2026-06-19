"""Headers / footers (v0.6.0 batch-2) — footer / slide-number / date.

The shared `HeadersFooters` wrapper at both scopes (slide override + master
default), with the batch2-spike footgun modeled in the fake: text/use_format read
back None while the element is hidden, and round-trip once shown.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, ReadOp, _edit_core, _read_core
from pptlive.cli.main import main

# -- library: slide scope ---------------------------------------------------


def test_read_default_all_hidden(deck) -> None:  # type: ignore[no-untyped-def]
    hf = deck.slides[1].headers_footers.read()
    assert hf["footer"]["visible"] is False
    assert hf["footer"]["text"] is None  # guarded: not readable while hidden
    assert hf["slide_number"]["visible"] is False
    assert hf["date"]["visible"] is False
    assert hf["display_on_title_slide"] is None  # master-only -> None on a slide


def test_set_footer_auto_shows_and_text_reads_back(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        hf = deck.slides[1].headers_footers.set_footer(text="Confidential")
    assert hf["footer"]["visible"] is True
    assert hf["footer"]["text"] == "Confidential"


def test_set_footer_hide(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        deck.slides[1].headers_footers.set_footer(text="x")
        hf = deck.slides[1].headers_footers.set_footer(visible=False)
    assert hf["footer"]["visible"] is False
    assert hf["footer"]["text"] is None  # hidden again -> guarded


def test_set_slide_number(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        hf = deck.slides[1].headers_footers.set_slide_number(True)
    assert hf["slide_number"]["visible"] is True


def test_set_date_fixed_text(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        hf = deck.slides[1].headers_footers.set_date(text="June 2026")
    assert hf["date"]["visible"] is True
    assert hf["date"]["text"] == "June 2026"
    assert hf["date"]["use_format"] is False


def test_set_date_auto_format(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        hf = deck.slides[1].headers_footers.set_date(fmt=14)
    assert hf["date"]["visible"] is True
    assert hf["date"]["format"] == 14
    assert hf["date"]["use_format"] is True


def test_set_date_text_and_format_conflict(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="not both"):
        deck.slides[1].headers_footers.set_date(text="x", fmt=14)


# -- library: master scope --------------------------------------------------


def test_master_headers_footers_read_and_set(deck) -> None:  # type: ignore[no-untyped-def]
    hf = deck.master.headers_footers.read()
    assert hf["footer"]["visible"] is False
    assert hf["display_on_title_slide"] is True  # master scope exposes it
    with deck.edit("hf"):
        out = deck.master.headers_footers.set_footer(text="ACME Corp")
    assert out["footer"]["text"] == "ACME Corp"


# -- batch ops --------------------------------------------------------------


def test_batch_set_headers_footers_slide(deck, ppt) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        out = _edit_core(
            deck,
            EditOp.SET_HEADERS_FOOTERS,
            {"slide": 1, "footer_text": "Draft", "slide_number_visible": True},
        )
    assert out["scope"] == "slide"
    assert out["headers_footers"]["footer"]["text"] == "Draft"
    assert out["headers_footers"]["slide_number"]["visible"] is True
    read = _read_core(ppt, ReadOp.HEADERS_FOOTERS, {"slide": 1})
    assert read["scope"] == "slide"
    assert read["headers_footers"]["footer"]["text"] == "Draft"


def test_batch_set_headers_footers_master(deck, ppt) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"):
        out = _edit_core(deck, EditOp.SET_HEADERS_FOOTERS, {"footer_text": "Master FT"})
    assert out["scope"] == "master"
    read = _read_core(ppt, ReadOp.HEADERS_FOOTERS, {})
    assert read["scope"] == "master"
    assert read["headers_footers"]["footer"]["text"] == "Master FT"


def test_batch_set_headers_footers_needs_a_field(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("hf"), pytest.raises(BatchOpError, match="at least one"):
        _edit_core(deck, EditOp.SET_HEADERS_FOOTERS, {"slide": 1})


# -- CLI --------------------------------------------------------------------


def test_cli_slide_set_footer_and_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    res = runner.invoke(main, ["--json", "slide", "set-footer", "--slide", "1", "--text", "Hi"])
    assert res.exit_code == 0, res.output
    assert json.loads(res.output)["headers_footers"]["footer"]["text"] == "Hi"
    read = runner.invoke(main, ["--json", "slide", "headers-footers", "1"])
    assert json.loads(read.output)["footer"]["text"] == "Hi"


def test_cli_slide_set_footer_no_args_usage_error(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["slide", "set-footer", "--slide", "1"])
    assert res.exit_code == 2  # click UsageError


def test_cli_master_set_footer(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "master", "set-footer", "--text", "Co"])
    assert res.exit_code == 0, res.output
    assert json.loads(res.output)["headers_footers"]["footer"]["text"] == "Co"


def test_cli_slide_number_toggle(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "slide", "slide-number", "--slide", "1", "--show"])
    assert res.exit_code == 0, res.output
    assert json.loads(res.output)["headers_footers"]["slide_number"]["visible"] is True

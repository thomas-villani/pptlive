"""Per-slide background (v0.4.0, the v1.2 cut) — solid override + revert.

The per-slide counterpart to v0.9's deck-wide master background; mirrors
`Master.set_background`. Verified against the fake deck.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, _edit_core
from pptlive.cli.main import main

# -- library: override + read -----------------------------------------------


def test_slide_inherits_master_by_default(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[1].background()["follows_master"] is True


def test_set_background_solid_color(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("t"):
        out = slide.set_background("#1A2B3C")
    assert out["follows_master"] is False
    assert out["type"] == "solid"
    assert out["color"] == "#1A2B3C"


def test_set_background_accepts_rgb_tuple(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("t"):
        out = slide.set_background((255, 0, 0))
    assert out["color"] == "#FF0000"


def test_follow_master_background_reverts(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("t"):
        slide.set_background("#1A2B3C")
    with deck.edit("t"):
        out = slide.follow_master_background()
    assert out["follows_master"] is True


def test_set_background_bad_color_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError):
        deck.slides[1].set_background("not-a-color")


def test_slide_read_includes_background(deck) -> None:  # type: ignore[no-untyped-def]
    grid = deck.slides[1].read()
    assert grid["background"]["follows_master"] is True


# -- batch op ---------------------------------------------------------------


def test_batch_slide_set_background_color(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"):
        out = _edit_core(deck, EditOp.SLIDE_SET_BACKGROUND, {"slide": 1, "color": "#102030"})
    assert out["background"]["color"] == "#102030"
    assert out["background"]["follows_master"] is False


def test_batch_slide_set_background_follow_master(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"):
        _edit_core(deck, EditOp.SLIDE_SET_BACKGROUND, {"slide": 1, "color": "#102030"})
    with deck.edit("t"):
        out = _edit_core(deck, EditOp.SLIDE_SET_BACKGROUND, {"slide": 1, "follow_master": True})
    assert out["background"]["follows_master"] is True


def test_batch_slide_set_background_needs_exactly_one(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"), pytest.raises(BatchOpError, match="exactly one"):
        _edit_core(deck, EditOp.SLIDE_SET_BACKGROUND, {"slide": 1})


# -- CLI --------------------------------------------------------------------


def test_cli_slide_set_background_color(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["slide", "set-background", "--slide", "1", "--color", "#1A2B3C"]
    )
    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["background"]["color"] == "#1A2B3C"
    assert payload["background"]["follows_master"] is False


def test_cli_slide_set_background_follow_master(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    runner.invoke(main, ["slide", "set-background", "--slide", "1", "--color", "#1A2B3C"])
    result = runner.invoke(main, ["slide", "set-background", "--slide", "1", "--follow-master"])
    assert result.exit_code == 0
    assert json.loads(result.output)["background"]["follows_master"] is True


def test_cli_slide_set_background_no_option_is_usage_error(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "set-background", "--slide", "1"])
    assert result.exit_code == 2  # click UsageError


def test_cli_slide_set_background_both_is_usage_error(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["slide", "set-background", "--slide", "1", "--color", "#fff", "--follow-master"],
    )
    assert result.exit_code == 2  # click UsageError

"""Slide transitions (v0.4.0, the v1.5 cut) — entry effect + duration + advance.

Curated `PpEntryEffect` friendly map + the auto-advance model (both
`AdvanceOnTime` and `AdvanceTime` together), proven against the fake deck.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, _edit_core
from pptlive.cli.main import main
from pptlive.constants import entry_effect_for, entry_effect_name

# -- constants helpers ------------------------------------------------------


def test_entry_effect_for_friendly_and_aliases() -> None:
    assert entry_effect_for("fade") == 1793
    assert entry_effect_for("none") == 0
    assert entry_effect_for("cover_left") == 1281
    # short alias -> default direction; case / separator insensitive
    assert entry_effect_for("cover") == 1281
    assert entry_effect_for("Cover Left") == 1281
    assert entry_effect_for("blinds-horizontal") == 769
    # raw int passes through (the escape hatch for exotic effects)
    assert entry_effect_for(3585) == 3585


def test_entry_effect_for_unknown_raises() -> None:
    with pytest.raises(ValueError, match="unknown transition effect"):
        entry_effect_for("teleport")


def test_entry_effect_name_roundtrip() -> None:
    assert entry_effect_name(1793) == "fade"
    assert entry_effect_name(0) == "none"
    assert entry_effect_name(999999) == "effect:999999"  # unnamed but valid-shaped


# -- library: set + read ----------------------------------------------------


def test_set_transition_effect(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("t"):
        out = slide.set_transition("fade")
    assert out["effect"] == "fade"
    assert deck.slides[1].transition()["effect"] == "fade"


def test_set_transition_duration_and_advance(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("t"):
        out = slide.set_transition("cut", duration=0.5, advance_after=3.0)
    assert out["duration"] == 0.5
    # advance_after sets BOTH the flag and the seconds (the spike finding).
    assert out["advance_on_time"] is True
    assert out["advance_time"] == 3.0


def test_set_transition_advance_on_click_toggle(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("t"):
        out = slide.set_transition(advance_on_click=False)
    assert out["advance_on_click"] is False


def test_set_transition_requires_an_argument(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="at least one"):
        deck.slides[1].set_transition()


def test_set_transition_unknown_effect_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown transition effect"):
        deck.slides[1].set_transition("warp")


def test_slide_read_includes_transition(deck) -> None:  # type: ignore[no-untyped-def]
    grid = deck.slides[1].read()
    assert grid["transition"]["effect"] == "none"  # default


# -- batch op ---------------------------------------------------------------


def test_batch_slide_set_transition(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"):
        out = _edit_core(
            deck, EditOp.SLIDE_SET_TRANSITION, {"slide": 1, "effect": "dissolve", "duration": 1.0}
        )
    assert out["transition"]["effect"] == "dissolve"
    assert out["transition"]["duration"] == 1.0


def test_batch_slide_set_transition_needs_an_option(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"), pytest.raises(BatchOpError, match="at least one"):
        _edit_core(deck, EditOp.SLIDE_SET_TRANSITION, {"slide": 1})


# -- CLI --------------------------------------------------------------------


def test_cli_slide_set_transition(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["slide", "set-transition", "--slide", "1", "--effect", "fade", "--advance-after", "5"],
    )
    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["transition"]["effect"] == "fade"
    assert payload["transition"]["advance_time"] == 5.0


def test_cli_slide_set_transition_no_options_is_usage_error(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "set-transition", "--slide", "1"])
    assert result.exit_code == 2  # click UsageError


def test_cli_slide_set_transition_bad_effect_rejected(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["slide", "set-transition", "--slide", "1", "--effect", "warp"]
    )
    assert result.exit_code == 2  # click rejects the invalid --effect choice


def test_set_transition_rejects_negative_durations(deck) -> None:  # type: ignore[no-untyped-def]
    # Durations are seconds; a negative is a clean ValueError before any COM.
    # advance_after=0 stays valid (keep the timing, require a click).
    with deck.edit("t"), pytest.raises(ValueError, match="duration"):
        deck.slides[1].set_transition("fade", duration=-1.0)
    with deck.edit("t"), pytest.raises(ValueError, match="advance_after"):
        deck.slides[1].set_transition("fade", advance_after=-2.0)

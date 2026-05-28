"""Live slide-show control (v0.6): `deck.show` + the CLI `show` group.

The fake presentation models `SlideShowSettings.Run()` -> a `SlideShowWindow`
whose `View` advances/jumps and carries a read/write `State`. Accessing
`SlideShowWindow` when nothing is running raises (as real COM does), so the
wrapper's "not running" detection and `SlideShowNotRunningError` are exercised
end to end.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.exceptions import SlideNotFoundError, SlideShowNotRunningError


def _json(result) -> object:  # type: ignore[no-untyped-def]
    return json.loads(result.output)


# ---------------------------------------------------------------------------
# Wrapper: state + lifecycle
# ---------------------------------------------------------------------------


def test_state_when_not_running(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.show.state()
    assert info == {
        "running": False,
        "state": "done",
        "state_code": 5,
        "current_slide": None,
        "position": None,
        "slide_count": 3,
    }
    assert deck.show.is_running() is False


def test_start_runs_from_top(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.show.start()
    assert info["running"] is True
    assert info["state"] == "running"
    assert info["current_slide"] == 1
    assert info["position"] == 1
    assert deck.show.is_running() is True


def test_start_from_slide(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.show.start(from_slide=3)
    assert info["current_slide"] == 3


def test_start_from_out_of_range_slide_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.show.start(from_slide=99)
    assert deck.show.is_running() is False  # never started


def test_start_idempotent_when_already_running(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start()
    deck.show.next()  # now on slide 2
    again = deck.show.start()  # no-op restart keeps position
    assert again["current_slide"] == 2
    # but start(from_slide=...) jumps an already-running show
    assert deck.show.start(from_slide=1)["current_slide"] == 1


def test_end_is_noop_when_not_running(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.show.end()
    assert info["running"] is False


def test_end_stops_a_running_show(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start()
    info = deck.show.end()
    assert info["running"] is False
    assert deck.show.is_running() is False


# ---------------------------------------------------------------------------
# Wrapper: navigation
# ---------------------------------------------------------------------------


def test_next_and_previous(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start()
    assert deck.show.next()["current_slide"] == 2
    assert deck.show.next()["current_slide"] == 3
    assert deck.show.previous()["current_slide"] == 2


def test_next_past_last_ends_show(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start(from_slide=3)
    info = deck.show.next()  # advancing past the final slide ends the show
    assert info["running"] is False


def test_goto_jumps(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start()
    assert deck.show.goto(3)["current_slide"] == 3


def test_goto_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start()
    with pytest.raises(SlideNotFoundError):
        deck.show.goto(0)


# ---------------------------------------------------------------------------
# Wrapper: blank-screen states
# ---------------------------------------------------------------------------


def test_black_white_resume(deck) -> None:  # type: ignore[no-untyped-def]
    deck.show.start()
    assert deck.show.black()["state"] == "black"
    assert deck.show.white()["state"] == "white"
    assert deck.show.resume()["state"] == "running"


# ---------------------------------------------------------------------------
# Wrapper: control verbs require a running show
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("verb", ["next", "previous", "black", "white", "resume"])
def test_control_verbs_require_running_show(deck, verb) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideShowNotRunningError):
        getattr(deck.show, verb)()


def test_goto_requires_running_show(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideShowNotRunningError):
        deck.show.goto(2)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def test_cli_show_state(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["show", "state"])
    assert result.exit_code == 0
    assert _json(result)["running"] is False


def test_cli_show_start_next_goto_end(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    assert runner.invoke(main, ["show", "start"]).exit_code == 0
    out = _json(runner.invoke(main, ["show", "next"]))
    assert out["current_slide"] == 2
    out = _json(runner.invoke(main, ["show", "goto", "--slide", "3"]))
    assert out["current_slide"] == 3
    assert _json(runner.invoke(main, ["show", "end"]))["running"] is False


def test_cli_show_start_from(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    out = _json(CliRunner().invoke(main, ["show", "start", "--from", "2"]))
    assert out["current_slide"] == 2


def test_cli_show_next_not_running_exit_1(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["show", "next"])
    assert result.exit_code == 1
    assert "no slide show is running" in result.output


def test_cli_show_goto_out_of_range_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    CliRunner().invoke(main, ["show", "start"])
    result = CliRunner().invoke(main, ["show", "goto", "--slide", "9"])
    assert result.exit_code == 2

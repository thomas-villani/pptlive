"""Shape animations (v0.10, the v1.5-rest cut) — entrance/exit effects.

The curated `MsoAnimEffect` / `MsoAnimTriggerType` friendly maps, `Shape.animate`
(entrance + exit + timing), per-shape and whole-slide `clear_animations`, the
`slide.animations()` read, and the batch / CLI / MCP surfaces — proven against the
fake deck (its `TimeLine.MainSequence` mirrors the live spike round-trip).
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, ReadOp, _edit_core, _read_core
from pptlive.cli.main import main
from pptlive.constants import (
    anim_effect_for,
    anim_effect_name,
    anim_trigger_for,
    anim_trigger_name,
)

# -- constants helpers ------------------------------------------------------


def test_anim_effect_for_friendly_and_aliases() -> None:
    assert anim_effect_for("fade") == 10
    assert anim_effect_for("appear") == 1
    assert anim_effect_for("fly_in") == 2
    # case / separator insensitive
    assert anim_effect_for("Fly In") == 2
    assert anim_effect_for("float-in") == 21
    # raw int passes through (the escape hatch for exotic effects)
    assert anim_effect_for(99) == 99


def test_anim_effect_for_unknown_raises() -> None:
    with pytest.raises(ValueError, match="unknown animation effect"):
        anim_effect_for("teleport")


def test_anim_effect_name_roundtrip() -> None:
    assert anim_effect_name(10) == "fade"
    assert anim_effect_name(1) == "appear"
    assert anim_effect_name(424242) == "effect:424242"  # unnamed but valid-shaped


def test_anim_trigger_for_and_name() -> None:
    assert anim_trigger_for("on_click") == 1
    assert anim_trigger_for("with_previous") == 2
    assert anim_trigger_for("after_previous") == 3
    assert anim_trigger_for(7) == 7
    assert anim_trigger_name(3) == "after_previous"
    with pytest.raises(ValueError, match="unknown animation trigger"):
        anim_trigger_for("someday")


# -- library: animate + read ------------------------------------------------


def _first_shape(deck):  # type: ignore[no-untyped-def]
    """The default deck's slide 1 has placeholders; return the first as a Shape."""
    return deck.slides[1].shapes[1]


def test_animate_basic(deck) -> None:  # type: ignore[no-untyped-def]
    sh = _first_shape(deck)
    with deck.edit("a"):
        out = sh.animate("fade")
    assert out["effect"] == "fade"
    assert out["exit"] is False
    assert out["trigger"] == "on_click"
    assert out["shapeid"] == sh.shapeid
    # it shows up on the slide's animation read
    anims = deck.slides[1].animations()
    assert len(anims) == 1
    assert anims[0]["seq_index"] == 1
    assert anims[0]["effect"] == "fade"


def test_animate_exit_with_timing(deck) -> None:  # type: ignore[no-untyped-def]
    sh = _first_shape(deck)
    with deck.edit("a"):
        out = sh.animate("fly_in", trigger="after_previous", duration=1.5, delay=0.75, exit=True)
    assert out["exit"] is True
    assert out["effect"] == "fly_in"
    assert out["trigger"] == "after_previous"
    assert out["duration"] == 1.5
    assert out["delay"] == 0.75


def test_animate_unknown_effect_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown animation effect"):
        deck.slides[1].shapes[1].animate("warpspeed")


def test_animate_multiple_then_clear_shape(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    s1, s2 = slide.shapes[1], slide.shapes[2]
    with deck.edit("a"):
        s1.animate("fade")
        s1.animate("zoom", exit=True)
        s2.animate("appear")
    assert len(slide.animations()) == 3
    # clearing s1 leaves s2's single effect
    with deck.edit("c"):
        removed = s1.clear_animations()
    assert removed == 2
    remaining = slide.animations()
    assert len(remaining) == 1
    assert remaining[0]["shapeid"] == s2.shapeid


def test_clear_animations_whole_slide(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    with deck.edit("a"):
        slide.shapes[1].animate("fade")
        slide.shapes[2].animate("appear")
    with deck.edit("c"):
        removed = slide.clear_animations()
    assert removed == 2
    assert slide.animations() == []


def test_slide_read_includes_animations(deck) -> None:  # type: ignore[no-untyped-def]
    grid = deck.slides[1].read()
    assert grid["animations"] == []  # default: none
    with deck.edit("a"):
        deck.slides[1].shapes[1].animate("wipe")
    assert deck.slides[1].read()["animations"][0]["effect"] == "wipe"


# -- batch ops --------------------------------------------------------------


def test_batch_shape_animate(deck) -> None:  # type: ignore[no-untyped-def]
    anchor = deck.slides[1].shapes[1].anchor_id
    with deck.edit("a"):
        out = _edit_core(
            deck, EditOp.SHAPE_ANIMATE, {"anchor_id": anchor, "effect": "zoom", "exit": True}
        )
    assert out["ok"] is True
    assert out["animation"]["effect"] == "zoom"
    assert out["animation"]["exit"] is True


def test_batch_shape_animate_requires_effect(deck) -> None:  # type: ignore[no-untyped-def]
    anchor = deck.slides[1].shapes[1].anchor_id
    with deck.edit("a"), pytest.raises(BatchOpError, match="requires `effect`"):
        _edit_core(deck, EditOp.SHAPE_ANIMATE, {"anchor_id": anchor})


def test_batch_clear_and_read(deck, ppt) -> None:  # type: ignore[no-untyped-def]
    anchor = deck.slides[1].shapes[1].anchor_id
    with deck.edit("a"):
        _edit_core(deck, EditOp.SHAPE_ANIMATE, {"anchor_id": anchor, "effect": "fade"})
    read = _read_core(ppt, ReadOp.ANIMATIONS, {"slide": 1})
    assert read["slide"] == 1
    assert len(read["animations"]) == 1
    with deck.edit("c"):
        out = _edit_core(deck, EditOp.SLIDE_CLEAR_ANIMATIONS, {"slide": 1})
    assert out["removed"] == 1


# -- CLI --------------------------------------------------------------------


def test_cli_shape_animate(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["shape", "animate", "--anchor-id", "shape:1:1", "--effect", "fade", "--duration", "1.0"],
    )
    assert result.exit_code == 0, result.output
    payload = json.loads(result.output)
    assert payload["animation"]["effect"] == "fade"
    assert payload["animation"]["duration"] == 1.0


def test_cli_shape_animate_exit_flag(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "animate", "--anchor-id", "shape:1:1", "--effect", "zoom", "--exit"]
    )
    assert result.exit_code == 0, result.output
    assert json.loads(result.output)["animation"]["exit"] is True


def test_cli_shape_animate_bad_effect_rejected(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "animate", "--anchor-id", "shape:1:1", "--effect", "warp"]
    )
    assert result.exit_code == 2  # click rejects the invalid --effect choice


def test_cli_slide_animations_and_clear(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    runner.invoke(main, ["shape", "animate", "--anchor-id", "shape:1:1", "--effect", "fade"])
    listing = runner.invoke(main, ["--json", "slide", "animations", "1"])
    assert listing.exit_code == 0, listing.output
    rows = json.loads(listing.output)
    assert len(rows) == 1 and rows[0]["effect"] == "fade"
    cleared = runner.invoke(main, ["slide", "clear-animations", "--slide", "1"])
    assert cleared.exit_code == 0, cleared.output
    assert json.loads(cleared.output)["removed"] == 1


def test_cli_shape_clear_animations(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    runner.invoke(main, ["shape", "animate", "--anchor-id", "shape:1:1", "--effect", "fade"])
    result = runner.invoke(main, ["shape", "clear-animations", "--anchor-id", "shape:1:1"])
    assert result.exit_code == 0, result.output
    assert json.loads(result.output)["removed"] == 1

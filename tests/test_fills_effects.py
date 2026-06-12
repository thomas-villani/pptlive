"""Advanced fills (gradient / picture / pattern) and shape effects (shadow /
glow / soft-edge / reflection) — the v1.2 styling-completion cut.

Spiked live in `scripts/fill_advanced_spike.py` / `scripts/effects_spike.py`; the
fake `_FakeShapeFill` + effect namespaces in conftest reproduce the verified COM
behaviour (legacy `Insert` stops read back sorted, `Type` discriminators, the
shadow/glow/soft-edge/reflection round-trips).
"""

from __future__ import annotations

import os
import tempfile

import pytest

from pptlive.constants import (
    gradient_style_for,
    gradient_style_name,
    pattern_for,
    pattern_name,
    preset_gradient_for,
)

# -- constants helpers ------------------------------------------------------


def test_gradient_style_for_and_name() -> None:
    assert gradient_style_for("horizontal") == 1
    assert gradient_style_for("Diagonal Up") == 3  # case/separator insensitive
    assert gradient_style_for(7) == 7  # raw int passthrough
    assert gradient_style_name(1) == "horizontal"
    assert gradient_style_name(999) == 999  # unknown int falls through


def test_gradient_style_for_unknown_raises() -> None:
    with pytest.raises(ValueError, match="unknown gradient style"):
        gradient_style_for("sideways")


def test_preset_gradient_for() -> None:
    assert preset_gradient_for("ocean") == 7
    assert preset_gradient_for("rainbow_ii") == 17
    assert preset_gradient_for(20) == 20
    with pytest.raises(ValueError, match="unknown preset gradient"):
        preset_gradient_for("lava")


def test_pattern_for_and_name() -> None:
    assert pattern_for("percent_50") == 7
    assert pattern_for("Dark Horizontal") == 13
    assert pattern_for(33) == 33  # raw int passthrough
    assert pattern_name(7) == "percent_50"
    with pytest.raises(ValueError, match="unknown pattern"):
        pattern_for("zigzag")


# -- gradient fill ----------------------------------------------------------


def test_two_color_gradient(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_gradient_fill(["#FF0000", "#0000FF"], style="vertical")
    fill = shape.to_dict()["fill"]
    assert fill["type"] == "gradient"
    assert fill["gradient_style"] == "vertical"
    assert fill["stops"] == [
        {"color": "#FF0000", "position": 0.0},
        {"color": "#0000FF", "position": 1.0},
    ]


def test_multi_stop_gradient_reads_back_sorted(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_gradient_fill(["#FF0000", "#00FF00", "#0000FF"], positions=[0.0, 0.25, 1.0])
    stops = shape.to_dict()["fill"]["stops"]
    # endpoints at 0/1 (Fore/Back), interior inserted at 0.25 — sorted by position
    assert [s["position"] for s in stops] == [0.0, 0.25, 1.0]
    assert [s["color"] for s in stops] == ["#FF0000", "#00FF00", "#0000FF"]


def test_one_color_gradient(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_gradient_fill(["#123456"], degree=0.3)
    assert shape.to_dict()["fill"]["type"] == "gradient"
    assert shape.com.Fill.GradientColorType == 1  # one-color
    assert shape.com.Fill.GradientDegree == 0.3


def test_preset_gradient(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_gradient_fill(preset="ocean")
    assert shape.to_dict()["fill"]["type"] == "gradient"
    assert shape.com.Fill.GradientColorType == 3  # preset


def test_gradient_requires_colors_or_preset(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="colors= or preset="):
        deck.slides[2].shapes[1].set_gradient_fill()


def test_gradient_bad_color_raises_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with pytest.raises(ValueError):
        shape.set_gradient_fill(["#FF0000", "not-a-color"])
    assert shape.com.Fill.Type != 3  # nothing applied


def test_gradient_position_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="between 0.0 and 1.0"):
        deck.slides[2].shapes[1].set_gradient_fill(["#FF0000", "#00FF00", "#0000FF"], positions=[0, 2, 1])


# -- picture fill -----------------------------------------------------------


def test_picture_fill(deck) -> None:  # type: ignore[no-untyped-def]
    fd, path = tempfile.mkstemp(suffix=".png")
    os.close(fd)
    try:
        shape = deck.slides[2].shapes[1]
        shape.set_picture_fill(path)
        assert shape.to_dict()["fill"]["type"] == "picture"
        assert shape.com.Fill.picture_path == os.path.abspath(path)
    finally:
        os.remove(path)


def test_picture_fill_missing_file_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(FileNotFoundError):
        deck.slides[2].shapes[1].set_picture_fill("does-not-exist-12345.png")


# -- pattern fill -----------------------------------------------------------


def test_pattern_fill(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_pattern_fill("percent_50", fore="#FF0000", back="#FFFFFF")
    fill = shape.to_dict()["fill"]
    assert fill["type"] == "patterned"
    assert fill["pattern"] == "percent_50"
    assert fill["color"] == "#FF0000"
    assert fill["back_color"] == "#FFFFFF"


def test_pattern_fill_bad_pattern_raises_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with pytest.raises(ValueError, match="unknown pattern"):
        shape.set_pattern_fill("zigzag", fore="#FF0000")
    assert shape.com.Fill.Type != 2


# -- effects ----------------------------------------------------------------


def test_set_shadow_round_trips(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_effect(shadow={"color": "#FF0000", "blur": 8, "transparency": 0.4, "offset_x": 4})
    effects = shape.to_dict()["effects"]
    assert effects["shadow"]["color"] == "#FF0000"
    assert effects["shadow"]["blur"] == 8.0
    assert effects["shadow"]["transparency"] == 0.4
    assert effects["shadow"]["offset_x"] == 4.0


def test_set_shadow_none_disables(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_effect(shadow={"color": "#000000"})
    assert shape.to_dict()["effects"] is not None
    shape.set_effect(shadow="none")
    assert shape.to_dict()["effects"] is None


def test_set_glow_round_trips(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_effect(glow={"color": "#00FF00", "radius": 10})
    glow = shape.to_dict()["effects"]["glow"]
    assert glow["color"] == "#00FF00"
    assert glow["radius"] == 10.0


def test_set_soft_edge_and_reflection(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    shape.set_effect(soft_edge=4, reflection=5)
    effects = shape.to_dict()["effects"]
    assert effects["soft_edge"]["type"] == 4
    assert effects["reflection"]["type"] == 5


def test_set_effect_requires_an_arg(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="at least one"):
        deck.slides[2].shapes[1].set_effect()


def test_set_effect_bad_color_raises_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with pytest.raises(ValueError):
        shape.set_effect(shadow={"color": "not-a-color"})
    # the shape stays shadow-less (Visible never flipped on)
    assert shape.to_dict()["effects"] is None


def test_plain_shape_has_no_effects(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[2].shapes[1].to_dict()["effects"] is None

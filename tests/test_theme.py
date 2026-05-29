"""Master / theme styling (v0.9): deck.theme + deck.master, CLI and wrappers.

Against the fake, the slide master carries a theme (12-slot palette + major/minor
fonts) and three text styles (title/body/default, 5 levels each) plus a background
fill. Writes round-trip through the same plain objects, so the politeness-free
deck-wide setters are proven here without PowerPoint.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.constants import (
    text_style_for,
    theme_color_for,
    theme_font_script_for,
    theme_font_slot_for,
)


def _json(result):  # type: ignore[no-untyped-def]
    return json.loads(result.output)


# -- constants / resolvers --------------------------------------------------


def test_text_style_for_and_unknown() -> None:
    assert text_style_for("Body") == 3
    assert text_style_for("title") == 2
    assert text_style_for("default") == 1
    with pytest.raises(ValueError, match="unknown text style"):
        text_style_for("caption")


def test_theme_color_for_aliases_and_unknown() -> None:
    assert theme_color_for("accent1") == 5
    assert theme_color_for("Accent 6") == 10
    assert theme_color_for("hlink") == 11  # alias
    assert theme_color_for("followed_hyperlink") == 12
    with pytest.raises(ValueError, match="unknown theme color slot"):
        theme_color_for("accent7")


def test_theme_font_resolvers() -> None:
    assert theme_font_slot_for("heading") == "major"
    assert theme_font_slot_for("body") == "minor"
    assert theme_font_script_for("latin") == 1
    assert theme_font_script_for("complex-script") == 3
    with pytest.raises(ValueError, match="unknown theme font"):
        theme_font_slot_for("caption")
    with pytest.raises(ValueError, match="unknown font script"):
        theme_font_script_for("klingon")


# -- Theme wrapper ----------------------------------------------------------


def test_theme_read_shape(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.theme.read()
    assert set(info) == {"colors", "fonts"}
    # 12 named slots in palette order
    assert list(info["colors"]) == [
        "dark1",
        "light1",
        "dark2",
        "light2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hyperlink",
        "followed_hyperlink",
    ]
    assert all(v.startswith("#") for v in info["colors"].values())
    assert set(info["fonts"]) == {"major", "minor"}


def test_theme_set_color_round_trip(deck) -> None:  # type: ignore[no-untyped-def]
    deck.theme.set_color("accent1", "#FF0000")
    assert deck.theme.read()["colors"]["accent1"] == "#FF0000"


def test_theme_set_color_accepts_tuple(deck) -> None:  # type: ignore[no-untyped-def]
    deck.theme.set_color("dark1", (0, 128, 255))
    assert deck.theme.read()["colors"]["dark1"] == "#0080FF"


def test_theme_set_color_unknown_slot_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown theme color slot"):
        deck.theme.set_color("accent9", "#000000")


def test_theme_set_font_round_trip(deck) -> None:  # type: ignore[no-untyped-def]
    deck.theme.set_font("major", "Georgia")
    deck.theme.set_font("minor", "Verdana")
    fonts = deck.theme.read()["fonts"]
    assert fonts["major"] == "Georgia"
    assert fonts["minor"] == "Verdana"


def test_theme_set_font_unknown_which_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown theme font"):
        deck.theme.set_font("caption", "Georgia")


# -- Master wrapper ---------------------------------------------------------


def test_master_read_shape(deck) -> None:  # type: ignore[no-untyped-def]
    info = deck.master.read()
    assert set(info["text_styles"]) == {"title", "body", "default"}
    body_levels = info["text_styles"]["body"]["levels"]
    assert len(body_levels) == 5
    assert body_levels[0]["level"] == 1
    assert {"font", "size", "bold", "italic", "underline", "color", "alignment"} <= set(
        body_levels[0]
    )
    assert set(info["background"]) == {"type", "color"}


def test_master_format_text_style_round_trip(deck) -> None:  # type: ignore[no-untyped-def]
    deck.master.format_text_style("body", 1, font="Georgia", size=32, bold=True, color="#112233")
    lvl = deck.master.read()["text_styles"]["body"]["levels"][0]
    assert lvl["font"] == "Georgia"
    assert lvl["size"] == 32.0
    assert lvl["bold"] is True
    assert lvl["color"] == "#112233"


def test_master_format_text_style_only_touches_its_level(deck) -> None:  # type: ignore[no-untyped-def]
    deck.master.format_text_style("title", 2, size=44)
    levels = deck.master.read()["text_styles"]["title"]["levels"]
    assert levels[1]["size"] == 44.0
    assert levels[0]["size"] != 44.0  # level 1 untouched


def test_master_format_text_style_bad_level_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="between 1 and 5"):
        deck.master.format_text_style("body", 6, size=10)


def test_master_format_text_style_unknown_style_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown text style"):
        deck.master.format_text_style("caption", 1, size=10)


def test_master_format_paragraph_style_round_trip(deck) -> None:  # type: ignore[no-untyped-def]
    deck.master.format_paragraph_style("title", 1, alignment="center")
    assert deck.master.read()["text_styles"]["title"]["levels"][0]["alignment"] == 2


def test_master_set_background_round_trip(deck) -> None:  # type: ignore[no-untyped-def]
    deck.master.set_background("#1F1F1F")
    bg = deck.master.read()["background"]
    assert bg["type"] == "solid"
    assert bg["color"] == "#1F1F1F"


def test_master_read_is_side_effect_free(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.master.read() == deck.master.read()


# -- CLI --------------------------------------------------------------------


def test_cli_theme_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["theme", "read"])
    assert result.exit_code == 0, result.output
    info = _json(result)
    assert "colors" in info and "fonts" in info


def test_cli_theme_set_color(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["theme", "set-color", "--slot", "accent1", "--color", "#C00000"]
    )
    assert result.exit_code == 0, result.output
    assert _json(result)["colors"]["accent1"] == "#C00000"


def test_cli_theme_set_color_rejects_bad_slot(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["theme", "set-color", "--slot", "accent9", "--color", "#000000"]
    )
    assert result.exit_code != 0  # click.Choice rejects it


def test_cli_theme_set_font(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["theme", "set-font", "--which", "major", "--name", "Georgia"]
    )
    assert result.exit_code == 0, result.output
    assert _json(result)["fonts"]["major"] == "Georgia"


def test_cli_master_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["master", "read"])
    assert result.exit_code == 0, result.output
    info = _json(result)
    assert set(info["text_styles"]) == {"title", "body", "default"}


def test_cli_master_format_text_style(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["master", "format-text-style", "--style", "body", "--level", "1", "--font", "Georgia"],
    )
    assert result.exit_code == 0, result.output
    assert _json(result)["text_styles"]["body"]["levels"][0]["font"] == "Georgia"


def test_cli_master_format_text_style_requires_an_option(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["master", "format-text-style", "--style", "body", "--level", "1"]
    )
    assert result.exit_code != 0
    assert "at least one" in result.output


def test_cli_master_set_background(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["master", "set-background", "--color", "#202020"])
    assert result.exit_code == 0, result.output
    assert _json(result)["background"]["color"] == "#202020"

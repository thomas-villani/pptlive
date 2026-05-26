"""CLI behaviour against the fake app: payloads, mutations, and exit codes."""

from __future__ import annotations

import json

from click.testing import CliRunner

from pptlive.cli.main import main


def _json(result) -> object:  # type: ignore[no-untyped-def]
    return json.loads(result.output)


def test_status_lists_decks_and_viewed_slide(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    result = CliRunner().invoke(main, ["status"])
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["viewed_slide"] == 2
    assert payload["decks"][0]["name"] == "Pitch.pptx"
    assert payload["decks"][0]["is_active"] is True


def test_status_not_running_exit_4(no_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["status"])
    assert result.exit_code == 4


def test_slides(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slides"])
    assert result.exit_code == 0
    rows = _json(result)
    assert [r["index"] for r in rows] == [1, 2, 3]


def test_outline(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["outline"])
    assert result.exit_code == 0
    items = _json(result)
    assert items[1]["bullets"] == ["Intro", "Demo", "Q&A"]


def test_slide_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "read", "2"])
    assert result.exit_code == 0
    grid = _json(result)
    assert grid["title"] == "Agenda"
    assert len(grid["shapes"]) == 3


def test_shapes(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shapes", "--slide", "2"])
    assert result.exit_code == 0
    rows = _json(result)
    assert rows[0]["anchor_id"] == "shape:2:1"


def test_read_anchor(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["read", "anchor", "--anchor-id", "ph:2:body"])
    assert result.exit_code == 0
    assert _json(result)["text"] == "Intro\rDemo\rQ&A"


def test_read_notes(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["read", "notes", "--slide", "1"])
    assert result.exit_code == 0
    assert _json(result)["text"] == "Lead with the vision."


def test_write_sets_text(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["write", "--anchor-id", "ph:2:title", "--text", "Q3 Results"]
    )
    assert result.exit_code == 0
    assert _json(result)["ok"] is True
    title = fake_powerpoint.ActivePresentation.Slides(2).Shapes(1)
    assert title.TextFrame.TextRange.Text == "Q3 Results"


def test_replace_sets_text(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["replace", "--anchor-id", "shape:3:1", "--text", "Replaced"])
    assert result.exit_code == 0
    box = fake_powerpoint.ActivePresentation.Slides(3).Shapes(1)
    assert box.TextFrame.TextRange.Text == "Replaced"


def test_write_to_frameless_shape_exit_6(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["write", "--anchor-id", "shape:2:3", "--text", "x"])
    assert result.exit_code == 6


def test_write_missing_anchor_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["write", "--anchor-id", "ph:2:banner", "--text", "x"])
    assert result.exit_code == 2


def test_read_missing_slide_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["read", "anchor", "--anchor-id", "shape:9:1"])
    assert result.exit_code == 2


def test_go_to(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    result = CliRunner().invoke(main, ["go-to", "--anchor-id", "shape:3:1"])
    assert result.exit_code == 0
    assert fake_powerpoint._viewed == 3


def test_text_output_mode(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["--text", "slides"])
    assert result.exit_code == 0
    assert "Welcome" in result.output
    # Not JSON in text mode.
    assert not result.output.lstrip().startswith("[")

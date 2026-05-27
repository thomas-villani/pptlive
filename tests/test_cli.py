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


def test_slide_layouts(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "layouts"])
    assert result.exit_code == 0
    rows = _json(result)
    assert rows[0]["name"] == "Title Slide"


def test_slide_add(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "add", "--layout", "two_content"])
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["ok"] is True
    assert payload["layout"] == "Two Content"
    assert fake_powerpoint.ActivePresentation.Slides.Count == 4


def test_slide_add_unknown_layout_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "add", "--layout", "bogus"])
    assert result.exit_code == 2
    # Deck untouched on a bad layout.
    assert fake_powerpoint.ActivePresentation.Slides.Count == 3


def test_slide_delete(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "delete", "--slide", "2"])
    assert result.exit_code == 0
    assert _json(result)["deleted"] == 2
    assert fake_powerpoint.ActivePresentation.Slides.Count == 2


def test_slide_delete_out_of_range_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "delete", "--slide", "9"])
    assert result.exit_code == 2


def test_slide_duplicate(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "duplicate", "--slide", "1"])
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["from"] == 1
    assert payload["index"] == 2
    assert fake_powerpoint.ActivePresentation.Slides.Count == 4


def test_slide_move(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "move", "--slide", "1", "--to", "3"])
    assert result.exit_code == 0
    assert _json(result)["index"] == 3
    assert fake_powerpoint.ActivePresentation.Slides(3).SlideID == 256


def test_slide_set_layout(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["slide", "set-layout", "--slide", "3", "--layout", "comparison"]
    )
    assert result.exit_code == 0
    assert _json(result)["layout"] == "Comparison"


def test_slide_lifecycle_fences_one_undo_entry(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # A mutating slide command runs through deck.edit() -> one StartNewUndoEntry.
    assert fake_powerpoint._undo_entries == 0
    result = CliRunner().invoke(main, ["slide", "add", "--layout", "blank"])
    assert result.exit_code == 0
    assert fake_powerpoint._undo_entries == 1


def test_shape_add_textbox(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "textbox", "--text", "Hi"]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["ok"] is True
    assert payload["type"] == "textbox"
    assert payload["text"] == "Hi"
    assert fake_powerpoint.ActivePresentation.Slides(3).Shapes.Count == 3  # was 2


def test_shape_add_autoshape(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "shape", "--shape-type", "oval"]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["type"] == "auto_shape"
    new = fake_powerpoint.ActivePresentation.Slides(3).Shapes(3)
    assert new.AutoShapeType == 9  # oval


def test_shape_add_picture(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    img = tmp_path / "pic.png"
    img.write_bytes(b"\x89PNG\r\n\x1a\n")
    result = CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "picture", "--path", str(img)]
    )
    assert result.exit_code == 0
    assert _json(result)["type"] == "picture"


def test_shape_add_picture_requires_path_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "add", "--slide", "3", "--kind", "picture"])
    assert result.exit_code == 2  # click UsageError
    assert fake_powerpoint.ActivePresentation.Slides(3).Shapes.Count == 2  # untouched


def test_shape_add_bad_slide_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "add", "--slide", "9", "--kind", "textbox"])
    assert result.exit_code == 2


def test_shape_move(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "move", "--anchor-id", "shape:2:3", "--left", "150", "--top", "160"]
    )
    assert result.exit_code == 0
    geo = _json(result)["geometry"]
    assert geo["left"] == 150.0 and geo["top"] == 160.0


def test_shape_move_requires_arg_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "move", "--anchor-id", "shape:2:3"])
    assert result.exit_code == 2  # click UsageError


def test_shape_resize(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "resize", "--anchor-id", "shape:2:3", "--width", "320"]
    )
    assert result.exit_code == 0
    assert _json(result)["geometry"]["width"] == 320.0


def test_shape_delete(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "delete", "--anchor-id", "shape:2:3"])
    assert result.exit_code == 0
    assert _json(result)["ok"] is True
    assert fake_powerpoint.ActivePresentation.Slides(2).Shapes.Count == 2  # was 3


def test_shape_delete_non_shape_anchor_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # notes:1 resolves to a Notes anchor, not a Shape.
    result = CliRunner().invoke(main, ["shape", "delete", "--anchor-id", "notes:1"])
    assert result.exit_code == 2


def test_shape_add_fences_one_undo_entry(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    assert fake_powerpoint._undo_entries == 0
    result = CliRunner().invoke(main, ["shape", "add", "--slide", "3", "--kind", "textbox"])
    assert result.exit_code == 0
    assert fake_powerpoint._undo_entries == 1


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

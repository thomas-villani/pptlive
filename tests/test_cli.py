"""CLI behaviour against the fake app: payloads, mutations, and exit codes."""

from __future__ import annotations

import json
import os
from pathlib import Path

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


def test_slide_add_bad_placeholders_json_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["slide", "add", "--layout", "blank", "--placeholders", "{not json}"]
    )
    assert result.exit_code == 2  # click UsageError on malformed JSON
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


def test_shape_add_autoshape_with_text(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # --text on an autoshape must be applied (it used to be silently dropped: only
    # the textbox branch passed it through, while MCP shape_add did set it).
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "shape",
            "--shape-type",
            "rectangle",
            "--text",
            "Hello",
        ],
    )
    assert result.exit_code == 0
    new = fake_powerpoint.ActivePresentation.Slides(3).Shapes(3)
    assert new.TextFrame.TextRange.Text == "Hello"


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


def test_shape_add_with_fill(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["shape", "add", "--slide", "3", "--kind", "shape", "--fill", "#FF0000", "--line", "none"],
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["fill"] == {
        "type": "solid",
        "color": "#FF0000",
        "visible": True,
        "transparency": 0.0,
    }
    assert payload["line"]["visible"] is False


def test_shape_fill_command(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["shape", "fill", "--anchor-id", "shape:2:1", "--fill", "#102030", "--line-width", "3"],
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["fill"] == {
        "type": "solid",
        "color": "#102030",
        "visible": True,
        "transparency": 0.0,
    }
    assert payload["line"]["weight"] == 3.0


def test_shape_fill_requires_an_option_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "fill", "--anchor-id", "shape:2:1"])
    assert result.exit_code == 2  # click UsageError


def test_shape_gradient_fill_command(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "gradient-fill",
            "--anchor-id",
            "shape:2:1",
            "--colors",
            "#FF0000,#0000FF",
            "--style",
            "vertical",
        ],
    )
    assert result.exit_code == 0
    fill = _json(result)["fill"]
    assert fill["type"] == "gradient"
    assert fill["gradient_style"] == "vertical"


def test_shape_gradient_fill_preset(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "gradient-fill", "--anchor-id", "shape:2:1", "--preset", "ocean"]
    )
    assert result.exit_code == 0
    assert _json(result)["fill"]["type"] == "gradient"


def test_shape_gradient_fill_requires_colors_or_preset(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "gradient-fill", "--anchor-id", "shape:2:1"])
    assert result.exit_code == 2


def test_shape_pattern_fill_command(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "pattern-fill",
            "--anchor-id",
            "shape:2:1",
            "--pattern",
            "percent_50",
            "--fore",
            "#FF0000",
            "--back",
            "#FFFFFF",
        ],
    )
    assert result.exit_code == 0
    fill = _json(result)["fill"]
    assert fill["type"] == "patterned"
    assert fill["pattern"] == "percent_50"


def test_shape_effect_command(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "effect",
            "--anchor-id",
            "shape:2:1",
            "--shadow",
            '{"color":"#FF0000","blur":8}',
            "--soft-edge",
            "4",
        ],
    )
    assert result.exit_code == 0
    effects = _json(result)["effects"]
    assert effects["shadow"]["color"] == "#FF0000"
    assert effects["soft_edge"]["type"] == 4


def test_shape_effect_none_disables(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    CliRunner().invoke(
        main, ["shape", "effect", "--anchor-id", "shape:2:1", "--glow", '{"color":"#00FF00"}']
    )
    result = CliRunner().invoke(
        main, ["shape", "effect", "--anchor-id", "shape:2:1", "--glow", "none"]
    )
    assert result.exit_code == 0
    assert _json(result)["effects"] is None


def test_shape_effect_requires_an_option(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "effect", "--anchor-id", "shape:2:1"])
    assert result.exit_code == 2


def test_shape_fill_transparency_command(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "fill",
            "--anchor-id",
            "shape:2:1",
            "--fill",
            "#FF0000",
            "--fill-transparency",
            "0.5",
        ],
    )
    assert result.exit_code == 0
    assert _json(result)["fill"]["transparency"] == 0.5


def test_shape_line_style_command(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "line-style",
            "--anchor-id",
            "shape:2:1",
            "--dash",
            "dash_dot",
            "--end-arrow",
            "triangle",
        ],
    )
    assert result.exit_code == 0
    line = _json(result)["line"]
    assert line["dash"] == "dash_dot"
    assert line["end_arrow"] == "triangle"


def test_shape_line_style_requires_an_option(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "line-style", "--anchor-id", "shape:2:1"])
    assert result.exit_code == 2


def test_shape_order_send_to_back(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "order", "--anchor-id", "shape:2:3", "--to", "back"]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["index"] == 1
    assert fake_powerpoint.ActivePresentation.Slides(2).Shapes(1).Name == "Picture 3"


def test_shape_order_bad_choice_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "order", "--anchor-id", "shape:2:1", "--to", "sideways"]
    )
    assert result.exit_code == 2  # click rejects the invalid --to choice


def test_shapeid_anchor_via_cli(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # shapeid:2:3 is Content Placeholder 2 (has a text frame); the read resolves it.
    result = CliRunner().invoke(main, ["read", "anchor", "--anchor-id", "shapeid:2:3"])
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["anchor_id"] == "shapeid:2:3"
    assert payload["text"] == "Intro\rDemo\rQ&A"


def test_slide_geometry_report_via_cli(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "geometry", "1"])
    assert result.exit_code == 0
    rep = _json(result)
    assert rep["slide_size"] == {"width": 960.0, "height": 540.0}
    # Slide 1's two placeholders share the default box, so they overlap.
    assert len(rep["overlaps"]) == 1
    assert rep["shapes"][0]["shapeid"] == "shapeid:1:2"


def test_shape_order_cli_echoes_shapeid(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "order", "--anchor-id", "shape:2:1", "--to", "back"]
    )
    assert result.exit_code == 0
    assert _json(result)["shapeid"] == "shapeid:2:2"


def test_paragraphs_list(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["paragraphs", "--anchor-id", "shape:2:2"])
    assert result.exit_code == 0
    rows = _json(result)
    assert [r["text"] for r in rows] == ["Intro", "Demo", "Q&A"]
    assert rows[1]["anchor_id"] == "para:2:2:2"


def test_paragraphs_non_shape_anchor_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["paragraphs", "--anchor-id", "notes:1"])
    assert result.exit_code == 2


def test_insert_after(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["insert", "--anchor-id", "para:2:2:3", "--text", "Wrap-up"])
    assert result.exit_code == 0
    assert _json(result)["where"] == "after"
    body = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2)
    assert body.TextFrame.TextRange.Text == "Intro\rDemo\rQ&A\rWrap-up"


def test_insert_before(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["insert", "--anchor-id", "para:2:2:1", "--text", "Title", "--before"]
    )
    assert result.exit_code == 0
    body = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2)
    assert body.TextFrame.TextRange.Text == "Title\rIntro\rDemo\rQ&A"


def test_format_paragraph(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "format-paragraph",
            "--anchor-id",
            "para:2:2:1",
            "--alignment",
            "center",
            "--indent-level",
            "2",
        ],
    )
    assert result.exit_code == 0
    para = (
        fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange.Paragraphs(1, 1)
    )
    assert int(para.ParagraphFormat.Alignment) == 2
    assert int(para.IndentLevel) == 2


def test_format_paragraph_no_options_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["format-paragraph", "--anchor-id", "para:2:2:1"])
    assert result.exit_code == 2  # click UsageError


def test_format_paragraph_line_spacing_points(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["format-paragraph", "--anchor-id", "para:2:2:1", "--line-spacing-points", "24"],
    )
    assert result.exit_code == 0
    pf = (
        fake_powerpoint.ActivePresentation.Slides(2)
        .Shapes(2)
        .TextFrame.TextRange.Paragraphs(1, 1)
        .ParagraphFormat
    )
    assert float(pf.SpaceWithin) == 24.0
    assert int(pf.LineRuleWithin) == 0  # points


def test_format_paragraph_line_spacing_guardrail_exit_1(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["format-paragraph", "--anchor-id", "para:2:2:1", "--line-spacing", "24"],
    )
    assert result.exit_code == 1  # rejected: a 24x multiple is almost surely points


def test_format_paragraph_line_spacing_force(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["format-paragraph", "--anchor-id", "para:2:2:1", "--line-spacing", "24", "--force"],
    )
    assert result.exit_code == 0


def test_set_paragraphs_json(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "set-paragraphs",
            "--anchor-id",
            "ph:2:body",
            "--json",
            '["Alpha", {"text": "Beta", "list_type": "numbered"}]',
        ],
    )
    assert result.exit_code == 0
    out = json.loads(result.output)
    assert out["paragraphs"] == ["para:2:2:1", "para:2:2:2"]
    body = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange
    assert body.Text == "Alpha\rBeta"


def test_set_paragraphs_needs_one_source(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["set-paragraphs", "--anchor-id", "ph:2:body"])
    assert result.exit_code == 2  # click UsageError


def test_exec_runs_a_batch_script(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    script = {
        "label": "Build slide 2",
        "ops": [
            {"op": "write", "anchor_id": "ph:2:title", "text": "Q3 Results"},
            {"op": "set_paragraphs", "anchor_id": "ph:2:body", "paragraphs": ["A", "B"]},
            {"op": "format", "anchor_id": "ph:2:title", "bold": True},
        ],
    }
    path = tmp_path / "ops.json"
    path.write_text(json.dumps(script), encoding="utf-8")
    result = CliRunner().invoke(main, ["exec", "--script", str(path)])
    assert result.exit_code == 0
    out = _json(result)
    assert out["ok"] is True and out["count"] == 3
    title = fake_powerpoint.ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange
    assert title.Text == "Q3 Results"
    assert title.Font.Bold != 0
    body = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange
    assert body.Text == "A\rB"


def test_exec_failing_op_maps_exit_code(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    # A missing anchor -> not_found -> exit 2; stop_on_error by default.
    script = {"ops": [{"op": "write", "anchor_id": "ph:2:banner", "text": "x"}]}
    path = tmp_path / "ops.json"
    path.write_text(json.dumps(script), encoding="utf-8")
    result = CliRunner().invoke(main, ["exec", "--script", str(path)])
    assert result.exit_code == 2
    out = _json(result)
    assert out["ok"] is False
    assert out["results"][0]["error"] == "not_found"


def test_exec_bad_script_is_usage_error(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    path = tmp_path / "bad.json"
    path.write_text('{"no_ops": true}', encoding="utf-8")
    result = CliRunner().invoke(main, ["exec", "--script", str(path)])
    assert result.exit_code == 2  # click UsageError (missing "ops" array)


def test_format_text(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "format-text",
            "--anchor-id",
            "ph:2:title",
            "--bold",
            "--size",
            "40",
            "--color",
            "#00FF00",
        ],
    )
    assert result.exit_code == 0
    title = fake_powerpoint.ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange
    assert int(title.Font.Bold) == -1
    assert float(title.Font.Size) == 40.0
    assert int(title.Font.Color.RGB) == 0x00FF00  # green = 65280


def test_format_text_no_options_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["format-text", "--anchor-id", "ph:2:title"])
    assert result.exit_code == 2  # click UsageError


def test_list_apply_and_remove(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    body = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange
    apply = CliRunner().invoke(
        main, ["list", "apply", "--anchor-id", "shape:2:2", "--type", "numbered"]
    )
    assert apply.exit_code == 0
    assert int(body.ParagraphFormat.Bullet.Visible) == -1
    assert int(body.ParagraphFormat.Bullet.Type) == 2
    remove = CliRunner().invoke(main, ["list", "remove", "--anchor-id", "shape:2:2"])
    assert remove.exit_code == 0
    assert int(body.ParagraphFormat.Bullet.Visible) == 0


def test_format_text_fences_one_undo_entry(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    assert fake_powerpoint._undo_entries == 0
    result = CliRunner().invoke(main, ["format-text", "--anchor-id", "ph:2:title", "--bold"])
    assert result.exit_code == 0
    assert fake_powerpoint._undo_entries == 1


def test_go_to(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 1
    result = CliRunner().invoke(main, ["go-to", "--anchor-id", "shape:3:1"])
    assert result.exit_code == 0
    assert fake_powerpoint._viewed == 3


def test_slide_export(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "slide2.png"
    result = CliRunner().invoke(
        main, ["slide", "export", "--slide", "2", "--out", str(out), "--width", "640"]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["path"] == str(out)
    assert os.path.isfile(payload["path"])


def test_slide_export_temp_default(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["slide", "export", "--slide", "2"])
    assert result.exit_code == 0
    path = _json(result)["path"]
    try:
        assert os.path.isfile(path)
    finally:
        os.remove(path)


def test_selection_cli_shapes(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Title 1")
    result = CliRunner().invoke(main, ["selection"])
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["type"] == "shapes"
    assert payload["anchor_id"] == "shape:2:1"


def test_selection_cli_none(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["selection"])
    assert result.exit_code == 0
    assert _json(result)["type"] == "none"


def test_shape_add_table(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "table", "--rows", "2", "--cols", "3"]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["ok"] is True
    assert payload["has_table"] is True
    assert payload["anchor_id"] == "shape:3:3"


def test_shape_add_table_requires_dims_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "add", "--slide", "3", "--kind", "table"])
    assert result.exit_code == 2  # click UsageError
    assert fake_powerpoint.ActivePresentation.Slides(3).Shapes.Count == 2  # untouched


def test_table_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # Add a table, fill a cell through its anchor, then read the grid.
    CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "table", "--rows", "2", "--cols", "2"]
    )
    CliRunner().invoke(main, ["write", "--anchor-id", "cell:3:3:1:1", "--text", "TL"])
    result = CliRunner().invoke(main, ["table", "read", "--slide", "3", "--shape", "3"])
    assert result.exit_code == 0
    grid = _json(result)
    assert grid["rows"] == 2 and grid["columns"] == 2
    assert grid["cells"][0][0]["text"] == "TL"
    assert grid["cells"][0][0]["anchor_id"] == "cell:3:3:1:1"


def test_table_read_no_table_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # shape:3:1 is the TextBox — no table.
    result = CliRunner().invoke(main, ["table", "read", "--slide", "3", "--shape", "1"])
    assert result.exit_code == 2


def test_table_add_row(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "table", "--rows", "1", "--cols", "2"]
    )
    result = CliRunner().invoke(
        main,
        ["table", "add-row", "--slide", "3", "--shape", "3", "--values", '["x", "y"]'],
    )
    assert result.exit_code == 0
    assert _json(result)["rows"] == 2
    # The new row's cells were filled.
    read = CliRunner().invoke(main, ["table", "read", "--slide", "3", "--shape", "3"])
    assert _json(read)["cells"][1] == [
        {"row": 2, "col": 1, "text": "x", "anchor_id": "cell:3:3:2:1"},
        {"row": 2, "col": 2, "text": "y", "anchor_id": "cell:3:3:2:2"},
    ]


def test_table_add_row_bad_values_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "table", "--rows", "1", "--cols", "2"]
    )
    result = CliRunner().invoke(
        main, ["table", "add-row", "--slide", "3", "--shape", "3", "--values", "not-json"]
    )
    assert result.exit_code == 2  # click UsageError


def test_table_delete_row(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "table", "--rows", "3", "--cols", "1"]
    )
    result = CliRunner().invoke(
        main, ["table", "delete-row", "--slide", "3", "--shape", "3", "--row", "2"]
    )
    assert result.exit_code == 0
    assert _json(result)["rows"] == 2


def test_table_fences_one_undo_entry(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "table", "--rows", "1", "--cols", "1"]
    )
    fake_powerpoint._undo_entries = 0
    result = CliRunner().invoke(main, ["table", "add-row", "--slide", "3", "--shape", "3"])
    assert result.exit_code == 0
    assert fake_powerpoint._undo_entries == 1


def test_text_output_mode(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["--text", "slides"])
    assert result.exit_code == 0
    assert "Welcome" in result.output
    # Not JSON in text mode.
    assert not result.output.lstrip().startswith("[")


# ---------------------------------------------------------------------------
# llm-help / install-skill / install-mcp (offline — no PowerPoint needed)
# ---------------------------------------------------------------------------


def test_llm_help_prints_cli_skill_body() -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["llm-help"])
    assert result.exit_code == 0
    out = result.output
    assert not out.lstrip().startswith("{")  # raw Markdown, not JSON
    assert out.lstrip().startswith("# pptlive (CLI)")
    assert "name: pptlive-cli" not in out  # frontmatter stripped
    assert "--anchor-id" in out and "Exit codes" in out


def test_llm_help_python_prints_python_skill_body() -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["llm-help", "--python"])
    assert result.exit_code == 0
    assert result.output.lstrip().startswith("# pptlive (Python API)")
    assert "import pptlive as pl" in result.output


def test_llm_help_ignores_json_flag() -> None:  # type: ignore[no-untyped-def]
    default = CliRunner().invoke(main, ["llm-help"])
    as_json = CliRunner().invoke(main, ["--json", "llm-help"])
    assert default.exit_code == 0 and as_json.exit_code == 0
    assert default.output == as_json.output
    assert not as_json.output.lstrip().startswith("{")


def test_install_skill_local_writes_both(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    result = CliRunner().invoke(main, ["install-skill"])
    assert result.exit_code == 0
    data = _json(result)
    assert data["ok"] is True and data["scope"] == "local"
    names = {r["name"] for r in data["installed"]}
    assert names == {"pptlive-cli", "pptlive-python"}
    cli = tmp_path / ".agents" / "skills" / "pptlive-cli" / "SKILL.md"
    py = tmp_path / ".agents" / "skills" / "pptlive-python" / "SKILL.md"
    assert cli.exists() and py.exists()
    body = cli.read_text(encoding="utf-8")
    assert body.startswith("---") and "name: pptlive-cli" in body  # frontmatter kept on disk


def test_install_skill_only_cli(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    result = CliRunner().invoke(main, ["install-skill", "--cli"])
    assert result.exit_code == 0
    assert [r["name"] for r in _json(result)["installed"]] == ["pptlive-cli"]
    assert (tmp_path / ".agents" / "skills" / "pptlive-cli" / "SKILL.md").exists()
    assert not (tmp_path / ".agents" / "skills" / "pptlive-python").exists()


def test_install_skill_system(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path))
    result = CliRunner().invoke(main, ["install-skill", "--python", "--system"])
    assert result.exit_code == 0
    assert _json(result)["scope"] == "system"
    assert (tmp_path / ".agents" / "skills" / "pptlive-python" / "SKILL.md").exists()


def test_install_skill_refuses_overwrite_without_force(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    assert CliRunner().invoke(main, ["install-skill", "--cli"]).exit_code == 0
    again = CliRunner().invoke(main, ["install-skill", "--cli"])
    assert again.exit_code == 1
    assert "force" in again.output.lower()
    assert CliRunner().invoke(main, ["install-skill", "--cli", "--force"]).exit_code == 0


def test_llm_help_matches_installed_skill_body(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    CliRunner().invoke(main, ["install-skill", "--cli"])
    installed = (tmp_path / ".agents" / "skills" / "pptlive-cli" / "SKILL.md").read_text(
        encoding="utf-8"
    )
    printed = CliRunner().invoke(main, ["llm-help"]).output
    # The printed guide is the installed skill minus its YAML frontmatter.
    assert installed.rstrip().endswith(printed.rstrip())


def test_install_mcp_print_emits_uvx_entry_and_writes_nothing(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    result = CliRunner().invoke(main, ["install-mcp", "--print"])
    assert result.exit_code == 0
    data = _json(result)
    assert data["mcpServers"]["pptlive"]["command"] == "uvx"
    assert data["mcpServers"]["pptlive"]["args"] == ["--from", "pptlive[mcp]", "pptlive-mcp"]
    # --print never touches the filesystem.
    assert not (tmp_path / ".mcp.json").exists()


def test_install_mcp_directory_form(tmp_path) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["install-mcp", "--print", "--directory", "C:/repo/pptlive"])
    entry = _json(result)["entry"]
    assert entry["command"] == "uv"
    assert entry["args"] == ["run", "--directory", "C:/repo/pptlive", "pptlive-mcp"]


def test_install_mcp_writes_config_and_merges(tmp_path) -> None:  # type: ignore[no-untyped-def]
    cfg = tmp_path / "claude_desktop_config.json"
    cfg.write_text(json.dumps({"mcpServers": {"other": {"command": "x"}}}), encoding="utf-8")
    result = CliRunner().invoke(main, ["install-mcp", "--config", str(cfg)])
    assert result.exit_code == 0
    assert _json(result)["action"] == "created"
    written = json.loads(cfg.read_text(encoding="utf-8"))
    # Existing servers are preserved; pptlive is added.
    assert written["mcpServers"]["other"] == {"command": "x"}
    assert written["mcpServers"]["pptlive"]["command"] == "uvx"


def test_install_mcp_refuses_overwrite_without_force(tmp_path) -> None:  # type: ignore[no-untyped-def]
    cfg = tmp_path / "cfg.json"
    assert CliRunner().invoke(main, ["install-mcp", "--config", str(cfg)]).exit_code == 0
    again = CliRunner().invoke(main, ["install-mcp", "--config", str(cfg)])
    assert again.exit_code == 1 and "force" in again.output.lower()
    forced = CliRunner().invoke(main, ["install-mcp", "--config", str(cfg), "--force"])
    assert forced.exit_code == 0 and _json(forced)["action"] == "updated"


def test_install_mcp_claude_code_writes_project_mcp_json(tmp_path, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    result = CliRunner().invoke(main, ["install-mcp", "--client", "claude-code"])
    assert result.exit_code == 0
    target = tmp_path / ".mcp.json"
    assert target.exists()
    assert (
        json.loads(target.read_text(encoding="utf-8"))["mcpServers"]["pptlive"]["command"] == "uvx"
    )

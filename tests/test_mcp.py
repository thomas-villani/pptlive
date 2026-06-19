"""Tests for the MCP server (`pptlive.mcp.server`), against the fake COM deck.

Skipped unless the optional `mcp` SDK is installed (`pptlive[mcp]` / the `dev`
extra). The MCP tools are plain module-level functions that each call `attach()`,
so the `fake_powerpoint` fixture's monkeypatch (which makes `attach()` return the
fake app) lets us call them directly — no MCP transport needed. State persists
across calls within a test because the monkeypatch hands back the same fake app
every time, exactly like the CLI tests.

The surface is five dispatch tools — `ppt_read` / `ppt_edit` / `ppt_render` /
`ppt_show` / `ppt_batch` — each taking an `op`; the per-op logic lives in
`_<tool>_core` helpers that `ppt_batch` reuses across one shared `attach()`.
"""

from __future__ import annotations

import asyncio
import base64
from typing import Any

import pytest

pytest.importorskip("mcp")

from mcp.server.fastmcp.exceptions import ToolError  # noqa: E402
from mcp.types import CallToolResult, ImageContent, TextContent  # noqa: E402

from pptlive.mcp.server import (  # noqa: E402
    EditOp,
    ReadOp,
    RenderOp,
    ShowOp,
    build_server,
    ppt_batch,
    ppt_edit,
    ppt_read,
    ppt_render,
    ppt_show,
)


def _op_enum(tool: Any) -> list[str]:
    """The allowed `op` values from a tool's input schema.

    The `op` arg is typed by a `StrEnum`, which pydantic renders as a `$ref` into
    `$defs` (rather than an inline `enum`), so resolve through the ref. Still a
    valid JSON-Schema enum — every MCP client dereferences it.
    """
    schema = tool.inputSchema
    op = schema["properties"]["op"]
    if "enum" in op:
        return list(op["enum"])
    ref = op["$ref"].rsplit("/", 1)[-1]  # "#/$defs/ReadOp" -> "ReadOp"
    return list(schema["$defs"][ref]["enum"])


# ---------------------------------------------------------------------------
# Server assembly
# ---------------------------------------------------------------------------


def test_build_server_registers_all_tools() -> None:
    srv = build_server()
    names = {t.name for t in asyncio.run(srv.list_tools())}
    assert names == {"ppt_read", "ppt_edit", "ppt_render", "ppt_show", "ppt_batch"}


def test_every_op_is_documented_in_its_tool_docstring() -> None:
    """The agent-facing docstring is the third place an op must appear (after the
    enum that types it and the dispatch registry keyed by it). The import-time
    `assert set(*_OPS) == set(*Op)` in server.py guards the enum↔registry pair;
    this guards the enum↔docstring pair, so the whole triplet can't drift."""
    for fn, enum in (
        (ppt_read, ReadOp),
        (ppt_edit, EditOp),
        (ppt_render, RenderOp),
        (ppt_show, ShowOp),
    ):
        doc = fn.__doc__ or ""
        missing = [op.value for op in enum if f'"{op.value}"' not in doc]
        assert not missing, f"{fn.__name__} docstring is missing ops: {missing}"


def test_guide_resources_serve_skill_bodies() -> None:
    srv = build_server()
    resources = {str(r.uri): r for r in asyncio.run(srv.list_resources())}
    assert "pptlive://guide" in resources
    assert "pptlive://guide/python" in resources
    cli = asyncio.run(srv.read_resource("pptlive://guide"))
    body = cli[0].content if isinstance(cli, list) else cli
    assert "# pptlive (CLI)" in body and "name: pptlive-cli" not in body  # frontmatter stripped


def test_tool_schema_marks_required_args() -> None:
    srv = build_server()
    tools = {t.name: t for t in asyncio.run(srv.list_tools())}
    # The dispatch arg `op` is the one required field on the read/edit tools.
    assert tools["ppt_read"].inputSchema["required"] == ["op"]
    assert tools["ppt_edit"].inputSchema["required"] == ["op"]
    # ppt_batch's one required arg is the command list.
    assert tools["ppt_batch"].inputSchema["required"] == ["commands"]
    # The StrEnum op surfaces as a schema enum (via $defs) so the agent gets the
    # valid choices.
    assert {"write", "format", "shape_add", "chart_set_data"} <= set(_op_enum(tools["ppt_edit"]))
    write_mode = tools["ppt_edit"].inputSchema["properties"]["mode"]
    assert write_mode["enum"] == ["set", "insert_after", "insert_before"]
    render_ops = set(_op_enum(tools["ppt_render"]))
    assert {"deck_snapshot", "deck_pdf", "save", "save_as"} <= render_ops


def test_structured_output_not_wrapped_under_result(fake_powerpoint: Any) -> None:
    # FastMCP wraps a *bare list / union* tool return under {"result": ...}; every
    # tool returns a plain dict so structured content passes through verbatim. This
    # drives the real call_tool dispatch (a live deck regression caught this).
    srv = build_server()
    _content, structured = asyncio.run(srv.call_tool("ppt_read", {"op": "slides"}))
    assert isinstance(structured, dict)
    assert "slides" in structured and "result" not in structured


# ---------------------------------------------------------------------------
# Reads (ppt_read op=...)
# ---------------------------------------------------------------------------


def test_status(fake_powerpoint: Any) -> None:
    out = ppt_read("status")
    assert [d["name"] for d in out["decks"]] == ["Pitch.pptx"]
    assert out["decks"][0]["is_active"] is True
    assert out["viewed_slide"] == 1


def test_slides(fake_powerpoint: Any) -> None:
    rows = ppt_read("slides")["slides"]
    assert [r["index"] for r in rows] == [1, 2, 3]
    assert rows[1]["title"] == "Agenda"


def test_outline(fake_powerpoint: Any) -> None:
    items = ppt_read("outline")["outline"]
    agenda = next(i for i in items if i["slide"] == 2)
    assert agenda["title"] == "Agenda"
    assert agenda["bullets"] == ["Intro", "Demo", "Q&A"]


def test_slide_read(fake_powerpoint: Any) -> None:
    grid = ppt_read("slide", slide=2)
    assert grid["index"] == 2
    names = [s["name"] for s in grid["shapes"]]
    assert "Title 1" in names and "Picture 3" in names
    picture = next(s for s in grid["shapes"] if s["name"] == "Picture 3")
    assert picture["has_table"] is False


def test_slide_read_requires_slide(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("slide")
    assert "invalid_args" in str(exc.value)


def test_read_anchor_includes_paragraphs(fake_powerpoint: Any) -> None:
    out = ppt_read("anchor", anchor_id="ph:2:body")
    # `text` is raw COM text — PowerPoint separates paragraphs with `\r` (the CLI
    # returns it unnormalized too); the structured `paragraphs` is the clean view.
    assert out["text"] == "Intro\rDemo\rQ&A"
    assert [p["text"] for p in out["paragraphs"]] == ["Intro", "Demo", "Q&A"]
    assert out["paragraphs"][0]["anchor_id"] == "para:2:2:1"


def test_read_anchor_paragraph_has_effective_font(fake_powerpoint: Any) -> None:
    # PPTLIVE-003: each paragraph carries the full effective font, not just bold.
    ppt_edit("format", anchor_id="para:2:2:1", bold=True, italic=True, font="Georgia", size=28)
    para = ppt_read("anchor", anchor_id="ph:2:body")["paragraphs"][0]
    font = para["font"]
    assert font["bold"] is True and font["italic"] is True
    assert font["font"] == "Georgia" and font["size"] == 28.0
    assert set(font) == {
        "bold",
        "italic",
        "underline",
        "size",
        "font",
        "color",
        "color_source",
        "theme_color",
    }


def test_read_text_frame_status(fake_powerpoint: Any) -> None:
    out = ppt_read("text_frame_status", anchor_id="ph:2:body")
    assert out["anchor_id"] == "ph:2:body"
    assert out["autosize"] == "text_to_fit_shape"
    assert out["overflow_risk"] == "low"
    assert out["margins"]["top"] == 3.6


def test_read_notes_has_no_paragraphs_key(fake_powerpoint: Any) -> None:
    out = ppt_read("anchor", anchor_id="notes:1")
    assert out["text"] == "Lead with the vision."
    assert "paragraphs" not in out  # Notes is a plain Anchor, not a Shape


def test_selection_none(fake_powerpoint: Any) -> None:
    assert ppt_read("selection")["type"] == "none"


def test_selection_resolves_here(fake_powerpoint: Any) -> None:
    fake_powerpoint._select_shapes("Title 1")
    out = ppt_read("selection")
    assert out["type"] == "shapes"
    assert out["shapes"][0]["anchor_id"] == "shape:1:1"


def test_layouts(fake_powerpoint: Any) -> None:
    layouts = ppt_read("layouts")["layouts"]
    assert any(row["name"] == "Title and Content" for row in layouts)


# ---------------------------------------------------------------------------
# Writes / formatting (ppt_edit op=...)
# ---------------------------------------------------------------------------


def test_write_set_then_read_back(fake_powerpoint: Any) -> None:
    out = ppt_edit("write", anchor_id="ph:2:title", text="New Title")
    assert out["ok"] is True
    assert ppt_read("anchor", anchor_id="ph:2:title")["text"] == "New Title"
    # The mutation went through deck.edit -> exactly one undo fence.
    assert fake_powerpoint._undo_entries == 1


def test_write_multiple_paragraphs(fake_powerpoint: Any) -> None:
    # `\r` is PowerPoint's paragraph break; set_text passes it straight through.
    ppt_edit("write", anchor_id="ph:2:body", text="One\rTwo")
    out = ppt_read("anchor", anchor_id="ph:2:body")
    assert [p["text"] for p in out["paragraphs"]] == ["One", "Two"]


def test_write_newline_splits_into_addressable_paragraphs(fake_powerpoint: Any) -> None:
    # PPTLIVE-001: `\n` must create real paragraphs (not soft breaks), so each
    # line is its own addressable `para:S:N:P`.
    ppt_edit("write", anchor_id="ph:2:body", text="A\nB\nC\nD")
    out = ppt_read("anchor", anchor_id="ph:2:body")
    assert [p["text"] for p in out["paragraphs"]] == ["A", "B", "C", "D"]
    assert ppt_read("anchor", anchor_id="para:2:2:2")["text"] == "B"


def test_write_crlf_collapses_to_one_paragraph_break(fake_powerpoint: Any) -> None:
    ppt_edit("write", anchor_id="ph:2:body", text="A\r\nB")
    out = ppt_read("anchor", anchor_id="ph:2:body")
    assert [p["text"] for p in out["paragraphs"]] == ["A", "B"]


def test_write_soft_break_stays_one_paragraph(fake_powerpoint: Any) -> None:
    # `\v` (SOFT_BREAK) is a within-paragraph soft line break — NOT a new paragraph.
    ppt_edit("write", anchor_id="ph:2:body", text="A\vB")
    out = ppt_read("anchor", anchor_id="ph:2:body")
    assert len(out["paragraphs"]) == 1
    assert out["paragraphs"][0]["text"] == "A\vB"


def test_write_insert_after_adds_paragraph(fake_powerpoint: Any) -> None:
    ppt_edit("write", anchor_id="ph:2:body", text="Cash runway: 30 months", mode="insert_after")
    paras = ppt_read("anchor", anchor_id="ph:2:body")["paragraphs"]
    assert [p["text"] for p in paras] == ["Intro", "Demo", "Q&A", "Cash runway: 30 months"]


def test_set_paragraphs_builds_a_bulleted_list(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "set_paragraphs",
        anchor_id="ph:2:body",
        paragraphs=[
            {"text": "Launch responsibly", "list_type": "bulleted"},
            {"text": "Measure impact", "list_type": "bulleted"},
            "Brief legal last",
        ],
    )
    assert out["ok"] is True
    assert out["paragraphs"] == ["para:2:2:1", "para:2:2:2", "para:2:2:3"]
    paras = ppt_read("anchor", anchor_id="ph:2:body")["paragraphs"]
    assert [p["text"] for p in paras] == [
        "Launch responsibly",
        "Measure impact",
        "Brief legal last",
    ]
    assert paras[0]["bullet"] == "bulleted"


def test_set_paragraphs_requires_a_list(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("set_paragraphs", anchor_id="ph:2:body", paragraphs=[])
    assert "invalid_args" in str(exc.value)


def test_format_text_sets_bold(fake_powerpoint: Any) -> None:
    out = ppt_edit("format", anchor_id="ph:2:title", bold=True, size=40.0)
    assert out["ok"] is True
    font = fake_powerpoint.ActivePresentation.Slides(2).Shapes(1).TextFrame.TextRange.Font
    assert font.Bold != 0
    assert font.Size == 40.0


def test_format_requires_an_option(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("format", anchor_id="ph:2:title")
    assert "invalid_args" in str(exc.value)


def test_format_line_spacing_points(fake_powerpoint: Any) -> None:
    out = ppt_edit("format", anchor_id="ph:2:body", line_spacing_points=24.0)
    assert out["ok"] is True
    pf = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange.ParagraphFormat
    assert pf.SpaceWithin == 24.0
    assert pf.LineRuleWithin == 0  # msoFalse -> exact points, not 24x


def test_format_line_spacing_guardrail(fake_powerpoint: Any) -> None:
    # The reviewer's footgun: a bare 24 as a multiple is rejected (use points).
    with pytest.raises(ToolError) as exc:
        ppt_edit("format", anchor_id="ph:2:body", line_spacing=24.0)
    assert "line_spacing_points" in str(exc.value)
    # ...unless explicitly forced.
    assert ppt_edit("format", anchor_id="ph:2:body", line_spacing=24.0, force=True)["ok"] is True


def test_format_warns_on_tiny_font(fake_powerpoint: Any) -> None:
    out = ppt_edit("format", anchor_id="ph:2:body", size=5.0)
    assert out["ok"] is True
    assert any("very small" in w for w in out["warnings"])


def test_format_no_warnings_when_clean(fake_powerpoint: Any) -> None:
    out = ppt_edit("format", anchor_id="ph:2:body", size=18.0)
    assert "warnings" not in out  # only present when there's something to flag


def test_text_reset_format(fake_powerpoint: Any) -> None:
    ppt_edit("format", anchor_id="ph:2:body", line_spacing_points=240.0, space_before=99.0)
    out = ppt_edit("text_reset_format", anchor_id="ph:2:body")
    assert out["ok"] is True
    pf = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange.ParagraphFormat
    assert pf.SpaceWithin == 1.0
    assert pf.LineRuleWithin == -1  # single (multiple)
    assert pf.SpaceBefore == 0.0


def test_shape_reset_layout(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_reset_layout", anchor_id="ph:2:body")
    assert out["ok"] is True
    assert out["restored"]["width"] == 828.0
    assert out["restored"]["font_size"] == 28.0


def test_shape_reset_layout_on_non_placeholder_errors(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_reset_layout", anchor_id="shape:3:1")  # a textbox
    assert "invalid_args" in str(exc.value)


def test_format_applies_and_removes_list(fake_powerpoint: Any) -> None:
    assert ppt_edit("format", anchor_id="ph:2:body", list_type="bulleted")["ok"] is True
    bullet = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange
    assert bullet.ParagraphFormat.Bullet.Visible != 0
    ppt_edit("format", anchor_id="ph:2:body", list_type="none")
    assert bullet.ParagraphFormat.Bullet.Visible == 0


# ---------------------------------------------------------------------------
# find / find_replace (ppt_read op=find, ppt_edit op=find_replace)
# ---------------------------------------------------------------------------


def test_find_returns_hits_with_resolvable_anchors(fake_powerpoint: Any) -> None:
    out = ppt_read("find", text="Welcome")
    assert out["count"] == 1
    assert out["matches"][0]["anchor_id"] == "shape:1:1"


def test_find_requires_text(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("find")
    assert "invalid_args" in str(exc.value)


def test_find_replace_single_applies(fake_powerpoint: Any) -> None:
    out = ppt_edit("find_replace", find="Welcome", text="Hello")
    assert out["ok"] is True and out["count"] == 1
    assert ppt_read("anchor", anchor_id="shape:1:1")["text"] == "Hello"
    assert fake_powerpoint._undo_entries == 1  # one undo fence


def test_find_replace_all(fake_powerpoint: Any) -> None:
    out = ppt_edit("find_replace", find="de", text="XX", replace_all=True)
    assert out["count"] == 2
    assert ppt_read("anchor", anchor_id="shape:1:2")["text"] == "A XXmo XXck"


def test_find_replace_ambiguous_raises(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("find_replace", find="de", text="X")
    assert "ambiguous" in str(exc.value)


def test_find_replace_zero_matches_raises_not_found(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("find_replace", find="nonexistent-zzz", text="X")
    assert "not_found" in str(exc.value)


# ---------------------------------------------------------------------------
# Slide lifecycle (ppt_edit op=slide_*)
# ---------------------------------------------------------------------------


def test_slide_add_and_delete(fake_powerpoint: Any) -> None:
    added = ppt_edit("slide_add", layout="blank")
    assert added["ok"] is True and added["index"] == 4
    assert len(ppt_read("slides")["slides"]) == 4
    ppt_edit("slide_delete", slide=4)
    assert len(ppt_read("slides")["slides"]) == 3


def test_slide_add_bad_placeholders_is_invalid_args(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("slide_add", layout="blank", placeholders={"body": {"bogus": 1}})
    assert "invalid_args" in str(exc.value)


def test_slide_duplicate(fake_powerpoint: Any) -> None:
    dup = ppt_edit("slide_duplicate", slide=1)
    assert dup["from"] == 1 and dup["index"] == 2
    assert len(ppt_read("slides")["slides"]) == 4


def test_slide_move(fake_powerpoint: Any) -> None:
    moved = ppt_edit("slide_move", slide=1, to=3)
    assert moved["index"] == 3


def test_slide_delete_missing_slide_is_tool_error(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("slide_delete")
    assert "invalid_args" in str(exc.value)


# ---------------------------------------------------------------------------
# Shapes (ppt_edit op=shape_* / set_alt)
# ---------------------------------------------------------------------------


def test_shape_add_textbox(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_add", slide=3, kind="textbox", text="Hello", left=72.0, top=72.0)
    assert out["ok"] is True
    assert out["type"] == "textbox"
    assert ppt_read("anchor", anchor_id=out["anchor_id"])["text"] == "Hello"


def test_shape_add_table(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_add", slide=3, kind="table", rows=2, cols=2)
    assert out["ok"] is True
    assert out["has_table"] is True


def test_shape_add_table_requires_dims(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_add", slide=3, kind="table")
    assert "invalid_args" in str(exc.value)


def test_shape_add_picture_requires_path(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_add", slide=3, kind="picture")
    assert "invalid_args" in str(exc.value)


def test_shape_move(fake_powerpoint: Any) -> None:
    added = ppt_edit("shape_add", slide=3, kind="textbox", text="x")
    out = ppt_edit("shape_move", anchor_id=added["anchor_id"], left=200.0, top=150.0)
    assert out["geometry"]["left"] == 200.0
    assert out["geometry"]["top"] == 150.0


def test_shape_resize(fake_powerpoint: Any) -> None:
    added = ppt_edit("shape_add", slide=3, kind="textbox", text="x")
    out = ppt_edit("shape_resize", anchor_id=added["anchor_id"], width=321.0)
    assert out["geometry"]["width"] == 321.0


def test_shape_delete(fake_powerpoint: Any) -> None:
    before = len(ppt_read("slide", slide=3)["shapes"])
    added = ppt_edit("shape_add", slide=3, kind="textbox", text="x")
    assert len(ppt_read("slide", slide=3)["shapes"]) == before + 1
    ppt_edit("shape_delete", anchor_id=added["anchor_id"])
    assert len(ppt_read("slide", slide=3)["shapes"]) == before


def test_shape_move_requires_anchor(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_move", left=10.0)
    assert "invalid_args" in str(exc.value)


def test_set_alt(fake_powerpoint: Any) -> None:
    out = ppt_edit("set_alt", anchor_id="shape:2:3", alt_text="Revenue chart")
    assert out["ok"] is True
    assert out["alt_text"] == "Revenue chart"
    assert ppt_read("slide", slide=2)["shapes"][2]["alt_text"] == "Revenue chart"


def test_set_alt_requires_alt_text(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("set_alt", anchor_id="shape:2:3")
    assert "invalid_args" in str(exc.value)


def test_shape_add_picture_with_alt(fake_powerpoint: Any, tmp_path: Any) -> None:
    img = tmp_path / "p.png"
    img.write_bytes(b"\x89PNG\r\n\x1a\n")
    out = ppt_edit("shape_add", slide=3, kind="picture", path=str(img), alt_text="Logo")
    assert out["alt_text"] == "Logo"


def test_shape_image_export(fake_powerpoint: Any, tmp_path: Any) -> None:
    out_path = tmp_path / "shape.png"
    # embed=False keeps the plain structured dict (path only).
    out = ppt_render("shape_image", anchor_id="shape:2:3", out=str(out_path), embed=False)
    assert out["ok"] is True
    assert out["anchor_id"] == "shape:2:3"
    assert out_path.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")


# -- fill / line (PPTLIVE-007), z-order (008), shapeid (010) -----------------


def test_format_sets_shape_fill(fake_powerpoint: Any) -> None:
    out = ppt_edit("format", anchor_id="shape:2:1", fill_color="#FF0000", line_color="none")
    assert out["ok"] is True
    d = ppt_read("slide", slide=2)["shapes"][0]
    assert d["fill"] == {
        "type": "solid",
        "color": "#FF0000",
        "visible": True,
        "transparency": 0.0,
    }
    assert d["line"]["visible"] is False


def test_format_fill_on_non_shape_anchor_errors(fake_powerpoint: Any) -> None:
    # A paragraph anchor has no fill — should be a clean invalid_args, not a 500.
    with pytest.raises(ToolError) as exc:
        ppt_edit("format", anchor_id="para:2:2:1", fill_color="#FF0000")
    assert "invalid_args" in str(exc.value)


def test_shape_add_with_fill(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_add", slide=3, kind="shape", shape_type="rectangle", fill_color="#102030")
    assert out["fill"] == {
        "type": "solid",
        "color": "#102030",
        "visible": True,
        "transparency": 0.0,
    }


def test_shape_gradient_fill(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "shape_gradient_fill",
        anchor_id="shape:2:1",
        colors=["#FF0000", "#0000FF"],
        gradient_style="vertical",
    )
    assert out["ok"] is True
    assert out["fill"]["type"] == "gradient"
    assert out["fill"]["gradient_style"] == "vertical"


def test_shape_gradient_fill_preset(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_gradient_fill", anchor_id="shape:2:1", preset="ocean")
    assert out["fill"]["type"] == "gradient"


def test_shape_gradient_fill_requires_colors_or_preset(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_gradient_fill", anchor_id="shape:2:1")
    assert "invalid_args" in str(exc.value)


def test_shape_pattern_fill(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "shape_pattern_fill", anchor_id="shape:2:1", pattern="percent_50", fore="#FF0000"
    )
    assert out["fill"]["type"] == "patterned"
    assert out["fill"]["pattern"] == "percent_50"


def test_shape_picture_fill_missing_file_errors(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError):
        ppt_edit("shape_picture_fill", anchor_id="shape:2:1", path="nope-12345.png")


def test_shape_set_effect(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "shape_set_effect",
        anchor_id="shape:2:1",
        shadow={"color": "#FF0000", "blur": 8},
        soft_edge=4,
    )
    assert out["effects"]["shadow"]["color"] == "#FF0000"
    assert out["effects"]["soft_edge"]["type"] == 4
    # the active effects also surface on a slide read
    assert ppt_read("slide", slide=2)["shapes"][0]["effects"]["shadow"]["blur"] == 8.0


def test_shape_set_effect_requires_an_arg(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_set_effect", anchor_id="shape:2:1")
    assert "invalid_args" in str(exc.value)


def test_format_sets_fill_transparency(fake_powerpoint: Any) -> None:
    ppt_edit("format", anchor_id="shape:2:1", fill_color="#FF0000", fill_transparency=0.4)
    assert ppt_read("slide", slide=2)["shapes"][0]["fill"]["transparency"] == 0.4


def test_shape_line_style(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "shape_line_style",
        anchor_id="shape:2:1",
        dash="dash_dot",
        end_arrow="triangle",
        end_arrow_size="large",
    )
    assert out["line"]["dash"] == "dash_dot"
    assert out["line"]["end_arrow"] == "triangle"
    # surfaces on a slide read too
    assert ppt_read("slide", slide=2)["shapes"][0]["line"]["dash"] == "dash_dot"


def test_shape_line_style_requires_an_arg(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_line_style", anchor_id="shape:2:1")
    assert "invalid_args" in str(exc.value)


def test_shape_line_style_bad_dash_errors(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_line_style", anchor_id="shape:2:1", dash="squiggle")
    assert "invalid_args" in str(exc.value)


def test_shape_order_send_to_back(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_order", anchor_id="shape:2:3", order="back")
    assert out["ok"] is True
    assert out["index"] == 1
    assert ppt_read("slide", slide=2)["shapes"][0]["name"] == "Picture 3"


def test_shape_order_requires_order(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("shape_order", anchor_id="shape:2:1")
    assert "invalid_args" in str(exc.value)


# ---------------------------------------------------------------------------
# shapeid everywhere: every shape read + mutation echoes the restack-proof
# `shapeid:S:ID`, so a chained edit survives the z-order drift it causes.
# ---------------------------------------------------------------------------


def test_slide_read_carries_shapeid(fake_powerpoint: Any) -> None:
    sh = ppt_read("slide", slide=2)["shapes"][0]  # Title 1, Shape.Id 2
    assert sh["shapeid"] == f"shapeid:2:{sh['id']}"
    assert sh["shapeid"] == "shapeid:2:2"


def test_shape_order_returns_stable_shapeid(fake_powerpoint: Any) -> None:
    # Sending Title 1 to the back shifts shape:S:N indices, but the returned
    # shapeid keeps reaching the same shape.
    out = ppt_edit("shape_order", anchor_id="shape:2:1", order="back")
    assert out["ok"] is True
    assert out["shapeid"] == "shapeid:2:2"  # Title 1 has Shape.Id 2
    assert ppt_read("anchor", anchor_id=out["shapeid"])["text"] == "Agenda"


def test_shape_move_echoes_shapeid(fake_powerpoint: Any) -> None:
    out = ppt_edit("shape_move", anchor_id="shape:2:3", left=12.0)
    assert out["shapeid"] == "shapeid:2:4"  # Picture 3 has Shape.Id 4


# ---------------------------------------------------------------------------
# Geometry report (read op "geometry"): a spatial sanity-check before rendering.
# ---------------------------------------------------------------------------


def test_geometry_report_flags_overlap(fake_powerpoint: Any) -> None:
    rep = ppt_read("geometry", slide=1)
    assert rep["slide"] == 1
    assert rep["slide_size"] == {"width": 960.0, "height": 540.0}
    assert len(rep["shapes"]) == 2
    assert all(not s["off_slide"] for s in rep["shapes"])
    # Both placeholders sit at the default (10,20,100,50) box -> they overlap.
    assert len(rep["overlaps"]) == 1
    ov = rep["overlaps"][0]
    assert {ov["a"], ov["b"]} == {"shape:1:1", "shape:1:2"}
    assert ov["area"] == 100.0 * 50.0
    assert rep["off_slide"] == []
    # Each box carries the stable shapeid for re-addressing.
    assert rep["shapes"][0]["shapeid"] == "shapeid:1:2"


def test_geometry_report_flags_off_slide(fake_powerpoint: Any) -> None:
    # Push a 100-wide shape to left=900 -> right edge 1000 > slide width 960.
    ppt_edit("shape_move", anchor_id="shape:1:1", left=900.0)
    rep = ppt_read("geometry", slide=1)
    moved = next(s for s in rep["shapes"] if s["anchor_id"] == "shape:1:1")
    assert moved["off_slide"] is True
    assert "shape:1:1" in rep["off_slide"]


def test_geometry_report_requires_slide(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("geometry")
    assert "invalid_args" in str(exc.value)


def test_shapeid_anchor_resolves_in_read(fake_powerpoint: Any) -> None:
    out = ppt_read("anchor", anchor_id="shapeid:2:3")  # Content Placeholder 2
    assert out["anchor_id"] == "shapeid:2:3"
    assert out["text"] == "Intro\rDemo\rQ&A"


def test_shapeid_format_survives_a_delete(fake_powerpoint: Any) -> None:
    # Style Picture 3 by its stable id after a delete that shifts z-order indices.
    ppt_edit("shape_delete", anchor_id="shape:2:1")  # delete Title 1
    out = ppt_edit("format", anchor_id="shapeid:2:4", fill_color="#345678")
    assert out["ok"] is True


# ---------------------------------------------------------------------------
# Tables — addressed by the table shape's anchor_id (shape:S:N) + cell anchors
# ---------------------------------------------------------------------------


def test_table_read_add_delete_row(fake_powerpoint: Any) -> None:
    added = ppt_edit("shape_add", slide=3, kind="table", rows=2, cols=2)
    aid = added["anchor_id"]
    grid = ppt_read("table", anchor_id=aid)
    assert grid["rows"] == 2 and grid["columns"] == 2

    after_add = ppt_edit("table_add_row", anchor_id=aid, values=["x", "y"])
    assert after_add["rows"] == 3
    after_del = ppt_edit("table_delete_row", anchor_id=aid, row=3)
    assert after_del["rows"] == 2


def test_table_cell_via_write_and_read(fake_powerpoint: Any) -> None:
    added = ppt_edit("shape_add", slide=3, kind="table", rows=2, cols=2)
    shape_n = int(added["anchor_id"].split(":")[2])  # shape:3:N -> N
    cell_id = f"cell:3:{shape_n}:1:1"
    ppt_edit("write", anchor_id=cell_id, text="Metric")
    assert ppt_read("anchor", anchor_id=cell_id)["text"] == "Metric"


def test_table_delete_row_requires_row(fake_powerpoint: Any) -> None:
    added = ppt_edit("shape_add", slide=3, kind="table", rows=2, cols=2)
    with pytest.raises(ToolError) as exc:
        ppt_edit("table_delete_row", anchor_id=added["anchor_id"])
    assert "invalid_args" in str(exc.value)


# ---------------------------------------------------------------------------
# Charts — addressed by the chart shape's anchor_id (shape:S:N)
# ---------------------------------------------------------------------------


def _add_chart(**kwargs: Any) -> str:
    added = ppt_edit("shape_add", slide=3, kind="chart", **kwargs)
    return str(added["anchor_id"])


def test_shape_add_chart_with_data(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "shape_add",
        slide=3,
        kind="chart",
        chart_type="line",
        categories=["Q1", "Q2"],
        series={"Rev": [10, 20]},
    )
    assert out["has_chart"] is True


def test_chart_read(fake_powerpoint: Any) -> None:
    aid = _add_chart(chart_type="pie", categories=["A", "B"], series={"S": [1, 2]})
    info = ppt_read("chart", anchor_id=aid)
    assert info["chart_type"] == "pie"
    assert info["categories"] == ["A", "B"]
    assert info["series"][0]["values"] == [1.0, 2.0]


def test_chart_set_type(fake_powerpoint: Any) -> None:
    aid = _add_chart()
    out = ppt_edit("chart_set_type", anchor_id=aid, chart_type="bar")
    assert out["chart_type"] == "bar_clustered"


def test_chart_set_data(fake_powerpoint: Any) -> None:
    aid = _add_chart()
    info = ppt_edit("chart_set_data", anchor_id=aid, categories=["X", "Y"], series={"A": [3, 4]})
    assert info["categories"] == ["X", "Y"]
    assert info["series"][0]["values"] == [3.0, 4.0]


def test_chart_set_type_requires_chart_type(fake_powerpoint: Any) -> None:
    aid = _add_chart()
    with pytest.raises(ToolError) as exc:
        ppt_edit("chart_set_type", anchor_id=aid)
    assert "invalid_args" in str(exc.value)


def test_chart_read_non_chart_is_not_found(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("chart", anchor_id="shape:2:3")  # a picture
    assert "not_found" in str(exc.value)


def test_chart_recolor_text(fake_powerpoint: Any) -> None:
    aid = _add_chart(categories=["Q1", "Q2"], series={"R": [1, 2]})
    out = ppt_edit("chart_recolor_text", anchor_id=aid, color="#FFFFFF")
    assert out["ok"] is True
    assert out["color"] == "#FFFFFF"
    assert "chart_area" in out["recolored"]


def test_chart_recolor_text_requires_color(fake_powerpoint: Any) -> None:
    aid = _add_chart()
    with pytest.raises(ToolError) as exc:
        ppt_edit("chart_recolor_text", anchor_id=aid)
    assert "invalid_args" in str(exc.value)


# ---------------------------------------------------------------------------
# SmartArt — addressed by the SmartArt shape's anchor_id (shape:S:N)
# ---------------------------------------------------------------------------


def _add_smartart(**kwargs: Any) -> str:
    added = ppt_edit("shape_add", slide=3, kind="smartart", **kwargs)
    return str(added["anchor_id"])


def test_shape_add_smartart_with_nodes(fake_powerpoint: Any) -> None:
    out = ppt_edit(
        "shape_add",
        slide=3,
        kind="smartart",
        smartart_kind="process",
        nodes=["Discover", "Design", "Build"],
    )
    assert out["has_smartart"] is True
    assert out["type"] == "smart_art"


def test_smartart_read(fake_powerpoint: Any) -> None:
    aid = _add_smartart(smartart_kind="cycle", nodes=["A", "B", "C"])
    info = ppt_read("smartart", anchor_id=aid)
    assert info["layout"] == "cycle"
    assert [n["text"] for n in info["nodes"]] == ["A", "B", "C"]


def test_smartart_set_nodes_tree(fake_powerpoint: Any) -> None:
    aid = _add_smartart(smartart_kind="orgchart")
    info = ppt_edit(
        "smartart_set_nodes",
        anchor_id=aid,
        nodes=[{"text": "CEO", "children": ["VP Eng", "VP Sales"]}],
    )
    assert info["nodes"][0]["text"] == "CEO"
    assert [c["text"] for c in info["nodes"][0]["children"]] == ["VP Eng", "VP Sales"]


def test_smartart_set_nodes_requires_nodes(fake_powerpoint: Any) -> None:
    aid = _add_smartart()
    with pytest.raises(ToolError) as exc:
        ppt_edit("smartart_set_nodes", anchor_id=aid)
    assert "invalid_args" in str(exc.value)


def test_smartart_read_non_smartart_is_not_found(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("smartart", anchor_id="shape:2:3")  # a picture
    assert "not_found" in str(exc.value)


def test_smartart_recolor_text(fake_powerpoint: Any) -> None:
    aid = _add_smartart(smartart_kind="process", nodes=["A", "B", "C"])
    out = ppt_edit("smartart_recolor_text", anchor_id=aid, color="#FFFFFF")
    assert out["ok"] is True
    assert out["color"] == "#FFFFFF"
    assert out["nodes_recolored"] == 3


def test_smartart_recolor_text_requires_color(fake_powerpoint: Any) -> None:
    aid = _add_smartart()
    with pytest.raises(ToolError) as exc:
        ppt_edit("smartart_recolor_text", anchor_id=aid)
    assert "invalid_args" in str(exc.value)


def test_read_op_enum_includes_smartart() -> None:
    srv = build_server()
    tools = {t.name: t for t in asyncio.run(srv.list_tools())}
    assert "smartart" in _op_enum(tools["ppt_read"])
    assert "smartart_set_nodes" in _op_enum(tools["ppt_edit"])
    edit_ops = set(_op_enum(tools["ppt_edit"]))
    assert {"chart_recolor_text", "smartart_recolor_text"} <= edit_ops


# ---------------------------------------------------------------------------
# Comments — ppt_read op=comments, ppt_edit comment_add/reply/delete
# ---------------------------------------------------------------------------


def test_read_comments_slide(fake_powerpoint: Any) -> None:
    out = ppt_read("comments", slide=1)
    assert out["slide"] == 1
    assert out["comments"][0]["text"] == "Tighten this headline."
    assert out["comments"][0]["replies"][0]["text"] == "Agreed — will do."


def test_read_comments_deck_rollup(fake_powerpoint: Any) -> None:
    out = ppt_read("comments")
    assert out["total"] == 1
    assert out["slides"][0]["slide"] == 1


def test_comment_add(fake_powerpoint: Any) -> None:
    out = ppt_edit("comment_add", slide=2, text="Please cite a source.")
    assert out["ok"] is True
    assert out["comment"]["text"] == "Please cite a source."
    assert ppt_read("comments", slide=2)["comments"][0]["text"] == "Please cite a source."


def test_comment_add_requires_text(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("comment_add", slide=2)
    assert "invalid_args" in str(exc.value)


def test_comment_reply(fake_powerpoint: Any) -> None:
    out = ppt_edit("comment_reply", slide=1, index=1, text="Done.")
    assert out["ok"] is True
    assert out["reply"]["text"] == "Done."
    assert [r["text"] for r in ppt_read("comments", slide=1)["comments"][0]["replies"]] == [
        "Agreed — will do.",
        "Done.",
    ]


def test_comment_reply_requires_index(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("comment_reply", slide=1, text="x")
    assert "invalid_args" in str(exc.value)


def test_comment_delete(fake_powerpoint: Any) -> None:
    out = ppt_edit("comment_delete", slide=1, index=1)
    assert out["ok"] is True
    assert ppt_read("comments")["total"] == 0


def test_comment_ops_in_enums() -> None:
    srv = build_server()
    tools = {t.name: t for t in asyncio.run(srv.list_tools())}
    read_ops = set(_op_enum(tools["ppt_read"]))
    edit_ops = set(_op_enum(tools["ppt_edit"]))
    assert "comments" in read_ops
    assert {"comment_add", "comment_reply", "comment_delete"} <= edit_ops


# ---------------------------------------------------------------------------
# Master / theme (ppt_read op=theme|master, ppt_edit op=theme_*/master_*)
# ---------------------------------------------------------------------------


def test_theme_read(fake_powerpoint: Any) -> None:
    info = ppt_read("theme")
    assert "colors" in info and "fonts" in info
    assert "accent1" in info["colors"]


def test_theme_set_color(fake_powerpoint: Any) -> None:
    info = ppt_edit("theme_set_color", slot="accent1", color="#C00000")
    assert info["colors"]["accent1"] == "#C00000"


def test_theme_set_color_requires_color(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("theme_set_color", slot="accent1")
    assert "invalid_args" in str(exc.value)


def test_theme_set_font(fake_powerpoint: Any) -> None:
    info = ppt_edit("theme_set_font", which="major", name="Georgia")
    assert info["fonts"]["major"] == "Georgia"


def test_master_read(fake_powerpoint: Any) -> None:
    info = ppt_read("master")
    assert set(info["text_styles"]) == {"title", "body", "default"}


def test_master_format_text_style(fake_powerpoint: Any) -> None:
    out = ppt_edit("master_format_text_style", style="body", level=1, font="Georgia", size=32)
    assert out["ok"] is True
    assert ppt_read("master")["text_styles"]["body"]["levels"][0]["font"] == "Georgia"


def test_master_format_text_style_level_defaults_to_one(fake_powerpoint: Any) -> None:
    # `level` is optional; omitting it targets level 1 (natural for `title`).
    out = ppt_edit("master_format_text_style", style="title", color="#156082")
    assert out["ok"] is True and out["level"] == 1
    assert ppt_read("master")["text_styles"]["title"]["levels"][0]["color"] == "#156082"


def test_master_format_text_style_requires_an_option(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_edit("master_format_text_style", style="body", level=1)
    assert "invalid_args" in str(exc.value)


def test_master_set_background(fake_powerpoint: Any) -> None:
    out = ppt_edit("master_set_background", color="#202020")
    assert out["ok"] is True
    # The result nests the changed resource under "background" (mirroring the
    # CLI and the read shape), not as bare top-level keys.
    assert out["background"]["color"] == "#202020"


def test_edit_op_enum_includes_theme_master() -> None:
    srv = build_server()
    tools = {t.name: t for t in asyncio.run(srv.list_tools())}
    read_ops = _op_enum(tools["ppt_read"])
    edit_ops = _op_enum(tools["ppt_edit"])
    assert {"theme", "master"} <= set(read_ops)
    assert {
        "theme_set_color",
        "theme_set_font",
        "master_format_text_style",
        "master_format_paragraph_style",
        "master_set_background",
    } <= set(edit_ops)


# ---------------------------------------------------------------------------
# Render + navigate (ppt_render op=...)
# ---------------------------------------------------------------------------


def test_slide_image_export(fake_powerpoint: Any, tmp_path: Any) -> None:
    out_path = tmp_path / "slide2.png"
    out = ppt_render("slide_image", slide=2, out=str(out_path), embed=False)
    assert out["ok"] is True
    assert out["format"] == "png"
    assert out_path.exists()


def _png_dims(data: bytes) -> tuple[int, int]:
    """Recover (width, height) from the stub PNG's IHDR the fake writes."""
    return (
        int.from_bytes(data[16:20], "big"),
        int.from_bytes(data[20:24], "big"),
    )


def test_slide_image_embeds_inline_and_keeps_path(fake_powerpoint: Any, tmp_path: Any) -> None:
    out_path = tmp_path / "slide2.png"
    # Default embed=True returns BOTH the inline image and the structured path.
    res = ppt_render("slide_image", slide=2, out=str(out_path))
    assert isinstance(res, CallToolResult)
    # structuredContent carries the path so a co-located filesystem tool still works.
    assert res.structuredContent is not None
    assert res.structuredContent["ok"] is True
    assert res.structuredContent["path"] == str(out_path)
    # the inline block is a real base64 image a vision model can see.
    images = [c for c in res.content if isinstance(c, ImageContent)]
    assert len(images) == 1
    assert images[0].mimeType == "image/png"
    decoded = base64.b64decode(images[0].data)
    assert decoded == out_path.read_bytes()
    assert decoded.startswith(b"\x89PNG\r\n\x1a\n")


def test_slide_image_embed_defaults_to_legible_width(fake_powerpoint: Any) -> None:
    # With no width/height and embed on, the render is capped to ~1024 px wide
    # so the inline block stays cheap — the fake echoes the requested size.
    res = ppt_render("slide_image", slide=2)
    assert isinstance(res, CallToolResult)
    image = next(c for c in res.content if isinstance(c, ImageContent))
    width, _height = _png_dims(base64.b64decode(image.data))
    assert width == 1024


def test_shape_image_embeds_with_correct_mime(fake_powerpoint: Any) -> None:
    res = ppt_render("shape_image", anchor_id="shape:2:3", fmt="jpg")
    assert isinstance(res, CallToolResult)
    image = next(c for c in res.content if isinstance(c, ImageContent))
    assert image.mimeType == "image/jpeg"


def test_deck_snapshot_embeds_one_image_per_slide(fake_powerpoint: Any) -> None:
    # The whole-deck vision read: one labelled image block per slide, inline.
    res = ppt_render("deck_snapshot")
    assert isinstance(res, CallToolResult)
    assert res.structuredContent is not None
    assert res.structuredContent["ok"] is True
    assert res.structuredContent["count"] == 3
    # structured result carries the written paths (no bytes — no double-encode).
    assert [img["slide"] for img in res.structuredContent["images"]] == [1, 2, 3]
    images = [c for c in res.content if isinstance(c, ImageContent)]
    assert len(images) == 3
    assert all(img.mimeType == "image/png" for img in images)
    # each image is preceded by a "slide N" text label.
    labels = [
        c.text for c in res.content if isinstance(c, TextContent) and c.text.startswith("slide ")
    ]
    assert labels == ["slide 1", "slide 2", "slide 3"]


def test_deck_snapshot_embed_defaults_max_dim(fake_powerpoint: Any) -> None:
    # With no max_dim and embed on, the cap defaults to ~1000 px (long edge) so
    # the whole deck stays cheap — the fake echoes the requested size in the PNG.
    res = ppt_render("deck_snapshot", slides="1")
    assert isinstance(res, CallToolResult)
    image = next(c for c in res.content if isinstance(c, ImageContent))
    width, _height = _png_dims(base64.b64decode(image.data))
    assert width == 1000


def test_deck_snapshot_respects_slides_span(fake_powerpoint: Any) -> None:
    res = ppt_render("deck_snapshot", slides="2-3")
    assert isinstance(res, CallToolResult)
    assert [img["slide"] for img in res.structuredContent["images"]] == [2, 3]


def test_deck_snapshot_bad_slides_is_invalid_args(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError, match="invalid_args"):
        ppt_render("deck_snapshot", slides="oops")


def test_slide_image_embed_survives_server_call_tool(fake_powerpoint: Any) -> None:
    # The real MCP path: FastMCP validates a tool's result against the output
    # schema it infers from `-> dict`. Returning a CallToolResult must pass that
    # gate (structuredContent is the dict; the image rides as a content block).
    srv = build_server()
    res = asyncio.run(srv.call_tool("ppt_render", {"op": "slide_image", "slide": 2}))
    assert isinstance(res, CallToolResult)
    assert res.structuredContent is not None
    assert res.structuredContent["ok"] is True and res.structuredContent["path"]
    assert any(isinstance(c, ImageContent) for c in res.content)


def test_navigate_never_embeds(fake_powerpoint: Any) -> None:
    # navigate produces no image, so it stays a plain dict even with embed on.
    out = ppt_render("navigate", anchor_id="shape:2:1")
    assert not isinstance(out, CallToolResult)
    assert out["ok"] is True


def test_navigate_moves_the_view(fake_powerpoint: Any) -> None:
    assert fake_powerpoint._viewed == 1
    out = ppt_render("navigate", anchor_id="shape:2:1")
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 2


# ---------------------------------------------------------------------------
# Save / export (ppt_render op=deck_pdf|save|save_as)
# ---------------------------------------------------------------------------


def test_deck_pdf_writes_and_never_embeds(fake_powerpoint: Any, tmp_path: Any) -> None:
    # A PDF carries `path` like an image op, but its format is "pdf" — it must NOT
    # be read back and mis-encoded as an inline image block.
    out = tmp_path / "deck.pdf"
    res = ppt_render("deck_pdf", out=str(out))
    assert not isinstance(res, CallToolResult)  # plain dict, no image embedding
    assert res["ok"] is True
    assert res["path"] == str(out.resolve())
    assert out.read_bytes().startswith(b"%PDF-")


def test_deck_pdf_requires_out(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError, match="requires `out`"):
        ppt_render("deck_pdf")


def test_save_clears_dirty_flag(fake_powerpoint: Any) -> None:
    fake_powerpoint.ActivePresentation.Saved = 0
    out = ppt_render("save")
    assert out["ok"] is True and out["saved"] is True
    assert fake_powerpoint.ActivePresentation.Saved == -1


def test_save_never_saved_is_error(fake_powerpoint: Any) -> None:
    fake_powerpoint.ActivePresentation.FullName = "Presentation1"
    with pytest.raises(ToolError, match="never been saved"):
        ppt_render("save")


def test_save_as_writes_pptx_and_rebinds(fake_powerpoint: Any, tmp_path: Any) -> None:
    out = tmp_path / "copy.pptx"
    res = ppt_render("save_as", out=str(out))
    assert res["ok"] is True
    assert res["path"] == str(out.resolve())
    assert out.read_bytes().startswith(b"PK\x03\x04")
    assert fake_powerpoint.ActivePresentation.FullName == str(out.resolve())


def test_save_as_refuses_overwrite(fake_powerpoint: Any, tmp_path: Any) -> None:
    out = tmp_path / "copy.pptx"
    out.write_bytes(b"old")
    with pytest.raises(ToolError, match="invalid_args"):
        ppt_render("save_as", out=str(out))
    assert out.read_bytes() == b"old"
    res = ppt_render("save_as", out=str(out), overwrite=True)
    assert res["ok"] is True
    assert out.read_bytes().startswith(b"PK\x03\x04")


def test_save_as_threads_format_param(fake_powerpoint: Any, tmp_path: Any) -> None:
    # save_as honors an explicit `format` and reports it back, matching the CLI's
    # --format (it used to hard-code "pptx" and ignore the param).
    out = tmp_path / "copy.pptx"
    res = ppt_render("save_as", out=str(out), save_format="pptx")
    assert res["ok"] is True
    assert res["format"] == "pptx"


def test_save_as_unknown_format_is_invalid_args(fake_powerpoint: Any, tmp_path: Any) -> None:
    with pytest.raises(ToolError, match="invalid_args"):
        ppt_render("save_as", out=str(tmp_path / "x.odp"), save_format="odp")


def test_deck_pdf_in_batch_does_not_break_image_embedding(
    fake_powerpoint: Any, tmp_path: Any
) -> None:
    # A batch mixing a slide_image (embeds) with a deck_pdf (no image): the PDF's
    # path must be skipped by _render_reply, the slide image still embedded.
    out_png = tmp_path / "s.png"
    out_pdf = tmp_path / "d.pdf"
    res = ppt_batch(
        [
            {"tool": "render", "op": "slide_image", "slide": 1, "out": str(out_png)},
            {"tool": "render", "op": "deck_pdf", "out": str(out_pdf)},
        ]
    )
    assert isinstance(res, CallToolResult)
    images = [c for c in res.content if isinstance(c, ImageContent)]
    assert len(images) == 1  # only the PNG, not the PDF
    assert out_pdf.read_bytes().startswith(b"%PDF-")
    assert res.structuredContent is not None
    assert all(r["ok"] for r in res.structuredContent["results"])


# ---------------------------------------------------------------------------
# Live slide show (ppt_show op=...)
# ---------------------------------------------------------------------------


def test_show_state_not_running(fake_powerpoint: Any) -> None:
    out = ppt_show("state")
    assert out["running"] is False
    assert out["state"] == "done"
    assert out["slide_count"] == 3


def test_show_start_then_navigate(fake_powerpoint: Any) -> None:
    started = ppt_show("start")
    assert started["running"] is True
    assert started["current_slide"] == 1
    assert ppt_show("next")["current_slide"] == 2
    assert ppt_show("goto", slide=3)["current_slide"] == 3
    assert ppt_show("previous")["current_slide"] == 2
    black = ppt_show("black")
    assert black["state"] == "black"
    assert ppt_show("resume")["state"] == "running"
    assert ppt_show("end")["running"] is False


def test_show_start_from_slide(fake_powerpoint: Any) -> None:
    out = ppt_show("start", slide=2)
    assert out["running"] is True
    assert out["current_slide"] == 2


def test_show_next_without_running_errors(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_show("next")
    assert "error" in str(exc.value)


# ---------------------------------------------------------------------------
# Batch (ppt_batch)
# ---------------------------------------------------------------------------


def test_batch_atomic_is_one_undo_entry(fake_powerpoint: Any) -> None:
    out = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "Q3 Results"},
            {"op": "write", "anchor_id": "ph:2:body", "text": "Up and to the right"},
        ]
    )
    assert out["ok"] is True
    assert out["count"] == 2
    assert [r["ok"] for r in out["results"]] == [True, True]
    assert ppt_read("anchor", anchor_id="ph:2:title")["text"] == "Q3 Results"
    # The whole batch is fenced into a single undo entry.
    assert fake_powerpoint._undo_entries == 1


def test_batch_non_atomic_is_one_entry_per_edit(fake_powerpoint: Any) -> None:
    out = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "A"},
            {"op": "write", "anchor_id": "ph:2:body", "text": "B"},
        ],
        atomic=False,
    )
    assert out["ok"] is True
    assert fake_powerpoint._undo_entries == 2


def test_batch_builds_a_slide(fake_powerpoint: Any) -> None:
    # Add a textbox then write into the shape the previous command created — the
    # anchor_id from a prior result isn't available, so target by the known z-order.
    before = len(ppt_read("slide", slide=3)["shapes"])
    out = ppt_batch(
        [
            {"op": "slide_add", "layout": "blank"},
            {"op": "shape_add", "slide": 3, "kind": "textbox", "text": "Title"},
        ]
    )
    assert out["ok"] is True
    assert out["results"][0]["result"]["index"] == 4  # slide_add lands at the end
    assert len(ppt_read("slide", slide=3)["shapes"]) == before + 1


def test_batch_defaults_tool_to_edit(fake_powerpoint: Any) -> None:
    # No "tool" key -> defaults to "edit".
    out = ppt_batch([{"op": "write", "anchor_id": "ph:2:title", "text": "X"}])
    assert out["results"][0]["tool"] == "edit"
    assert out["ok"] is True


def test_batch_mixed_read_and_edit(fake_powerpoint: Any) -> None:
    out = ppt_batch(
        [
            {"tool": "edit", "op": "write", "anchor_id": "ph:2:title", "text": "Hi"},
            {"tool": "read", "op": "anchor", "anchor_id": "ph:2:title"},
        ]
    )
    assert out["ok"] is True
    assert out["results"][1]["result"]["text"] == "Hi"


def test_batch_embeds_render_image_with_summary(fake_powerpoint: Any) -> None:
    # Build then look in one round trip: the edit summary rides as structured
    # content while the rendered slide comes back as an inline image block.
    res = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "Q3"},
            {"tool": "render", "op": "slide_image", "slide": 2},
        ]
    )
    assert isinstance(res, CallToolResult)
    assert res.structuredContent is not None
    assert res.structuredContent["ok"] is True
    assert res.structuredContent["count"] == 2
    images = [c for c in res.content if isinstance(c, ImageContent)]
    assert len(images) == 1
    assert base64.b64decode(images[0].data).startswith(b"\x89PNG\r\n\x1a\n")


def test_batch_render_embed_false_is_plain_dict(fake_powerpoint: Any) -> None:
    out = ppt_batch(
        [{"tool": "render", "op": "slide_image", "slide": 2}],
        embed=False,
    )
    assert not isinstance(out, CallToolResult)
    assert out["ok"] is True


def test_batch_stops_on_first_error(fake_powerpoint: Any) -> None:
    out = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "ok"},
            {"op": "write", "anchor_id": "ph:9:title", "text": "bad slide"},
            {"op": "write", "anchor_id": "ph:2:body", "text": "never runs"},
        ]
    )
    assert out["ok"] is False
    assert out["count"] == 2  # stopped after the failure; third never ran
    assert out["results"][1]["ok"] is False
    assert out["results"][1]["error"] == "not_found"
    # The body was never touched.
    assert ppt_read("anchor", anchor_id="ph:2:body")["text"] == "Intro\rDemo\rQ&A"


def test_batch_continue_on_error_reports_each(fake_powerpoint: Any) -> None:
    out = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "ok"},
            {"op": "write", "anchor_id": "ph:9:title", "text": "bad"},
            {"op": "write", "anchor_id": "ph:2:body", "text": "still runs"},
        ],
        stop_on_error=False,
    )
    assert out["ok"] is False
    assert out["count"] == 3
    assert [r["ok"] for r in out["results"]] == [True, False, True]
    assert ppt_read("anchor", anchor_id="ph:2:body")["text"] == "still runs"


def test_batch_unknown_op_is_invalid_args(fake_powerpoint: Any) -> None:
    out = ppt_batch([{"op": "frobnicate"}], stop_on_error=False)
    assert out["results"][0]["ok"] is False
    assert out["results"][0]["error"] == "invalid_args"


def test_batch_empty_list_is_tool_error(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_batch([])
    assert "invalid_args" in str(exc.value)


# ---------------------------------------------------------------------------
# Tool discoverability (PPTLIVE-002): a search for the product name must surface
# every tool, so each tool's description carries the word "PowerPoint".
# ---------------------------------------------------------------------------


def test_every_tool_description_mentions_powerpoint() -> None:
    for tool in (ppt_read, ppt_edit, ppt_render, ppt_show, ppt_batch):
        assert "PowerPoint" in (tool.__doc__ or ""), f"{tool.__name__} omits 'PowerPoint'"


# ---------------------------------------------------------------------------
# View preservation (the politeness guarantee) across batches — regression
# coverage for the "jumps back to slide 1 after an action" report.
# ---------------------------------------------------------------------------


def test_batch_atomic_preserves_viewed_slide(fake_powerpoint: Any) -> None:
    # User is parked on slide 3; an atomic batch of edits must leave them there.
    fake_powerpoint._viewed = 3
    out = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "Q3"},
            {"op": "write", "anchor_id": "ph:2:body", "text": "up"},
        ]
    )
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 3


def test_batch_navigate_is_not_snapped_back(fake_powerpoint: Any) -> None:
    # A deliberate `navigate` inside an atomic batch (which opens one EditScope)
    # must survive the scope's view restore — not be reverted to the pre-batch
    # slide. Without the allow_view_move opt-out it would snap back to slide 3.
    fake_powerpoint._viewed = 3
    out = ppt_batch(
        [
            {"tool": "edit", "op": "write", "anchor_id": "ph:2:title", "text": "Q3"},
            {"tool": "render", "op": "navigate", "anchor_id": "shape:2:1"},
        ]
    )
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 2  # navigate target, not snapped back to 3


def test_single_edit_preserves_viewed_slide(fake_powerpoint: Any) -> None:
    fake_powerpoint._viewed = 4
    ppt_edit("write", anchor_id="ph:2:title", text="hello")
    assert fake_powerpoint._viewed == 4


# ---------------------------------------------------------------------------
# "Follow the work" (default on): an authoring batch that ADDS a slide ends on
# the slide it built, instead of snapping back to slide 1 every batch.
# ---------------------------------------------------------------------------


def test_batch_follows_the_work_to_new_slide(fake_powerpoint: Any) -> None:
    # User is on slide 1; the batch adds slide 4 and builds on it. The view
    # should END on slide 4 (the work), not snap back to slide 1.
    fake_powerpoint._viewed = 1
    out = ppt_batch(
        [
            {"op": "slide_add", "layout": "blank"},
            {"op": "shape_add", "slide": 4, "kind": "textbox", "text": "Built here"},
        ]
    )
    assert out["ok"] is True
    assert out["results"][0]["result"]["index"] == 4  # new slide at the end
    assert fake_powerpoint._viewed == 4  # followed the work, not snapped back to 1


def test_batch_pure_edit_does_not_follow(fake_powerpoint: Any) -> None:
    # No slide added -> not an authoring batch -> the polite restore still wins.
    fake_powerpoint._viewed = 1
    out = ppt_batch(
        [
            {"op": "write", "anchor_id": "ph:2:title", "text": "Q3"},
            {"op": "write", "anchor_id": "ph:2:body", "text": "up"},
        ]
    )
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 1  # preserved (edits only touched slide 2)


def test_batch_follow_view_false_restores(fake_powerpoint: Any) -> None:
    # The per-call opt-out forces the polite snap-back even on an authoring batch.
    fake_powerpoint._viewed = 1
    out = ppt_batch(
        [
            {"op": "slide_add", "layout": "blank"},
            {"op": "shape_add", "slide": 4, "kind": "textbox", "text": "x"},
        ],
        follow_view=False,
    )
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 1  # snapped back to the pre-batch slide


def test_batch_follow_view_env_off(fake_powerpoint: Any, monkeypatch: Any) -> None:
    # PPTLIVE_VIEW_FOLLOW=0 disables follow without a per-call flag.
    monkeypatch.setenv("PPTLIVE_VIEW_FOLLOW", "0")
    fake_powerpoint._viewed = 1
    out = ppt_batch(
        [
            {"op": "slide_add", "layout": "blank"},
            {"op": "shape_add", "slide": 4, "kind": "textbox", "text": "x"},
        ]
    )
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 1


def test_batch_navigate_wins_over_follow(fake_powerpoint: Any) -> None:
    # A deliberate navigate inside an authoring batch is respected — follow must
    # not override it with the new slide.
    fake_powerpoint._viewed = 1
    out = ppt_batch(
        [
            {"op": "slide_add", "layout": "blank"},
            {"tool": "render", "op": "navigate", "anchor_id": "shape:2:1"},
        ]
    )
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 2  # navigate target, not the new slide 4


# ---------------------------------------------------------------------------
# Error taxonomy -> ToolError category tokens (the CLI exit-code analog)
# ---------------------------------------------------------------------------


def test_not_found_maps_to_tool_error(fake_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("anchor", anchor_id="ph:9:title")
    assert "not_found" in str(exc.value)


def test_no_text_frame_maps_to_tool_error(fake_powerpoint: Any) -> None:
    # shape:2:3 is the Picture — no text frame.
    with pytest.raises(ToolError) as exc:
        ppt_edit("write", anchor_id="shape:2:3", text="nope")
    assert "no_text_frame" in str(exc.value)


def test_not_running_maps_to_tool_error(no_powerpoint: Any) -> None:
    with pytest.raises(ToolError) as exc:
        ppt_read("status")
    assert "not_running" in str(exc.value)

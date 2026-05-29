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
from typing import Any

import pytest

pytest.importorskip("mcp")

from mcp.server.fastmcp.exceptions import ToolError  # noqa: E402

from pptlive.mcp.server import (  # noqa: E402
    build_server,
    ppt_batch,
    ppt_edit,
    ppt_read,
    ppt_render,
    ppt_show,
)

# ---------------------------------------------------------------------------
# Server assembly
# ---------------------------------------------------------------------------


def test_build_server_registers_all_tools() -> None:
    srv = build_server()
    names = {t.name for t in asyncio.run(srv.list_tools())}
    assert names == {"ppt_read", "ppt_edit", "ppt_render", "ppt_show", "ppt_batch"}


def test_tool_schema_marks_required_args() -> None:
    srv = build_server()
    tools = {t.name: t for t in asyncio.run(srv.list_tools())}
    # The dispatch arg `op` is the one required field on the read/edit tools.
    assert tools["ppt_read"].inputSchema["required"] == ["op"]
    assert tools["ppt_edit"].inputSchema["required"] == ["op"]
    # ppt_batch's one required arg is the command list.
    assert tools["ppt_batch"].inputSchema["required"] == ["commands"]
    # Literal -> enum surfaces in the schema so the agent gets valid choices.
    edit_op = tools["ppt_edit"].inputSchema["properties"]["op"]
    assert {"write", "format", "shape_add", "chart_set_data"} <= set(edit_op["enum"])
    write_mode = tools["ppt_edit"].inputSchema["properties"]["mode"]
    assert write_mode["enum"] == ["set", "insert_after", "insert_before"]


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


def test_write_insert_after_adds_paragraph(fake_powerpoint: Any) -> None:
    ppt_edit("write", anchor_id="ph:2:body", text="Cash runway: 30 months", mode="insert_after")
    paras = ppt_read("anchor", anchor_id="ph:2:body")["paragraphs"]
    assert [p["text"] for p in paras] == ["Intro", "Demo", "Q&A", "Cash runway: 30 months"]


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


def test_format_applies_and_removes_list(fake_powerpoint: Any) -> None:
    assert ppt_edit("format", anchor_id="ph:2:body", list_type="bulleted")["ok"] is True
    bullet = fake_powerpoint.ActivePresentation.Slides(2).Shapes(2).TextFrame.TextRange
    assert bullet.ParagraphFormat.Bullet.Visible != 0
    ppt_edit("format", anchor_id="ph:2:body", list_type="none")
    assert bullet.ParagraphFormat.Bullet.Visible == 0


# ---------------------------------------------------------------------------
# Slide lifecycle (ppt_edit op=slide_*)
# ---------------------------------------------------------------------------


def test_slide_add_and_delete(fake_powerpoint: Any) -> None:
    added = ppt_edit("slide_add", layout="blank")
    assert added["ok"] is True and added["index"] == 4
    assert len(ppt_read("slides")["slides"]) == 4
    ppt_edit("slide_delete", slide=4)
    assert len(ppt_read("slides")["slides"]) == 3


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
    out = ppt_render("shape_image", anchor_id="shape:2:3", out=str(out_path))
    assert out["ok"] is True
    assert out["anchor_id"] == "shape:2:3"
    assert out_path.read_bytes().startswith(b"\x89PNG\r\n\x1a\n")


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


def test_read_op_enum_includes_smartart() -> None:
    srv = build_server()
    tools = {t.name: t for t in asyncio.run(srv.list_tools())}
    assert "smartart" in tools["ppt_read"].inputSchema["properties"]["op"]["enum"]
    assert "smartart_set_nodes" in tools["ppt_edit"].inputSchema["properties"]["op"]["enum"]


# ---------------------------------------------------------------------------
# Render + navigate (ppt_render op=...)
# ---------------------------------------------------------------------------


def test_slide_image_export(fake_powerpoint: Any, tmp_path: Any) -> None:
    out_path = tmp_path / "slide2.png"
    out = ppt_render("slide_image", slide=2, out=str(out_path))
    assert out["ok"] is True
    assert out["format"] == "png"
    assert out_path.exists()


def test_navigate_moves_the_view(fake_powerpoint: Any) -> None:
    assert fake_powerpoint._viewed == 1
    out = ppt_render("navigate", anchor_id="shape:2:1")
    assert out["ok"] is True
    assert fake_powerpoint._viewed == 2


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

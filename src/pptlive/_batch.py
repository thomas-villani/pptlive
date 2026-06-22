"""Op-dispatch core for the pptlive MCP surface — fastmcp-free.

This module holds the whole dispatch layer the MCP server and the CLI `exec` verb
share: the four op enums, the handler registries + decorators, every `_<tool>_*`
handler, the `_<tool>_core` dispatchers, and `run_batch`. It deliberately does
**not** import `mcp`/`fastmcp` — invalid arguments raise the native `BatchOpError`
(not `mcp`'s `ToolError`) — so the base CLI can drive a batch script without the
optional `[mcp]` extra installed. `mcp/server.py` wraps these in FastMCP tools and
maps `BatchOpError` back to a `ToolError`; `cli/commands.py exec` maps it to exit 1.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Callable
from contextlib import nullcontext
from enum import StrEnum
from typing import Any

from ._anchors import LINE_SPACING_MULTIPLE_MAX, SOFT_BREAK
from ._presentation import Presentation
from ._shapes import Shape
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    BatchOpError,
    NoTextFrameError,
    PowerPointBusyError,
    PowerPointNotRunningError,
    PptliveError,
    PresentationNotFoundError,
)


def _error_code(exc: PptliveError) -> str:
    # Order mirrors cli.main._exit_for: NoTextFrameError before the generic
    # AnchorNotFoundError (which covers SlideNotFoundError / LayoutNotFoundError).
    if isinstance(exc, BatchOpError):
        return "invalid_args"
    if isinstance(exc, NoTextFrameError):
        return "no_text_frame"
    if isinstance(exc, AnchorNotFoundError):
        return "not_found"
    if isinstance(exc, AmbiguousMatchError):
        return "ambiguous"
    if isinstance(exc, PowerPointBusyError):
        return "busy"
    if isinstance(exc, PowerPointNotRunningError):
        return "not_running"
    if isinstance(exc, PresentationNotFoundError):
        return "not_found"
    return "error"


def _require(condition: Any, message: str) -> None:
    """Guard a required argument; surfaces as a ToolError, not a 500."""
    if not condition:
        raise BatchOpError(f"invalid_args: {message}")


def _pick_deck(ppt: Any, doc: str | None) -> Presentation:
    """The active presentation, or the one named `doc` (mirrors the CLI's --doc)."""
    if doc is None:
        return ppt.presentations.active
    return ppt.presentations[doc]


def _resolve_shape(deck: Presentation, anchor_id: str | None) -> Shape:
    """Resolve a shape/placeholder anchor to a `Shape` (else a not_found ToolError)."""
    _require(anchor_id is not None, "this op requires `anchor_id`")
    assert anchor_id is not None
    anchor = deck.anchor_by_id(anchor_id)
    if not isinstance(anchor, Shape):
        raise AnchorNotFoundError("shape", anchor_id)
    return anchor


# ===========================================================================


class ReadOp(StrEnum):
    STATUS = "status"
    SLIDES = "slides"
    OUTLINE = "outline"
    SLIDE = "slide"
    GEOMETRY = "geometry"
    ANIMATIONS = "animations"
    SECTIONS = "sections"
    HEADERS_FOOTERS = "headers_footers"
    ANCHOR = "anchor"
    TEXT_FRAME_STATUS = "text_frame_status"
    SELECTION = "selection"
    FIND = "find"
    TABLE = "table"
    CHART = "chart"
    SMARTART = "smartart"
    COMMENTS = "comments"
    THEME = "theme"
    MASTER = "master"
    LAYOUTS = "layouts"


class EditOp(StrEnum):
    WRITE = "write"
    SET_PARAGRAPHS = "set_paragraphs"
    FIND_REPLACE = "find_replace"
    FORMAT = "format"
    TEXT_RESET_FORMAT = "text_reset_format"
    SLIDE_ADD = "slide_add"
    SLIDE_DELETE = "slide_delete"
    SLIDE_DUPLICATE = "slide_duplicate"
    SLIDE_MOVE = "slide_move"
    SET_LAYOUT = "set_layout"
    SHAPE_ADD = "shape_add"
    SHAPE_MOVE = "shape_move"
    SHAPE_RESIZE = "shape_resize"
    SHAPE_DELETE = "shape_delete"
    SHAPE_ORDER = "shape_order"
    SHAPE_ANIMATE = "shape_animate"
    SHAPE_CLEAR_ANIMATIONS = "shape_clear_animations"
    SLIDE_CLEAR_ANIMATIONS = "slide_clear_animations"
    SHAPE_RESET_LAYOUT = "shape_reset_layout"
    SHAPE_GRADIENT_FILL = "shape_gradient_fill"
    SHAPE_PICTURE_FILL = "shape_picture_fill"
    SHAPE_SET_PICTURE = "shape_set_picture"
    SHAPE_PATTERN_FILL = "shape_pattern_fill"
    SHAPE_SET_EFFECT = "shape_set_effect"
    SHAPE_LINE_STYLE = "shape_line_style"
    SHAPE_SET_HYPERLINK = "shape_set_hyperlink"
    SHAPE_REMOVE_HYPERLINK = "shape_remove_hyperlink"
    SLIDE_SET_TRANSITION = "slide_set_transition"
    SLIDE_SET_BACKGROUND = "slide_set_background"
    SECTION_ADD = "section_add"
    SECTION_RENAME = "section_rename"
    SECTION_DELETE = "section_delete"
    SECTION_MOVE = "section_move"
    SET_HEADERS_FOOTERS = "set_headers_footers"
    SET_ALT = "set_alt"
    TABLE_ADD_ROW = "table_add_row"
    TABLE_DELETE_ROW = "table_delete_row"
    TABLE_ADD_COLUMN = "table_add_column"
    TABLE_DELETE_COLUMN = "table_delete_column"
    TABLE_SET_FILL = "table_set_fill"
    TABLE_SET_BORDER = "table_set_border"
    CHART_SET_TYPE = "chart_set_type"
    CHART_SET_DATA = "chart_set_data"
    CHART_RECOLOR_TEXT = "chart_recolor_text"
    SMARTART_SET_NODES = "smartart_set_nodes"
    SMARTART_RECOLOR_TEXT = "smartart_recolor_text"
    SMARTART_FORMAT_NODE = "smartart_format_node"
    COMMENT_ADD = "comment_add"
    COMMENT_REPLY = "comment_reply"
    COMMENT_DELETE = "comment_delete"
    THEME_SET_COLOR = "theme_set_color"
    THEME_SET_FONT = "theme_set_font"
    MASTER_FORMAT_TEXT_STYLE = "master_format_text_style"
    MASTER_FORMAT_PARAGRAPH_STYLE = "master_format_paragraph_style"
    MASTER_SET_BACKGROUND = "master_set_background"


class RenderOp(StrEnum):
    SLIDE_IMAGE = "slide_image"
    DECK_SNAPSHOT = "deck_snapshot"
    SHAPE_IMAGE = "shape_image"
    DECK_PDF = "deck_pdf"
    SAVE = "save"
    SAVE_AS = "save_as"
    NAVIGATE = "navigate"


class ShowOp(StrEnum):
    STATE = "state"
    START = "start"
    END = "end"
    NEXT = "next"
    PREVIOUS = "previous"
    GOTO = "goto"
    BLACK = "black"
    WHITE = "white"
    RESUME = "resume"


# ===========================================================================
# Op registries — one op = one handler = one registry key. read/render handlers
# take the app handle `(ppt, p)` and pick their own deck (status reads across
# all decks, so it never picks one); edit/show handlers take an already-picked
# `(deck, p)` — for edit, under the caller's open `deck.edit(...)` scope — so
# ppt_batch can run them all under one shared attach + undo fence.
# ===========================================================================

ReadHandler = Callable[[Any, dict[str, Any]], dict[str, Any]]
EditHandler = Callable[[Presentation, dict[str, Any]], dict[str, Any]]
RenderHandler = Callable[[Any, dict[str, Any]], dict[str, Any]]
ShowHandler = Callable[[Presentation, dict[str, Any]], dict[str, Any]]

READ_OPS: dict[ReadOp, ReadHandler] = {}
EDIT_OPS: dict[EditOp, EditHandler] = {}
RENDER_OPS: dict[RenderOp, RenderHandler] = {}
SHOW_OPS: dict[ShowOp, ShowHandler] = {}


def read_op(op: ReadOp) -> Callable[[ReadHandler], ReadHandler]:
    def reg(fn: ReadHandler) -> ReadHandler:
        READ_OPS[op] = fn
        return fn

    return reg


def edit_op(op: EditOp) -> Callable[[EditHandler], EditHandler]:
    def reg(fn: EditHandler) -> EditHandler:
        EDIT_OPS[op] = fn
        return fn

    return reg


def render_op(op: RenderOp) -> Callable[[RenderHandler], RenderHandler]:
    def reg(fn: RenderHandler) -> RenderHandler:
        RENDER_OPS[op] = fn
        return fn

    return reg


def show_op(op: ShowOp) -> Callable[[ShowHandler], ShowHandler]:
    def reg(fn: ShowHandler) -> ShowHandler:
        SHOW_OPS[op] = fn
        return fn

    return reg


# ===========================================================================
# Read ops (never move the view).
# ===========================================================================


@read_op(ReadOp.STATUS)
def _read_status(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return {"decks": ppt.presentations.list(), "viewed_slide": ppt.viewed_slide_index()}


@read_op(ReadOp.SLIDES)
def _read_slides(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return {"slides": _pick_deck(ppt, p.get("doc")).slides.list()}


@read_op(ReadOp.OUTLINE)
def _read_outline(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return {"outline": _pick_deck(ppt, p.get("doc")).outline()}


@read_op(ReadOp.LAYOUTS)
def _read_layouts(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return {"layouts": _pick_deck(ppt, p.get("doc")).layouts()}


@read_op(ReadOp.SELECTION)
def _read_selection(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return _pick_deck(ppt, p.get("doc")).selection().to_dict()


@read_op(ReadOp.FIND)
def _read_find(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("text") is not None, "read op='find' requires `text`")
    deck = _pick_deck(ppt, p.get("doc"))
    matches = deck.find(p["text"], scope=p.get("scope"))
    return {"count": len(matches), "matches": matches}


@read_op(ReadOp.SLIDE)
def _read_slide(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "read op='slide' requires `slide`")
    return _pick_deck(ppt, p.get("doc")).slides[p["slide"]].read()


@read_op(ReadOp.GEOMETRY)
def _read_geometry(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "read op='geometry' requires `slide`")
    return _pick_deck(ppt, p.get("doc")).slides[p["slide"]].geometry_report()


@read_op(ReadOp.ANIMATIONS)
def _read_animations(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "read op='animations' requires `slide`")
    deck = _pick_deck(ppt, p.get("doc"))
    return {"slide": p["slide"], "animations": deck.slides[p["slide"]].animations()}


@read_op(ReadOp.SECTIONS)
def _read_sections(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return {"sections": _pick_deck(ppt, p.get("doc")).sections.list()}


def _headers_footers_for(deck: Presentation, slide: int | None) -> Any:
    """The slide's HeadersFooters (when `slide` given) or the master's default."""
    if slide is not None:
        return deck.slides[slide].headers_footers
    return deck.master.headers_footers


@read_op(ReadOp.HEADERS_FOOTERS)
def _read_headers_footers(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    scope = "slide" if p.get("slide") is not None else "master"
    hf = _headers_footers_for(deck, p.get("slide"))
    return {"scope": scope, "slide": p.get("slide"), "headers_footers": hf.read()}


@read_op(ReadOp.ANCHOR)
def _read_anchor(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("anchor_id") is not None, "read op='anchor' requires `anchor_id`")
    anchor = _pick_deck(ppt, p.get("doc")).anchor_by_id(p["anchor_id"])
    payload: dict[str, Any] = {
        "anchor_id": anchor.anchor_id,
        "kind": anchor.kind,
        "text": anchor.text,
    }
    paragraphs = getattr(anchor, "paragraphs", None)
    if paragraphs is not None:
        payload["paragraphs"] = paragraphs.list()
    return payload


@read_op(ReadOp.TEXT_FRAME_STATUS)
def _read_text_frame_status(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(_pick_deck(ppt, p.get("doc")), p.get("anchor_id"))
    return {"anchor_id": sh.anchor_id, **sh.text_frame_status().to_dict()}


@read_op(ReadOp.TABLE)
def _read_table(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return _resolve_shape(_pick_deck(ppt, p.get("doc")), p.get("anchor_id")).table.read()


@read_op(ReadOp.CHART)
def _read_chart(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return _resolve_shape(_pick_deck(ppt, p.get("doc")), p.get("anchor_id")).chart.read()


@read_op(ReadOp.SMARTART)
def _read_smartart(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return _resolve_shape(_pick_deck(ppt, p.get("doc")), p.get("anchor_id")).smartart.read()


@read_op(ReadOp.COMMENTS)
def _read_comments(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    if p.get("slide") is not None:
        return {"slide": p["slide"], "comments": deck.slides[p["slide"]].comments.list()}
    return deck.comments()


@read_op(ReadOp.THEME)
def _read_theme(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return _pick_deck(ppt, p.get("doc")).theme.read()


@read_op(ReadOp.MASTER)
def _read_master(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    return _pick_deck(ppt, p.get("doc")).master.read()


def _read_core(ppt: Any, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Read-only dispatch (never moves the view)."""
    try:
        key = ReadOp(op)
    except ValueError as exc:
        raise BatchOpError(f"invalid_args: unknown read op {op!r}") from exc
    return READ_OPS[key](ppt, p)


# ===========================================================================
# Edit ops (each runs under the caller's open `deck.edit(...)` undo fence).
# ===========================================================================


@edit_op(EditOp.WRITE)
def _edit_write(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("anchor_id") is not None, "edit op='write' requires `anchor_id`")
    _require(p.get("text") is not None, "edit op='write' requires `text`")
    anchor = deck.anchor_by_id(p["anchor_id"])
    mode = p.get("mode") or "set"
    if mode == "set":
        anchor.set_text(p["text"])
    elif mode == "insert_after":
        anchor.insert_paragraph_after(p["text"])
    elif mode == "insert_before":
        anchor.insert_paragraph_before(p["text"])
    else:
        raise BatchOpError(f"invalid_args: unknown write mode {mode!r}")
    return {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind, "mode": mode}


@edit_op(EditOp.SET_PARAGRAPHS)
def _edit_set_paragraphs(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("anchor_id") is not None, "edit op='set_paragraphs' requires `anchor_id`")
    paragraphs = p.get("paragraphs")
    _require(
        isinstance(paragraphs, list) and len(paragraphs) > 0,
        "edit op='set_paragraphs' requires a non-empty `paragraphs` list",
    )
    assert isinstance(paragraphs, list)  # narrowed by _require above
    anchor = deck.anchor_by_id(p["anchor_id"])
    new_ids = anchor.set_paragraphs(paragraphs)
    return {"ok": True, "anchor_id": anchor.anchor_id, "paragraphs": new_ids}


@edit_op(EditOp.TEXT_RESET_FORMAT)
def _edit_text_reset_format(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("anchor_id") is not None, "edit op='text_reset_format' requires `anchor_id`")
    anchor = deck.anchor_by_id(p["anchor_id"])
    anchor.reset_format()
    return {"ok": True, "anchor_id": anchor.anchor_id}


@edit_op(EditOp.FIND_REPLACE)
def _edit_find_replace(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("find") is not None, "edit op='find_replace' requires `find`")
    _require(
        p.get("text") is not None,
        "edit op='find_replace' requires `text` (the replacement)",
    )
    applied = deck.find_replace(
        p["find"],
        p["text"],
        scope=p.get("scope"),
        all=bool(p.get("replace_all")),
        occurrence=p.get("occurrence"),
    )
    return {"ok": True, "count": len(applied), "replacements": applied}


@edit_op(EditOp.FORMAT)
def _edit_format(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("anchor_id") is not None, "edit op='format' requires `anchor_id`")
    font_opts = ("bold", "italic", "underline", "size", "font", "color")
    para_opts = (
        "alignment",
        "space_before",
        "space_after",
        "space_before_lines",
        "space_after_lines",
        "line_spacing",
        "line_spacing_points",
        "indent_level",
    )
    fill_opts = ("fill_color", "line_color", "line_width", "fill_transparency", "line_transparency")
    list_type = p.get("list_type")
    _require(
        any(p.get(k) is not None for k in font_opts + para_opts + fill_opts)
        or list_type is not None,
        "edit op='format' needs at least one font, paragraph, fill, or list option",
    )
    anchor = deck.anchor_by_id(p["anchor_id"])
    if any(p.get(k) is not None for k in font_opts):
        anchor.format_text(
            bold=p.get("bold"),
            italic=p.get("italic"),
            underline=p.get("underline"),
            size=p.get("size"),
            font=p.get("font"),
            color=p.get("color"),
        )
    if any(p.get(k) is not None for k in para_opts):
        anchor.format_paragraph(
            alignment=p.get("alignment"),
            space_before=p.get("space_before"),
            space_after=p.get("space_after"),
            space_before_lines=p.get("space_before_lines"),
            space_after_lines=p.get("space_after_lines"),
            line_spacing=p.get("line_spacing"),
            line_spacing_points=p.get("line_spacing_points"),
            indent_level=p.get("indent_level"),
            force=bool(p.get("force", False)),
        )
    if any(p.get(k) is not None for k in fill_opts):
        if not isinstance(anchor, Shape):
            raise BatchOpError(
                "invalid_args: fill_color/line_color/line_width need a shape anchor "
                f"(shape:/shapeid:/ph:), got {p['anchor_id']!r}"
            )
        anchor.set_fill(
            fill=p.get("fill_color"),
            line=p.get("line_color"),
            line_width=p.get("line_width"),
            fill_transparency=p.get("fill_transparency"),
            line_transparency=p.get("line_transparency"),
        )
    if list_type == "none":
        anchor.remove_list()
    elif list_type is not None:
        anchor.apply_list(list_type, character=p.get("bullet_char"))
    result: dict[str, Any] = {"ok": True, "anchor_id": anchor.anchor_id}
    warnings = _format_warnings(anchor, p)
    if warnings:
        result["warnings"] = warnings
    return result


def _format_warnings(anchor: Any, p: dict[str, Any]) -> list[str]:
    """Non-fatal advisories for a `format` edit — flag suspicious-but-applied inputs.

    The structured-I/O contract has room for a `warnings` array (the gpt-5.4
    review's "even just returning warnings would help"): a forced large line-spacing
    multiple, a tiny font, or a list applied to a single soft-break paragraph (where
    each line won't become its own bullet — `set_paragraphs` is the fix).
    """
    warnings: list[str] = []
    size = p.get("size")
    if size is not None and float(size) < 8:
        warnings.append(f"font size {size}pt is very small (<8pt) and may be unreadable")
    line_spacing = p.get("line_spacing")
    if line_spacing is not None and float(line_spacing) > LINE_SPACING_MULTIPLE_MAX:
        warnings.append(
            f"line_spacing={line_spacing} is a large multiple "
            f"({line_spacing}x line height); did you mean line_spacing_points={line_spacing}?"
        )
    if p.get("list_type") in ("bulleted", "numbered"):
        try:
            if anchor.paragraph_count() == 1 and SOFT_BREAK in anchor.text:
                warnings.append(
                    "a list was applied to a single paragraph with soft line breaks — "
                    "each line won't be its own bullet; use op='set_paragraphs' instead"
                )
        except Exception:
            pass
    return warnings


@edit_op(EditOp.SLIDE_ADD)
def _edit_slide_add(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    placeholders = p.get("placeholders")
    try:
        new = deck.slides.add(
            layout=p.get("layout"), index=p.get("index"), placeholders=placeholders
        )
    except ValueError as exc:
        raise BatchOpError(f"invalid_args: {exc}") from exc
    result: dict[str, Any] = {
        "ok": True,
        "index": new.index,
        "id": new.id,
        "layout": new.layout_name,
    }
    if placeholders:
        # Echo the placeholders' resulting geometry so the agent can confirm the fit.
        result["placeholders"] = {
            kind: _resolve_shape(deck, f"ph:{new.index}:{kind}").geometry() for kind in placeholders
        }
    return result


@edit_op(EditOp.SLIDE_DELETE)
def _edit_slide_delete(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='slide_delete' requires `slide`")
    deck.slides[p["slide"]].delete()
    return {"ok": True, "deleted": p["slide"]}


@edit_op(EditOp.SLIDE_DUPLICATE)
def _edit_slide_duplicate(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='slide_duplicate' requires `slide`")
    new = deck.slides[p["slide"]].duplicate()
    return {"ok": True, "index": new.index, "id": new.id, "from": p["slide"]}


@edit_op(EditOp.SLIDE_MOVE)
def _edit_slide_move(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='slide_move' requires `slide`")
    _require(p.get("to") is not None, "edit op='slide_move' requires `to`")
    moved = deck.slides[p["slide"]].move_to(p["to"])
    return {"ok": True, "index": moved.index, "id": moved.id}


@edit_op(EditOp.SET_LAYOUT)
def _edit_set_layout(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='set_layout' requires `slide`")
    _require(p.get("layout") is not None, "edit op='set_layout' requires `layout`")
    target = deck.slides[p["slide"]]
    target.set_layout(p["layout"])
    return {"ok": True, "index": p["slide"], "layout": target.layout_name}


@edit_op(EditOp.SHAPE_ADD)
def _edit_shape_add(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='shape_add' requires `slide`")
    kind = p.get("kind")
    _require(kind is not None, "edit op='shape_add' requires `kind`")
    geom = {k: p.get(k) for k in ("left", "top", "width", "height")}
    shapes = deck.slides[p["slide"]].shapes
    fill_kw = {
        "fill": p.get("fill_color"),
        "line": p.get("line_color"),
        "line_width": p.get("line_width"),
    }
    if kind == "textbox":
        created = shapes.add_textbox(p.get("text") or "", **fill_kw, **geom)
    elif kind == "shape":
        created = shapes.add_shape(p.get("shape_type") or "rectangle", **fill_kw, **geom)
        if p.get("text"):
            created.set_text(p["text"])
    elif kind == "table":
        _require(
            p.get("rows") is not None and p.get("cols") is not None,
            "edit shape_add kind='table' requires `rows` and `cols`",
        )
        created = shapes.add_table(p["rows"], p["cols"], **geom)
    elif kind == "chart":
        created = shapes.add_chart(
            p.get("chart_type") or "column",
            p.get("categories"),
            p.get("series"),
            **geom,
        )
    elif kind == "smartart":
        created = shapes.add_smartart(p.get("smartart_kind") or "process", p.get("nodes"), **geom)
    elif kind == "picture":
        _require(p.get("path") is not None, "edit shape_add kind='picture' requires `path`")
        created = shapes.add_picture(p["path"], alt_text=p.get("alt_text"), **geom)
    else:
        raise BatchOpError(f"invalid_args: unknown shape kind {kind!r}")
    return {"ok": True, **created.to_dict()}


@edit_op(EditOp.SHAPE_MOVE)
def _edit_shape_move(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        p.get("left") is not None or p.get("top") is not None,
        "edit op='shape_move' requires `left`/`top`",
    )
    sh.move(left=p.get("left"), top=p.get("top"))
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "geometry": sh.geometry()}


@edit_op(EditOp.SHAPE_RESIZE)
def _edit_shape_resize(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        p.get("width") is not None or p.get("height") is not None,
        "edit op='shape_resize' requires `width`/`height`",
    )
    sh.resize(width=p.get("width"), height=p.get("height"))
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "geometry": sh.geometry()}


@edit_op(EditOp.SET_ALT)
def _edit_set_alt(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(p.get("alt_text") is not None, "edit op='set_alt' requires `alt_text`")
    sh.set_alt_text(p["alt_text"])
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "alt_text": sh.alt_text}


@edit_op(EditOp.SHAPE_DELETE)
def _edit_shape_delete(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    info = {"anchor_id": sh.anchor_id, "name": sh.name, "id": sh.shape_id}
    sh.delete()
    return {"ok": True, **info}


@edit_op(EditOp.SHAPE_ORDER)
def _edit_shape_order(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        p.get("order") is not None,
        "edit op='shape_order' requires `order` (front/back/forward/backward)",
    )
    new_index = sh.reorder(p["order"])
    return {
        "ok": True,
        "anchor_id": sh.anchor_id,
        "shapeid": sh.shapeid,
        "name": sh.name,
        "index": new_index,
    }


@edit_op(EditOp.SHAPE_ANIMATE)
def _edit_shape_animate(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(p.get("effect") is not None, "edit op='shape_animate' requires `effect`")
    animation = sh.animate(
        p["effect"],
        trigger=p.get("trigger") or "on_click",
        duration=p.get("duration"),
        delay=p.get("delay"),
        exit=bool(p.get("exit", False)),
    )
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "animation": animation}


@edit_op(EditOp.SHAPE_CLEAR_ANIMATIONS)
def _edit_shape_clear_animations(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    removed = sh.clear_animations()
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "removed": removed}


@edit_op(EditOp.SLIDE_CLEAR_ANIMATIONS)
def _edit_slide_clear_animations(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='slide_clear_animations' requires `slide`")
    removed = deck.slides[p["slide"]].clear_animations()
    return {"ok": True, "index": p["slide"], "removed": removed}


@edit_op(EditOp.SHAPE_RESET_LAYOUT)
def _edit_shape_reset_layout(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    restored = sh.reset_to_layout()
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "restored": restored}


@edit_op(EditOp.SHAPE_GRADIENT_FILL)
def _edit_shape_gradient_fill(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        p.get("colors") is not None or p.get("preset") is not None,
        "edit op='shape_gradient_fill' requires `colors` or `preset`",
    )
    sh.set_gradient_fill(
        p.get("colors"),
        positions=p.get("positions"),
        style=p.get("gradient_style") or "horizontal",
        variant=p.get("variant") or 1,
        degree=p.get("degree"),
        preset=p.get("preset"),
    )
    return {
        "ok": True,
        "anchor_id": sh.anchor_id,
        "shapeid": sh.shapeid,
        "fill": sh.to_dict().get("fill"),
    }


@edit_op(EditOp.SHAPE_PICTURE_FILL)
def _edit_shape_picture_fill(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(p.get("path") is not None, "edit op='shape_picture_fill' requires `path`")
    sh.set_picture_fill(p["path"])
    return {
        "ok": True,
        "anchor_id": sh.anchor_id,
        "shapeid": sh.shapeid,
        "fill": sh.to_dict().get("fill"),
    }


@edit_op(EditOp.SHAPE_SET_PICTURE)
def _edit_shape_set_picture(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(p.get("path") is not None, "edit op='shape_set_picture' requires `path`")
    new = sh.set_picture(p["path"], alt_text=p.get("alt_text"))
    return {
        "ok": True,
        "anchor_id": new.anchor_id,
        "shapeid": new.shapeid,
        "geometry": new.geometry(),
    }


@edit_op(EditOp.SHAPE_PATTERN_FILL)
def _edit_shape_pattern_fill(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(p.get("pattern") is not None, "edit op='shape_pattern_fill' requires `pattern`")
    _require(p.get("fore") is not None, "edit op='shape_pattern_fill' requires `fore` color")
    sh.set_pattern_fill(p["pattern"], fore=p["fore"], back=p.get("back"))
    return {
        "ok": True,
        "anchor_id": sh.anchor_id,
        "shapeid": sh.shapeid,
        "fill": sh.to_dict().get("fill"),
    }


@edit_op(EditOp.SHAPE_SET_EFFECT)
def _edit_shape_set_effect(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        any(p.get(k) is not None for k in ("shadow", "glow", "soft_edge", "reflection")),
        "edit op='shape_set_effect' needs at least one of shadow / glow / soft_edge / reflection",
    )
    sh.set_effect(
        shadow=p.get("shadow"),
        glow=p.get("glow"),
        soft_edge=p.get("soft_edge"),
        reflection=p.get("reflection"),
    )
    return {
        "ok": True,
        "anchor_id": sh.anchor_id,
        "shapeid": sh.shapeid,
        "effects": sh.to_dict().get("effects"),
    }


@edit_op(EditOp.SHAPE_LINE_STYLE)
def _edit_shape_line_style(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        any(
            p.get(k) is not None
            for k in ("dash", "begin_arrow", "end_arrow", "begin_arrow_size", "end_arrow_size")
        ),
        "edit op='shape_line_style' needs at least one of dash / begin_arrow / end_arrow / "
        "begin_arrow_size / end_arrow_size",
    )
    sh.set_line_style(
        dash=p.get("dash"),
        begin_arrow=p.get("begin_arrow"),
        end_arrow=p.get("end_arrow"),
        begin_arrow_size=p.get("begin_arrow_size"),
        end_arrow_size=p.get("end_arrow_size"),
    )
    return {
        "ok": True,
        "anchor_id": sh.anchor_id,
        "shapeid": sh.shapeid,
        "line": sh.to_dict().get("line"),
    }


@edit_op(EditOp.SHAPE_SET_HYPERLINK)
def _edit_shape_set_hyperlink(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        (p.get("url") is None) != (p.get("slide") is None),
        "edit op='shape_set_hyperlink' requires exactly one of `url` or `slide`",
    )
    link = sh.set_hyperlink(url=p.get("url"), slide=p.get("slide"), screen_tip=p.get("screen_tip"))
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "hyperlink": link}


@edit_op(EditOp.SHAPE_REMOVE_HYPERLINK)
def _edit_shape_remove_hyperlink(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    sh.remove_hyperlink()
    return {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "hyperlink": None}


@edit_op(EditOp.SLIDE_SET_TRANSITION)
def _edit_slide_set_transition(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='slide_set_transition' requires `slide`")
    _require(
        any(
            p.get(k) is not None
            for k in ("effect", "duration", "advance_after", "advance_on_click")
        ),
        "edit op='slide_set_transition' needs at least one of effect / duration / "
        "advance_after / advance_on_click",
    )
    target = deck.slides[p["slide"]]
    trans = target.set_transition(
        p.get("effect"),
        duration=p.get("duration"),
        advance_after=p.get("advance_after"),
        advance_on_click=p.get("advance_on_click"),
    )
    return {"ok": True, "index": p["slide"], "transition": trans}


@edit_op(EditOp.SLIDE_SET_BACKGROUND)
def _edit_slide_set_background(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='slide_set_background' requires `slide`")
    follow_master = bool(p.get("follow_master", False))
    _require(
        (p.get("color") is not None) != follow_master,
        "edit op='slide_set_background' requires exactly one of `color` or `follow_master`",
    )
    target = deck.slides[p["slide"]]
    bg = target.follow_master_background() if follow_master else target.set_background(p["color"])
    return {"ok": True, "index": p["slide"], "background": bg}


@edit_op(EditOp.SECTION_ADD)
def _edit_section_add(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("name") is not None, "edit op='section_add' requires `name`")
    row = deck.sections.add(p["name"], before_slide=p.get("before_slide"))
    return {"ok": True, "section": row}


@edit_op(EditOp.SECTION_RENAME)
def _edit_section_rename(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("section") is not None, "edit op='section_rename' requires `section`")
    _require(p.get("name") is not None, "edit op='section_rename' requires `name`")
    row = deck.sections.rename(p["section"], p["name"])
    return {"ok": True, "section": row}


@edit_op(EditOp.SECTION_DELETE)
def _edit_section_delete(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("section") is not None, "edit op='section_delete' requires `section`")
    return {
        "ok": True,
        **deck.sections.delete(p["section"], delete_slides=bool(p.get("delete_slides", False))),
    }


@edit_op(EditOp.SECTION_MOVE)
def _edit_section_move(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("section") is not None, "edit op='section_move' requires `section`")
    _require(p.get("to") is not None, "edit op='section_move' requires `to`")
    row = deck.sections.move(p["section"], p["to"])
    return {"ok": True, "section": row}


@edit_op(EditOp.SET_HEADERS_FOOTERS)
def _edit_set_headers_footers(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    footer_keys = ("footer_text", "footer_visible")
    date_keys = ("date_visible", "date_text", "date_format")
    _require(
        any(p.get(k) is not None for k in (*footer_keys, "slide_number_visible", *date_keys)),
        "edit op='set_headers_footers' needs at least one of footer_text / footer_visible / "
        "slide_number_visible / date_visible / date_text / date_format",
    )
    scope = "slide" if p.get("slide") is not None else "master"
    hf = _headers_footers_for(deck, p.get("slide"))
    if any(p.get(k) is not None for k in footer_keys):
        hf.set_footer(text=p.get("footer_text"), visible=p.get("footer_visible"))
    if p.get("slide_number_visible") is not None:
        hf.set_slide_number(bool(p["slide_number_visible"]))
    if any(p.get(k) is not None for k in date_keys):
        hf.set_date(
            visible=p.get("date_visible"), text=p.get("date_text"), fmt=p.get("date_format")
        )
    return {"ok": True, "scope": scope, "slide": p.get("slide"), "headers_footers": hf.read()}


@edit_op(EditOp.TABLE_ADD_ROW)
def _edit_table_add_row(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    table = _resolve_shape(deck, p.get("anchor_id")).table
    table.add_row(p.get("values"))
    return {"ok": True, "anchor_id": table.shape.anchor_id, "rows": table.row_count}


@edit_op(EditOp.TABLE_DELETE_ROW)
def _edit_table_delete_row(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    table = _resolve_shape(deck, p.get("anchor_id")).table
    _require(p.get("row") is not None, "edit op='table_delete_row' requires `row`")
    table.delete_row(p["row"])
    return {"ok": True, "anchor_id": table.shape.anchor_id, "rows": table.row_count}


@edit_op(EditOp.TABLE_ADD_COLUMN)
def _edit_table_add_column(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    table = _resolve_shape(deck, p.get("anchor_id")).table
    table.add_column(p.get("values"), before=p.get("before"))
    return {"ok": True, "anchor_id": table.shape.anchor_id, "columns": table.column_count}


@edit_op(EditOp.TABLE_DELETE_COLUMN)
def _edit_table_delete_column(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    table = _resolve_shape(deck, p.get("anchor_id")).table
    _require(p.get("column") is not None, "edit op='table_delete_column' requires `column`")
    table.delete_column(p["column"])
    return {"ok": True, "anchor_id": table.shape.anchor_id, "columns": table.column_count}


@edit_op(EditOp.TABLE_SET_FILL)
def _edit_table_set_fill(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    table = _resolve_shape(deck, p.get("anchor_id")).table
    _require(p.get("fill") is not None, "edit op='table_set_fill' requires `fill`")
    n = table.set_fill(
        p["fill"],
        rows=p.get("rows"),
        cols=p.get("cols"),
        transparency=p.get("fill_transparency"),
    )
    return {"ok": True, "anchor_id": table.shape.anchor_id, "cells": n}


@edit_op(EditOp.TABLE_SET_BORDER)
def _edit_table_set_border(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    table = _resolve_shape(deck, p.get("anchor_id")).table
    _require(
        any(p.get(k) is not None for k in ("color", "weight", "dash", "visible")),
        "edit op='table_set_border' requires at least one of `color`, `weight`, `dash`, `visible`",
    )
    n = table.set_border(
        color=p.get("color"),
        weight=p.get("weight"),
        dash=p.get("dash"),
        edges=p.get("edges") or "all",
        rows=p.get("rows"),
        cols=p.get("cols"),
        visible=p.get("visible"),
    )
    return {"ok": True, "anchor_id": table.shape.anchor_id, "cells": n}


@edit_op(EditOp.CHART_SET_TYPE)
def _edit_chart_set_type(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    chart = _resolve_shape(deck, p.get("anchor_id")).chart
    _require(p.get("chart_type") is not None, "edit op='chart_set_type' requires `chart_type`")
    chart.set_type(p["chart_type"])
    return {"ok": True, "anchor_id": chart.shape.anchor_id, "chart_type": chart.chart_type}


@edit_op(EditOp.CHART_SET_DATA)
def _edit_chart_set_data(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    chart = _resolve_shape(deck, p.get("anchor_id")).chart
    _require(
        p.get("categories") is not None and p.get("series") is not None,
        "edit op='chart_set_data' requires `categories` and `series`",
    )
    chart.set_data(p["categories"], p["series"])
    return chart.read()


@edit_op(EditOp.CHART_RECOLOR_TEXT)
def _edit_chart_recolor_text(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    chart = _resolve_shape(deck, p.get("anchor_id")).chart
    _require(p.get("color") is not None, "edit op='chart_recolor_text' requires `color`")
    return chart.recolor_text(p["color"])


@edit_op(EditOp.SMARTART_SET_NODES)
def _edit_smartart_set_nodes(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sa = _resolve_shape(deck, p.get("anchor_id")).smartart
    _require(p.get("nodes") is not None, "edit op='smartart_set_nodes' requires `nodes`")
    sa.set_nodes(p["nodes"])
    return sa.read()


@edit_op(EditOp.SMARTART_RECOLOR_TEXT)
def _edit_smartart_recolor_text(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sa = _resolve_shape(deck, p.get("anchor_id")).smartart
    _require(p.get("color") is not None, "edit op='smartart_recolor_text' requires `color`")
    return sa.recolor_text(p["color"])


@edit_op(EditOp.SMARTART_FORMAT_NODE)
def _edit_smartart_format_node(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sa = _resolve_shape(deck, p.get("anchor_id")).smartart
    _require(
        p.get("node_index") is not None, "edit op='smartart_format_node' requires `node_index`"
    )
    return sa.format_node(
        p["node_index"],
        bold=p.get("bold"),
        italic=p.get("italic"),
        underline=p.get("underline"),
        size=p.get("size"),
        font=p.get("font"),
        color=p.get("color"),
    )


@edit_op(EditOp.COMMENT_ADD)
def _edit_comment_add(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='comment_add' requires `slide`")
    _require(p.get("text") is not None, "edit op='comment_add' requires `text`")
    kwargs: dict[str, Any] = {"author": p.get("author"), "initials": p.get("initials")}
    if p.get("left") is not None:
        kwargs["left"] = p["left"]
    if p.get("top") is not None:
        kwargs["top"] = p["top"]
    c = deck.slides[p["slide"]].comments.add(p["text"], **kwargs)
    return {"ok": True, "slide": p["slide"], "comment": c.to_dict()}


@edit_op(EditOp.COMMENT_REPLY)
def _edit_comment_reply(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='comment_reply' requires `slide`")
    _require(p.get("index") is not None, "edit op='comment_reply' requires `index`")
    _require(p.get("text") is not None, "edit op='comment_reply' requires `text`")
    rep = deck.slides[p["slide"]].comments[p["index"]].reply(p["text"])
    return {"ok": True, "slide": p["slide"], "parent": p["index"], "reply": rep.to_dict()}


@edit_op(EditOp.COMMENT_DELETE)
def _edit_comment_delete(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "edit op='comment_delete' requires `slide`")
    _require(p.get("index") is not None, "edit op='comment_delete' requires `index`")
    deck.slides[p["slide"]].comments[p["index"]].delete()
    return {"ok": True, "slide": p["slide"], "index": p["index"]}


@edit_op(EditOp.THEME_SET_COLOR)
def _edit_theme_set_color(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slot") is not None, "edit op='theme_set_color' requires `slot`")
    _require(p.get("color") is not None, "edit op='theme_set_color' requires `color`")
    deck.theme.set_color(p["slot"], p["color"])
    return deck.theme.read()


@edit_op(EditOp.THEME_SET_FONT)
def _edit_theme_set_font(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("which") is not None, "edit op='theme_set_font' requires `which`")
    _require(p.get("name") is not None, "edit op='theme_set_font' requires `name`")
    deck.theme.set_font(p["which"], p["name"], script=p.get("script") or "latin")
    return deck.theme.read()


@edit_op(EditOp.MASTER_FORMAT_TEXT_STYLE)
def _edit_master_format_text_style(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("style") is not None, "edit op='master_format_text_style' requires `style`")
    level = 1 if p.get("level") is None else p["level"]  # default outline level
    style_font_opts = ("bold", "italic", "underline", "size", "font", "color")
    _require(
        any(p.get(k) is not None for k in style_font_opts),
        "edit op='master_format_text_style' needs at least one font option",
    )
    deck.master.format_text_style(
        p["style"],
        level,
        bold=p.get("bold"),
        italic=p.get("italic"),
        underline=p.get("underline"),
        size=p.get("size"),
        font=p.get("font"),
        color=p.get("color"),
    )
    return {"ok": True, "style": p["style"], "level": level}


@edit_op(EditOp.MASTER_FORMAT_PARAGRAPH_STYLE)
def _edit_master_format_paragraph_style(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("style") is not None, "edit op='master_format_paragraph_style' requires `style`")
    level = 1 if p.get("level") is None else p["level"]  # default outline level
    style_para_opts = ("alignment", "space_before", "space_after", "line_spacing")
    _require(
        any(p.get(k) is not None for k in style_para_opts),
        "edit op='master_format_paragraph_style' needs at least one paragraph option",
    )
    deck.master.format_paragraph_style(
        p["style"],
        level,
        alignment=p.get("alignment"),
        space_before=p.get("space_before"),
        space_after=p.get("space_after"),
        line_spacing=p.get("line_spacing"),
    )
    return {"ok": True, "style": p["style"], "level": level}


@edit_op(EditOp.MASTER_SET_BACKGROUND)
def _edit_master_set_background(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("color") is not None, "edit op='master_set_background' requires `color`")
    deck.master.set_background(p["color"])
    return {"ok": True, "background": deck.master.read().get("background", {})}


def _edit_core(deck: Presentation, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Mutation dispatch. The caller MUST have an open `deck.edit(...)` scope."""
    try:
        key = EditOp(op)
    except ValueError as exc:
        raise BatchOpError(f"invalid_args: unknown edit op {op!r}") from exc
    return EDIT_OPS[key](deck, p)


#: Batch commands that deliberately move what the user sees — a `render navigate`
#: or any `show` control verb (everything but the read-only `state`). When one of
#: these runs inside an atomic batch, the scope must NOT snap the view back.
_MOVING_SHOW_OPS = frozenset(
    {"start", "end", "next", "previous", "goto", "black", "white", "resume"}
)


def _moves_view(tool: str, op: str | None) -> bool:
    """True if a batch command intentionally changes the viewed slide / screen."""
    return (tool == "render" and op == "navigate") or (tool == "show" and op in _MOVING_SHOW_OPS)


# --- "Follow the work" view policy -----------------------------------------
# When an atomic batch ADDS a slide, the polite snap-back to the pre-batch view
# is the wrong default: the user is watching the agent author, and wants to end
# up looking at the slide that was just built — not get yanked back to slide 1
# every batch. So an *authoring* batch (one that ran a slide_add/slide_duplicate)
# leaves the view on the last slide it touched instead of restoring. Pure-edit
# batches keep the polite restore. Toggle off with `PPTLIVE_VIEW_FOLLOW=0` (env)
# or the per-call `follow_view` flag.

#: Edit ops that mark a batch as "authoring" — their presence enables view-follow.
_VIEW_FOLLOW_ADD_OPS = frozenset({"slide_add", "slide_duplicate"})

#: Edit ops whose target slide should NOT become the follow focus (deletions —
#: you don't want the view to chase a slide/shape that no longer exists, and the
#: index has shifted anyway).
_VIEW_FOLLOW_SKIP_OPS = frozenset({"slide_delete", "shape_delete", "comment_delete"})

#: Edit ops whose result `index` is a *slide* index (vs. shape_add, whose
#: `to_dict()` carries a z-order `index` that must not be read as a slide).
_SLIDE_INDEX_RESULT_OPS = frozenset({"slide_add", "slide_duplicate", "slide_move", "set_layout"})


def _env_view_follow_default() -> bool:
    """The default view-follow setting from `PPTLIVE_VIEW_FOLLOW` (on unless disabled)."""
    val = os.environ.get("PPTLIVE_VIEW_FOLLOW")
    if val is None:
        return True
    return val.strip().lower() not in ("0", "false", "no", "off", "")


def _anchor_slide_index(anchor_id: Any) -> int | None:
    """The leading 1-based slide index of a hierarchical anchor id, or None.

    Every text/shape anchor is slide-first (`shape:S:N`, `ph:S:KIND`, `para:S:N:P`,
    `cell:S:N:R:C`, `notes:S`, `shapeid:S:ID`, `comments:S`), so the slide is the
    field right after the prefix.
    """
    if not isinstance(anchor_id, str):
        return None
    parts = anchor_id.split(":")
    if len(parts) >= 2 and parts[0] in (
        "shape",
        "shapeid",
        "ph",
        "para",
        "cell",
        "notes",
        "comments",
    ):
        try:
            return int(parts[1])
        except ValueError:
            return None
    return None


def _edit_focus_slide(op: str, p: dict[str, Any], result: dict[str, Any]) -> int | None:
    """The slide an edit op 'landed on', for view-follow — None = leave focus as-is.

    Prefers a slide-returning op's result `index`, then an explicit `slide` param,
    then the slide parsed off an `anchor_id`. Deletions are skipped so the view
    never chases a just-removed slide/shape.
    """
    if op in _VIEW_FOLLOW_SKIP_OPS:
        return None
    if op in _SLIDE_INDEX_RESULT_OPS:
        idx = result.get("index")
        return idx if isinstance(idx, int) else None
    slide = p.get("slide")
    if isinstance(slide, int):
        return slide
    return _anchor_slide_index(p.get("anchor_id"))


def _parse_slide_selector(slides: Any) -> int | tuple[int, int] | None:
    """Parse a `deck_snapshot` `slides` arg into a `snapshot()` selector.

    Accepts `None` (all slides), an `int` (single 1-based slide), or a string —
    `"3"` (single) or `"2-4"` (inclusive span). A malformed string is an
    `invalid_args` ToolError.
    """
    if slides is None:
        return None
    if isinstance(slides, bool):
        raise BatchOpError("invalid_args: `slides` must be an int or a span string like '2-4'")
    if isinstance(slides, int):
        return slides
    s = str(slides).strip()
    try:
        if "-" in s:
            a, _, b = s.partition("-")
            return int(a), int(b)
        return int(s)
    except ValueError as e:
        raise BatchOpError(
            f"invalid_args: `slides` must be an int or a span like '2-4', got {slides!r}"
        ) from e


# ===========================================================================
# Render ops — render-to-image + the one deliberate view move (`navigate`).
# Each picks its own deck; `fmt` defaults to PNG (save_as overrides with its own
# file format).
# ===========================================================================


@render_op(RenderOp.SLIDE_IMAGE)
def _render_slide_image(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    fmt = p.get("fmt") or "png"
    _require(p.get("slide") is not None, "render op='slide_image' requires `slide`")
    path = deck.slides[p["slide"]].export_image(
        p.get("out"), width=p.get("width"), height=p.get("height"), fmt=fmt
    )
    return {"ok": True, "slide": p["slide"], "path": str(path), "format": fmt}


@render_op(RenderOp.DECK_SNAPSHOT)
def _render_deck_snapshot(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    fmt = p.get("fmt") or "png"
    selector = _parse_slide_selector(p.get("slides"))
    out_dir = tempfile.mkdtemp(prefix="pptlive_snap_")
    base = os.path.join(out_dir, f"snap.{fmt}")
    snaps = deck.snapshot(
        base,
        slides=selector,
        fmt=fmt,
        max_dim=p.get("max_dim"),
        width=p.get("width"),
        height=p.get("height"),
    )
    return {
        "ok": True,
        "count": len(snaps),
        "format": fmt,
        "max_dim": p.get("max_dim"),
        "width": p.get("width"),
        "height": p.get("height"),
        "images": [{"slide": s.slide, "path": str(s.path), "format": fmt} for s in snaps],
    }


@render_op(RenderOp.SHAPE_IMAGE)
def _render_shape_image(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    fmt = p.get("fmt") or "png"
    sh = _resolve_shape(deck, p.get("anchor_id"))
    path = sh.export_image(p.get("out"), fmt=fmt)
    return {"ok": True, "anchor_id": sh.anchor_id, "path": str(path), "format": fmt}


@render_op(RenderOp.DECK_PDF)
def _render_deck_pdf(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    _require(p.get("out") is not None, "render op='deck_pdf' requires `out` (the PDF path)")
    written = deck.export_pdf(p["out"])
    # `format: "pdf"` keeps _render_reply from mis-embedding the PDF as an image.
    return {"ok": True, "path": written, "format": "pdf"}


@render_op(RenderOp.SAVE)
def _render_save(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    saved_path = _pick_deck(ppt, p.get("doc")).save()
    return {"ok": True, "path": saved_path, "saved": True, "format": "pptx"}


@render_op(RenderOp.SAVE_AS)
def _render_save_as(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    _require(p.get("out") is not None, "render op='save_as' requires `out` (the .pptx path)")
    fmt = str(p.get("save_format", "pptx"))
    try:
        written = deck.save_as(p["out"], fmt=fmt, overwrite=bool(p.get("overwrite", False)))
    except (FileExistsError, ValueError) as exc:
        raise BatchOpError(f"invalid_args: {exc}") from exc
    return {"ok": True, "path": written, "format": fmt}


@render_op(RenderOp.NAVIGATE)
def _render_navigate(ppt: Any, p: dict[str, Any]) -> dict[str, Any]:
    deck = _pick_deck(ppt, p.get("doc"))
    _require(p.get("anchor_id") is not None, "render op='navigate' requires `anchor_id`")
    anchor = deck.anchor_by_id(p["anchor_id"])
    deck.go_to(anchor, select=p.get("select", True))
    return {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind}


def _render_core(ppt: Any, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Render-to-image + the one deliberate view move (`navigate`)."""
    try:
        key = RenderOp(op)
    except ValueError as exc:
        raise BatchOpError(f"invalid_args: unknown render op {op!r}") from exc
    return RENDER_OPS[key](ppt, p)


# ===========================================================================
# Show ops — live slide-show control (deliberately drives the user's screen).
# ===========================================================================


@show_op(ShowOp.STATE)
def _show_state(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.state()


@show_op(ShowOp.START)
def _show_start(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.start(from_slide=p.get("slide"))


@show_op(ShowOp.END)
def _show_end(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.end()


@show_op(ShowOp.NEXT)
def _show_next(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.next()


@show_op(ShowOp.PREVIOUS)
def _show_previous(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.previous()


@show_op(ShowOp.GOTO)
def _show_goto(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    _require(p.get("slide") is not None, "show op='goto' requires `slide`")
    return deck.show.goto(p["slide"])


@show_op(ShowOp.BLACK)
def _show_black(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.black()


@show_op(ShowOp.WHITE)
def _show_white(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.white()


@show_op(ShowOp.RESUME)
def _show_resume(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    return deck.show.resume()


def _show_core(deck: Presentation, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Live slide-show control (deliberately drives the user's screen)."""
    try:
        key = ShowOp(op)
    except ValueError as exc:
        raise BatchOpError(f"invalid_args: unknown show op {op!r}") from exc
    return SHOW_OPS[key](deck, p)


# Drift guard: every enum member must have a registered handler. Turns a missing
# op (or an enum/registry mismatch) into an import-time failure instead of a
# silent gap the agent only discovers at call time.
assert set(READ_OPS) == set(ReadOp), f"READ_OPS missing {set(ReadOp) - set(READ_OPS)}"
assert set(EDIT_OPS) == set(EditOp), f"EDIT_OPS missing {set(EditOp) - set(EDIT_OPS)}"
assert set(RENDER_OPS) == set(RenderOp), f"RENDER_OPS missing {set(RenderOp) - set(RENDER_OPS)}"
assert set(SHOW_OPS) == set(ShowOp), f"SHOW_OPS missing {set(ShowOp) - set(SHOW_OPS)}"


# ===========================================================================
# Batch runner — shared by the MCP `ppt_batch` tool and the CLI `exec` verb.
# ===========================================================================


def run_batch(
    handle: Any,
    deck: Presentation,
    commands: list[dict[str, Any]],
    *,
    doc: str | None = None,
    atomic: bool = True,
    stop_on_error: bool = True,
    follow_view: bool | None = None,
    label: str = "batch",
) -> list[dict[str, Any]]:
    """Dispatch a list of `{tool, op, ...params}` commands against one connection.

    The fastmcp-free heart of `ppt_batch` (the MCP tool keeps image embedding on
    top; the CLI `exec` verb calls this directly). `tool` is "read"/"edit"/"render"/
    "show" (default "edit"); the rest of each dict is that op's params. With
    `atomic` the whole batch is fenced into one undo entry (one Ctrl-Z); a
    view-moving op opts the scope out of the view restore. On a per-op failure the
    entry records `{ok: False, error, message}` and, if `stop_on_error`, the run
    stops (earlier mutations stay applied — undo grouping, not a transaction).
    Returns the list of per-op result entries; the caller owns `attach()`.

    `follow_view` ("follow the work"): when an atomic batch *adds* a slide, leave
    the view on the last slide it touched rather than snapping back to the
    pre-batch view — the right default while an agent authors. `None` (default)
    takes the `PPTLIVE_VIEW_FOLLOW` env default (on); pass `True`/`False` to force
    it per call. A deliberate `navigate`/`show` op still wins (its own view move is
    respected, never overridden by follow).
    """
    follow = follow_view if follow_view is not None else _env_view_follow_default()
    has_edit = any(cmd.get("tool", "edit") == "edit" for cmd in commands)
    scope: Any = deck.edit(label) if (atomic and has_edit) else nullcontext()
    results: list[dict[str, Any]] = []
    added_slide = False  # batch ran a slide_add/slide_duplicate -> authoring
    moved_by_op = False  # a deliberate navigate/show already claimed the view
    focus_slide: int | None = None  # last slide an edit op touched
    with scope as edit_scope:
        for i, cmd in enumerate(commands):
            tool = cmd.get("tool", "edit")
            op = cmd.get("op")
            p = {k: v for k, v in cmd.items() if k not in ("tool", "op")}
            p["doc"] = doc
            entry: dict[str, Any] = {"index": i, "tool": tool, "op": op}
            try:
                _require(op is not None, f"command #{i} is missing `op`")
                assert op is not None
                if tool == "read":
                    result = _read_core(handle, op, p)
                elif tool == "edit":
                    if atomic:
                        result = _edit_core(deck, op, p)
                    else:
                        with deck.edit(f"{label} #{i} {op}"):
                            result = _edit_core(deck, op, p)
                elif tool == "render":
                    result = _render_core(handle, op, p)
                elif tool == "show":
                    result = _show_core(deck, op, p)
                else:
                    raise BatchOpError(f"command #{i} unknown tool {tool!r}")
                if _moves_view(tool, op) and edit_scope is not None:
                    edit_scope.allow_view_move()
                    moved_by_op = True
                if tool == "edit":
                    if op in _VIEW_FOLLOW_ADD_OPS:
                        added_slide = True
                    fs = _edit_focus_slide(op, p, result)
                    if fs is not None:
                        focus_slide = fs
                entry.update(ok=True, result=result)
            except (PptliveError, ValueError, FileNotFoundError) as exc:
                # Library handlers raise bare ValueError (e.g. format_paragraph
                # line_spacing > 5, empty set_paragraphs) and FileNotFoundError
                # (picture sourcing) without wrapping them in BatchOpError. Record
                # those as invalid_args per-command — mirroring the single-op MCP
                # path (_mcp_errors) — so one bad op doesn't abort the whole batch
                # or escape the per-op contract.
                code = _error_code(exc) if isinstance(exc, PptliveError) else "invalid_args"
                entry.update(ok=False, error=code, message=str(exc))
                results.append(entry)
                if stop_on_error:
                    break
                continue
            results.append(entry)
        # Follow the work: an authoring batch ends on the slide it last built,
        # unless a deliberate navigate/show already moved the view (that wins).
        if (
            follow
            and added_slide
            and not moved_by_op
            and focus_slide is not None
            and edit_scope is not None
        ):
            try:
                deck.go_to(deck.slides[focus_slide])
                edit_scope.allow_view_move()
            except Exception:
                pass
    return results

"""The pptlive MCP server — a small set of dispatch tools over the pptlive API.

Design notes (the things that make this safe and faithful):

* **Threading.** FastMCP calls a *sync* tool function directly on its asyncio
  event-loop thread (no thread-pool offload — verified against the SDK and in
  `scripts/mcp_spike.py`). So every tool's `attach()` runs on one consistent
  thread, and the COM apartment is initialised **once** for that thread and kept
  open for the life of the process (see `_com.com_apartment`). The original
  design re-`CoUninitialize`d after each call, on the assumption that a balanced
  per-call cycle was STA-safe — but that was wrong: repeated `CoUninitialize` on
  the long-lived event-loop thread dropped PowerPoint's automation connection
  (snapping its view back to slide 1) and eventually segfaulted (diagnosed
  2026-05-29). With the apartment held open, each tool still re-`attach()`es
  (cheap `GetActiveObject`, so we never cache a COM proxy across calls and stay
  robust to the user closing/reopening a deck) but COM itself is never torn down
  mid-session. The only cost is that a COM call briefly blocks the loop, which is
  fine for a single user driving PowerPoint serially. Tools are therefore
  deliberately **sync**.

* **Politeness + atomic undo come for free.** Tools wrap the same public API the
  CLI does, so reads don't move the view and every mutation goes through
  `deck.edit(label)` — preserving the user's viewed slide + Selection and
  fencing the change into a single Ctrl-Z.

* **Dispatch surface, not 1:1.** Five tools, four of which take an `op`-style
  argument (`ppt_read` / `ppt_edit` / `ppt_render` / `ppt_show`) instead of one
  tool per CLI subcommand — a much smaller surface for the agent's tool picker
  and far less schema resident in its context. `ppt_batch` runs a *list* of
  those same ops against one `attach()`, with all edits fenced into a single
  undo entry. The full CLI is still there for humans.

  Each op's logic lives in a `_<tool>_core(handle, op, params)` helper that does
  no `attach()` of its own; the public tool wraps the core in `attach()` (+ an
  `edit()` fence for `ppt_edit`), and `ppt_batch` reuses the very same cores
  across one shared `attach()`.

* **Errors mirror the CLI's exit-code taxonomy.** A `PptliveError` is re-raised
  as an MCP `ToolError` whose message carries a stable category token
  (`not_found` / `ambiguous` / `busy` / `not_running` / `no_text_frame` /
  `invalid_args` / `error`) — the string analog of the CLI's exit codes — so the
  agent can branch on failure. Inside `ppt_batch` the same tokens are reported
  per-command instead of aborting the batch.
"""

from __future__ import annotations

import base64
import json
import os
import tempfile
from collections.abc import Callable, Iterator
from contextlib import contextmanager, nullcontext
from enum import StrEnum
from pathlib import Path
from typing import Any, Literal

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.types import CallToolResult, ImageContent, TextContent

from .. import attach
from .._guide import skill_body
from .._presentation import Presentation
from .._shapes import Shape
from ..exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    NoTextFrameError,
    PowerPointBusyError,
    PowerPointNotRunningError,
    PptliveError,
    PresentationNotFoundError,
)

# ---------------------------------------------------------------------------
# Error mapping — the string analog of cli/main.py's _exit_for exit codes.
# ---------------------------------------------------------------------------


def _error_code(exc: PptliveError) -> str:
    # Order mirrors cli.main._exit_for: NoTextFrameError before the generic
    # AnchorNotFoundError (which covers SlideNotFoundError / LayoutNotFoundError).
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


@contextmanager
def _mcp_errors() -> Iterator[None]:
    """Re-raise a PptliveError as a ToolError carrying its taxonomy category.

    Wraps the whole `with attach() as ppt: ...` body, so an attach-time
    `PowerPointNotRunningError` is mapped too.
    """
    try:
        yield
    except PptliveError as exc:
        raise ToolError(f"{type(exc).__name__} ({_error_code(exc)}): {exc}") from exc


def _require(condition: Any, message: str) -> None:
    """Guard a required argument; surfaces as a ToolError, not a 500."""
    if not condition:
        raise ToolError(f"invalid_args: {message}")


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


# ---------------------------------------------------------------------------
# Returning rendered images *through* the MCP call (not just a filesystem path).
#
# A render writes a PNG to the Windows box, but a hosted/remote client (e.g.
# claude.ai talking to a local bundle) runs the model in a separate sandbox and
# can't open that path. The fix is the MCP image content block: we base64 the
# bytes back inline so a vision model sees the slide regardless of where the
# file lives. We return *both* — the image block AND the structured `path` —
# because a co-located filesystem tool (a local coding agent) still wants the
# path; never depend on the path alone.
#
# Caveat worth stating plainly: an image block is only as good as the host. A
# good host turns it into a native image (cost ~= w*h/750 tokens); a poor one
# inlines the base64 as text (tens of thousands of tokens). So we render small
# by default (`_EMBED_DEFAULT_WIDTH`) — legible for text-heavy slides without
# letting a whole deck blow out the context window. Pass `width`/`height` to
# override, or `embed=False` to get the path only.
# ---------------------------------------------------------------------------

#: Default long-edge pixels for an embedded slide image when the caller gives no
#: size — ~1024 px stays legible while keeping the encoded block cheap.
_EMBED_DEFAULT_WIDTH = 1024

#: Default `max_dim` long-edge cap for an embedded deck snapshot — kept a touch
#: smaller than a single-slide render since a whole deck multiplies the cost.
_EMBED_DEFAULT_MAX_DIM = 1000

_MIME_BY_FMT = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
}


def _image_block(path: str, fmt: str) -> ImageContent:
    """Read a just-rendered image file and wrap its bytes as MCP ImageContent."""
    data = base64.b64encode(Path(path).read_bytes()).decode("ascii")
    mime = _MIME_BY_FMT.get(fmt.lower(), "application/octet-stream")
    return ImageContent(type="image", data=data, mimeType=mime)


def _render_reply(results: list[dict[str, Any]], structured: dict[str, Any]) -> Any:
    """Build the tool reply for render output that may carry images.

    `results` is the list of per-op result dicts that *might* hold a `path` +
    `format` (slide_image / shape_image), or an `images` list of `{slide, path,
    format}` (deck_snapshot — each prefixed by a "slide N" text label so the model
    knows which is which); `structured` is the dict to surface as the call's
    structuredContent (the single op's result, or the batch summary). Returns a
    `CallToolResult` (image block(s) + JSON text + structured content) when any
    image is present, else just `structured` for FastMCP to serialize.
    """
    blocks: list[Any] = []
    for r in results:
        if not isinstance(r, dict):
            continue
        # Only embed paths whose format is an actual image — a deck_pdf / save /
        # save_as result also carries `path`, but its `format` ("pdf"/"pptx") is
        # not an image, so it's reported in the structured content only, never
        # read back and mis-encoded as an image block.
        if r.get("path") and r.get("format", "png").lower() in _MIME_BY_FMT:
            blocks.append(_image_block(r["path"], r.get("format", "png")))
        for img in r.get("images") or []:
            if not (isinstance(img, dict) and img.get("path")):
                continue
            if img.get("slide") is not None:
                blocks.append(TextContent(type="text", text=f"slide {img['slide']}"))
            blocks.append(_image_block(img["path"], img.get("format", "png")))
    if not blocks:
        return structured
    text = TextContent(type="text", text=json.dumps(structured, indent=2, default=str))
    return CallToolResult(content=[text, *blocks], structuredContent=structured)


# ===========================================================================
# Op vocabulary — one StrEnum per tool is the SINGLE source of truth for that
# tool's op list. The tool's `op:` parameter is typed by the enum (so FastMCP
# derives the same enum schema the agent sees today), and the dispatch registry
# below is keyed by it. An import-time assertion (after the cores) makes a
# missing handler a hard error, so the Literal / dispatch / docstring
# triplication can no longer silently drift.
# ===========================================================================


class ReadOp(StrEnum):
    STATUS = "status"
    SLIDES = "slides"
    OUTLINE = "outline"
    SLIDE = "slide"
    ANCHOR = "anchor"
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
    FIND_REPLACE = "find_replace"
    FORMAT = "format"
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
    SET_ALT = "set_alt"
    TABLE_ADD_ROW = "table_add_row"
    TABLE_DELETE_ROW = "table_delete_row"
    CHART_SET_TYPE = "chart_set_type"
    CHART_SET_DATA = "chart_set_data"
    CHART_RECOLOR_TEXT = "chart_recolor_text"
    SMARTART_SET_NODES = "smartart_set_nodes"
    SMARTART_RECOLOR_TEXT = "smartart_recolor_text"
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
        raise ToolError(f"invalid_args: unknown read op {op!r}") from exc
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
        raise ToolError(f"invalid_args: unknown write mode {mode!r}")
    return {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind, "mode": mode}


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
    para_opts = ("alignment", "space_before", "space_after", "line_spacing", "indent_level")
    fill_opts = ("fill_color", "line_color", "line_width")
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
            line_spacing=p.get("line_spacing"),
            indent_level=p.get("indent_level"),
        )
    if any(p.get(k) is not None for k in fill_opts):
        if not isinstance(anchor, Shape):
            raise ToolError(
                "invalid_args: fill_color/line_color/line_width need a shape anchor "
                f"(shape:/shapeid:/ph:), got {p['anchor_id']!r}"
            )
        anchor.set_fill(
            fill=p.get("fill_color"),
            line=p.get("line_color"),
            line_width=p.get("line_width"),
        )
    if list_type == "none":
        anchor.remove_list()
    elif list_type is not None:
        anchor.apply_list(list_type, character=p.get("bullet_char"))
    return {"ok": True, "anchor_id": anchor.anchor_id}


@edit_op(EditOp.SLIDE_ADD)
def _edit_slide_add(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    new = deck.slides.add(layout=p.get("layout"), index=p.get("index"))
    return {"ok": True, "index": new.index, "id": new.id, "layout": new.layout_name}


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
        raise ToolError(f"invalid_args: unknown shape kind {kind!r}")
    return {"ok": True, **created.to_dict()}


@edit_op(EditOp.SHAPE_MOVE)
def _edit_shape_move(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        p.get("left") is not None or p.get("top") is not None,
        "edit op='shape_move' requires `left`/`top`",
    )
    sh.move(left=p.get("left"), top=p.get("top"))
    return {"ok": True, "anchor_id": sh.anchor_id, "geometry": sh.geometry()}


@edit_op(EditOp.SHAPE_RESIZE)
def _edit_shape_resize(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(
        p.get("width") is not None or p.get("height") is not None,
        "edit op='shape_resize' requires `width`/`height`",
    )
    sh.resize(width=p.get("width"), height=p.get("height"))
    return {"ok": True, "anchor_id": sh.anchor_id, "geometry": sh.geometry()}


@edit_op(EditOp.SET_ALT)
def _edit_set_alt(deck: Presentation, p: dict[str, Any]) -> dict[str, Any]:
    sh = _resolve_shape(deck, p.get("anchor_id"))
    _require(p.get("alt_text") is not None, "edit op='set_alt' requires `alt_text`")
    sh.set_alt_text(p["alt_text"])
    return {"ok": True, "anchor_id": sh.anchor_id, "alt_text": sh.alt_text}


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
    return {"ok": True, "anchor_id": sh.anchor_id, "name": sh.name, "index": new_index}


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
        raise ToolError(f"invalid_args: unknown edit op {op!r}") from exc
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


def _parse_slide_selector(slides: Any) -> int | tuple[int, int] | None:
    """Parse a `deck_snapshot` `slides` arg into a `snapshot()` selector.

    Accepts `None` (all slides), an `int` (single 1-based slide), or a string —
    `"3"` (single) or `"2-4"` (inclusive span). A malformed string is an
    `invalid_args` ToolError.
    """
    if slides is None:
        return None
    if isinstance(slides, bool):
        raise ToolError("invalid_args: `slides` must be an int or a span string like '2-4'")
    if isinstance(slides, int):
        return slides
    s = str(slides).strip()
    try:
        if "-" in s:
            a, _, b = s.partition("-")
            return int(a), int(b)
        return int(s)
    except ValueError as e:
        raise ToolError(
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
    snaps = deck.snapshot(base, slides=selector, fmt=fmt, max_dim=p.get("max_dim"))
    return {
        "ok": True,
        "count": len(snaps),
        "format": fmt,
        "max_dim": p.get("max_dim"),
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
        raise ToolError(f"invalid_args: {exc}") from exc
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
        raise ToolError(f"invalid_args: unknown render op {op!r}") from exc
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
        raise ToolError(f"invalid_args: unknown show op {op!r}") from exc
    return SHOW_OPS[key](deck, p)


# Drift guard: every enum member must have a registered handler. Turns a missing
# op (or an enum/registry mismatch) into an import-time failure instead of a
# silent gap the agent only discovers at call time.
assert set(READ_OPS) == set(ReadOp), f"READ_OPS missing {set(ReadOp) - set(READ_OPS)}"
assert set(EDIT_OPS) == set(EditOp), f"EDIT_OPS missing {set(EditOp) - set(EDIT_OPS)}"
assert set(RENDER_OPS) == set(RenderOp), f"RENDER_OPS missing {set(RenderOp) - set(RENDER_OPS)}"
assert set(SHOW_OPS) == set(ShowOp), f"SHOW_OPS missing {set(ShowOp) - set(SHOW_OPS)}"


# ===========================================================================
# Public tools — the typed schema the agent sees. Each wraps a core in
# attach() (+ an edit() fence for mutations).
# ===========================================================================


def ppt_read(
    op: ReadOp,
    anchor_id: str | None = None,
    slide: int | None = None,
    text: str | None = None,
    scope: str | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Read the live PowerPoint deck — inspect slides, shapes, anchors, tables,
    charts, SmartArt, theme/master, and the user's selection. Always
    side-effect-free, never moves the user's view. `op`:
    - "status": open presentations, which is active, and the slide in view. Start here.
    - "slides": every slide — index (1-based), id, layout, title, shape count, has-notes.
    - "outline": each slide's title + body bullets (the fastest read of the deck's text).
    - "slide": one slide in full — every shape with its anchor_id, name, id, type,
      geometry (points), text, and whether it holds a table/chart. Needs `slide`.
      This is how you discover the anchor_ids to target with `ppt_edit`.
    - "anchor": the text of any text anchor (`anchor_id`): `ph:S:KIND` (placeholder,
      e.g. ph:2:title), `shape:S:N` (Nth shape by z-order), `para:S:N:P`,
      `cell:S:N:R:C` (table cell), `notes:S`, or `here:` (the user's selection).
      Returns text plus a `paragraphs` breakdown — each paragraph carries its
      effective `font` (`bold`/`italic`/`underline` as true/false/"mixed", `size`,
      `font` name, `color` `#RRGGBB` or null for a theme/auto color). These are
      *rendered* values; COM doesn't expose a per-run "directly set vs inherited"
      flag (only color distinguishes a literal RGB from an inherited theme color).
    - "selection": what the user currently has selected, resolved to anchors.
    - "find": every fuzzy occurrence of `text` (smart-quote / whitespace tolerant)
      across the deck — each hit a `{anchor_id, start, length, text, context}`,
      where `anchor_id` is a resolvable text anchor and `start` is the 0-based char
      offset within it. Optional `scope` restricts the search to a `slide:S` or any
      text anchor id. Pair with `ppt_edit` op="find_replace" to act on the hits.
    - "table" / "chart" / "smartart": the grid / chart data / node tree of the
      table-, chart-, or SmartArt shape at `anchor_id` (a `shape:S:N`).
    - "comments": review comments. Pass `slide` for one slide's comments (each with
      its reply thread); omit `slide` for a deck-wide roll-up
      `{total, slides:[{slide, comments:[...]}]}`. Each comment is
      `{index, author, initials, text, datetime, left, top, replies:[...]}`; address
      one for reply/delete by its `slide` + 1-based `index`.
    - "theme": the deck-wide palette (12 slots, e.g. accent1) + heading/body fonts.
    - "master": the master text styles (title/body/default, 5 levels each) + background.
    - "layouts": the layout names that `ppt_edit` slide_add/set_layout accept.

    `doc` targets a presentation by name (default: the active one)."""
    with _mcp_errors(), attach() as ppt:
        return _read_core(
            ppt,
            op,
            {"anchor_id": anchor_id, "slide": slide, "text": text, "scope": scope, "doc": doc},
        )


def ppt_edit(
    op: EditOp,
    anchor_id: str | None = None,
    text: str | None = None,
    mode: Literal["set", "insert_after", "insert_before"] = "set",
    find: str | None = None,
    scope: str | None = None,
    replace_all: bool = False,
    occurrence: int | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | None = None,
    alignment: Literal["left", "center", "right", "justify", "distribute"] | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
    indent_level: int | None = None,
    fill_color: str | None = None,
    line_color: str | None = None,
    line_width: float | None = None,
    order: Literal["front", "back", "forward", "backward"] | None = None,
    list_type: Literal["bulleted", "numbered", "none"] | None = None,
    bullet_char: str | None = None,
    slide: int | None = None,
    to: int | None = None,
    layout: str | None = None,
    index: int | None = None,
    kind: Literal["textbox", "shape", "picture", "table", "chart", "smartart"] | None = None,
    shape_type: str = "rectangle",
    path: str | None = None,
    rows: int | None = None,
    cols: int | None = None,
    chart_type: str | None = None,
    categories: list[str] | None = None,
    series: dict[str, list[float]] | None = None,
    smartart_kind: str | None = None,
    nodes: list[Any] | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    alt_text: str | None = None,
    values: list[str] | None = None,
    row: int | None = None,
    slot: str | None = None,
    which: Literal["major", "minor"] | None = None,
    name: str | None = None,
    script: Literal["latin", "east_asian", "complex_script"] = "latin",
    style: Literal["title", "body", "default"] | None = None,
    level: int | None = None,
    author: str | None = None,
    initials: str | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Edit the live PowerPoint deck — write/format text, add or arrange slides
    and shapes, find-and-replace, and apply theme/master styling. Every call is
    ONE undo entry and preserves the user's view.
    Geometry is in points (1 inch = 72 pt). `op`:

    Text & formatting (target a text anchor via `anchor_id`):
    - "write": write `text`. mode="set" replaces the whole anchor; embed `\\n` (or
      `\\r`) to start a new paragraph (each line becomes its own addressable
      `para:S:N:P`), or `\\v` for a soft line break within one paragraph.
      "insert_after"/"insert_before" add a paragraph relative to the anchor instead.
    - "find_replace": fuzzy-locate `find` across the deck and rewrite the matched
      spans with `text` (only the span changes, so run formatting is preserved).
      Scope with `scope` (a `slide:S` / anchor id). One match auto-applies; for
      several pass `replace_all=true` or `occurrence` (1-based) — otherwise it
      errors `ambiguous`. Zero matches errors `not_found`. Use op="find" first to
      preview the hits.
    - "format": font (`bold`/`italic`/`underline`/`size`/`font`/`color` — `color` is
      *font* color), paragraph (`alignment`/`space_before`/`space_after`/`line_spacing`
      in points/`indent_level` 1-5), shape fill/border on a shape anchor
      (`fill_color`/`line_color` — a hex or "none" for transparent/no border;
      `line_width` in points), and/or list (`list_type` "bulleted"/"numbered", or
      "none" to strip; `bullet_char` for a custom bullet). PowerPoint has no named
      styles, so this direct formatting is its "apply a style". Pass at least one
      option.

    Slide lifecycle:
    - "slide_add": add a slide (`layout` name, optional 1-based `index`; default end).
    - "slide_delete" / "slide_duplicate": delete / duplicate slide `slide`.
    - "slide_move": move slide `slide` to position `to`.
    - "set_layout": re-apply layout `layout` to slide `slide`.

    Shapes (create on `slide`; move/resize/delete/order/tag by `anchor_id`):
    - "shape_add": add `kind`="textbox" (with `text`), "shape" (autoshape via
      `shape_type`, e.g. "star", optional `text`), "picture" (`path`, optional
      `alt_text`), "table" (`rows`+`cols`), "chart" (`chart_type`, optional
      `categories`+`series`), or "smartart" (`smartart_kind` e.g. "process"/
      "cycle"/"orgchart", optional `nodes`). Optional `left`/`top`/`width`/`height`;
      textbox/shape also take `fill_color`/`line_color` (hex or "none") + `line_width`.
    - "shape_move": move to absolute `left`/`top`. "shape_resize": set `width`/`height`.
    - "shape_order": restack by `order`="front"/"back"/"forward"/"backward" (e.g.
      send a new background panel to the back, behind existing content).
    - "shape_delete": delete it. "set_alt": set `alt_text` (a drift-proof handle).
      Address a shape that must survive a delete/restack by `shapeid:S:ID` (the
      stable `id` from any shape listing), not the positional `shape:S:N`.

    Tables, charts & SmartArt (target the shape by its `anchor_id`, a shape:S:N):
    - "table_add_row": append a row, optionally filled from `values`.
    - "table_delete_row": delete 1-based `row`.
    - "chart_set_type": change chart kind to `chart_type` (e.g. "line"/"pie"/"bar").
    - "chart_set_data": replace `categories` + `series` (a {name:[values]} map).
      Series are plotted in insertion order; note bar charts render series
      bottom-to-top, so the first series sits at the bottom (Excel/PowerPoint
      convention, not a reorder).
    - "chart_recolor_text": set `color` on EVERY shown chart text element (legend,
      axis tick labels, title, data labels) at once — the coarse fix when inherited
      black chart text is invisible on a custom background. A chart has no text
      anchor, so this is the only color path for its internal text.
    - "smartart_set_nodes": replace the diagram's `nodes` — a list of strings
      (flat) and/or {text, children} objects (nested; tree layouts take one root).
    - "smartart_recolor_text": set `color` on EVERY node's label at once — the
      coarse fix when inherited black node text is invisible on a custom
      background. A SmartArt diagram has no text anchor, so this is its only text
      color path.

    To edit a table cell's text, write to its `cell:S:N:R:C` anchor with op="write".

    Review comments (threaded, per-slide; addressed by `slide` + 1-based `index`):
    - "comment_add": add a comment to `slide` with body `text` (optional `left`/`top`
      anchor point in points). Binds to the signed-in Office account — the shown
      author/initials follow that account, not the optional `author`/`initials`
      (those reach only the legacy fallback on a deck with no comments to source an
      identity from). Use op="comments" to find a comment's `index`.
    - "comment_reply": add a threaded reply (`text`) to comment `index` on `slide`.
    - "comment_delete": delete comment `index` on `slide` (takes its replies too).
    There is no resolve/reopen op — comment resolution state is not COM-readable.

    Deck-wide styling (global — restyles every inheriting slide; no `anchor_id`):
    - "theme_set_color": set palette `slot` (e.g. "accent1"/"dark1"/"hyperlink") to `color`.
    - "theme_set_font": set the `which`="major" (headings) or "minor" (body) typeface
      to `name` (optional `script`, default "latin").
    - "master_format_text_style": font (`bold`/`italic`/`underline`/`size`/`font`/`color`)
      on master text `style` ("title"/"body"/"default") + outline `level` (1-5,
      default 1 — the natural choice for `title`).
    - "master_format_paragraph_style": paragraph (`alignment`/`space_before`/
      `space_after`/`line_spacing`) on `style` + `level` (1-5, default 1).
    - "master_set_background": set the master background to solid `color`.

    `doc` targets a presentation by name (default: the active one)."""
    params = {
        "anchor_id": anchor_id,
        "text": text,
        "mode": mode,
        "find": find,
        "scope": scope,
        "replace_all": replace_all,
        "occurrence": occurrence,
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "size": size,
        "font": font,
        "color": color,
        "alignment": alignment,
        "space_before": space_before,
        "space_after": space_after,
        "line_spacing": line_spacing,
        "indent_level": indent_level,
        "fill_color": fill_color,
        "line_color": line_color,
        "line_width": line_width,
        "order": order,
        "list_type": list_type,
        "bullet_char": bullet_char,
        "slide": slide,
        "to": to,
        "layout": layout,
        "index": index,
        "kind": kind,
        "shape_type": shape_type,
        "path": path,
        "rows": rows,
        "cols": cols,
        "chart_type": chart_type,
        "categories": categories,
        "series": series,
        "smartart_kind": smartart_kind,
        "nodes": nodes,
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "alt_text": alt_text,
        "values": values,
        "row": row,
        "slot": slot,
        "which": which,
        "name": name,
        "script": script,
        "style": style,
        "level": level,
        "author": author,
        "initials": initials,
    }
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        with deck.edit(f"MCP: {op}"):
            return _edit_core(deck, op, params)


def ppt_render(
    op: RenderOp,
    slide: int | None = None,
    anchor_id: str | None = None,
    out: str | None = None,
    width: int | None = None,
    height: int | None = None,
    fmt: Literal["png", "jpg", "jpeg", "gif", "bmp"] = "png",
    select: bool = True,
    embed: bool = True,
    doc: str | None = None,
    slides: str | None = None,
    max_dim: int | None = None,
    overwrite: bool = False,
    save_format: str = "pptx",
) -> Any:
    """Render a PowerPoint slide or shape to an image a vision model can see, or
    move the user's view to a slide/shape. `op`:
    - "slide_image": render slide `slide` (1-based) to an image and return it BOTH
      as an inline image the vision model can *see* (render -> look -> iterate) AND
      as an absolute file `path` in the structured result — so it works whether you
      run co-located with the file or in a remote sandbox that can't open the path.
      Renders the current unsaved state; polite (does not move the view). `out`
      defaults to a temp file; pass `width`/`height` (pixels; one is enough — the
      other follows the aspect ratio). Default render is ~1024 px on the long edge
      to keep the inline image cheap; override with `width`/`height`.
    - "deck_snapshot": render the WHOLE deck (or `slides`) to one inline image per
      slide so you can SEE every slide at once cheaply — the token-aware "did my
      styling land across all slides" read. `max_dim` caps each slide's long edge
      in pixels (only ever lowering resolution); since every slide shares one
      geometry the cap is a uniform, predictable per-slide budget (defaults to
      ~1000 px when embedding). `slides` selects what to render: a single 1-based
      slide ("3") or an inclusive span ("2-4"); omit for the whole deck. Each slide
      comes back as a "slide N" label + image block; the structured result lists the
      written file `path`s. Polite (does not move the view).
    - "shape_image": render *just* the shape at `anchor_id` (cropped to its bounds,
      native pixel size) — so a vision model can see one picture/diagram alone.
      Same dual return (inline image + `path`). Polite. `out` defaults to a temp file.
    - "deck_pdf": export the whole deck to a PDF at `out` (required) — the "hand
      back a deliverable" path. A read: a pixel-faithful render of the current
      unsaved state that does NOT rebind the working file or clear its dirty flag.
      Overwrites an existing PDF. Returns the written `path` (no inline image).
    - "save": save the deck to its existing file (explicit; pptlive never
      auto-saves). Errors if the deck has never been saved — use "save_as" first.
    - "save_as": save the deck to `out` (required, a .pptx path) and REBIND the
      working file to it (the open deck becomes that file). `save_format` is the
      target format (default "pptx"). Refuses to clobber an existing file unless
      `overwrite=True`. For PDF use "deck_pdf" (a read).
    - "navigate": move the user's view to `anchor_id`'s slide — a deliberate,
      opt-in view move (every other tool leaves the view alone). With `select=True`
      (default), also selects the target shape. Use only when asked to be taken
      somewhere.

    `embed` (default True) returns the rendered image inline so a remote model can
    see it; set False for the path only (smaller reply when a local tool reads the
    file). Note: whether the inline image reaches the model depends on the MCP host
    — most desktop hosts forward it as a native image, but some only pass the path.
    `fmt` is the image format. `doc` targets a presentation by name."""
    if op == "slide_image" and embed and width is None and height is None:
        width = _EMBED_DEFAULT_WIDTH
    if op == "deck_snapshot" and embed and max_dim is None:
        max_dim = _EMBED_DEFAULT_MAX_DIM
    params = {
        "slide": slide,
        "anchor_id": anchor_id,
        "out": out,
        "width": width,
        "height": height,
        "fmt": fmt,
        "select": select,
        "doc": doc,
        "slides": slides,
        "max_dim": max_dim,
        "overwrite": overwrite,
        "save_format": save_format,
    }
    with _mcp_errors(), attach() as ppt:
        result = _render_core(ppt, op, params)
        if embed and op in ("slide_image", "deck_snapshot", "shape_image"):
            return _render_reply([result], result)
        return result


def ppt_show(
    op: ShowOp = ShowOp.STATE,
    slide: int | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Drive a live PowerPoint slide show — the presenter's clicker (start, next,
    previous, goto, black/white, end). Unlike every other
    mutating tool, this deliberately controls what's on the user's screen. `op`:
    - "state": is a show running, and which slide is on screen (read-only).
    - "start": begin the show (optional 1-based `slide` to start on).
    - "end": exit the show (no-op if none is running).
    - "next" / "previous": advance / step back one build or slide.
    - "goto": jump to 1-based `slide`.
    - "black" / "white": blank the screen; "resume" returns to the slide.

    Every op returns the resulting show state. The control verbs (next/previous/
    goto/black/white/resume) need a running show — they error otherwise."""
    with _mcp_errors(), attach() as ppt:
        return _show_core(_pick_deck(ppt, doc), op, {"slide": slide})


def ppt_batch(
    commands: list[dict[str, Any]],
    doc: str | None = None,
    atomic: bool = True,
    stop_on_error: bool = True,
    embed: bool = True,
) -> Any:
    """Run a list of ops against one PowerPoint connection — the way to build or
    restructure a slide without a round-trip per change. Each command is a dict:

        {"tool": "edit", "op": "write", "anchor_id": "ph:2:title", "text": "Q3 Results"}

    `tool` is "read" | "edit" | "render" | "show" (default "edit"); the remaining
    keys are exactly the parameters of that tool (minus `doc`). All commands target
    the same deck (`doc`).

    `atomic` (default True): all `edit` commands are fenced into a SINGLE undo entry
    — the whole batch is one Ctrl-Z. This is undo *grouping*, not a transaction:
    if a command fails partway, earlier mutations stay applied (pair with
    `stop_on_error`). With atomic=False each edit is its own undo entry.

    `stop_on_error` (default True): stop at the first failing command. With False,
    every command runs and failures are reported in place.

    `embed` (default True): any `render` slide_image/shape_image commands also
    return their rendered image inline (in addition to the `path` in the structured
    result), so "build a slide, then look at it" works in one round trip even from a
    remote sandbox. Set False for paths only. Inline slide images default to
    ~1024 px on the long edge; a render command's own `width`/`height` overrides that.

    Returns `{"ok": <all succeeded>, "atomic", "count", "results": [...]}` where each
    result is `{"index", "tool", "op", "ok", "result"}` on success or
    `{..., "ok": false, "error": <category>, "message"}` on failure (same category
    tokens as the other tools' ToolErrors). When `embed` surfaces images, the reply
    carries those image blocks alongside this summary as its structured content."""
    _require(
        isinstance(commands, list) and len(commands) > 0,
        "ppt_batch requires a non-empty `commands` list",
    )
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        has_edit = any(cmd.get("tool", "edit") == "edit" for cmd in commands)
        scope: Any = (
            deck.edit(f"MCP: batch ({len(commands)} ops)")
            if (atomic and has_edit)
            else nullcontext()
        )
        results: list[dict[str, Any]] = []
        with scope as edit_scope:
            for i, cmd in enumerate(commands):
                tool = cmd.get("tool", "edit")
                op = cmd.get("op")
                p = {k: v for k, v in cmd.items() if k not in ("tool", "op")}
                p["doc"] = doc
                if (
                    embed
                    and tool == "render"
                    and op == "slide_image"
                    and p.get("width") is None
                    and p.get("height") is None
                ):
                    p["width"] = _EMBED_DEFAULT_WIDTH
                if (
                    embed
                    and tool == "render"
                    and op == "deck_snapshot"
                    and p.get("max_dim") is None
                ):
                    p["max_dim"] = _EMBED_DEFAULT_MAX_DIM
                entry: dict[str, Any] = {"index": i, "tool": tool, "op": op}
                try:
                    _require(op is not None, f"command #{i} is missing `op`")
                    assert op is not None
                    if tool == "read":
                        result = _read_core(ppt, op, p)
                    elif tool == "edit":
                        if atomic:
                            result = _edit_core(deck, op, p)
                        else:
                            with deck.edit(f"MCP: batch #{i} {op}"):
                                result = _edit_core(deck, op, p)
                    elif tool == "render":
                        result = _render_core(ppt, op, p)
                    elif tool == "show":
                        result = _show_core(deck, op, p)
                    else:
                        raise ToolError(f"invalid_args: command #{i} unknown tool {tool!r}")
                    # A deliberate view-move inside an atomic batch must survive the
                    # scope's restore — otherwise a `navigate`/`show` is snapped back
                    # to the pre-batch slide on exit. Opt the whole scope out of the
                    # view restore once such a command runs (mirrors how a bare
                    # `go_to` requires `allow_view_move()` inside a `deck.edit`).
                    if _moves_view(tool, op) and edit_scope is not None:
                        edit_scope.allow_view_move()
                    entry.update(ok=True, result=result)
                except (PptliveError, ToolError) as exc:
                    code = _error_code(exc) if isinstance(exc, PptliveError) else "invalid_args"
                    entry.update(ok=False, error=code, message=str(exc))
                    results.append(entry)
                    if stop_on_error:
                        break
                    continue
                results.append(entry)
        summary = {
            "ok": all(r["ok"] for r in results),
            "atomic": atomic,
            "count": len(results),
            "results": results,
        }
        if not embed:
            return summary
        rendered = [
            r["result"]
            for r in results
            if r["ok"] and r["tool"] == "render" and isinstance(r.get("result"), dict)
        ]
        return _render_reply(rendered, summary)


# ---------------------------------------------------------------------------
# Server assembly
# ---------------------------------------------------------------------------

_TOOLS: list[Callable[..., Any]] = [
    ppt_read,
    ppt_edit,
    ppt_render,
    ppt_show,
    ppt_batch,
]


_INSTRUCTIONS = (
    "Drive the PowerPoint deck the user has open right now. Five dispatch tools "
    "(ppt_read / ppt_edit / ppt_render / ppt_show / ppt_batch), each taking an "
    "`op`. Address content with hierarchical anchors (`ph:S:KIND`, `shape:S:N`, "
    "`para:S:N:P`, `cell:S:N:R:C`, `notes:S`); reads never move the view and every "
    "edit is one Ctrl-Z. Read the `pptlive://guide` resource for the full op "
    "vocabulary and anchor model (`pptlive://guide/python` for the Python API)."
)


def build_server(name: str = "pptlive") -> FastMCP:
    """Construct a FastMCP server with every pptlive tool registered.

    Kept as a factory (rather than only a module-level singleton) so tests can
    stand up a fresh server, and so an embedder can mount the tools elsewhere.
    """
    srv = FastMCP(name, instructions=_INSTRUCTIONS)
    for fn in _TOOLS:
        srv.add_tool(fn)

    @srv.resource("pptlive://guide", mime_type="text/markdown")
    def guide() -> str:
        """The full pptlive agent guide: anchor model, every verb, the op vocabulary."""
        return skill_body("cli")

    @srv.resource("pptlive://guide/python", mime_type="text/markdown")
    def guide_python() -> str:
        """The pptlive Python-API guide (`import pptlive as pl`)."""
        return skill_body("python")

    return srv


# The singleton the `pptlive-mcp` entry point runs.
server = build_server()

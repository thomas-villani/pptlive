"""The pptlive MCP server — a curated set of tools over the pptlive API.

Design notes (the things that make this safe and faithful):

* **Threading.** FastMCP calls a *sync* tool function directly on its asyncio
  event-loop thread (no thread-pool offload — verified against the SDK and in
  `scripts/mcp_spike.py`). So each tool's `attach()` does its whole
  `CoInitialize -> work -> CoUninitialize` cycle on one consistent thread per
  call — the same shape as a one-shot CLI invocation, just repeated in a
  long-lived process. That is STA-safe. The only cost is that a COM call briefly
  blocks the loop, which is fine for a single user driving PowerPoint serially.
  Tools are therefore deliberately **sync**, and never cache a COM object across
  calls (each tool re-`attach()`es).

* **Politeness + atomic undo come for free.** Tools wrap the same public API the
  CLI does, so reads don't move the view and every mutation goes through
  `deck.edit(label)` — preserving the user's viewed slide + Selection and
  fencing the change into a single Ctrl-Z.

* **Curated, not 1:1.** ~14 tools, several of which take a `verb`-style argument
  (`op` / `mode`) instead of one tool per CLI subcommand — a smaller surface for
  the agent's tool picker. The full CLI is still there for humans.

* **Errors mirror the CLI's exit-code taxonomy.** A `PptliveError` is re-raised
  as an MCP `ToolError` whose message carries a stable category token
  (`not_found` / `ambiguous` / `busy` / `not_running` / `no_text_frame` /
  `error`) — the string analog of the CLI's exit codes — so the agent can branch
  on failure.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator
from contextlib import contextmanager
from typing import Any, Literal

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError

from .. import attach
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


def _resolve_shape(deck: Presentation, anchor_id: str) -> Shape:
    """Resolve a shape/placeholder anchor to a `Shape` (else a not_found ToolError)."""
    anchor = deck.anchor_by_id(anchor_id)
    if not isinstance(anchor, Shape):
        raise AnchorNotFoundError("shape", anchor_id)
    return anchor


# ===========================================================================
# Reads (side-effect-free; never move the view)
# ===========================================================================


def ppt_status() -> dict[str, Any]:
    """List the open PowerPoint presentations, which one is active, and the slide
    currently in view. Use this first to see what the user has open. Read-only."""
    with _mcp_errors(), attach() as ppt:
        return {
            "decks": ppt.presentations.list(),
            "viewed_slide": ppt.viewed_slide_index(),
        }


def ppt_slides(doc: str | None = None) -> dict[str, Any]:
    """List every slide in the deck: index (1-based), id, layout, title, shape
    count, and whether it has speaker notes. Returns `{"slides": [...]}`. Read-only.

    `doc` targets a presentation by name (default: the active one)."""
    with _mcp_errors(), attach() as ppt:
        return {"slides": _pick_deck(ppt, doc).slides.list()}


def ppt_outline(doc: str | None = None) -> dict[str, Any]:
    """The deck's outline — each slide's title and body bullets (the Outline-view
    analog). Returns `{"outline": [...]}`; the fastest way to read the deck's text
    content. Read-only."""
    with _mcp_errors(), attach() as ppt:
        return {"outline": _pick_deck(ppt, doc).outline()}


def ppt_slide_read(slide: int, doc: str | None = None) -> dict[str, Any]:
    """Read one slide in full: every shape with its anchor_id, name, id, type,
    geometry (points), text, and whether it holds a table. `slide` is 1-based.
    This is how you discover the anchor_ids to target with other tools. Read-only."""
    with _mcp_errors(), attach() as ppt:
        return _pick_deck(ppt, doc).slides[slide].read()


def ppt_read(anchor_id: str, doc: str | None = None) -> dict[str, Any]:
    """Read the text of any text anchor. `anchor_id` is one of:
    `ph:S:KIND` (placeholder, e.g. ph:2:title), `shape:S:N` (Nth shape by
    z-order), `para:S:N:P` (a paragraph), `cell:S:N:R:C` (a table cell),
    `notes:S` (speaker notes), or `here:` (the user's current selection).

    Returns the text plus, for anchors that hold multiple paragraphs, a
    `paragraphs` breakdown (each with its own para:S:N:P anchor_id, indent level,
    and bullet). Read-only."""
    with _mcp_errors(), attach() as ppt:
        anchor = _pick_deck(ppt, doc).anchor_by_id(anchor_id)
        payload: dict[str, Any] = {
            "anchor_id": anchor.anchor_id,
            "kind": anchor.kind,
            "text": anchor.text,
        }
        paragraphs = getattr(anchor, "paragraphs", None)
        if paragraphs is not None:
            payload["paragraphs"] = paragraphs.list()
        return payload


def ppt_selection(doc: str | None = None) -> dict[str, Any]:
    """Report what the user currently has selected, resolved to anchors — the
    selected shapes (`shape:S:N`) or a text caret (`para:S:N:P`), plus the single
    `anchor_id` that the `here:` anchor resolves to. Use this to act on "the thing
    the user is looking at". Read-only (does not change the selection)."""
    with _mcp_errors(), attach() as ppt:
        return _pick_deck(ppt, doc).selection().to_dict()


# ===========================================================================
# Edits (each its own deck.edit block => preserves the view + one Ctrl-Z)
# ===========================================================================


def ppt_write(
    anchor_id: str,
    text: str,
    mode: Literal["set", "insert_after", "insert_before"] = "set",
    doc: str | None = None,
) -> dict[str, Any]:
    """Write text to a text anchor (one undo entry; preserves the viewed slide).

    `mode="set"` (default) replaces the anchor's whole text — embed `\\n` for
    multiple paragraphs. `mode="insert_after"` / `"insert_before"` adds a new
    paragraph relative to the anchor instead of replacing it. `anchor_id` accepts
    ph:/shape:/para:/cell:/notes:/here: forms (see ppt_read)."""
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        anchor = deck.anchor_by_id(anchor_id)
        with deck.edit(f"MCP: write {anchor_id} ({mode})"):
            if mode == "set":
                anchor.set_text(text)
            elif mode == "insert_after":
                anchor.insert_paragraph_after(text)
            else:
                anchor.insert_paragraph_before(text)
        return {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind, "mode": mode}


def ppt_format(
    anchor_id: str,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | None = None,
    alignment: Literal["left", "center", "right", "justify"] | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
    indent_level: int | None = None,
    list_type: Literal["bulleted", "numbered", "none"] | None = None,
    bullet_char: str | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Format a text anchor: font (bold/italic/underline/size/font/color),
    paragraph (alignment/spacing in points/indent_level 1-5), and/or list
    formatting (list_type "bulleted"/"numbered", or "none" to strip bullets;
    bullet_char sets a custom bullet). PowerPoint has no named paragraph styles,
    so this direct font/paragraph formatting is its "apply a style".

    Pass at least one option. Applies in a single undo entry; preserves the view."""
    font_opts = (bold, italic, underline, size, font, color)
    para_opts = (alignment, space_before, space_after, line_spacing, indent_level)
    _require(
        any(v is not None for v in font_opts + para_opts) or list_type is not None,
        "ppt_format needs at least one font, paragraph, or list option",
    )
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        anchor = deck.anchor_by_id(anchor_id)
        with deck.edit(f"MCP: format {anchor_id}"):
            if any(v is not None for v in font_opts):
                anchor.format_text(
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    size=size,
                    font=font,
                    color=color,
                )
            if any(v is not None for v in para_opts):
                anchor.format_paragraph(
                    alignment=alignment,
                    space_before=space_before,
                    space_after=space_after,
                    line_spacing=line_spacing,
                    indent_level=indent_level,
                )
            if list_type == "none":
                anchor.remove_list()
            elif list_type is not None:
                anchor.apply_list(list_type, character=bullet_char)
        return {"ok": True, "anchor_id": anchor.anchor_id}


def ppt_slide_op(
    op: Literal["add", "delete", "duplicate", "move", "set_layout", "layouts"],
    slide: int | None = None,
    to: int | None = None,
    layout: str | None = None,
    index: int | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Slide lifecycle. `op`:
    - "layouts": the layout names add/set_layout accept, as `{"layouts": [...]}`
      (read-only).
    - "add": add a slide (`layout` name, optional 1-based `index`; default end).
    - "delete": delete slide `slide`.
    - "duplicate": duplicate slide `slide` (copy lands just after it).
    - "move": move slide `slide` to position `to`.
    - "set_layout": re-apply layout `layout` to slide `slide`.
    All mutating ops are one undo entry and preserve the user's view."""
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        if op == "layouts":
            return {"layouts": deck.layouts()}
        if op == "add":
            with deck.edit(f"MCP: add slide ({layout or 'default'})"):
                new = deck.slides.add(layout=layout, index=index)
            return {"ok": True, "index": new.index, "id": new.id, "layout": new.layout_name}

        _require(slide is not None, f"slide_op op={op!r} requires `slide`")
        assert slide is not None
        target = deck.slides[slide]
        if op == "delete":
            with deck.edit(f"MCP: delete slide {slide}"):
                target.delete()
            return {"ok": True, "deleted": slide}
        if op == "duplicate":
            with deck.edit(f"MCP: duplicate slide {slide}"):
                new = target.duplicate()
            return {"ok": True, "index": new.index, "id": new.id, "from": slide}
        if op == "move":
            _require(to is not None, "slide_op op='move' requires `to`")
            assert to is not None
            with deck.edit(f"MCP: move slide {slide} -> {to}"):
                moved = target.move_to(to)
            return {"ok": True, "index": moved.index, "id": moved.id}
        # op == "set_layout"
        _require(layout is not None, "slide_op op='set_layout' requires `layout`")
        assert layout is not None
        with deck.edit(f"MCP: set layout of slide {slide}"):
            target.set_layout(layout)
        return {"ok": True, "index": slide, "layout": target.layout_name}


def ppt_shape_op(
    op: Literal["add", "move", "resize", "delete"],
    slide: int | None = None,
    anchor_id: str | None = None,
    kind: Literal["textbox", "shape", "picture", "table"] | None = None,
    text: str | None = None,
    shape_type: str = "rectangle",
    path: str | None = None,
    rows: int | None = None,
    cols: int | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Create or place shapes (geometry in points; 1 inch = 72 pt). `op`:
    - "add": add a shape on `slide`. `kind`="textbox" (with `text`), "shape"
      (autoshape geometry via `shape_type`, e.g. "star"/"rectangle", optional
      `text`), "picture" (embed the image at `path`), or "table" (needs `rows`
      and `cols`). Optional `left`/`top`/`width`/`height`.
    - "move": move the shape at `anchor_id` to absolute `left`/`top`.
    - "resize": set the shape at `anchor_id` to `width`/`height`.
    - "delete": delete the shape at `anchor_id`.
    One undo entry; preserves the view. Returns the shape's anchor_id, name,
    type, and geometry."""
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        if op == "add":
            _require(slide is not None, "shape_op op='add' requires `slide`")
            _require(kind is not None, "shape_op op='add' requires `kind`")
            assert slide is not None
            shapes = deck.slides[slide].shapes
            with deck.edit(f"MCP: add {kind} on slide {slide}"):
                if kind == "textbox":
                    new = shapes.add_textbox(
                        text or "", left=left, top=top, width=width, height=height
                    )
                elif kind == "shape":
                    new = shapes.add_shape(
                        shape_type, left=left, top=top, width=width, height=height
                    )
                    if text:
                        new.set_text(text)
                elif kind == "table":
                    _require(
                        rows is not None and cols is not None,
                        "shape_op kind='table' requires `rows` and `cols`",
                    )
                    assert rows is not None and cols is not None
                    new = shapes.add_table(
                        rows, cols, left=left, top=top, width=width, height=height
                    )
                else:  # picture
                    _require(path is not None, "shape_op kind='picture' requires `path`")
                    assert path is not None
                    new = shapes.add_picture(path, left=left, top=top, width=width, height=height)
            return {"ok": True, **new.to_dict()}

        _require(anchor_id is not None, f"shape_op op={op!r} requires `anchor_id`")
        assert anchor_id is not None
        sh = _resolve_shape(deck, anchor_id)
        if op == "move":
            _require(left is not None or top is not None, "shape_op op='move' requires left/top")
            with deck.edit(f"MCP: move {anchor_id}"):
                sh.move(left=left, top=top)
        elif op == "resize":
            _require(
                width is not None or height is not None,
                "shape_op op='resize' requires width/height",
            )
            with deck.edit(f"MCP: resize {anchor_id}"):
                sh.resize(width=width, height=height)
        else:  # delete
            info = {"anchor_id": sh.anchor_id, "name": sh.name, "id": sh.shape_id}
            with deck.edit(f"MCP: delete {anchor_id}"):
                sh.delete()
            return {"ok": True, **info}
        return {"ok": True, "anchor_id": sh.anchor_id, "geometry": sh.geometry()}


def ppt_table(
    op: Literal["read", "add_row", "delete_row"],
    slide: int,
    shape: int,
    values: list[str] | None = None,
    row: int | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Read or edit a table. A table is a shape, so address it by `slide` (1-based)
    and `shape` (1-based z-order index); cells are `cell:S:N:R:C` anchors you can
    target with ppt_read / ppt_write / ppt_format. `op`:
    - "read": return the grid — rows, columns, and each cell with its anchor_id
      and text (read-only).
    - "add_row": append a row, optionally filling it left-to-right from `values`.
    - "delete_row": delete 1-based `row`.
    Mutating ops are one undo entry."""
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        table = deck.slides[slide].shapes[shape].table
        if op == "read":
            return table.read()
        if op == "add_row":
            with deck.edit(f"MCP: add row to table shape:{slide}:{shape}"):
                table.add_row(values)
            return {"ok": True, "anchor_id": table.shape.anchor_id, "rows": table.row_count}
        # delete_row
        _require(row is not None, "table op='delete_row' requires `row`")
        assert row is not None
        with deck.edit(f"MCP: delete row {row} from table shape:{slide}:{shape}"):
            table.delete_row(row)
        return {"ok": True, "anchor_id": table.shape.anchor_id, "rows": table.row_count}


# ===========================================================================
# Render + navigate
# ===========================================================================


def ppt_export(
    slide: int,
    out: str | None = None,
    width: int | None = None,
    height: int | None = None,
    fmt: Literal["png", "jpg", "jpeg", "gif", "bmp"] = "png",
    doc: str | None = None,
) -> dict[str, Any]:
    """Render slide `slide` (1-based) to an image file and return its absolute path
    — so a vision model can *see* the slide it just built (export -> look ->
    iterate). Renders the current unsaved state; polite (doesn't move the view).
    `out` defaults to a temp file. Pass `width` and/or `height` (pixels) for a
    specific size (one is enough — the other follows the slide's aspect ratio)."""
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        path = deck.slides[slide].export_image(out, width=width, height=height, fmt=fmt)
        return {"ok": True, "slide": slide, "path": str(path), "format": fmt}


def ppt_show(
    op: Literal[
        "state", "start", "end", "next", "previous", "goto", "black", "white", "resume"
    ] = "state",
    slide: int | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Drive a live slide show — the presenter's clicker. Unlike every other
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
        sh = _pick_deck(ppt, doc).show
        if op == "state":
            return sh.state()
        if op == "start":
            return sh.start(from_slide=slide)
        if op == "end":
            return sh.end()
        if op == "next":
            return sh.next()
        if op == "previous":
            return sh.previous()
        if op == "goto":
            _require(slide is not None, "show op='goto' requires `slide`")
            assert slide is not None
            return sh.goto(slide)
        if op == "black":
            return sh.black()
        if op == "white":
            return sh.white()
        return sh.resume()  # op == "resume"


def ppt_navigate(anchor_id: str, select: bool = True, doc: str | None = None) -> dict[str, Any]:
    """Move the user's view to an anchor's slide — a deliberate, opt-in view move
    (every *other* tool is polite and leaves the view alone). With `select=True`
    (default), also selects the target shape. Use only when the user asks to be
    taken somewhere."""
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        anchor = deck.anchor_by_id(anchor_id)
        deck.go_to(anchor, select=select)
        return {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind}


# ---------------------------------------------------------------------------
# Server assembly
# ---------------------------------------------------------------------------

_TOOLS: list[Callable[..., Any]] = [
    ppt_status,
    ppt_slides,
    ppt_outline,
    ppt_slide_read,
    ppt_read,
    ppt_selection,
    ppt_write,
    ppt_format,
    ppt_slide_op,
    ppt_shape_op,
    ppt_table,
    ppt_export,
    ppt_show,
    ppt_navigate,
]


def build_server(name: str = "pptlive") -> FastMCP:
    """Construct a FastMCP server with every pptlive tool registered.

    Kept as a factory (rather than only a module-level singleton) so tests can
    stand up a fresh server, and so an embedder can mount the tools elsewhere.
    """
    srv = FastMCP(name)
    for fn in _TOOLS:
        srv.add_tool(fn)
    return srv


# The singleton the `pptlive-mcp` entry point runs.
server = build_server()

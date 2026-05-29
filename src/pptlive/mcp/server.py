"""The pptlive MCP server — a small set of dispatch tools over the pptlive API.

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

from collections.abc import Callable, Iterator
from contextlib import contextmanager, nullcontext
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


def _resolve_shape(deck: Presentation, anchor_id: str | None) -> Shape:
    """Resolve a shape/placeholder anchor to a `Shape` (else a not_found ToolError)."""
    _require(anchor_id is not None, "this op requires `anchor_id`")
    assert anchor_id is not None
    anchor = deck.anchor_by_id(anchor_id)
    if not isinstance(anchor, Shape):
        raise AnchorNotFoundError("shape", anchor_id)
    return anchor


# ===========================================================================
# Op cores — pure dispatch over an already-attached handle (no attach/edit
# bracketing of their own, so ppt_batch can call them under one shared scope).
# ===========================================================================


def _read_core(ppt: Any, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Read-only dispatch (never moves the view)."""
    if op == "status":
        return {
            "decks": ppt.presentations.list(),
            "viewed_slide": ppt.viewed_slide_index(),
        }
    deck = _pick_deck(ppt, p.get("doc"))
    if op == "slides":
        return {"slides": deck.slides.list()}
    if op == "outline":
        return {"outline": deck.outline()}
    if op == "layouts":
        return {"layouts": deck.layouts()}
    if op == "selection":
        return deck.selection().to_dict()
    if op == "slide":
        _require(p.get("slide") is not None, "read op='slide' requires `slide`")
        return deck.slides[p["slide"]].read()
    if op == "anchor":
        _require(p.get("anchor_id") is not None, "read op='anchor' requires `anchor_id`")
        anchor = deck.anchor_by_id(p["anchor_id"])
        payload: dict[str, Any] = {
            "anchor_id": anchor.anchor_id,
            "kind": anchor.kind,
            "text": anchor.text,
        }
        paragraphs = getattr(anchor, "paragraphs", None)
        if paragraphs is not None:
            payload["paragraphs"] = paragraphs.list()
        return payload
    if op == "table":
        return _resolve_shape(deck, p.get("anchor_id")).table.read()
    if op == "chart":
        return _resolve_shape(deck, p.get("anchor_id")).chart.read()
    raise ToolError(f"invalid_args: unknown read op {op!r}")


def _edit_core(deck: Presentation, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Mutation dispatch. The caller MUST have an open `deck.edit(...)` scope."""
    # -- text --------------------------------------------------------------
    if op == "write":
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

    if op == "format":
        _require(p.get("anchor_id") is not None, "edit op='format' requires `anchor_id`")
        font_opts = ("bold", "italic", "underline", "size", "font", "color")
        para_opts = ("alignment", "space_before", "space_after", "line_spacing", "indent_level")
        list_type = p.get("list_type")
        _require(
            any(p.get(k) is not None for k in font_opts + para_opts) or list_type is not None,
            "edit op='format' needs at least one font, paragraph, or list option",
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
        if list_type == "none":
            anchor.remove_list()
        elif list_type is not None:
            anchor.apply_list(list_type, character=p.get("bullet_char"))
        return {"ok": True, "anchor_id": anchor.anchor_id}

    # -- slide lifecycle ---------------------------------------------------
    if op == "slide_add":
        new = deck.slides.add(layout=p.get("layout"), index=p.get("index"))
        return {"ok": True, "index": new.index, "id": new.id, "layout": new.layout_name}
    if op in ("slide_delete", "slide_duplicate", "slide_move", "set_layout"):
        _require(p.get("slide") is not None, f"edit op={op!r} requires `slide`")
        target = deck.slides[p["slide"]]
        if op == "slide_delete":
            target.delete()
            return {"ok": True, "deleted": p["slide"]}
        if op == "slide_duplicate":
            new = target.duplicate()
            return {"ok": True, "index": new.index, "id": new.id, "from": p["slide"]}
        if op == "slide_move":
            _require(p.get("to") is not None, "edit op='slide_move' requires `to`")
            moved = target.move_to(p["to"])
            return {"ok": True, "index": moved.index, "id": moved.id}
        # set_layout
        _require(p.get("layout") is not None, "edit op='set_layout' requires `layout`")
        target.set_layout(p["layout"])
        return {"ok": True, "index": p["slide"], "layout": target.layout_name}

    # -- shapes ------------------------------------------------------------
    if op == "shape_add":
        _require(p.get("slide") is not None, "edit op='shape_add' requires `slide`")
        kind = p.get("kind")
        _require(kind is not None, "edit op='shape_add' requires `kind`")
        geom = {k: p.get(k) for k in ("left", "top", "width", "height")}
        shapes = deck.slides[p["slide"]].shapes
        if kind == "textbox":
            created = shapes.add_textbox(p.get("text") or "", **geom)
        elif kind == "shape":
            created = shapes.add_shape(p.get("shape_type") or "rectangle", **geom)
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
        elif kind == "picture":
            _require(p.get("path") is not None, "edit shape_add kind='picture' requires `path`")
            created = shapes.add_picture(p["path"], alt_text=p.get("alt_text"), **geom)
        else:
            raise ToolError(f"invalid_args: unknown shape kind {kind!r}")
        return {"ok": True, **created.to_dict()}

    if op in ("shape_move", "shape_resize", "shape_delete", "set_alt"):
        sh = _resolve_shape(deck, p.get("anchor_id"))
        if op == "shape_move":
            _require(
                p.get("left") is not None or p.get("top") is not None,
                "edit op='shape_move' requires `left`/`top`",
            )
            sh.move(left=p.get("left"), top=p.get("top"))
            return {"ok": True, "anchor_id": sh.anchor_id, "geometry": sh.geometry()}
        if op == "shape_resize":
            _require(
                p.get("width") is not None or p.get("height") is not None,
                "edit op='shape_resize' requires `width`/`height`",
            )
            sh.resize(width=p.get("width"), height=p.get("height"))
            return {"ok": True, "anchor_id": sh.anchor_id, "geometry": sh.geometry()}
        if op == "set_alt":
            _require(p.get("alt_text") is not None, "edit op='set_alt' requires `alt_text`")
            sh.set_alt_text(p["alt_text"])
            return {"ok": True, "anchor_id": sh.anchor_id, "alt_text": sh.alt_text}
        # shape_delete
        info = {"anchor_id": sh.anchor_id, "name": sh.name, "id": sh.shape_id}
        sh.delete()
        return {"ok": True, **info}

    # -- tables (addressed by the table shape's anchor_id) -----------------
    if op in ("table_add_row", "table_delete_row"):
        table = _resolve_shape(deck, p.get("anchor_id")).table
        if op == "table_add_row":
            table.add_row(p.get("values"))
        else:
            _require(p.get("row") is not None, "edit op='table_delete_row' requires `row`")
            table.delete_row(p["row"])
        return {"ok": True, "anchor_id": table.shape.anchor_id, "rows": table.row_count}

    # -- charts (addressed by the chart shape's anchor_id) -----------------
    if op == "chart_set_type":
        chart = _resolve_shape(deck, p.get("anchor_id")).chart
        _require(p.get("chart_type") is not None, "edit op='chart_set_type' requires `chart_type`")
        chart.set_type(p["chart_type"])
        return {"ok": True, "anchor_id": chart.shape.anchor_id, "chart_type": chart.chart_type}
    if op == "chart_set_data":
        chart = _resolve_shape(deck, p.get("anchor_id")).chart
        _require(
            p.get("categories") is not None and p.get("series") is not None,
            "edit op='chart_set_data' requires `categories` and `series`",
        )
        chart.set_data(p["categories"], p["series"])
        return chart.read()

    raise ToolError(f"invalid_args: unknown edit op {op!r}")


def _render_core(ppt: Any, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Render-to-image + the one deliberate view move (`navigate`)."""
    deck = _pick_deck(ppt, p.get("doc"))
    fmt = p.get("fmt") or "png"
    if op == "slide_image":
        _require(p.get("slide") is not None, "render op='slide_image' requires `slide`")
        path = deck.slides[p["slide"]].export_image(
            p.get("out"), width=p.get("width"), height=p.get("height"), fmt=fmt
        )
        return {"ok": True, "slide": p["slide"], "path": str(path), "format": fmt}
    if op == "shape_image":
        sh = _resolve_shape(deck, p.get("anchor_id"))
        path = sh.export_image(p.get("out"), fmt=fmt)
        return {"ok": True, "anchor_id": sh.anchor_id, "path": str(path), "format": fmt}
    if op == "navigate":
        _require(p.get("anchor_id") is not None, "render op='navigate' requires `anchor_id`")
        anchor = deck.anchor_by_id(p["anchor_id"])
        deck.go_to(anchor, select=p.get("select", True))
        return {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind}
    raise ToolError(f"invalid_args: unknown render op {op!r}")


def _show_core(deck: Presentation, op: str, p: dict[str, Any]) -> dict[str, Any]:
    """Live slide-show control (deliberately drives the user's screen)."""
    sh = deck.show
    if op == "state":
        return sh.state()
    if op == "start":
        return sh.start(from_slide=p.get("slide"))
    if op == "end":
        return sh.end()
    if op == "next":
        return sh.next()
    if op == "previous":
        return sh.previous()
    if op == "goto":
        _require(p.get("slide") is not None, "show op='goto' requires `slide`")
        return sh.goto(p["slide"])
    if op == "black":
        return sh.black()
    if op == "white":
        return sh.white()
    if op == "resume":
        return sh.resume()
    raise ToolError(f"invalid_args: unknown show op {op!r}")


# ===========================================================================
# Public tools — the typed schema the agent sees. Each wraps a core in
# attach() (+ an edit() fence for mutations).
# ===========================================================================


def ppt_read(
    op: Literal[
        "status", "slides", "outline", "slide", "anchor", "selection", "table", "chart", "layouts"
    ],
    anchor_id: str | None = None,
    slide: int | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Read the deck — always side-effect-free, never moves the user's view. `op`:
    - "status": open presentations, which is active, and the slide in view. Start here.
    - "slides": every slide — index (1-based), id, layout, title, shape count, has-notes.
    - "outline": each slide's title + body bullets (the fastest read of the deck's text).
    - "slide": one slide in full — every shape with its anchor_id, name, id, type,
      geometry (points), text, and whether it holds a table/chart. Needs `slide`.
      This is how you discover the anchor_ids to target with `ppt_edit`.
    - "anchor": the text of any text anchor (`anchor_id`): `ph:S:KIND` (placeholder,
      e.g. ph:2:title), `shape:S:N` (Nth shape by z-order), `para:S:N:P`,
      `cell:S:N:R:C` (table cell), `notes:S`, or `here:` (the user's selection).
      Returns text plus, for multi-paragraph anchors, a `paragraphs` breakdown.
    - "selection": what the user currently has selected, resolved to anchors.
    - "table" / "chart": the grid / chart data of the table-or-chart shape at
      `anchor_id` (a `shape:S:N`).
    - "layouts": the layout names that `ppt_edit` slide_add/set_layout accept.

    `doc` targets a presentation by name (default: the active one)."""
    with _mcp_errors(), attach() as ppt:
        return _read_core(ppt, op, {"anchor_id": anchor_id, "slide": slide, "doc": doc})


def ppt_edit(
    op: Literal[
        "write",
        "format",
        "slide_add",
        "slide_delete",
        "slide_duplicate",
        "slide_move",
        "set_layout",
        "shape_add",
        "shape_move",
        "shape_resize",
        "shape_delete",
        "set_alt",
        "table_add_row",
        "table_delete_row",
        "chart_set_type",
        "chart_set_data",
    ],
    anchor_id: str | None = None,
    text: str | None = None,
    mode: Literal["set", "insert_after", "insert_before"] = "set",
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
    slide: int | None = None,
    to: int | None = None,
    layout: str | None = None,
    index: int | None = None,
    kind: Literal["textbox", "shape", "picture", "table", "chart"] | None = None,
    shape_type: str = "rectangle",
    path: str | None = None,
    rows: int | None = None,
    cols: int | None = None,
    chart_type: str | None = None,
    categories: list[str] | None = None,
    series: dict[str, list[float]] | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    alt_text: str | None = None,
    values: list[str] | None = None,
    row: int | None = None,
    doc: str | None = None,
) -> dict[str, Any]:
    """Mutate the deck. Every call is ONE undo entry and preserves the user's view.
    Geometry is in points (1 inch = 72 pt). `op`:

    Text & formatting (target a text anchor via `anchor_id`):
    - "write": write `text`. mode="set" replaces the whole anchor (embed `\\n` for
      multiple paragraphs); "insert_after"/"insert_before" add a paragraph instead.
    - "format": font (`bold`/`italic`/`underline`/`size`/`font`/`color`), paragraph
      (`alignment`/`space_before`/`space_after`/`line_spacing` in points/`indent_level`
      1-5), and/or list (`list_type` "bulleted"/"numbered", or "none" to strip;
      `bullet_char` for a custom bullet). PowerPoint has no named styles, so this
      direct formatting is its "apply a style". Pass at least one option.

    Slide lifecycle:
    - "slide_add": add a slide (`layout` name, optional 1-based `index`; default end).
    - "slide_delete" / "slide_duplicate": delete / duplicate slide `slide`.
    - "slide_move": move slide `slide` to position `to`.
    - "set_layout": re-apply layout `layout` to slide `slide`.

    Shapes (create on `slide`; move/resize/delete/tag by `anchor_id`):
    - "shape_add": add `kind`="textbox" (with `text`), "shape" (autoshape via
      `shape_type`, e.g. "star", optional `text`), "picture" (`path`, optional
      `alt_text`), "table" (`rows`+`cols`), or "chart" (`chart_type`, optional
      `categories`+`series`). Optional `left`/`top`/`width`/`height`.
    - "shape_move": move to absolute `left`/`top`. "shape_resize": set `width`/`height`.
    - "shape_delete": delete it. "set_alt": set `alt_text` (a drift-proof handle).

    Tables & charts (target the table/chart shape by its `anchor_id`, a shape:S:N):
    - "table_add_row": append a row, optionally filled from `values`.
    - "table_delete_row": delete 1-based `row`.
    - "chart_set_type": change chart kind to `chart_type` (e.g. "line"/"pie"/"bar").
    - "chart_set_data": replace `categories` + `series` (a {name:[values]} map).

    To edit a table cell's text, write to its `cell:S:N:R:C` anchor with op="write".
    `doc` targets a presentation by name (default: the active one)."""
    params = {
        "anchor_id": anchor_id,
        "text": text,
        "mode": mode,
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
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "alt_text": alt_text,
        "values": values,
        "row": row,
    }
    with _mcp_errors(), attach() as ppt:
        deck = _pick_deck(ppt, doc)
        with deck.edit(f"MCP: {op}"):
            return _edit_core(deck, op, params)


def ppt_render(
    op: Literal["slide_image", "shape_image", "navigate"],
    slide: int | None = None,
    anchor_id: str | None = None,
    out: str | None = None,
    width: int | None = None,
    height: int | None = None,
    fmt: Literal["png", "jpg", "jpeg", "gif", "bmp"] = "png",
    select: bool = True,
    doc: str | None = None,
) -> dict[str, Any]:
    """Render to an image, or move the user's view. `op`:
    - "slide_image": render slide `slide` (1-based) to an image file; return its
      absolute path — so a vision model can *see* the slide it just built
      (render -> look -> iterate). Renders the current unsaved state; polite (does
      not move the view). `out` defaults to a temp file; pass `width`/`height`
      (pixels; one is enough — the other follows the aspect ratio).
    - "shape_image": render *just* the shape at `anchor_id` (cropped to its bounds,
      native pixel size) — so a vision model can see one picture/diagram alone.
      Polite. `out` defaults to a temp file.
    - "navigate": move the user's view to `anchor_id`'s slide — a deliberate,
      opt-in view move (every other tool leaves the view alone). With `select=True`
      (default), also selects the target shape. Use only when asked to be taken
      somewhere.

    `fmt` is the image format. `doc` targets a presentation by name."""
    params = {
        "slide": slide,
        "anchor_id": anchor_id,
        "out": out,
        "width": width,
        "height": height,
        "fmt": fmt,
        "select": select,
        "doc": doc,
    }
    with _mcp_errors(), attach() as ppt:
        return _render_core(ppt, op, params)


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
        return _show_core(_pick_deck(ppt, doc), op, {"slide": slide})


def ppt_batch(
    commands: list[dict[str, Any]],
    doc: str | None = None,
    atomic: bool = True,
    stop_on_error: bool = True,
) -> dict[str, Any]:
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

    Returns `{"ok": <all succeeded>, "atomic", "count", "results": [...]}` where each
    result is `{"index", "tool", "op", "ok", "result"}` on success or
    `{..., "ok": false, "error": <category>, "message"}` on failure (same category
    tokens as the other tools' ToolErrors)."""
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
        with scope:
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
                    entry.update(ok=True, result=result)
                except (PptliveError, ToolError) as exc:
                    code = _error_code(exc) if isinstance(exc, PptliveError) else "invalid_args"
                    entry.update(ok=False, error=code, message=str(exc))
                    results.append(entry)
                    if stop_on_error:
                        break
                    continue
                results.append(entry)
        return {
            "ok": all(r["ok"] for r in results),
            "atomic": atomic,
            "count": len(results),
            "results": results,
        }


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

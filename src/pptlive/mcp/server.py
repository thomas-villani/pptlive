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
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Literal

from mcp.server.fastmcp import FastMCP
from mcp.server.fastmcp.exceptions import ToolError
from mcp.types import CallToolResult, ImageContent, TextContent

from .. import attach
from .._batch import (
    EditOp,
    ReadOp,
    RenderOp,
    ShowOp,
    _edit_core,
    _error_code,
    _pick_deck,
    _read_core,
    _render_core,
    _require,
    _show_core,
    run_batch,
)
from .._guide import skill_body
from ..exceptions import BatchOpError, PptliveError

# ---------------------------------------------------------------------------
# Error mapping — the string analog of cli/main.py's _exit_for exit codes.
# ---------------------------------------------------------------------------


@contextmanager
def _mcp_errors() -> Iterator[None]:
    """Re-raise a PptliveError as a ToolError carrying its taxonomy category.

    Wraps the whole `with attach() as ppt: ...` body, so an attach-time
    `PowerPointNotRunningError` is mapped too.
    """
    try:
        yield
    except BatchOpError as exc:
        # A dispatch-layer invalid-args error — its message already carries the
        # "invalid_args: …" text the agent branches on; surface it verbatim.
        raise ToolError(str(exc)) from exc
    except PptliveError as exc:
        raise ToolError(f"{type(exc).__name__} ({_error_code(exc)}): {exc}") from exc
    except (ValueError, FileNotFoundError) as exc:
        # Library-level input validation (e.g. a line_spacing multiple > 5, an
        # out-of-range indent level, a missing picture/picture-fill path) — surface
        # as invalid_args, not a 500.
        raise ToolError(f"invalid_args: {exc}") from exc


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
# Public tools — the typed schema the agent sees. Each wraps a `_<tool>_core`
# (from `_dispatch`) in attach() (+ an edit() fence for mutations); `op` is typed
# by the enum so FastMCP derives the schema the agent sees.
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
    - "geometry": a spatial sanity-check of a slide (needs `slide`) — the slide size
      (points), each shape's bounding `box` + an `off_slide` flag, and the list of
      `overlaps` (shape pairs whose boxes intersect, biggest first). Run it after
      placing shapes to catch overlaps / off-edge shapes *without* a render. Axis-
      aligned boxes; rotation isn't accounted for (each shape carries its `rotation`).
    - "animations": slide `slide`'s shape animations in play order — each a
      `{seq_index, shapeid, shape, effect, exit, trigger, duration, delay}` mapping
      an effect back to its target shape. Needs `slide`. Empty if the slide has none.
    - "sections": the deck's sections (named slide spans) in order — each a
      `{index, name, first_slide, slide_count}`. Address a section for edit by its
      1-based `index` (ppt_edit section_rename/section_delete/section_move).
    - "headers_footers": the footer / slide-number / date settings — pass `slide`
      for that slide's, or omit it for the deck-wide master default. Returns
      `{footer:{visible,text}, slide_number:{visible}, date:{visible,text,format,
      use_format}, display_on_title_slide}` (text/use_format read null when hidden).
    - "anchor": the text of any text anchor (`anchor_id`): `ph:S:KIND` (placeholder,
      e.g. ph:2:title), `shape:S:N` (Nth shape by z-order), `para:S:N:P`,
      `cell:S:N:R:C` (table cell), `notes:S`, or `here:` (the user's selection).
      Returns text plus a `paragraphs` breakdown — each paragraph carries its
      effective `font` (`bold`/`italic`/`underline` as true/false/"mixed", `size`,
      `font` name, `color` `#RRGGBB` or null for a theme/auto color, plus
      `color_source` "direct"/"theme"/"mixed" and `theme_color` the inherited slot
      name when themed). `color_source` is the "is this color set on the run or
      cascaded from the theme/master?" tell — so a surprise color traces to its
      origin. The other font attrs are still *rendered* values (COM resolves the
      cascade before we see them and exposes no directly-set flag beyond color).
    - "text_frame_status": autofit diagnostics for the shape at `anchor_id` —
      `{autosize, word_wrap, margins:{left,right,top,bottom}, overflow_risk}`. The
      read to run when text looks clipped or overflowing: `overflow_risk` is
      "possible" when autosize is off (text can clip), "low" when an autofit mode
      is active. A coarse, mode-derived flag (no measured extent).
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
    paragraphs: list[Any] | None = None,
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
    space_before_lines: float | None = None,
    space_after_lines: float | None = None,
    line_spacing: float | None = None,
    line_spacing_points: float | None = None,
    indent_level: int | None = None,
    force: bool = False,
    fill_color: str | None = None,
    line_color: str | None = None,
    line_width: float | None = None,
    fill_transparency: float | None = None,
    line_transparency: float | None = None,
    colors: list[Any] | None = None,
    positions: list[float] | None = None,
    gradient_style: str | None = None,
    variant: int | None = None,
    degree: float | None = None,
    preset: str | None = None,
    pattern: str | None = None,
    fore: str | None = None,
    back: str | None = None,
    shadow: Any | None = None,
    glow: Any | None = None,
    soft_edge: int | None = None,
    reflection: int | None = None,
    dash: str | None = None,
    begin_arrow: str | None = None,
    end_arrow: str | None = None,
    begin_arrow_size: str | None = None,
    end_arrow_size: str | None = None,
    url: str | None = None,
    screen_tip: str | None = None,
    effect: str | None = None,
    duration: float | None = None,
    advance_after: float | None = None,
    advance_on_click: bool | None = None,
    follow_master: bool = False,
    order: Literal["front", "back", "forward", "backward"] | None = None,
    trigger: Literal["on_click", "with_previous", "after_previous"] | None = None,
    delay: float | None = None,
    exit: bool = False,
    section: int | None = None,
    before_slide: int | None = None,
    delete_slides: bool = False,
    footer_text: str | None = None,
    footer_visible: bool | None = None,
    slide_number_visible: bool | None = None,
    date_visible: bool | None = None,
    date_text: str | None = None,
    date_format: int | None = None,
    list_type: Literal["bulleted", "numbered", "none"] | None = None,
    bullet_char: str | None = None,
    slide: int | None = None,
    to: int | None = None,
    layout: str | None = None,
    index: int | None = None,
    placeholders: dict[str, dict[str, float]] | None = None,
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
    - "set_paragraphs": replace the anchor's text with `paragraphs` — a list whose
      items are strings or `{"text", "list_type"?, "indent_level"?, "alignment"?,
      "line_spacing"?/"line_spacing_points"?, "size"?, "bold"?, ...}` objects. Each
      item becomes exactly one bullet/paragraph (a newline inside an item is a soft
      break, never a split) with its formatting applied — the safe way to author a
      list without relying on `\\n` inference. Returns the new `para:` anchor ids.
    - "find_replace": fuzzy-locate `find` across the deck and rewrite the matched
      spans with `text` (only the span changes, so run formatting is preserved).
      Scope with `scope` (a `slide:S` / anchor id). One match auto-applies; for
      several pass `replace_all=true` or `occurrence` (1-based) — otherwise it
      errors `ambiguous`. Zero matches errors `not_found`. Use op="find" first to
      preview the hits.
    - "format": font (`bold`/`italic`/`underline`/`size`/`font`/`color` — `color` is
      *font* color), paragraph, shape fill/border on a shape anchor
      (`fill_color`/`line_color` — a hex or "none" for transparent/no border;
      `line_width` in points; `fill_transparency`/`line_transparency` — a 0..1 alpha,
      0 opaque, 1 fully transparent, the partial-alpha knob distinct from "none"),
      and/or list (`list_type` "bulleted"/"numbered", or
      "none" to strip; `bullet_char` for a custom bullet). Paragraph spacing is
      **unit-explicit** — `line_spacing` is a MULTIPLE (1.5), `line_spacing_points`
      is EXACT POINTS (24); `space_before`/`space_after` are points and
      `space_before_lines`/`space_after_lines` are multiples (pass one of each pair).
      A `line_spacing` multiple > 5 is rejected unless `force=true` (it's almost
      always a points-vs-multiple mix-up — use `line_spacing_points`). Plus
      `alignment` and `indent_level` (1-5). PowerPoint has no named styles, so this
      direct formatting is its "apply a style". Pass at least one option.
    - "text_reset_format": reset the anchor's paragraph spacing to clean
      single-spaced defaults (zero before/after, indent 1) — the recovery for a
      line-spacing spiral that pushed text off the slide. Does NOT reset font
      size/typeface (PowerPoint has no inheritance primitive); for a placeholder's
      geometry + default font size use "shape_reset_layout".

    Slide lifecycle:
    - "slide_add": add a slide (`layout` name, optional 1-based `index`; default end).
      Optional `placeholders` repositions the layout's placeholders in the same op —
      `{KIND: {left, top, width, height}}` (points, any subset), KIND as in `ph:S:KIND`
      (e.g. {"body": {"left": 40, "width": 440}} for a left-half content area beside a
      right-side panel). Saves the add-then-resize fix-up; returns each adjusted
      placeholder's resulting geometry. Read op="geometry" gives the slide size to size from.
    - "slide_delete" / "slide_duplicate": delete / duplicate slide `slide`.
    - "slide_move": move slide `slide` to position `to`.
    - "set_layout": re-apply layout `layout` to slide `slide`.
    - "slide_set_transition": set slide `slide`'s entrance transition — `effect`
      (e.g. "fade"/"cut"/"dissolve"/"cover_left"), `duration` (seconds),
      `advance_after` (auto-advance after N seconds), `advance_on_click` (bool).
      Pass at least one.
    - "slide_set_background": give slide `slide` its own solid `color` background
      (overriding the master), or pass `follow_master`=true to revert to the master
      background. Exactly one of the two.

    Sections (named spans of slides; addressed by 1-based `section` index — read
    them with ppt_read op="sections"):
    - "section_add": add a section named `name`. `before_slide` (1-based) is the
      slide it starts at (the natural form); adding the first section in front of a
      later slide auto-creates a leading "Default Section". Omit `before_slide` to
      append an empty trailing section.
    - "section_rename": rename section `section` to `name`.
    - "section_delete": delete section `section`. Keeps its slides by default (drops
      only the boundary); pass `delete_slides`=true to delete the slides too.
    - "section_move": move section `section` (and the slides it spans) to position `to`.

    Headers / footers (footer text, auto slide number, date — set on one slide via
    `slide`, or the deck-wide master default when `slide` is omitted; read with
    ppt_read op="headers_footers"):
    - "set_headers_footers": set any of `footer_text` / `footer_visible`,
      `slide_number_visible`, `date_visible` / `date_text` (a fixed date) /
      `date_format` (a raw PpDateTimeFormat int for an auto-updating date — mutually
      exclusive with `date_text`). Setting footer/date text auto-shows it. Pass
      `slide` for a per-slide override; omit it to set the master default for every
      inheriting slide.

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
    - "shape_reset_layout": restore a placeholder's geometry (and default font
      size) from its slide layout — the fix for a placeholder manually moved/resized
      or shrunk to an unreadable font. Errors if the shape isn't a placeholder.
    - "shape_delete": delete it. "set_alt": set `alt_text` (a drift-proof handle).
      Address a shape that must survive a delete/restack by `shapeid:S:ID` (the
      stable `id` from any shape listing), not the positional `shape:S:N`.
    - "shape_set_hyperlink": make the shape a clickable link — pass EXACTLY one of
      `url` (external URL/file/mailto) or `slide` (1-based in-deck jump, e.g. a
      "back to agenda" button); optional `screen_tip` hover text. A shape needs no
      text frame to carry a link. "shape_remove_hyperlink": clear the link.

    Animations (an entrance/exit effect on a shape; play order is add order):
    - "shape_animate": animate the shape at `anchor_id`. `effect` is the animation
      ("fade"/"appear"/"fly_in"/"float_in"/"wipe"/"zoom"/"grow_turn"/"swivel"/
      "wheel"/"split"); `trigger` is when it fires ("on_click" default/
      "with_previous"/"after_previous"); `duration` (seconds) and `delay` (seconds)
      tune the timing; `exit`=true animates the shape OUT instead of in (the
      "disappear" case). Each call adds one effect (a shape can have several).
    - "shape_clear_animations": remove every animation targeting the shape at
      `anchor_id`. "slide_clear_animations": wipe ALL animations on slide `slide`.
      Use ppt_read op="animations" (needs `slide`) to see a slide's effects first.

    Advanced fills & effects (target the shape by `anchor_id`; distinct from the
    solid `fill_color`/`line_color` on op="format"):
    - "shape_gradient_fill": gradient fill. Pass `colors` (list of hex/`[r,g,b]`) —
      one=one-color (optional `degree` 0..1 brightness), two=two-color, three+=
      multi-stop with optional `positions` (floats 0..1 placing the interior stops) —
      OR `preset` (a named ramp: "ocean"/"fire"/"rainbow"/…). `gradient_style`
      ("horizontal"/"vertical"/"diagonal_up"/…) and `variant` (1-4) set the sweep.
    - "shape_picture_fill": fill with an image at `path` (resolved to absolute).
    - "shape_pattern_fill": two-color pattern — `pattern` (e.g. "percent_50",
      "trellis", "dark_horizontal"), `fore` color, optional `back` color.
    - "shape_set_effect": shadow / glow / soft-edge / reflection. `shadow` and `glow`
      are objects ({color?, transparency?, blur?, size?, offset_x?, offset_y?} /
      {color?, radius?, transparency?}); `soft_edge` is a 0-6 int preset and
      `reflection` a 0-9 int (0 = off). Pass "none" for `shadow`/`glow` to remove it.
      Active effects read back under each shape's `effects`.
    - "shape_line_style": line `dash` ("solid"/"dash"/"round_dot"/"dash_dot"/
      "long_dash"/…) and/or arrowheads. `begin_arrow`/`end_arrow` are styles
      ("none"/"triangle"/"open"/"stealth"/"diamond"/"oval"); `begin_arrow_size`/
      `end_arrow_size` are "small"/"medium"/"large". Arrowheads apply to
      lines/connectors only (a closed shape errors — use `dash` there). Reads back
      under each shape's `line` (`dash`/`begin_arrow`/`end_arrow`).

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
        "paragraphs": paragraphs,
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
        "space_before_lines": space_before_lines,
        "space_after_lines": space_after_lines,
        "line_spacing": line_spacing,
        "line_spacing_points": line_spacing_points,
        "indent_level": indent_level,
        "force": force,
        "fill_color": fill_color,
        "line_color": line_color,
        "line_width": line_width,
        "fill_transparency": fill_transparency,
        "line_transparency": line_transparency,
        "colors": colors,
        "positions": positions,
        "gradient_style": gradient_style,
        "variant": variant,
        "degree": degree,
        "preset": preset,
        "pattern": pattern,
        "fore": fore,
        "back": back,
        "shadow": shadow,
        "glow": glow,
        "soft_edge": soft_edge,
        "reflection": reflection,
        "dash": dash,
        "begin_arrow": begin_arrow,
        "end_arrow": end_arrow,
        "begin_arrow_size": begin_arrow_size,
        "end_arrow_size": end_arrow_size,
        "url": url,
        "screen_tip": screen_tip,
        "effect": effect,
        "duration": duration,
        "advance_after": advance_after,
        "advance_on_click": advance_on_click,
        "follow_master": follow_master,
        "order": order,
        "trigger": trigger,
        "delay": delay,
        "exit": exit,
        "section": section,
        "before_slide": before_slide,
        "delete_slides": delete_slides,
        "footer_text": footer_text,
        "footer_visible": footer_visible,
        "slide_number_visible": slide_number_visible,
        "date_visible": date_visible,
        "date_text": date_text,
        "date_format": date_format,
        "list_type": list_type,
        "bullet_char": bullet_char,
        "slide": slide,
        "to": to,
        "layout": layout,
        "index": index,
        "placeholders": placeholders,
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
      ~1000 px when embedding). For an exact size pass `width`/`height` instead (one
      is enough — the other follows the aspect ratio); it overrides `max_dim`, and
      passing both is an error. (PowerPoint has no JPEG-quality knob, so pixel
      dimensions are the only render-cost lever.) `slides` selects what to render: a
      single 1-based slide ("3") or an inclusive span ("2-4"); omit for the whole
      deck. Each slide comes back as a "slide N" label + image block; the structured
      result lists the written file `path`s. Polite (does not move the view).
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
    # Default the embed cap only when the caller gave no explicit size at all —
    # an explicit width/height is the other (conflicting) size lever.
    if op == "deck_snapshot" and embed and max_dim is None and width is None and height is None:
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
    follow_view: bool | None = None,
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

    `follow_view` ("follow the work", default on): when this batch ADDS a slide, the
    user's view ends on the last slide the batch built instead of snapping back to
    the pre-batch slide — so an authoring session isn't yanked back to slide 1 after
    every batch. Pure-edit batches (no slide added) still preserve the view. Pass
    False to force the polite snap-back; leave unset to take the `PPTLIVE_VIEW_FOLLOW`
    env default. A deliberate `render navigate` / `show` op always wins over follow.

    Returns `{"ok": <all succeeded>, "atomic", "count", "results": [...]}` where each
    result is `{"index", "tool", "op", "ok", "result"}` on success or
    `{..., "ok": false, "error": <category>, "message"}` on failure (same category
    tokens as the other tools' ToolErrors). When `embed` surfaces images, the reply
    carries those image blocks alongside this summary as its structured content."""
    with _mcp_errors(), attach() as ppt:
        _require(
            isinstance(commands, list) and len(commands) > 0,
            "ppt_batch requires a non-empty `commands` list",
        )
        deck = _pick_deck(ppt, doc)
        # Inline-embed defaults: a render slide_image/deck_snapshot in an embedding
        # batch gets a sensible size unless the command set one. Inject into the
        # command copies so the shared `run_batch` stays render-agnostic.
        prepared: list[dict[str, Any]] = []
        for cmd in commands:
            c = dict(cmd)
            tool, op = c.get("tool", "edit"), c.get("op")
            if embed and tool == "render":
                if op == "slide_image" and c.get("width") is None and c.get("height") is None:
                    c["width"] = _EMBED_DEFAULT_WIDTH
                elif op == "deck_snapshot" and c.get("max_dim") is None:
                    c["max_dim"] = _EMBED_DEFAULT_MAX_DIM
            prepared.append(c)
        results = run_batch(
            ppt,
            deck,
            prepared,
            doc=doc,
            atomic=atomic,
            stop_on_error=stop_on_error,
            follow_view=follow_view,
            label=f"MCP: batch ({len(commands)} ops)",
        )
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

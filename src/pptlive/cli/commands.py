"""CLI subcommands wired against the pptlive library.

v0 surface: status, slides, outline, slide read, shapes, read (anchor/notes),
write, replace, go-to. v0.1 adds the slide-lifecycle verbs under the `slide`
group: add, delete, duplicate, move, set-layout, and layouts. v0.2 adds the
`shape` group: add (textbox/shape/picture), move, resize, delete. v0.3 adds text
structure: paragraphs, insert, format-paragraph, format-text, and the `list`
group (apply/remove). v0.4 adds `slide export` (render a slide to an image) and
`selection` (what the user has selected, resolved to anchors; targetable via
`here:`). v0.5 adds the `table` group (read/add-row/delete-row) plus `shape add
--kind table`; cells are `cell:S:N:R:C` anchors. v0.6 adds the `show` group
(start/end/next/prev/goto/black/white/resume/state) for live slide-show control.
v0.7 adds `shape export` (render one shape to an image), `shape set-alt`, and
`--alt-text` on `shape add` (a drift-proof re-identification handle).
find/replace arrives in a later stage.
"""

from __future__ import annotations

import json
from typing import Any

import click

from .. import attach
from .._presentation import Presentation
from .._shapes import Shape
from ..constants import (
    ALIGNMENT_CHOICES,
    AUTOSHAPE_CHOICES,
    IMAGE_FORMAT_CHOICES,
    LIST_TYPE_CHOICES,
    SHAPE_IMAGE_FORMAT_CHOICES,
)
from ..exceptions import AnchorNotFoundError, PowerPointNotRunningError
from .main import _run, emit


def _pick_deck(ppt: Any, doc_name: str | None) -> Presentation:
    if doc_name is None:
        return ppt.presentations.active
    return ppt.presentations[doc_name]


# ---------------------------------------------------------------------------
# Text formatters (used when --text is selected)
# ---------------------------------------------------------------------------


def _fmt_status(info: dict[str, Any]) -> str:
    decks = info.get("decks") or []
    if not decks:
        return "(no presentations open)"
    width = max(len(str(d.get("name", ""))) for d in decks)
    lines = []
    for d in decks:
        marker = "*" if d.get("is_active") else " "
        lines.append(f"{marker} {str(d.get('name', '')):<{width}}  {d.get('path', '')}")
    viewed = info.get("viewed_slide")
    if viewed is not None:
        lines.append(f"viewing slide {viewed}")
    return "\n".join(lines)


def _fmt_slides(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no slides)"
    lines = []
    for r in rows:
        title = r.get("title") or "(untitled)"
        layout = r.get("layout") or "?"
        notes = " [notes]" if r.get("has_notes") else ""
        lines.append(
            f"{r.get('index'):>3}. {title}  [{layout}, {r.get('shape_count', 0)} shapes]{notes}"
        )
    return "\n".join(lines)


def _fmt_outline(items: list[dict[str, Any]]) -> str:
    if not items:
        return "(no slides)"
    lines = []
    for it in items:
        title = it.get("title") or "(untitled)"
        lines.append(f"{it.get('slide'):>3}. {title}")
        for bullet in it.get("bullets") or []:
            lines.append(f"      • {bullet}")
    return "\n".join(lines)


def _fmt_shapes(shapes: list[dict[str, Any]]) -> str:
    if not shapes:
        return "(no shapes)"
    lines = []
    for s in shapes:
        ph = f" ph={s['placeholder']}" if s.get("placeholder") else ""
        alt = f" alt={s['alt_text']!r}" if s.get("alt_text") else ""
        text = s.get("text")
        snippet = ""
        if text:
            flat = text.replace("\r", " / ").replace("\n", " / ").replace("\v", " ")
            snippet = "  " + (flat if len(flat) <= 60 else flat[:57] + "…")
        lines.append(
            f"[{s['anchor_id']}] {s.get('name', '')!r} ({s.get('type', '?')}{ph}{alt}){snippet}"
        )
    return "\n".join(lines)


def _fmt_slide_read(grid: dict[str, Any]) -> str:
    head = (
        f"slide {grid.get('index')} (id {grid.get('id')}) "
        f"layout={grid.get('layout') or '?'} title={grid.get('title') or '(untitled)'!r}"
    )
    return head + "\n" + _fmt_shapes(grid.get("shapes") or [])


def _fmt_layouts(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no layouts)"
    return "\n".join(f"{r.get('index'):>3}. {r.get('name')}" for r in rows)


def _fmt_geometry(geo: dict[str, float] | None) -> str:
    if not geo:
        return "(no geometry)"
    return (
        f"left={geo['left']:g} top={geo['top']:g} width={geo['width']:g} height={geo['height']:g}"
    )


def _fmt_paragraphs(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no paragraphs)"
    lines = []
    for r in rows:
        bullet = "" if r.get("bullet") in (None, "none") else f" •{r['bullet']}"
        lvl = r.get("indent_level")
        indent = "  " * (int(lvl) - 1) if isinstance(lvl, int) and lvl > 1 else ""
        lines.append(f"[{r['anchor_id']}] {indent}{r.get('text', '')!r}{bullet}")
    return "\n".join(lines)


def _fmt_table_read(grid: dict[str, Any]) -> str:
    head = (
        f"table at {grid.get('anchor_id')} (slide {grid.get('slide')}, "
        f"{grid.get('rows')}x{grid.get('columns')})"
    )
    lines = [head]
    for row in grid.get("cells") or []:
        cells = [str(cell.get("text", "")).replace("\r", " / ").replace("\v", " ") for cell in row]
        lines.append("  | " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _fmt_selection(info: dict[str, Any]) -> str:
    kind = info.get("type")
    slide = info.get("slide")
    if kind == "none":
        return "(nothing selected)"
    if kind == "slides":
        return f"slide {slide} selected"
    if kind == "shapes":
        shapes = info.get("shapes") or []
        names = ", ".join(f"{s.get('name')!r} [{s.get('anchor_id')}]" for s in shapes)
        return f"slide {slide}: {len(shapes)} shape(s) selected — {names}"
    if kind == "text":
        snippet = (info.get("text") or "").replace("\r", " / ").replace("\v", " ")
        return f"slide {slide}: text caret in {info.get('anchor_id')} — {snippet!r}"
    return str(info)


def _fmt_show(info: dict[str, Any]) -> str:
    if not info.get("running"):
        return "(no slide show running)"
    parts = [f"slide {info.get('current_slide')}/{info.get('slide_count')}"]
    state = info.get("state")
    if state and state != "running":
        parts.append(f"[{state}]")
    return "show running — " + " ".join(parts)


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(slides_cmd)
    group.add_command(outline)
    group.add_command(slide)
    group.add_command(shapes_cmd)
    group.add_command(shape)
    group.add_command(read)
    group.add_command(write)
    group.add_command(replace)
    group.add_command(paragraphs_cmd)
    group.add_command(insert)
    group.add_command(format_paragraph)
    group.add_command(format_text)
    group.add_command(list_cmd)
    group.add_command(table)
    group.add_command(selection_cmd)
    group.add_command(show)
    group.add_command(go_to)


# ---------------------------------------------------------------------------
# status
# ---------------------------------------------------------------------------


@click.command(name="status")
@click.pass_context
def status(ctx: click.Context) -> None:
    """List open presentations, which one is active, and the slide in view."""

    def go() -> None:
        try:
            with attach() as ppt:
                info = {
                    "decks": ppt.presentations.list(),
                    "viewed_slide": ppt.viewed_slide_index(),
                }
                emit(info, as_text=not ctx.obj["as_json"], text=_fmt_status(info))
        except PowerPointNotRunningError:
            info = {"decks": [], "viewed_slide": None}
            emit(info, as_text=not ctx.obj["as_json"], text=_fmt_status(info))
            raise

    _run(ctx, go)


# ---------------------------------------------------------------------------
# slides
# ---------------------------------------------------------------------------


@click.command(name="slides")
@click.pass_context
def slides_cmd(ctx: click.Context) -> None:
    """List every slide: index, id, layout, title, shape count, has-notes."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            rows = deck.slides.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_slides(rows))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# outline
# ---------------------------------------------------------------------------


@click.command(name="outline")
@click.pass_context
def outline(ctx: click.Context) -> None:
    """Print the title + body bullets of every slide (the Outline-view analog)."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            items = deck.outline()
            emit(items, as_text=not ctx.obj["as_json"], text=_fmt_outline(items))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# slide read S
# ---------------------------------------------------------------------------


@click.group(name="slide")
def slide() -> None:
    """Slide reads + lifecycle: read, add, delete, duplicate, move, set-layout, layouts."""


@slide.command(name="read")
@click.argument("index", type=int)
@click.pass_context
def slide_read(ctx: click.Context, index: int) -> None:
    """Read every shape on slide INDEX: anchor_id, name, type, geometry, text."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            grid = deck.slides[index].read()
            emit(grid, as_text=not ctx.obj["as_json"], text=_fmt_slide_read(grid))

    _run(ctx, go)


@slide.command(name="layouts")
@click.pass_context
def slide_layouts(ctx: click.Context) -> None:
    """List the deck's slide layouts (the names `add`/`set-layout` accept)."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            rows = deck.layouts()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_layouts(rows))

    _run(ctx, go)


@slide.command(name="export")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--out",
    "out",
    type=click.Path(dir_okay=False),
    default=None,
    help="Output image path (default: a temp file you can then read).",
)
@click.option("--width", type=int, default=None, help="Output width (pixels).")
@click.option("--height", type=int, default=None, help="Output height (pixels).")
@click.option(
    "--format",
    "fmt",
    type=click.Choice(IMAGE_FORMAT_CHOICES),
    default="png",
    show_default=True,
    help="Image format.",
)
@click.pass_context
def slide_export(
    ctx: click.Context,
    slide_index: int,
    out: str | None,
    width: int | None,
    height: int | None,
    fmt: str,
) -> None:
    """Render a slide to an image file — so a vision model can *see* it.

    Renders the slide's current (unsaved) state; polite (doesn't move the view).
    Prints the absolute path; pass `--width`/`--height` for a specific pixel size
    (one is enough — the other follows the slide's aspect ratio).
    """

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            path = deck.slides[slide_index].export_image(out, width=width, height=height, fmt=fmt)
            payload = {
                "ok": True,
                "slide": slide_index,
                "path": str(path),
                "format": fmt,
                "width": width,
                "height": height,
            }
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"exported slide {slide_index} -> {path}",
            )

    _run(ctx, go)


@slide.command(name="add")
@click.option(
    "--layout",
    default=None,
    help="Layout name (default: title_and_content). See `slide layouts` for names.",
)
@click.option(
    "--index", "index", type=int, default=None, help="1-based insertion position (default: end)."
)
@click.pass_context
def slide_add(ctx: click.Context, layout: str | None, index: int | None) -> None:
    """Add a slide; print its index, id, and layout."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            with deck.edit(f"CLI: add slide ({layout or 'default'})"):
                new = deck.slides.add(layout=layout, index=index)
            payload = {"ok": True, "index": new.index, "id": new.id, "layout": new.layout_name}
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"added slide {new.index} ({new.layout_name})",
            )

    _run(ctx, go)


@slide.command(name="delete")
@click.option(
    "--slide", "slide_index", type=int, required=True, help="1-based slide index to delete."
)
@click.pass_context
def slide_delete(ctx: click.Context, slide_index: int) -> None:
    """Delete a slide."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            target = deck.slides[slide_index]  # exit 2 here if out of range
            with deck.edit(f"CLI: delete slide {slide_index}"):
                target.delete()
            emit(
                {"ok": True, "deleted": slide_index},
                as_text=not ctx.obj["as_json"],
                text=f"deleted slide {slide_index}",
            )

    _run(ctx, go)


@slide.command(name="duplicate")
@click.option(
    "--slide", "slide_index", type=int, required=True, help="1-based slide index to duplicate."
)
@click.pass_context
def slide_duplicate(ctx: click.Context, slide_index: int) -> None:
    """Duplicate a slide; print the copy's index and id."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            target = deck.slides[slide_index]
            with deck.edit(f"CLI: duplicate slide {slide_index}"):
                new = target.duplicate()
            payload = {"ok": True, "index": new.index, "id": new.id, "from": slide_index}
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"duplicated slide {slide_index} -> {new.index}",
            )

    _run(ctx, go)


@slide.command(name="move")
@click.option(
    "--slide", "slide_index", type=int, required=True, help="1-based slide index to move."
)
@click.option("--to", "to_index", type=int, required=True, help="1-based destination position.")
@click.pass_context
def slide_move(ctx: click.Context, slide_index: int, to_index: int) -> None:
    """Move a slide to a new position."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            target = deck.slides[slide_index]
            with deck.edit(f"CLI: move slide {slide_index} -> {to_index}"):
                moved = target.move_to(to_index)
            payload = {"ok": True, "index": moved.index, "id": moved.id}
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"moved slide to position {moved.index}",
            )

    _run(ctx, go)


@slide.command(name="set-layout")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--layout", required=True, help="Layout name to apply. See `slide layouts` for names."
)
@click.pass_context
def slide_set_layout(ctx: click.Context, slide_index: int, layout: str) -> None:
    """Re-apply a slide's layout by name."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            target = deck.slides[slide_index]
            with deck.edit(f"CLI: set layout of slide {slide_index}"):
                target.set_layout(layout)
            payload = {"ok": True, "index": slide_index, "layout": target.layout_name}
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"slide {slide_index} layout -> {target.layout_name}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# shapes --slide S
# ---------------------------------------------------------------------------


@click.command(name="shapes")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.pass_context
def shapes_cmd(ctx: click.Context, slide_index: int) -> None:
    """List the shapes on a slide (anchor_id, name, id, type, geometry, text)."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            rows = deck.slides[slide_index].shapes.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_shapes(rows))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# shape add|move|resize|delete
# ---------------------------------------------------------------------------


@click.group(name="shape")
def shape() -> None:
    """Create + place shapes: add, move, resize, delete (geometry in points)."""


def _resolve_shape(deck: Presentation, anchor_id: str) -> Shape:
    """Resolve a shape/placeholder anchor to a `Shape` (else exit-2 not found)."""
    anchor = deck.anchor_by_id(anchor_id)
    if not isinstance(anchor, Shape):
        raise AnchorNotFoundError("shape", anchor_id)
    return anchor


@shape.command(name="add")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--kind",
    type=click.Choice(["textbox", "shape", "picture", "table"]),
    required=True,
    help="What to add.",
)
@click.option("--text", "text", default=None, help="Initial text (textbox/shape).")
@click.option("--rows", type=int, default=None, help="Row count (required for --kind table).")
@click.option("--cols", type=int, default=None, help="Column count (required for --kind table).")
@click.option(
    "--path",
    "path",
    type=click.Path(exists=True, dir_okay=False),
    default=None,
    help="Image file to embed (required for --kind picture).",
)
@click.option(
    "--shape-type",
    "shape_type",
    type=click.Choice(AUTOSHAPE_CHOICES),
    default="rectangle",
    show_default=True,
    help="Autoshape geometry (for --kind shape).",
)
@click.option("--left", type=float, default=None, help="Left edge (points).")
@click.option("--top", type=float, default=None, help="Top edge (points).")
@click.option("--width", type=float, default=None, help="Width (points).")
@click.option("--height", type=float, default=None, help="Height (points).")
@click.option(
    "--alt-text",
    "alt_text",
    default=None,
    help="Alternative text for a picture (a drift-proof re-identification handle).",
)
@click.pass_context
def shape_add(
    ctx: click.Context,
    slide_index: int,
    kind: str,
    text: str | None,
    rows: int | None,
    cols: int | None,
    path: str | None,
    shape_type: str,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    alt_text: str | None,
) -> None:
    """Add a shape to a slide; print its anchor_id, name, type, and geometry."""

    def go() -> None:
        if kind == "picture" and not path:
            raise click.UsageError("shape add --kind picture requires --path")
        if kind == "table" and (rows is None or cols is None):
            raise click.UsageError("shape add --kind table requires --rows and --cols")
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            shapes = deck.slides[slide_index].shapes  # exit 2 if slide out of range
            with deck.edit(f"CLI: add {kind} on slide {slide_index}"):
                if kind == "textbox":
                    new = shapes.add_textbox(
                        text or "", left=left, top=top, width=width, height=height
                    )
                elif kind == "shape":
                    new = shapes.add_shape(
                        shape_type, left=left, top=top, width=width, height=height
                    )
                elif kind == "table":
                    assert rows is not None and cols is not None  # guarded above
                    new = shapes.add_table(
                        rows, cols, left=left, top=top, width=width, height=height
                    )
                else:  # picture
                    assert path is not None  # guarded above
                    new = shapes.add_picture(
                        path, left=left, top=top, width=width, height=height, alt_text=alt_text
                    )
            payload = {"ok": True, **new.to_dict()}
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"added {payload['type']} {payload['name']!r} at {payload['anchor_id']}",
            )

    _run(ctx, go)


@shape.command(name="move")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to move (shape:S:N or ph:S:KIND)."
)
@click.option("--left", type=float, default=None, help="New left edge (points).")
@click.option("--top", type=float, default=None, help="New top edge (points).")
@click.pass_context
def shape_move(ctx: click.Context, anchor_id: str, left: float | None, top: float | None) -> None:
    """Move a shape to an absolute position (points)."""

    def go() -> None:
        if left is None and top is None:
            raise click.UsageError("shape move requires --left and/or --top")
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            sh = _resolve_shape(deck, anchor_id)  # exit 2 if missing / not a shape
            with deck.edit(f"CLI: move {anchor_id}"):
                sh.move(left=left, top=top)
            geo = sh.geometry()
            emit(
                {"ok": True, "anchor_id": sh.anchor_id, "geometry": geo},
                as_text=not ctx.obj["as_json"],
                text=f"moved {sh.anchor_id}: {_fmt_geometry(geo)}",
            )

    _run(ctx, go)


@shape.command(name="resize")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to resize (shape:S:N or ph:S:KIND)."
)
@click.option("--width", type=float, default=None, help="New width (points).")
@click.option("--height", type=float, default=None, help="New height (points).")
@click.pass_context
def shape_resize(
    ctx: click.Context, anchor_id: str, width: float | None, height: float | None
) -> None:
    """Set a shape's size (points)."""

    def go() -> None:
        if width is None and height is None:
            raise click.UsageError("shape resize requires --width and/or --height")
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            sh = _resolve_shape(deck, anchor_id)
            with deck.edit(f"CLI: resize {anchor_id}"):
                sh.resize(width=width, height=height)
            geo = sh.geometry()
            emit(
                {"ok": True, "anchor_id": sh.anchor_id, "geometry": geo},
                as_text=not ctx.obj["as_json"],
                text=f"resized {sh.anchor_id}: {_fmt_geometry(geo)}",
            )

    _run(ctx, go)


@shape.command(name="delete")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to delete (shape:S:N or ph:S:KIND)."
)
@click.pass_context
def shape_delete(ctx: click.Context, anchor_id: str) -> None:
    """Delete a shape from its slide."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            sh = _resolve_shape(deck, anchor_id)
            info = {"anchor_id": sh.anchor_id, "name": sh.name, "id": sh.shape_id}
            with deck.edit(f"CLI: delete {anchor_id}"):
                sh.delete()
            emit(
                {"ok": True, **info},
                as_text=not ctx.obj["as_json"],
                text=f"deleted {info['anchor_id']} ({info['name']!r})",
            )

    _run(ctx, go)


@shape.command(name="export")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to export (shape:S:N or ph:S:KIND)."
)
@click.option(
    "--out",
    "out",
    type=click.Path(dir_okay=False),
    default=None,
    help="Output image path (default: a temp file you can then read).",
)
@click.option(
    "--format",
    "fmt",
    type=click.Choice(SHAPE_IMAGE_FORMAT_CHOICES),
    default="png",
    show_default=True,
    help="Image format.",
)
@click.pass_context
def shape_export(ctx: click.Context, anchor_id: str, out: str | None, fmt: str) -> None:
    """Render a single shape to an image file — so a vision model can *see* it.

    The per-shape complement to `slide export`: crops to the shape's bounds at
    its native pixel size (no size override — `Shape.Export` doesn't honor one
    reliably). Polite (doesn't move the view). Prints the absolute path.
    """

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            sh = _resolve_shape(deck, anchor_id)
            path = sh.export_image(out, fmt=fmt)
            payload = {"ok": True, "anchor_id": sh.anchor_id, "path": str(path), "format": fmt}
            emit(
                payload,
                as_text=not ctx.obj["as_json"],
                text=f"exported {sh.anchor_id} -> {path}",
            )

    _run(ctx, go)


@shape.command(name="set-alt")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to tag (shape:S:N or ph:S:KIND)."
)
@click.option("--alt-text", "alt_text", required=True, help="Alternative (accessibility) text.")
@click.pass_context
def shape_set_alt(ctx: click.Context, anchor_id: str, alt_text: str) -> None:
    """Set a shape's alternative text — a drift-proof re-identification handle."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            sh = _resolve_shape(deck, anchor_id)
            with deck.edit(f"CLI: set alt-text {anchor_id}"):
                sh.set_alt_text(alt_text)
            emit(
                {"ok": True, "anchor_id": sh.anchor_id, "alt_text": sh.alt_text},
                as_text=not ctx.obj["as_json"],
                text=f"set alt-text on {sh.anchor_id}: {alt_text!r}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# text structure (v0.3): paragraphs, insert, format-paragraph, format-text, list
# ---------------------------------------------------------------------------


@click.command(name="paragraphs")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape whose paragraphs to list (shape:/ph:)."
)
@click.pass_context
def paragraphs_cmd(ctx: click.Context, anchor_id: str) -> None:
    """List a shape's paragraphs: anchor_id (para:S:N:P), text, indent, bullet."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            sh = _resolve_shape(deck, anchor_id)
            rows = sh.paragraphs.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_paragraphs(rows))

    _run(ctx, go)


@click.command(name="insert")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to insert relative to.")
@click.option("--text", "text", required=True, help="Paragraph text to insert.")
@click.option(
    "--after/--before",
    "after",
    default=True,
    show_default=True,
    help="Insert the new paragraph after (default) or before the anchor.",
)
@click.pass_context
def insert(ctx: click.Context, anchor_id: str, text: str, after: bool) -> None:
    """Insert a new paragraph before/after a text anchor (para:/ph:/shape:/notes:)."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            with deck.edit(f"CLI: insert paragraph at {anchor_id}"):
                if after:
                    anchor.insert_paragraph_after(text)
                else:
                    anchor.insert_paragraph_before(text)
            emit(
                {
                    "ok": True,
                    "anchor_id": anchor.anchor_id,
                    "where": "after" if after else "before",
                },
                as_text=not ctx.obj["as_json"],
                text=f"inserted paragraph {'after' if after else 'before'} {anchor.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="format-paragraph")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to format.")
@click.option("--alignment", type=click.Choice(ALIGNMENT_CHOICES), default=None, help="Alignment.")
@click.option("--space-before", type=float, default=None, help="Space before (points).")
@click.option("--space-after", type=float, default=None, help="Space after (points).")
@click.option("--line-spacing", type=float, default=None, help="Line spacing (multiple, e.g. 1.5).")
@click.option(
    "--indent-level", type=click.IntRange(1, 5), default=None, help="Outline/bullet level (1-5)."
)
@click.pass_context
def format_paragraph(
    ctx: click.Context,
    anchor_id: str,
    alignment: str | None,
    space_before: float | None,
    space_after: float | None,
    line_spacing: float | None,
    indent_level: int | None,
) -> None:
    """Set paragraph formatting (alignment, spacing, indent level) on a text anchor."""

    def go() -> None:
        if all(
            v is None for v in (alignment, space_before, space_after, line_spacing, indent_level)
        ):
            raise click.UsageError("format-paragraph requires at least one formatting option")
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            with deck.edit(f"CLI: format paragraph {anchor_id}"):
                anchor.format_paragraph(
                    alignment=alignment,
                    space_before=space_before,
                    space_after=space_after,
                    line_spacing=line_spacing,
                    indent_level=indent_level,
                )
            emit(
                {"ok": True, "anchor_id": anchor.anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"formatted {anchor.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="format-text")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to format.")
@click.option("--bold/--no-bold", "bold", default=None, help="Bold on/off.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italic on/off.")
@click.option("--underline/--no-underline", "underline", default=None, help="Underline on/off.")
@click.option("--size", type=float, default=None, help="Font size (points).")
@click.option("--font", "font", default=None, help="Font name (e.g. 'Arial').")
@click.option("--color", "color", default=None, help="Font color, '#RRGGBB'.")
@click.pass_context
def format_text(
    ctx: click.Context,
    anchor_id: str,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    size: float | None,
    font: str | None,
    color: str | None,
) -> None:
    """Set font formatting (bold/italic/size/font/color) on a text anchor.

    PowerPoint's analog of `style apply` — it has no named paragraph styles, so
    styling is direct font formatting.
    """

    def go() -> None:
        if all(v is None for v in (bold, italic, underline, size, font, color)):
            raise click.UsageError("format-text requires at least one formatting option")
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            with deck.edit(f"CLI: format text {anchor_id}"):
                anchor.format_text(
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    size=size,
                    font=font,
                    color=color,
                )
            emit(
                {"ok": True, "anchor_id": anchor.anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"formatted text of {anchor.anchor_id}",
            )

    _run(ctx, go)


@click.group(name="list")
def list_cmd() -> None:
    """List/bullet formatting: apply (bulleted/numbered) or remove."""


@list_cmd.command(name="apply")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to make a list.")
@click.option(
    "--type",
    "list_type",
    type=click.Choice(LIST_TYPE_CHOICES),
    default="bulleted",
    show_default=True,
    help="List type.",
)
@click.option("--char", "char", default=None, help="Custom bullet character (bulleted only).")
@click.pass_context
def list_apply(ctx: click.Context, anchor_id: str, list_type: str, char: str | None) -> None:
    """Turn a text anchor's paragraphs into a bulleted or numbered list."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            with deck.edit(f"CLI: apply {list_type} list to {anchor_id}"):
                anchor.apply_list(list_type, character=char)
            emit(
                {"ok": True, "anchor_id": anchor.anchor_id, "type": list_type},
                as_text=not ctx.obj["as_json"],
                text=f"applied {list_type} list to {anchor.anchor_id}",
            )

    _run(ctx, go)


@list_cmd.command(name="remove")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor to strip list formatting."
)
@click.pass_context
def list_remove(ctx: click.Context, anchor_id: str) -> None:
    """Strip bullets / numbering from a text anchor's paragraphs."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            with deck.edit(f"CLI: remove list from {anchor_id}"):
                anchor.remove_list()
            emit(
                {"ok": True, "anchor_id": anchor.anchor_id},
                as_text=not ctx.obj["as_json"],
                text=f"removed list from {anchor.anchor_id}",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# table read | add-row | delete-row  (a table is a shape: address it by slide + z-order)
# ---------------------------------------------------------------------------


@click.group(name="table")
def table() -> None:
    """Read + edit tables (a table is a shape; cells are anchors: cell:S:N:R:C)."""


def _resolve_table(deck: Presentation, slide_index: int, shape_index: int) -> Any:
    """Resolve the table on slide S, shape N (z-order). Exit 2 if no such table."""
    return deck.slides[slide_index].shapes[shape_index].table


@table.command(name="read")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.pass_context
def table_read(ctx: click.Context, slide_index: int, shape_index: int) -> None:
    """Read a table as a grid of cells, each carrying its cell:S:N:R:C anchor."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            grid = _resolve_table(deck, slide_index, shape_index).read()
            emit(grid, as_text=not ctx.obj["as_json"], text=_fmt_table_read(grid))

    _run(ctx, go)


@table.command(name="add-row")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--values", "values", default=None, help="Optional JSON array of cell values for the new row."
)
@click.pass_context
def table_add_row(
    ctx: click.Context, slide_index: int, shape_index: int, values: str | None
) -> None:
    """Append a row to a table (one Ctrl-Z)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--values must be a JSON array")

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            t = _resolve_table(deck, slide_index, shape_index)
            with deck.edit(f"CLI: add row to table shape:{slide_index}:{shape_index}"):
                t.add_row(parsed)
            emit(
                {"ok": True, "anchor_id": t.shape.anchor_id, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"added row to {t.shape.anchor_id} (now {t.row_count} rows)",
            )

    _run(ctx, go)


@table.command(name="delete-row")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option("--row", "row", type=int, required=True, help="1-based row to delete.")
@click.pass_context
def table_delete_row(ctx: click.Context, slide_index: int, shape_index: int, row: int) -> None:
    """Delete a row from a table (one Ctrl-Z)."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            t = _resolve_table(deck, slide_index, shape_index)
            with deck.edit(f"CLI: delete row {row} from table shape:{slide_index}:{shape_index}"):
                t.delete_row(row)
            emit(
                {"ok": True, "anchor_id": t.shape.anchor_id, "rows": t.row_count},
                as_text=not ctx.obj["as_json"],
                text=f"deleted row {row} from {t.shape.anchor_id} (now {t.row_count} rows)",
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# read anchor --anchor-id ID   |   read notes --slide S
# ---------------------------------------------------------------------------


@click.group(name="read")
def read() -> None:
    """Read text from the deck. `read anchor --anchor-id …` reads any text anchor
    (ph:/shape:/notes:); `read notes --slide S` is sugar for `--anchor-id notes:S`."""


@read.command(name="anchor")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Anchor to read (e.g. ph:3:title, shape:3:2, notes:3).",
)
@click.pass_context
def read_anchor(ctx: click.Context, anchor_id: str) -> None:
    """Read the text of any text anchor."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            text = anchor.text
            emit(
                {"anchor_id": anchor.anchor_id, "kind": anchor.kind, "text": text},
                as_text=not ctx.obj["as_json"],
                text=text,
            )

    _run(ctx, go)


@read.command(name="notes")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.pass_context
def read_notes(ctx: click.Context, slide_index: int) -> None:
    """Read the speaker notes of slide INDEX."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            notes = deck.slides[slide_index].notes
            text = notes.text
            emit(
                {"anchor_id": notes.anchor_id, "kind": notes.kind, "text": text},
                as_text=not ctx.obj["as_json"],
                text=text,
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# write --anchor-id ID --text "..."
# ---------------------------------------------------------------------------


def _set_text(ctx: click.Context, anchor_id: str, text: str, label: str) -> None:
    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            with deck.edit(label):
                anchor.set_text(text)
            emit(
                {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind},
                as_text=not ctx.obj["as_json"],
                text=f"wrote {anchor.anchor_id}",
            )

    _run(ctx, go)


@click.command(name="write")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor to set (ph:/shape:/notes:)."
)
@click.option("--text", "text", required=True, help="New text (embed \\n for paragraphs).")
@click.pass_context
def write(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Set the text of a text anchor (preserves the viewed slide; one Ctrl-Z)."""
    _set_text(ctx, anchor_id, text, f"CLI: write {anchor_id}")


# ---------------------------------------------------------------------------
# replace --anchor-id ID --text "..."
# ---------------------------------------------------------------------------


@click.command(name="replace")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor whose contents to replace."
)
@click.option("--text", "text", required=True, help="Replacement text (embed \\n for paragraphs).")
@click.pass_context
def replace(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Replace the entire text of a text anchor.

    In v0 this is the anchor-addressed form (identical effect to `write`). Fuzzy
    `replace --find OLD --text NEW` arrives with the find/replace stage.
    """
    _set_text(ctx, anchor_id, text, f"CLI: replace {anchor_id}")


# ---------------------------------------------------------------------------
# selection
# ---------------------------------------------------------------------------


@click.command(name="selection")
@click.pass_context
def selection_cmd(ctx: click.Context) -> None:
    """Report the user's current selection, resolved to anchors.

    A polite read (it doesn't change the selection): the selected shapes as
    `shape:S:N`, or a text caret as `para:S:N:P`, with the single targetable
    `anchor_id` that `--anchor-id here:` resolves to.
    """

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            info = deck.selection().to_dict()
            emit(info, as_text=not ctx.obj["as_json"], text=_fmt_selection(info))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# show — live slide-show control
# ---------------------------------------------------------------------------


@click.group(name="show")
def show() -> None:
    """Live slide-show control: start, end, next, prev, goto, black, white, resume, state.

    These deliberately drive what the user sees on screen (unlike the polite edit
    verbs). `show state` is the read; the control verbs all print the resulting
    state. They need a running show (start one with `show start`) — `show next`
    et al. exit 1 if none is running."""


def _show_action(ctx: click.Context, fn: Any, text: str | None = None) -> None:
    """Run a show verb `fn(show)` and emit the resulting state dict."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            info = fn(deck.show)
            emit(
                info,
                as_text=not ctx.obj["as_json"],
                text=text if text is not None else _fmt_show(info),
            )

    _run(ctx, go)


@show.command(name="start")
@click.option(
    "--from", "from_slide", type=int, default=None, help="1-based slide to start on (default: top)."
)
@click.pass_context
def show_start(ctx: click.Context, from_slide: int | None) -> None:
    """Start the slide show (optionally on a given slide)."""
    _show_action(ctx, lambda s: s.start(from_slide=from_slide))


@show.command(name="end")
@click.pass_context
def show_end(ctx: click.Context) -> None:
    """End the slide show (no-op if none is running)."""
    _show_action(ctx, lambda s: s.end())


@show.command(name="next")
@click.pass_context
def show_next(ctx: click.Context) -> None:
    """Advance to the next build/slide."""
    _show_action(ctx, lambda s: s.next())


@show.command(name="prev")
@click.pass_context
def show_prev(ctx: click.Context) -> None:
    """Step back to the previous build/slide."""
    _show_action(ctx, lambda s: s.previous())


@show.command(name="goto")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide to jump to.")
@click.pass_context
def show_goto(ctx: click.Context, slide_index: int) -> None:
    """Jump the running show to a slide."""
    _show_action(ctx, lambda s: s.goto(slide_index))


@show.command(name="black")
@click.pass_context
def show_black(ctx: click.Context) -> None:
    """Blank the screen to black (resume with `show resume`)."""
    _show_action(ctx, lambda s: s.black())


@show.command(name="white")
@click.pass_context
def show_white(ctx: click.Context) -> None:
    """Blank the screen to white (resume with `show resume`)."""
    _show_action(ctx, lambda s: s.white())


@show.command(name="resume")
@click.pass_context
def show_resume(ctx: click.Context) -> None:
    """Resume from a black/white blank screen."""
    _show_action(ctx, lambda s: s.resume())


@show.command(name="state")
@click.pass_context
def show_state(ctx: click.Context) -> None:
    """Report whether a show is running and which slide is on screen (read-only)."""
    _show_action(ctx, lambda s: s.state())


# ---------------------------------------------------------------------------
# go-to --anchor-id ID
# ---------------------------------------------------------------------------


@click.command(name="go-to")
@click.option("--anchor-id", "anchor_id", required=True, help="Anchor to move the user's view to.")
@click.option(
    "--select/--no-select",
    "select",
    default=True,
    show_default=True,
    help="Select the target shape after jumping to its slide.",
)
@click.pass_context
def go_to(ctx: click.Context, anchor_id: str, select: bool) -> None:
    """Move the user's view to an anchor's slide (deliberate, opt-in view move)."""

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            anchor = deck.anchor_by_id(anchor_id)
            deck.go_to(anchor, select=select)
            emit(
                {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind},
                as_text=not ctx.obj["as_json"],
                text=f"moved view to {anchor.anchor_id}",
            )

    _run(ctx, go)

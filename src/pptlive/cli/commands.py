"""CLI subcommands wired against the pptlive library.

v0 surface: status, slides, outline, slide read, shapes, read (anchor/notes),
write, replace, go-to. Slide lifecycle, shape geometry, find/replace, and the
`show` group arrive in later stages.
"""

from __future__ import annotations

from typing import Any

import click

from .. import attach
from .._presentation import Presentation
from ..exceptions import PowerPointNotRunningError
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
        text = s.get("text")
        snippet = ""
        if text:
            flat = text.replace("\r", " / ").replace("\n", " / ").replace("\v", " ")
            snippet = "  " + (flat if len(flat) <= 60 else flat[:57] + "…")
        lines.append(
            f"[{s['anchor_id']}] {s.get('name', '')!r} ({s.get('type', '?')}{ph}){snippet}"
        )
    return "\n".join(lines)


def _fmt_slide_read(grid: dict[str, Any]) -> str:
    head = (
        f"slide {grid.get('index')} (id {grid.get('id')}) "
        f"layout={grid.get('layout') or '?'} title={grid.get('title') or '(untitled)'!r}"
    )
    return head + "\n" + _fmt_shapes(grid.get("shapes") or [])


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(slides_cmd)
    group.add_command(outline)
    group.add_command(slide)
    group.add_command(shapes_cmd)
    group.add_command(read)
    group.add_command(write)
    group.add_command(replace)
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
    """Slide-level reads (lifecycle verbs — add/delete/duplicate — arrive in v0.1)."""


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

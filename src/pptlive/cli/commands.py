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

import base64
import functools
import json
import os
import sys
from collections.abc import Callable
from pathlib import Path
from typing import Any

import click

from .. import attach
from .._batch import run_batch
from .._guide import bundled_skill, skill_body, skill_name
from .._presentation import Presentation
from .._shapes import Shape
from ..constants import (
    ALIGN_CHOICES,
    ALIGNMENT_CHOICES,
    ANIM_EFFECT_CHOICES,
    ANIM_TRIGGER_CHOICES,
    ARROWHEAD_SIZE_CHOICES,
    ARROWHEAD_STYLE_CHOICES,
    AUTOSHAPE_CHOICES,
    CHART_TYPE_CHOICES,
    CONNECTOR_CHOICES,
    DASH_STYLE_CHOICES,
    DISTRIBUTE_CHOICES,
    ENTRY_EFFECT_CHOICES,
    GRADIENT_STYLE_CHOICES,
    IMAGE_FORMAT_CHOICES,
    LIST_TYPE_CHOICES,
    PATTERN_CHOICES,
    PRESET_GRADIENT_CHOICES,
    RELATIVE_TO_CHOICES,
    SAVE_FORMAT_CHOICES,
    SHAPE_IMAGE_FORMAT_CHOICES,
    SMARTART_CHOICES,
    TEXT_STYLE_CHOICES,
    THEME_COLOR_CHOICES,
    THEME_FONT_SCRIPT_CHOICES,
    THEME_FONT_SLOTS,
    ZORDER_CHOICES,
)
from ..exceptions import EXIT_CODE_FOR, AnchorNotFoundError, PowerPointNotRunningError
from .main import EXIT_OTHER, _run, emit


def _pick_deck(ppt: Any, doc_name: str | None) -> Presentation:
    if doc_name is None:
        return ppt.presentations.active
    return ppt.presentations[doc_name]


def _deck_command(fn: Callable[..., None]) -> Callable[..., None]:
    """Wrap a command body ``fn(ctx, deck, **opts)`` in the standard scaffold.

    Owns the envelope every deck command repeats: open PowerPoint (``attach``),
    pick the ``--doc`` deck (``_pick_deck``), and route ``PptliveError`` through
    the ``_run`` error boundary (stderr + exit code). The body receives the
    resolved ``deck`` and opens its own ``with deck.edit(label):`` fence for
    mutations — those labels are hand-tuned/dynamic and the read-before-fence /
    emit-after-fence shape is the meaningful part, so the fence stays in the body.

    Replaces the per-command ``def go(): with attach() as ppt: deck =
    _pick_deck(...); ...; _run(ctx, go)`` boilerplate. Apply it as the innermost
    decorator (directly above ``def``); ``@click.option``s stay above it.
    Commands that don't pick a deck (``status``, the offline ``llm-help`` /
    ``install-*``) stay un-decorated.
    """

    @functools.wraps(fn)
    def wrapper(ctx: click.Context, **kwargs: Any) -> None:
        def go() -> None:
            with attach() as ppt:
                deck = _pick_deck(ppt, ctx.obj["doc_name"])
                fn(ctx, deck, **kwargs)

        _run(ctx, go)

    return click.pass_context(wrapper)


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
        # `saved` is present from v1.1; a False (or missing on an old payload)
        # value flags unsaved changes the user may want persisted.
        dirty = " (unsaved)" if d.get("saved") is False else ""
        lines.append(f"{marker} {str(d.get('name', '')):<{width}}  {d.get('path', '')}{dirty}")
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


def _fmt_find(matches: list[dict[str, Any]]) -> str:
    if not matches:
        return "(no matches)"
    lines = []
    for m in matches:
        loc = f"{m['anchor_id']} @{m['start']}+{m['length']}"
        lines.append(f"{loc:<28}  {m.get('context', m['text'])}")
    return "\n".join(lines)


def _fmt_replace_summary(replacements: list[dict[str, Any]]) -> str:
    n = len(replacements)
    return f"replaced {n} occurrence{'s' if n != 1 else ''}"


def _fmt_comment(c: dict[str, Any], *, indent: str = "") -> list[str]:
    """Render one comment (and its reply thread) as indented text lines."""
    when = c.get("datetime") or ""
    head = f"{indent}[{c.get('index')}] {c.get('author') or '?'}"
    if when:
        head += f" ({when})"
    lines = [f"{head}: {c.get('text') or ''}"]
    for reply in c.get("replies") or []:
        lines.extend(_fmt_comment(reply, indent=indent + "    ↳ "))
    return lines


def _fmt_comment_list(payload: Any) -> str:
    """Text view of a per-slide list or a deck-wide `{total, slides:[...]}` rollup."""
    if isinstance(payload, dict):  # deck-wide rollup
        slides = payload.get("slides") or []
        if not slides:
            return "(no comments)"
        lines: list[str] = []
        for entry in slides:
            lines.append(f"slide {entry.get('slide')}:")
            for c in entry.get("comments") or []:
                lines.extend(_fmt_comment(c, indent="  "))
        return "\n".join(lines)
    # per-slide list
    if not payload:
        return "(no comments)"
    lines = []
    for c in payload:
        lines.extend(_fmt_comment(c))
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


def _fmt_chart_read(info: dict[str, Any]) -> str:
    head = (
        f"chart at {info.get('anchor_id')} (slide {info.get('slide')}, "
        f"type={info.get('chart_type')})"
    )
    lines = [head, "  categories: " + ", ".join(str(c) for c in info.get("categories") or [])]
    for s in info.get("series") or []:
        vals = ", ".join(str(v) for v in s.get("values") or [])
        lines.append(f"  {s.get('name')!r}: {vals}")
    return "\n".join(lines)


def _parse_categories(raw: str | None) -> list[str] | None:
    """Parse --categories: a JSON array, else a comma-separated list."""
    if raw is None:
        return None
    raw = raw.strip()
    if raw.startswith("["):
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--categories must be a JSON array or CSV: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--categories JSON must be an array")
        return [str(c) for c in parsed]
    return [c.strip() for c in raw.split(",") if c.strip()]


def _parse_series(raw: str | None) -> Any:
    """Parse --series: a JSON object {name:[values]} or array of [name,[values]]."""
    if raw is None:
        return None
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--series must be JSON (object or array): {e}") from e
    if isinstance(parsed, dict):
        return parsed
    if isinstance(parsed, list):
        out: list[tuple[str, list[float]]] = []
        for item in parsed:
            if not (isinstance(item, (list, tuple)) and len(item) == 2):
                raise click.UsageError("--series array items must be [name, [values]]")
            name, values = item
            if not isinstance(values, list):
                raise click.UsageError("--series values must be a JSON array")
            out.append((str(name), [float(v) for v in values]))
        return out
    raise click.UsageError("--series must be a JSON object or array")


def _parse_nodes(raw: str | None) -> Any:
    """Parse --nodes: a JSON array of strings and/or {text, children} objects.

    `["A", "B"]` (flat) or `[{"text": "CEO", "children": ["VP Eng"]}]` (tree). The
    structure is validated by `SmartArt.set_nodes`; here we only require a JSON
    array.
    """
    if raw is None:
        return None
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        raise click.UsageError(f"--nodes must be a JSON array: {e}") from e
    if not isinstance(parsed, list):
        raise click.UsageError("--nodes must be a JSON array of strings and/or objects")
    return parsed


def _fmt_smartart_read(info: dict[str, Any]) -> str:
    head = (
        f"smartart at {info.get('anchor_id')} (slide {info.get('slide')}, "
        f"layout={info.get('layout')}, {info.get('node_count')} nodes)"
    )
    lines = [head]

    def walk(nodes: list[dict[str, Any]], depth: int) -> None:
        for n in nodes:
            lines.append("  " * (depth + 1) + f"- {n.get('text')!r}")
            walk(n.get("children") or [], depth + 1)

    walk(info.get("nodes") or [], 0)
    return "\n".join(lines)


def _fmt_theme_read(info: dict[str, Any]) -> str:
    lines = ["theme colors:"]
    for slot, hexv in (info.get("colors") or {}).items():
        lines.append(f"  {slot:<19} {hexv}")
    fonts = info.get("fonts") or {}
    lines.append("theme fonts:")
    lines.append(f"  major (headings)    {fonts.get('major')}")
    lines.append(f"  minor (body)        {fonts.get('minor')}")
    return "\n".join(lines)


def _fmt_master_read(info: dict[str, Any]) -> str:
    lines = []
    for style, body in (info.get("text_styles") or {}).items():
        lines.append(f"{style} style:")
        for lvl in body.get("levels") or []:
            bits = []
            if lvl.get("font") is not None:
                bits.append(str(lvl["font"]))
            if lvl.get("size") is not None:
                bits.append(f"{lvl['size']}pt")
            if lvl.get("bold"):
                bits.append("bold")
            if lvl.get("color") is not None:
                bits.append(str(lvl["color"]))
            lines.append(f"  L{lvl.get('level')}: " + " ".join(bits))
    bg = info.get("background") or {}
    lines.append(f"background: {bg.get('type')} {bg.get('color') or ''}".rstrip())
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


def _parse_slides_range(
    slide: int | None, slides_range: str | None
) -> int | tuple[int, int] | None:
    """Resolve the CLI `--slide` / `--slides` options to a snapshot selector.

    `--slide N` → `N`; `--slides A-B` → `(A, B)`; neither → `None` (all slides).
    Giving both, or a malformed `--slides`, is a `ValueError` (a bad-input error).
    """
    if slide is not None and slides_range is not None:
        raise ValueError("pass at most one of --slide and --slides")
    if slide is not None:
        return slide
    if slides_range is None:
        return None
    start, sep, end = slides_range.partition("-")
    if not sep:
        raise ValueError(f"--slides must be a span like '2-4', got {slides_range!r}")
    try:
        return int(start), int(end)
    except ValueError as e:
        raise ValueError(f"--slides must be a span like '2-4', got {slides_range!r}") from e


@click.command(name="snapshot")
@click.option("--slide", "slide", type=int, default=None, help="Render a single 1-based slide.")
@click.option(
    "--slides", "slides_range", default=None, help="Render an inclusive slide span, e.g. '2-4'."
)
@click.option(
    "--out",
    "out",
    type=click.Path(dir_okay=False, path_type=Path),
    default=None,
    help="Write the PNG here. Multiple slides are written as <stem>-s<N><suffix>. "
    "Without --out, base64 PNG data is returned inline in the JSON.",
)
@click.option(
    "--max-dim",
    "max_dim",
    type=int,
    default=None,
    help="Cap each slide's long edge to this many pixels (only ever lowers resolution). "
    "The lever for a cheap whole-deck layout check — ~1000 stays legible at a "
    "fraction of the tokens; a uniform per-slide cost across the deck.",
)
@click.option(
    "--width",
    "width",
    type=int,
    default=None,
    help="Exact per-slide width in pixels (overrides --max-dim; height follows the "
    "aspect ratio if omitted). Can't be combined with --max-dim.",
)
@click.option(
    "--height",
    "height",
    type=int,
    default=None,
    help="Exact per-slide height in pixels (overrides --max-dim; width follows the "
    "aspect ratio if omitted). Can't be combined with --max-dim.",
)
@click.option(
    "--format",
    "fmt",
    type=click.Choice(IMAGE_FORMAT_CHOICES),
    default="png",
    show_default=True,
    help="Image format. Note: PowerPoint exposes no JPEG-quality knob, so pixel "
    "dimensions are the only render-cost lever.",
)
@_deck_command
def snapshot_cmd(
    ctx: click.Context,
    deck: Presentation,
    slide: int | None,
    slides_range: str | None,
    out: Path | None,
    max_dim: int | None,
    width: int | None,
    height: int | None,
    fmt: str,
) -> None:
    """Render slides to PNG so a vision model can *see* the whole deck cheaply.

    The token-cost-aware read: `--max-dim` caps each slide's long edge in pixels,
    giving a predictable, uniform per-slide budget — the lever for "render the
    whole deck and check my styling landed" without full-resolution bloat. For an
    exact size pass `--width`/`--height` instead (overrides `--max-dim`; one is
    enough, the other follows the aspect ratio). Renders the current unsaved state;
    polite (doesn't move the view). With `--out` the PNGs are written (single → that
    path, multiple → `<stem>-s<N><suffix>`) and the JSON reports each `path`;
    without it, base64 PNG data is returned inline.
    """
    # A bad --slides range or snapshot arg raises ValueError; the shared _run
    # boundary classifies it to a clean exit 1 (like every other deck command),
    # so no hand-rolled try/except here.
    selector = _parse_slides_range(slide, slides_range)
    snaps = deck.snapshot(
        out, slides=selector, fmt=fmt, max_dim=max_dim, width=width, height=height
    )
    images = [
        {
            "slide": s.slide,
            "bytes": len(s.image),
            **(
                {"path": str(s.path)}
                if s.path is not None
                else {"base64": base64.b64encode(s.image).decode("ascii")}
            ),
        }
        for s in snaps
    ]
    sel_text = (
        f"slide {slide}"
        if slide is not None
        else (f"slides {slides_range}" if slides_range else "all slides")
    )
    payload = {
        "ok": True,
        "selector": sel_text,
        "count": len(snaps),
        "format": fmt,
        "max_dim": max_dim,
        "width": width,
        "height": height,
        "images": images,
    }
    written = [str(s.path) for s in snaps if s.path is not None]
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=(
            f"snapshot: {len(snaps)} slide(s) -> " + ", ".join(written)
            if written
            else f"snapshot: {len(snaps)} slide(s) (base64 inline)"
        ),
    )


# ---------------------------------------------------------------------------
# save | save-as PATH [--format] [--overwrite] | export-pdf PATH
#   Explicit-only file output (pptlive never auto-saves). save/save-as persist
#   the working .pptx; export-pdf is a read (no rebind, dirty flag preserved).
# ---------------------------------------------------------------------------


@click.command(name="save")
@_deck_command
def save_cmd(ctx: click.Context, deck: Presentation) -> None:
    """Save the deck to its existing file (explicit; pptlive never auto-saves).

    Fails (exit 1) if the deck has never been saved — use `save-as PATH` first.
    """
    path = deck.save()
    emit(
        {"ok": True, "path": path, "saved": True},
        as_text=not ctx.obj["as_json"],
        text=f"saved {path}",
    )


@click.command(name="save-as")
@click.argument("path", type=click.Path(dir_okay=False, path_type=Path))
@click.option(
    "--format",
    "fmt",
    type=click.Choice(SAVE_FORMAT_CHOICES),
    default="pptx",
    show_default=True,
    help="Output format (PDF is `export-pdf`).",
)
@click.option(
    "--overwrite",
    is_flag=True,
    default=False,
    help="Allow overwriting an existing file (default: refuse).",
)
@_deck_command
def save_as_cmd(
    ctx: click.Context, deck: Presentation, path: Path, fmt: str, overwrite: bool
) -> None:
    """Save the deck to PATH and rebind the working file to it (explicit).

    After this the open deck *is* PATH (its name/path follow), matching
    PowerPoint's Save-As. Refuses to clobber an existing file unless `--overwrite`.
    For PDF, use `export-pdf` (a read — it doesn't rebind the working file).
    """
    try:
        written = deck.save_as(path, fmt=fmt, overwrite=overwrite)
    except FileExistsError as exc:
        click.echo(f"error: {exc}", err=True)
        sys.exit(1)
    emit(
        {"ok": True, "path": written, "format": fmt},
        as_text=not ctx.obj["as_json"],
        text=f"saved {written}",
    )


@click.command(name="export-pdf")
@click.argument("path", type=click.Path(dir_okay=False, path_type=Path))
@_deck_command
def export_pdf_cmd(ctx: click.Context, deck: Presentation, path: Path) -> None:
    """Export the deck to a PDF at PATH (the "hand back a deliverable" path).

    A pixel-faithful render of the deck's current (unsaved) state via PowerPoint's
    PDF engine. A read: it neither rebinds the working file nor clears its dirty
    flag, so your `.pptx` is untouched. Overwrites an existing PDF.
    """
    written = deck.export_pdf(path)
    emit(
        {"ok": True, "path": written},
        as_text=not ctx.obj["as_json"],
        text=f"exported {written}",
    )


# ---------------------------------------------------------------------------
# media add | export-video PATH | video-status
#   Insert audio/video narration (embed by default; autoplay + pace the slide),
#   then export the deck to MP4. export-video wraps the async CreateVideo: it
#   blocks by default (poll to Done) and `video-status` polls a --no-wait export.
# ---------------------------------------------------------------------------


@click.group(name="media")
def media() -> None:
    """Insert audio/video clips (embed by default; autoplay + pace the slide)."""


@media.command(name="add")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--kind", type=click.Choice(["audio", "video"]), required=True, help="Media kind to insert."
)
@click.option(
    "--path",
    "path",
    type=click.Path(exists=True, dir_okay=False),
    required=True,
    help="Audio/video file to embed (or link with --link).",
)
@click.option("--left", type=float, default=None, help="Left edge (points).")
@click.option("--top", type=float, default=None, help="Top edge (points).")
@click.option("--width", type=float, default=None, help="Width (points).")
@click.option("--height", type=float, default=None, help="Height (points).")
@click.option(
    "--link/--no-link",
    "link",
    default=False,
    show_default=True,
    help="Link the file on disk instead of embedding it (shrinks the deck).",
)
@click.option(
    "--autoplay/--no-autoplay",
    "autoplay",
    default=True,
    show_default=True,
    help="Play the clip automatically on slide entry.",
)
@click.option(
    "--hide-icon/--no-hide-icon",
    "hide_icon",
    default=True,
    show_default=True,
    help="Hide the speaker icon while not playing (audio only).",
)
@click.option(
    "--pace-slide/--no-pace-slide",
    "pace_slide",
    default=True,
    show_default=True,
    help="Auto-advance the slide to the clip's length (paces an exported video).",
)
@click.option(
    "--alt-text",
    "alt_text",
    default=None,
    help="Alternative text (a drift-proof re-identification handle).",
)
@_deck_command
def media_add(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    kind: str,
    path: str,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    link: bool,
    autoplay: bool,
    hide_icon: bool,
    pace_slide: bool,
    alt_text: str | None,
) -> None:
    """Insert audio/video on a slide; print its anchor_id, name, type, and media info."""
    slide = deck.slides[slide_index]  # exit 2 if slide out of range
    with deck.edit(f"CLI: add {kind} on slide {slide_index}"):
        if kind == "audio":
            new = slide.add_audio(
                path,
                left=left,
                top=top,
                width=width,
                height=height,
                link=link,
                autoplay=autoplay,
                hide_icon=hide_icon,
                pace_slide=pace_slide,
                alt_text=alt_text,
            )
        else:  # video
            new = slide.add_video(
                path,
                left=left,
                top=top,
                width=width,
                height=height,
                link=link,
                autoplay=autoplay,
                pace_slide=pace_slide,
                alt_text=alt_text,
            )
    payload = {"ok": True, **new.to_dict()}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"added {kind} {payload['name']!r} at {payload['anchor_id']}",
    )


@media.command(name="set")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Media clip to adjust (shape:S:N / shapeid:S:ID / ph:S:KIND).",
)
@click.option(
    "--muted/--no-muted",
    "muted",
    default=None,
    help="Silence the clip (--no-muted un-mutes).",
)
@click.option("--volume", type=float, default=None, help="Playback volume, 0.0-1.0.")
@click.option(
    "--start", "trim_start", type=float, default=None, help="Trim: clip start, in seconds."
)
@click.option("--end", "trim_end", type=float, default=None, help="Trim: clip end, in seconds.")
@_deck_command
def media_set(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    muted: bool | None,
    volume: float | None,
    trim_start: float | None,
    trim_end: float | None,
) -> None:
    """Set playback options on a media clip: mute, volume, and the trim window.

    `--start`/`--end` are the trim window in seconds (omit an edge to keep it).
    At least one option is required; the shape must hold audio/video.
    """
    sh = _resolve_shape(deck, anchor_id)  # exit 2 if missing / not a shape
    with deck.edit(f"CLI: set media playback on {anchor_id}"):
        media = sh.set_media_playback(muted=muted, volume=volume, start=trim_start, end=trim_end)
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "media": media},
        as_text=not ctx.obj["as_json"],
        text=f"set media on {sh.anchor_id}",
    )


@click.command(name="export-video")
@click.argument("path", type=click.Path(dir_okay=False, path_type=Path))
@click.option(
    "--resolution", type=int, default=720, show_default=True, help="Vertical pixel height."
)
@click.option("--fps", type=int, default=30, show_default=True, help="Frames per second.")
@click.option("--quality", type=int, default=85, show_default=True, help="Encode quality 0-100.")
@click.option(
    "--default-slide-duration",
    "default_slide_duration",
    type=float,
    default=5.0,
    show_default=True,
    help="Seconds per slide that has no timing.",
)
@click.option(
    "--use-timings/--no-use-timings",
    "use_timings",
    default=True,
    show_default=True,
    help="Honor per-slide timings + narration.",
)
@click.option(
    "--wait/--no-wait",
    "wait",
    default=True,
    show_default=True,
    help="Block until the encode finishes (--no-wait returns immediately; poll video-status).",
)
@click.option(
    "--timeout", type=float, default=600.0, show_default=True, help="Max seconds to wait."
)
@_deck_command
def export_video_cmd(
    ctx: click.Context,
    deck: Presentation,
    path: Path,
    resolution: int,
    fps: int,
    quality: int,
    default_slide_duration: float,
    use_timings: bool,
    wait: bool,
    timeout: float,
) -> None:
    """Export the deck to an MP4 at PATH (the narrated-video deliverable).

    Wraps PowerPoint's async CreateVideo: blocks until the encode finishes by
    default. A read — it doesn't rebind your `.pptx`. With --no-wait it returns the
    in-flight status immediately; poll `video-status` until it reports `done`.
    """
    result = deck.export_video(
        path,
        use_timings=use_timings,
        default_slide_duration=default_slide_duration,
        resolution=resolution,
        fps=fps,
        quality=quality,
        wait=wait,
        timeout=timeout,
    )
    emit(
        result.to_dict(),
        as_text=not ctx.obj["as_json"],
        text=f"video {result.status}: {result.path}",
    )


@click.command(name="video-status")
@_deck_command
def video_status_cmd(ctx: click.Context, deck: Presentation) -> None:
    """Poll the async video-export status (none/queued/in_progress/done/failed)."""
    result = deck.video_status()
    emit(
        result.to_dict(),
        as_text=not ctx.obj["as_json"],
        text=f"video status: {result.status}",
    )


def register(group: click.Group) -> None:
    group.add_command(status)
    group.add_command(slides_cmd)
    group.add_command(outline)
    group.add_command(slide)
    group.add_command(snapshot_cmd)
    group.add_command(save_cmd)
    group.add_command(save_as_cmd)
    group.add_command(export_pdf_cmd)
    group.add_command(media)
    group.add_command(export_video_cmd)
    group.add_command(video_status_cmd)
    group.add_command(shapes_cmd)
    group.add_command(shape)
    group.add_command(link)
    group.add_command(read)
    group.add_command(write)
    group.add_command(set_paragraphs)
    group.add_command(exec_script)
    group.add_command(find)
    group.add_command(replace)
    group.add_command(paragraphs_cmd)
    group.add_command(insert)
    group.add_command(format_paragraph)
    group.add_command(format_text)
    group.add_command(reset_format)
    group.add_command(list_cmd)
    group.add_command(table)
    group.add_command(chart)
    group.add_command(smartart)
    group.add_command(comment)
    group.add_command(section)
    group.add_command(theme)
    group.add_command(master)
    group.add_command(selection_cmd)
    group.add_command(show)
    group.add_command(go_to)
    group.add_command(llm_help_cmd)
    group.add_command(install_skill_cmd)
    group.add_command(install_mcp_cmd)


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
@_deck_command
def slides_cmd(ctx: click.Context, deck: Presentation) -> None:
    """List every slide: index, id, layout, title, shape count, has-notes."""
    rows = deck.slides.list()
    emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_slides(rows))


# ---------------------------------------------------------------------------
# outline
# ---------------------------------------------------------------------------


@click.command(name="outline")
@_deck_command
def outline(ctx: click.Context, deck: Presentation) -> None:
    """Print the title + body bullets of every slide (the Outline-view analog)."""
    items = deck.outline()
    emit(items, as_text=not ctx.obj["as_json"], text=_fmt_outline(items))


# ---------------------------------------------------------------------------
# slide read S
# ---------------------------------------------------------------------------


@click.group(name="slide")
def slide() -> None:
    """Slide reads + lifecycle: read, geometry, add, delete, duplicate, move, set-layout, layouts."""


@slide.command(name="read")
@click.argument("index", type=int)
@_deck_command
def slide_read(ctx: click.Context, deck: Presentation, index: int) -> None:
    """Read every shape on slide INDEX: anchor_id, name, type, geometry, text."""
    grid = deck.slides[index].read()
    emit(grid, as_text=not ctx.obj["as_json"], text=_fmt_slide_read(grid))


@slide.command(name="geometry")
@click.argument("index", type=int)
@_deck_command
def slide_geometry(ctx: click.Context, deck: Presentation, index: int) -> None:
    """Spatial map of slide INDEX: slide size, per-shape boxes, overlaps, off-slide.

    The geometry sanity-check to run after placing shapes — catch overlaps or an
    off-edge shape without a render. Axis-aligned boxes (rotation not accounted
    for). A read; the user's view doesn't move.
    """
    report = deck.slides[index].geometry_report()
    n_over, n_off = len(report["overlaps"]), len(report["off_slide"])
    emit(
        report,
        as_text=not ctx.obj["as_json"],
        text=(
            f"slide {index}: {len(report['shapes'])} shapes, {n_over} overlap(s), {n_off} off-slide"
        ),
    )


@slide.command(name="animations")
@click.argument("index", type=int)
@_deck_command
def slide_animations(ctx: click.Context, deck: Presentation, index: int) -> None:
    """List slide INDEX's shape animations in play order (effect, trigger, target).

    Each row maps an effect back to its target `shapeid` so you can see *what*
    animates *how* without a render. A read; the user's view doesn't move.
    """
    rows = deck.slides[index].animations()
    text = (
        "\n".join(
            f"{r['seq_index']}. {r['shape']} ({r['shapeid']}): "
            f"{'exit ' if r['exit'] else ''}{r['effect']} / {r['trigger']}"
            for r in rows
        )
        or f"slide {index}: no animations"
    )
    emit(rows, as_text=not ctx.obj["as_json"], text=text)


@slide.command(name="clear-animations")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@_deck_command
def slide_clear_animations(ctx: click.Context, deck: Presentation, slide_index: int) -> None:
    """Remove every animation from a slide; print how many effects were deleted."""
    with deck.edit(f"CLI: clear animations on slide {slide_index}"):
        removed = deck.slides[slide_index].clear_animations()
    emit(
        {"ok": True, "index": slide_index, "removed": removed},
        as_text=not ctx.obj["as_json"],
        text=f"slide {slide_index}: cleared {removed} animation(s)",
    )


@slide.command(name="layouts")
@_deck_command
def slide_layouts(ctx: click.Context, deck: Presentation) -> None:
    """List the deck's slide layouts (the names `add`/`set-layout` accept)."""
    rows = deck.layouts()
    emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_layouts(rows))


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
@_deck_command
def slide_export(
    ctx: click.Context,
    deck: Presentation,
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


@slide.command(name="add")
@click.option(
    "--layout",
    default=None,
    help="Layout name (default: title_and_content). See `slide layouts` for names.",
)
@click.option(
    "--index", "index", type=int, default=None, help="1-based insertion position (default: end)."
)
@click.option(
    "--placeholders",
    "placeholders_json",
    default=None,
    help=(
        'JSON map repositioning placeholders, e.g. \'{"body": {"left": 40, "width": 440}}\' '
        "(points, any subset of left/top/width/height; KIND as in ph:S:KIND) — "
        "add + resize in one op (a left-half content area beside a right-side panel)."
    ),
)
@_deck_command
def slide_add(
    ctx: click.Context,
    deck: Presentation,
    layout: str | None,
    index: int | None,
    placeholders_json: str | None,
) -> None:
    """Add a slide; print its index, id, and layout."""
    placeholders = None
    if placeholders_json is not None:
        try:
            placeholders = json.loads(placeholders_json)
        except json.JSONDecodeError as exc:
            raise click.UsageError(f"invalid JSON in --placeholders: {exc}") from exc
    with deck.edit(f"CLI: add slide ({layout or 'default'})"):
        new = deck.slides.add(layout=layout, index=index, placeholders=placeholders)
    payload: dict[str, Any] = {
        "ok": True,
        "index": new.index,
        "id": new.id,
        "layout": new.layout_name,
    }
    if placeholders:
        payload["placeholders"] = {
            kind: _resolve_shape(deck, f"ph:{new.index}:{kind}").geometry() for kind in placeholders
        }
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"added slide {new.index} ({new.layout_name})",
    )


@slide.command(name="delete")
@click.option(
    "--slide", "slide_index", type=int, required=True, help="1-based slide index to delete."
)
@_deck_command
def slide_delete(ctx: click.Context, deck: Presentation, slide_index: int) -> None:
    """Delete a slide."""
    target = deck.slides[slide_index]  # exit 2 here if out of range
    with deck.edit(f"CLI: delete slide {slide_index}"):
        target.delete()
    emit(
        {"ok": True, "deleted": slide_index},
        as_text=not ctx.obj["as_json"],
        text=f"deleted slide {slide_index}",
    )


@slide.command(name="duplicate")
@click.option(
    "--slide", "slide_index", type=int, required=True, help="1-based slide index to duplicate."
)
@_deck_command
def slide_duplicate(ctx: click.Context, deck: Presentation, slide_index: int) -> None:
    """Duplicate a slide; print the copy's index and id."""
    target = deck.slides[slide_index]
    with deck.edit(f"CLI: duplicate slide {slide_index}"):
        new = target.duplicate()
    payload = {"ok": True, "index": new.index, "id": new.id, "from": slide_index}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"duplicated slide {slide_index} -> {new.index}",
    )


@slide.command(name="move")
@click.option(
    "--slide", "slide_index", type=int, required=True, help="1-based slide index to move."
)
@click.option("--to", "to_index", type=int, required=True, help="1-based destination position.")
@_deck_command
def slide_move(ctx: click.Context, deck: Presentation, slide_index: int, to_index: int) -> None:
    """Move a slide to a new position."""
    target = deck.slides[slide_index]
    with deck.edit(f"CLI: move slide {slide_index} -> {to_index}"):
        moved = target.move_to(to_index)
    payload = {"ok": True, "index": moved.index, "id": moved.id}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"moved slide to position {moved.index}",
    )


@slide.command(name="set-layout")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--layout", required=True, help="Layout name to apply. See `slide layouts` for names."
)
@_deck_command
def slide_set_layout(ctx: click.Context, deck: Presentation, slide_index: int, layout: str) -> None:
    """Re-apply a slide's layout by name."""
    target = deck.slides[slide_index]
    with deck.edit(f"CLI: set layout of slide {slide_index}"):
        target.set_layout(layout)
    payload = {"ok": True, "index": slide_index, "layout": target.layout_name}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"slide {slide_index} layout -> {target.layout_name}",
    )


@slide.command(name="set-transition")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--effect",
    "effect",
    type=click.Choice(ENTRY_EFFECT_CHOICES),
    default=None,
    help="Entrance transition effect.",
)
@click.option("--duration", type=float, default=None, help="Transition length in seconds.")
@click.option(
    "--advance-after",
    "advance_after",
    type=float,
    default=None,
    help="Auto-advance after N seconds (sets the timer + flag).",
)
@click.option(
    "--on-click/--no-on-click",
    "advance_on_click",
    default=None,
    help="Whether a click advances the slide.",
)
@_deck_command
def slide_set_transition(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    effect: str | None,
    duration: float | None,
    advance_after: float | None,
    advance_on_click: bool | None,
) -> None:
    """Set a slide's entrance transition (effect / duration / auto-advance)."""
    if effect is None and duration is None and advance_after is None and advance_on_click is None:
        raise click.UsageError(
            "slide set-transition requires --effect, --duration, --advance-after, "
            "and/or --on-click/--no-on-click"
        )
    target = deck.slides[slide_index]
    with deck.edit(f"CLI: set transition of slide {slide_index}"):
        trans = target.set_transition(
            effect,
            duration=duration,
            advance_after=advance_after,
            advance_on_click=advance_on_click,
        )
    emit(
        {"ok": True, "index": slide_index, "transition": trans},
        as_text=not ctx.obj["as_json"],
        text=f"slide {slide_index} transition -> {trans['effect']}",
    )


@slide.command(name="set-background")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--color", "color", default=None, help="Solid background color (#RRGGBB). Overrides the master."
)
@click.option(
    "--follow-master",
    "follow_master",
    is_flag=True,
    default=False,
    help="Drop the per-slide override; inherit the master background.",
)
@_deck_command
def slide_set_background(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    color: str | None,
    follow_master: bool,
) -> None:
    """Give a slide its own solid background color, or revert it to the master."""
    if (color is None) == (not follow_master):
        raise click.UsageError(
            "slide set-background requires exactly one of --color or --follow-master"
        )
    target = deck.slides[slide_index]
    with deck.edit(f"CLI: set background of slide {slide_index}"):
        if follow_master:
            bg = target.follow_master_background()
        else:
            assert color is not None  # narrowed by the exactly-one check above
            bg = target.set_background(color)
    dest = "master" if follow_master else color
    emit(
        {"ok": True, "index": slide_index, "background": bg},
        as_text=not ctx.obj["as_json"],
        text=f"slide {slide_index} background -> {dest}",
    )


# ---------------------------------------------------------------------------
# headers / footers — footer / slide-number / date (slide override + master default)
# ---------------------------------------------------------------------------


def _fmt_hf(hf: dict[str, Any]) -> str:
    f, s, d = hf["footer"], hf["slide_number"], hf["date"]
    return (
        f"footer: {'on' if f['visible'] else 'off'}"
        + (f" {f['text']!r}" if f["text"] else "")
        + f" | slide#: {'on' if s['visible'] else 'off'}"
        + f" | date: {'on' if d['visible'] else 'off'}"
        + (f" {d['text']!r}" if d["text"] else "")
    )


@slide.command(name="headers-footers")
@click.argument("index", type=int)
@_deck_command
def slide_headers_footers(ctx: click.Context, deck: Presentation, index: int) -> None:
    """Read slide INDEX's footer / slide-number / date settings."""
    hf = deck.slides[index].headers_footers.read()
    emit(hf, as_text=not ctx.obj["as_json"], text=_fmt_hf(hf))


@slide.command(name="set-footer")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--text", "text", default=None, help="Footer text (setting it auto-shows the footer)."
)
@click.option("--show/--hide", "visible", default=None, help="Show or hide the footer.")
@_deck_command
def slide_set_footer(
    ctx: click.Context, deck: Presentation, slide_index: int, text: str | None, visible: bool | None
) -> None:
    """Set slide S's footer text and/or visibility."""
    if text is None and visible is None:
        raise click.UsageError("slide set-footer requires --text and/or --show/--hide")
    with deck.edit(f"CLI: set footer on slide {slide_index}"):
        hf = deck.slides[slide_index].headers_footers.set_footer(text=text, visible=visible)
    emit(
        {"ok": True, "slide": slide_index, "headers_footers": hf},
        as_text=not ctx.obj["as_json"],
        text=_fmt_hf(hf),
    )


@slide.command(name="slide-number")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option("--show/--hide", "visible", default=True, help="Show or hide the slide number.")
@_deck_command
def slide_slide_number(
    ctx: click.Context, deck: Presentation, slide_index: int, visible: bool
) -> None:
    """Show or hide slide S's auto slide number."""
    with deck.edit(f"CLI: slide number on slide {slide_index}"):
        hf = deck.slides[slide_index].headers_footers.set_slide_number(visible)
    emit(
        {"ok": True, "slide": slide_index, "headers_footers": hf},
        as_text=not ctx.obj["as_json"],
        text=_fmt_hf(hf),
    )


@slide.command(name="set-date")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option("--text", "text", default=None, help="Fixed date text (auto-shows the date).")
@click.option(
    "--format",
    "fmt",
    type=int,
    default=None,
    help="Auto-updating date format (PpDateTimeFormat int).",
)
@click.option("--show/--hide", "visible", default=None, help="Show or hide the date.")
@_deck_command
def slide_set_date(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    text: str | None,
    fmt: int | None,
    visible: bool | None,
) -> None:
    """Set slide S's date placeholder (fixed --text or auto --format)."""
    if text is None and fmt is None and visible is None:
        raise click.UsageError("slide set-date requires --text, --format, and/or --show/--hide")
    with deck.edit(f"CLI: set date on slide {slide_index}"):
        hf = deck.slides[slide_index].headers_footers.set_date(visible=visible, text=text, fmt=fmt)
    emit(
        {"ok": True, "slide": slide_index, "headers_footers": hf},
        as_text=not ctx.obj["as_json"],
        text=_fmt_hf(hf),
    )


# ---------------------------------------------------------------------------
# shapes --slide S
# ---------------------------------------------------------------------------


@click.command(name="shapes")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@_deck_command
def shapes_cmd(ctx: click.Context, deck: Presentation, slide_index: int) -> None:
    """List the shapes on a slide (anchor_id, name, id, type, geometry, text)."""
    rows = deck.slides[slide_index].shapes.list()
    emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_shapes(rows))


# ---------------------------------------------------------------------------
# shape add|move|resize|delete
# ---------------------------------------------------------------------------


@click.group(name="shape")
def shape() -> None:
    """Create, place & style shapes: add, move, resize, delete, fill, order (points)."""


def _resolve_shape(deck: Presentation, anchor_id: str) -> Shape:
    """Resolve a shape/placeholder anchor to a `Shape` (else exit-2 not found)."""
    anchor = deck.anchor_by_id(anchor_id)
    if not isinstance(anchor, Shape):
        raise AnchorNotFoundError("shape", anchor_id)
    return anchor


def _resolve_shapes(deck: Presentation, anchors: str) -> list[Shape]:
    """Resolve a comma-separated list of shape anchors to `Shape` handles."""
    ids = [a.strip() for a in anchors.split(",") if a.strip()]
    if not ids:
        raise click.UsageError("--anchors needs at least one shape anchor")
    return [_resolve_shape(deck, a) for a in ids]


@shape.command(name="add")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--kind",
    type=click.Choice(["textbox", "shape", "picture", "table", "chart", "smartart"]),
    required=True,
    help="What to add.",
)
@click.option("--text", "text", default=None, help="Initial text (textbox/shape).")
@click.option("--rows", type=int, default=None, help="Row count (required for --kind table).")
@click.option("--cols", type=int, default=None, help="Column count (required for --kind table).")
@click.option(
    "--chart-type",
    "chart_type",
    type=click.Choice(CHART_TYPE_CHOICES),
    default="column",
    show_default=True,
    help="Chart kind (for --kind chart).",
)
@click.option(
    "--categories",
    "categories",
    default=None,
    help="Chart category labels: a JSON array or comma-separated list (--kind chart).",
)
@click.option(
    "--series",
    "series",
    default=None,
    help='Chart series as a JSON object {"name":[values]} or array of [name,[values]] '
    "(--kind chart).",
)
@click.option(
    "--smartart-kind",
    "smartart_kind",
    type=click.Choice(SMARTART_CHOICES),
    default="process",
    show_default=True,
    help="SmartArt layout (for --kind smartart).",
)
@click.option(
    "--nodes",
    "nodes",
    default=None,
    help="SmartArt nodes: a JSON array of strings and/or {text, children} objects "
    "(--kind smartart).",
)
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
    "--fill",
    "fill",
    default=None,
    help='Fill color (#RRGGBB) or "none" for transparent (textbox/shape).',
)
@click.option(
    "--line",
    "line",
    default=None,
    help='Border color (#RRGGBB) or "none" for no border (textbox/shape).',
)
@click.option(
    "--line-width", "line_width", type=float, default=None, help="Border weight in points."
)
@click.option(
    "--alt-text",
    "alt_text",
    default=None,
    help="Alternative text for a picture (a drift-proof re-identification handle).",
)
@_deck_command
def shape_add(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    kind: str,
    text: str | None,
    rows: int | None,
    cols: int | None,
    path: str | None,
    shape_type: str,
    chart_type: str,
    categories: str | None,
    series: str | None,
    smartart_kind: str,
    nodes: str | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
    fill: str | None,
    line: str | None,
    line_width: float | None,
    alt_text: str | None,
) -> None:
    """Add a shape to a slide; print its anchor_id, name, type, and geometry."""
    if kind == "picture" and not path:
        raise click.UsageError("shape add --kind picture requires --path")
    if kind == "table" and (rows is None or cols is None):
        raise click.UsageError("shape add --kind table requires --rows and --cols")
    cats = _parse_categories(categories)
    ser = _parse_series(series)
    if (cats is None) != (ser is None):
        raise click.UsageError("shape add --kind chart needs both --categories and --series")
    sa_nodes = _parse_nodes(nodes)
    shapes = deck.slides[slide_index].shapes  # exit 2 if slide out of range
    with deck.edit(f"CLI: add {kind} on slide {slide_index}"):
        if kind == "textbox":
            new = shapes.add_textbox(
                text or "",
                left=left,
                top=top,
                width=width,
                height=height,
                fill=fill,
                line=line,
                line_width=line_width,
            )
        elif kind == "shape":
            new = shapes.add_shape(
                shape_type,
                left=left,
                top=top,
                width=width,
                height=height,
                fill=fill,
                line=line,
                line_width=line_width,
            )
            if text:
                new.set_text(text)
        elif kind == "table":
            assert rows is not None and cols is not None  # guarded above
            new = shapes.add_table(rows, cols, left=left, top=top, width=width, height=height)
        elif kind == "chart":
            new = shapes.add_chart(
                chart_type, cats, ser, left=left, top=top, width=width, height=height
            )
        elif kind == "smartart":
            new = shapes.add_smartart(
                smartart_kind, sa_nodes, left=left, top=top, width=width, height=height
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


@shape.command(name="move")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to move (shape:S:N or ph:S:KIND)."
)
@click.option("--left", type=float, default=None, help="New left edge (points).")
@click.option("--top", type=float, default=None, help="New top edge (points).")
@_deck_command
def shape_move(
    ctx: click.Context, deck: Presentation, anchor_id: str, left: float | None, top: float | None
) -> None:
    """Move a shape to an absolute position (points)."""
    if left is None and top is None:
        raise click.UsageError("shape move requires --left and/or --top")
    sh = _resolve_shape(deck, anchor_id)  # exit 2 if missing / not a shape
    with deck.edit(f"CLI: move {anchor_id}"):
        sh.move(left=left, top=top)
    geo = sh.geometry()
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "geometry": geo},
        as_text=not ctx.obj["as_json"],
        text=f"moved {sh.anchor_id}: {_fmt_geometry(geo)}",
    )


@shape.command(name="resize")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to resize (shape:S:N or ph:S:KIND)."
)
@click.option("--width", type=float, default=None, help="New width (points).")
@click.option("--height", type=float, default=None, help="New height (points).")
@_deck_command
def shape_resize(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    width: float | None,
    height: float | None,
) -> None:
    """Set a shape's size (points)."""
    if width is None and height is None:
        raise click.UsageError("shape resize requires --width and/or --height")
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: resize {anchor_id}"):
        sh.resize(width=width, height=height)
    geo = sh.geometry()
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "geometry": geo},
        as_text=not ctx.obj["as_json"],
        text=f"resized {sh.anchor_id}: {_fmt_geometry(geo)}",
    )


@shape.command(name="delete")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to delete (shape:S:N or ph:S:KIND)."
)
@_deck_command
def shape_delete(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Delete a shape from its slide."""
    sh = _resolve_shape(deck, anchor_id)
    info = {"anchor_id": sh.anchor_id, "name": sh.name, "id": sh.shape_id}
    with deck.edit(f"CLI: delete {anchor_id}"):
        sh.delete()
    emit(
        {"ok": True, **info},
        as_text=not ctx.obj["as_json"],
        text=f"deleted {info['anchor_id']} ({info['name']!r})",
    )


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
@_deck_command
def shape_export(
    ctx: click.Context, deck: Presentation, anchor_id: str, out: str | None, fmt: str
) -> None:
    """Render a single shape to an image file — so a vision model can *see* it.

    The per-shape complement to `slide export`: crops to the shape's bounds at
    its native pixel size (no size override — `Shape.Export` doesn't honor one
    reliably). Polite (doesn't move the view). Prints the absolute path.
    """
    sh = _resolve_shape(deck, anchor_id)
    path = sh.export_image(out, fmt=fmt)
    payload = {"ok": True, "anchor_id": sh.anchor_id, "path": str(path), "format": fmt}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"exported {sh.anchor_id} -> {path}",
    )


@shape.command(name="set-alt")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to tag (shape:S:N or ph:S:KIND)."
)
@click.option("--alt-text", "alt_text", required=True, help="Alternative (accessibility) text.")
@_deck_command
def shape_set_alt(ctx: click.Context, deck: Presentation, anchor_id: str, alt_text: str) -> None:
    """Set a shape's alternative text — a drift-proof re-identification handle."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: set alt-text {anchor_id}"):
        sh.set_alt_text(alt_text)
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "alt_text": sh.alt_text},
        as_text=not ctx.obj["as_json"],
        text=f"set alt-text on {sh.anchor_id}: {alt_text!r}",
    )


@shape.command(name="fill")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to style (shape:S:N / shapeid:S:ID / ph).",
)
@click.option(
    "--fill", "fill", default=None, help='Fill color (#RRGGBB) or "none" for transparent.'
)
@click.option(
    "--line", "line", default=None, help='Border color (#RRGGBB) or "none" for no border.'
)
@click.option(
    "--line-width", "line_width", type=float, default=None, help="Border weight in points."
)
@click.option(
    "--fill-transparency",
    "fill_transparency",
    type=float,
    default=None,
    help="Fill alpha 0..1 (0 opaque, 1 fully transparent).",
)
@click.option(
    "--line-transparency",
    "line_transparency",
    type=float,
    default=None,
    help="Border alpha 0..1 (0 opaque, 1 fully transparent).",
)
@_deck_command
def shape_fill(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    fill: str | None,
    line: str | None,
    line_width: float | None,
    fill_transparency: float | None,
    line_transparency: float | None,
) -> None:
    """Set a shape's fill and/or line (border) color — distinct from font color."""
    if (
        fill is None
        and line is None
        and line_width is None
        and fill_transparency is None
        and line_transparency is None
    ):
        raise click.UsageError(
            "shape fill requires --fill, --line, --line-width, "
            "--fill-transparency, and/or --line-transparency"
        )
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: fill {anchor_id}"):
        sh.set_fill(
            fill=fill,
            line=line,
            line_width=line_width,
            fill_transparency=fill_transparency,
            line_transparency=line_transparency,
        )
    payload = {"ok": True, **sh.to_dict()}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"styled {sh.anchor_id} (fill={payload['fill']}, line={payload['line']})",
    )


def _split_csv(value: str) -> list[str]:
    """Split a comma-separated CLI value into trimmed, non-empty tokens."""
    return [tok.strip() for tok in value.split(",") if tok.strip()]


@shape.command(name="gradient-fill")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to fill (shape:S:N / shapeid:S:ID / ph)."
)
@click.option(
    "--colors",
    "colors",
    default=None,
    help='Comma-separated colors, e.g. "#1a73e8,#fff". 1=one-color, 2=two-color, 3+=multi-stop.',
)
@click.option(
    "--preset",
    "preset",
    type=click.Choice(PRESET_GRADIENT_CHOICES),
    default=None,
    help='A named gradient ramp ("ocean", "fire", "rainbow", …) instead of --colors.',
)
@click.option(
    "--positions",
    "positions",
    default=None,
    help="Comma-separated stop positions 0..1 (same count as --colors; places interior stops).",
)
@click.option(
    "--style",
    "style",
    type=click.Choice(GRADIENT_STYLE_CHOICES),
    default="horizontal",
    show_default=True,
    help="Gradient direction.",
)
@click.option(
    "--variant", "variant", type=int, default=1, show_default=True, help="Shading variant 1-4."
)
@click.option(
    "--degree",
    "degree",
    type=float,
    default=None,
    help="One-color gradient brightness 0..1 (single-color only).",
)
@_deck_command
def shape_gradient_fill(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    colors: str | None,
    preset: str | None,
    positions: str | None,
    style: str,
    variant: int,
    degree: float | None,
) -> None:
    """Give a shape a gradient fill (two-color / multi-stop / one-color / preset)."""
    if colors is None and preset is None:
        raise click.UsageError("shape gradient-fill requires --colors or --preset")
    color_list = _split_csv(colors) if colors else None
    pos_list = [float(p) for p in _split_csv(positions)] if positions else None
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: gradient fill {anchor_id}"):
        sh.set_gradient_fill(
            color_list,
            positions=pos_list,
            style=style,
            variant=variant,
            degree=degree,
            preset=preset,
        )
    payload = {"ok": True, **sh.to_dict()}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"gradient-filled {sh.anchor_id} (type={payload['fill'] and payload['fill'].get('type')})",
    )


@shape.command(name="picture-fill")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to fill (shape:S:N / shapeid:S:ID / ph)."
)
@click.option(
    "--path", "path", required=True, help="Image file to fill with (resolved to absolute)."
)
@_deck_command
def shape_picture_fill(ctx: click.Context, deck: Presentation, anchor_id: str, path: str) -> None:
    """Fill a shape with an image."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: picture fill {anchor_id}"):
        sh.set_picture_fill(path)
    payload = {"ok": True, **sh.to_dict()}
    emit(payload, as_text=not ctx.obj["as_json"], text=f"picture-filled {sh.anchor_id}")


@shape.command(name="set-picture")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Picture to re-source (shape:S:N / shapeid:S:ID / name).",
)
@click.option(
    "--path", "path", required=True, help="New image file (embedded; resolved to absolute)."
)
@click.option(
    "--alt-text",
    "alt_text",
    default=None,
    help="Override the carried-over alt text (defaults to keeping the old picture's).",
)
@_deck_command
def shape_set_picture(
    ctx: click.Context, deck: Presentation, anchor_id: str, path: str, alt_text: str | None
) -> None:
    """Re-source a picture in place (swap the image, keep box/name/alt/z-order)."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: set picture {anchor_id}"):
        new = sh.set_picture(path, alt_text=alt_text)
    payload = {"ok": True, **new.to_dict()}
    emit(payload, as_text=not ctx.obj["as_json"], text=f"re-sourced picture -> {new.shapeid}")


@shape.command(name="pattern-fill")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to fill (shape:S:N / shapeid:S:ID / ph)."
)
@click.option(
    "--pattern",
    "pattern",
    type=click.Choice(PATTERN_CHOICES),
    required=True,
    help="Pattern name (percent_50, trellis, dark_horizontal, …).",
)
@click.option("--fore", "fore", required=True, help="Pattern foreground color (#RRGGBB).")
@click.option("--back", "back", default=None, help="Pattern background color (#RRGGBB).")
@_deck_command
def shape_pattern_fill(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    pattern: str,
    fore: str,
    back: str | None,
) -> None:
    """Give a shape a two-color pattern fill."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: pattern fill {anchor_id}"):
        sh.set_pattern_fill(pattern, fore=fore, back=back)
    payload = {"ok": True, **sh.to_dict()}
    emit(payload, as_text=not ctx.obj["as_json"], text=f"pattern-filled {sh.anchor_id} ({pattern})")


def _effect_arg(raw: str | None, *, as_dict: bool, label: str) -> Any:
    """Parse an --shadow/--glow (JSON dict) or --soft-edge/--reflection (int) effect arg.

    `"none"` (any case) means turn the effect off; otherwise a JSON dict for the
    dict-effects or an int preset for the others. Returns None when unset.
    """
    if raw is None:
        return None
    if raw.strip().lower() == "none":
        return "none"
    if as_dict:
        try:
            val = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise click.UsageError(
                f'--{label} must be JSON like \'{{"color":"#000"}}\' or "none"'
            ) from exc
        if not isinstance(val, dict):
            raise click.UsageError(f"--{label} JSON must be an object, got {type(val).__name__}")
        return val
    try:
        return int(raw)
    except ValueError as exc:
        raise click.UsageError(f'--{label} must be an int preset or "none"') from exc


@shape.command(name="effect")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to style (shape:S:N / shapeid:S:ID / ph).",
)
@click.option(
    "--shadow",
    "shadow",
    default=None,
    help='JSON dict {color,transparency,blur,size,offset_x,offset_y} or "none".',
)
@click.option(
    "--glow", "glow", default=None, help='JSON dict {color,radius,transparency} or "none".'
)
@click.option("--soft-edge", "soft_edge", default=None, help='Soft-edge preset 0-6 or "none".')
@click.option("--reflection", "reflection", default=None, help='Reflection preset 0-9 or "none".')
@_deck_command
def shape_effect(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    shadow: str | None,
    glow: str | None,
    soft_edge: str | None,
    reflection: str | None,
) -> None:
    """Set a shape's effects — shadow / glow / soft-edge / reflection."""
    shadow_arg = _effect_arg(shadow, as_dict=True, label="shadow")
    glow_arg = _effect_arg(glow, as_dict=True, label="glow")
    soft_arg = _effect_arg(soft_edge, as_dict=False, label="soft-edge")
    refl_arg = _effect_arg(reflection, as_dict=False, label="reflection")
    if shadow_arg is None and glow_arg is None and soft_arg is None and refl_arg is None:
        raise click.UsageError(
            "shape effect requires --shadow, --glow, --soft-edge, and/or --reflection"
        )
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: effects {anchor_id}"):
        sh.set_effect(shadow=shadow_arg, glow=glow_arg, soft_edge=soft_arg, reflection=refl_arg)
    payload = {"ok": True, **sh.to_dict()}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"set effects on {sh.anchor_id} (effects={payload.get('effects')})",
    )


@shape.command(name="line-style")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to style (shape:S:N / shapeid:S:ID / ph).",
)
@click.option(
    "--dash", "dash", type=click.Choice(DASH_STYLE_CHOICES), default=None, help="Line dash pattern."
)
@click.option(
    "--begin-arrow",
    "begin_arrow",
    type=click.Choice(ARROWHEAD_STYLE_CHOICES),
    default=None,
    help="Start arrowhead (lines/connectors only).",
)
@click.option(
    "--end-arrow",
    "end_arrow",
    type=click.Choice(ARROWHEAD_STYLE_CHOICES),
    default=None,
    help="End arrowhead (lines/connectors only).",
)
@click.option(
    "--begin-arrow-size",
    "begin_arrow_size",
    type=click.Choice(ARROWHEAD_SIZE_CHOICES),
    default=None,
    help="Start arrowhead size.",
)
@click.option(
    "--end-arrow-size",
    "end_arrow_size",
    type=click.Choice(ARROWHEAD_SIZE_CHOICES),
    default=None,
    help="End arrowhead size.",
)
@_deck_command
def shape_line_style(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    dash: str | None,
    begin_arrow: str | None,
    end_arrow: str | None,
    begin_arrow_size: str | None,
    end_arrow_size: str | None,
) -> None:
    """Set a shape's line dash pattern and/or arrowheads (arrowheads: lines/connectors only)."""
    if (
        dash is None
        and begin_arrow is None
        and end_arrow is None
        and begin_arrow_size is None
        and end_arrow_size is None
    ):
        raise click.UsageError(
            "shape line-style requires --dash, --begin-arrow, --end-arrow, "
            "--begin-arrow-size, and/or --end-arrow-size"
        )
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: line-style {anchor_id}"):
        sh.set_line_style(
            dash=dash,
            begin_arrow=begin_arrow,
            end_arrow=end_arrow,
            begin_arrow_size=begin_arrow_size,
            end_arrow_size=end_arrow_size,
        )
    payload = {"ok": True, **sh.to_dict()}
    emit(
        payload,
        as_text=not ctx.obj["as_json"],
        text=f"styled line on {sh.anchor_id} (line={payload['line']})",
    )


@shape.command(name="order")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to restack (shape:S:N / shapeid:S:ID / ph).",
)
@click.option(
    "--to",
    "to",
    type=click.Choice(ZORDER_CHOICES),
    required=True,
    help="front / back / forward / backward.",
)
@_deck_command
def shape_order(ctx: click.Context, deck: Presentation, anchor_id: str, to: str) -> None:
    """Restack a shape in the slide z-order; print its new 1-based position."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: order {anchor_id} {to}"):
        new_index = sh.reorder(to)
    emit(
        {
            "ok": True,
            "anchor_id": sh.anchor_id,
            "shapeid": sh.shapeid,
            "name": sh.name,
            "index": new_index,
        },
        as_text=not ctx.obj["as_json"],
        text=f"sent {sh.anchor_id} to {to} (now z-index {new_index}); stable handle {sh.shapeid}",
    )


@shape.command(name="animate")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to animate (shape:S:N / shapeid:S:ID / ph).",
)
@click.option(
    "--effect",
    "effect",
    type=click.Choice(ANIM_EFFECT_CHOICES),
    default="fade",
    show_default=True,
    help="Animation effect.",
)
@click.option(
    "--trigger",
    "trigger",
    type=click.Choice(ANIM_TRIGGER_CHOICES),
    default="on_click",
    show_default=True,
    help="When the animation fires.",
)
@click.option("--duration", type=float, default=None, help="Animation length in seconds.")
@click.option("--delay", type=float, default=None, help="Start delay in seconds.")
@click.option(
    "--exit", "exit_", is_flag=True, default=False, help="Animate the shape OUT (exit effect)."
)
@_deck_command
def shape_animate(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    effect: str,
    trigger: str,
    duration: float | None,
    delay: float | None,
    exit_: bool,
) -> None:
    """Add an entrance (or --exit) animation to a shape; print the effect."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: animate {anchor_id} {effect}"):
        animation = sh.animate(effect, trigger=trigger, duration=duration, delay=delay, exit=exit_)
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "animation": animation},
        as_text=not ctx.obj["as_json"],
        text=(
            f"animated {sh.anchor_id}: {'exit ' if animation['exit'] else ''}"
            f"{animation['effect']} / {animation['trigger']}"
        ),
    )


@shape.command(name="clear-animations")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape whose animations to clear (shape:S:N / shapeid:S:ID / ph).",
)
@_deck_command
def shape_clear_animations(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Remove all animations targeting a shape; print how many were deleted."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: clear animations on {anchor_id}"):
        removed = sh.clear_animations()
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "shapeid": sh.shapeid, "removed": removed},
        as_text=not ctx.obj["as_json"],
        text=f"{sh.anchor_id}: cleared {removed} animation(s)",
    )


@shape.command(name="set-link")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to link (shape:S:N / shapeid:S:ID / ph).",
)
@click.option("--url", "url", default=None, help="External link target (URL / mailto / file path).")
@click.option(
    "--slide", "slide_index", type=int, default=None, help="In-deck jump to this 1-based slide."
)
@click.option("--screen-tip", "screen_tip", default=None, help="Hover tooltip (optional).")
@_deck_command
def shape_set_link(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    url: str | None,
    slide_index: int | None,
    screen_tip: str | None,
) -> None:
    """Make a shape a clickable hyperlink — an external --url or an in-deck --slide jump."""
    if (url is None) == (slide_index is None):
        raise click.UsageError("shape set-link requires exactly one of --url or --slide")
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: set link {anchor_id}"):
        link = sh.set_hyperlink(url=url, slide=slide_index, screen_tip=screen_tip)
    dest = url if url is not None else f"slide {slide_index}"
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "hyperlink": link},
        as_text=not ctx.obj["as_json"],
        text=f"linked {sh.anchor_id} -> {dest}",
    )


@shape.command(name="remove-link")
@click.option(
    "--anchor-id",
    "anchor_id",
    required=True,
    help="Shape to unlink (shape:S:N / shapeid:S:ID / ph).",
)
@_deck_command
def shape_remove_link(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Remove a shape's mouse-click hyperlink (a no-op if it has none)."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: remove link {anchor_id}"):
        sh.remove_hyperlink()
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "hyperlink": None},
        as_text=not ctx.obj["as_json"],
        text=f"unlinked {sh.anchor_id}",
    )


@shape.command(name="group")
@click.option(
    "--anchors",
    "anchors",
    required=True,
    help="Shapes to group: comma-separated shape:S:N / shapeid:S:ID / names (>= 2).",
)
@_deck_command
def shape_group(ctx: click.Context, deck: Presentation, anchors: str) -> None:
    """Group two or more shapes on one slide into a single group shape."""
    members = _resolve_shapes(deck, anchors)
    with deck.edit("CLI: group shapes"):
        group = members[0].slide.shapes.group(members)
    emit(
        {
            "ok": True,
            "anchor_id": group.anchor_id,
            "shapeid": group.shapeid,
            "group_item_ids": group.to_dict().get("group_item_ids", []),
        },
        as_text=not ctx.obj["as_json"],
        text=f"grouped {len(members)} shapes -> {group.shapeid}",
    )


@shape.command(name="ungroup")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Group shape to ungroup (shape:/shapeid:/name)."
)
@_deck_command
def shape_ungroup(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Ungroup a group shape; the freed members keep their original ids."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: ungroup {anchor_id}"):
        children = sh.ungroup()
    ids = [c.shapeid for c in children]
    emit(
        {"ok": True, "ungrouped": ids},
        as_text=not ctx.obj["as_json"],
        text=f"ungrouped into {len(ids)} shapes",
    )


@shape.command(name="align")
@click.option("--anchors", "anchors", required=True, help="Shapes to align (comma-separated).")
@click.option(
    "--how",
    "how",
    required=True,
    type=click.Choice(ALIGN_CHOICES),
    help="left/center/right (horizontal) or top/middle/bottom (vertical).",
)
@click.option(
    "--relative-to",
    "relative_to",
    default="slide",
    type=click.Choice(RELATIVE_TO_CHOICES),
    help="Align to the slide (default) or to the selection.",
)
@_deck_command
def shape_align(
    ctx: click.Context, deck: Presentation, anchors: str, how: str, relative_to: str
) -> None:
    """Align a set of shapes to a common edge / center."""
    members = _resolve_shapes(deck, anchors)
    with deck.edit("CLI: align shapes"):
        members[0].slide.shapes.align(members, how, relative_to=relative_to)
    emit(
        {"ok": True, "how": how, "shapeids": [m.shapeid for m in members]},
        as_text=not ctx.obj["as_json"],
        text=f"aligned {len(members)} shapes ({how}, vs {relative_to})",
    )


@shape.command(name="distribute")
@click.option("--anchors", "anchors", required=True, help="Shapes to distribute (comma-separated).")
@click.option(
    "--how",
    "how",
    required=True,
    type=click.Choice(DISTRIBUTE_CHOICES),
    help="horizontal or vertical even spacing.",
)
@click.option(
    "--relative-to",
    "relative_to",
    default="slide",
    type=click.Choice(RELATIVE_TO_CHOICES),
    help="Distribute across the slide (default) or within the selection.",
)
@_deck_command
def shape_distribute(
    ctx: click.Context, deck: Presentation, anchors: str, how: str, relative_to: str
) -> None:
    """Space three or more shapes evenly on one axis."""
    members = _resolve_shapes(deck, anchors)
    with deck.edit("CLI: distribute shapes"):
        members[0].slide.shapes.distribute(members, how, relative_to=relative_to)
    emit(
        {"ok": True, "how": how, "shapeids": [m.shapeid for m in members]},
        as_text=not ctx.obj["as_json"],
        text=f"distributed {len(members)} shapes ({how})",
    )


@shape.command(name="connect")
@click.option(
    "--type",
    "connector_type",
    default="straight",
    type=click.Choice(CONNECTOR_CHOICES),
    help="straight / elbow / curved.",
)
@click.option("--begin", "begin", default=None, help="Shape anchor glued to the start end.")
@click.option("--end", "end", default=None, help="Shape anchor glued to the finish end.")
@click.option(
    "--begin-site", "begin_site", type=int, default=1, help="Start connection site (1-based)."
)
@click.option(
    "--end-site", "end_site", type=int, default=1, help="Finish connection site (1-based)."
)
@click.option(
    "--slide",
    "slide_index",
    type=int,
    default=None,
    help="Slide for the geometry form (when no --begin/--end).",
)
@click.option("--left", "left", type=float, default=None, help="Geometry form: start x (points).")
@click.option("--top", "top", type=float, default=None, help="Geometry form: start y (points).")
@click.option(
    "--width", "width", type=float, default=None, help="Geometry form: span width (points)."
)
@click.option(
    "--height", "height", type=float, default=None, help="Geometry form: span height (points)."
)
@_deck_command
def shape_connect(
    ctx: click.Context,
    deck: Presentation,
    connector_type: str,
    begin: str | None,
    end: str | None,
    begin_site: int,
    end_site: int,
    slide_index: int | None,
    left: float | None,
    top: float | None,
    width: float | None,
    height: float | None,
) -> None:
    """Draw a connector — glue it to two shapes (--begin/--end) or a free geometry box."""
    begin_sh = _resolve_shape(deck, begin) if begin else None
    end_sh = _resolve_shape(deck, end) if end else None
    if begin_sh is not None:
        coll = begin_sh.slide.shapes
    elif slide_index is not None:
        coll = deck.slides[slide_index].shapes
    else:
        raise click.UsageError(
            "shape connect needs --begin and --end shapes, or --slide for the geometry form"
        )
    with deck.edit("CLI: add connector"):
        conn = coll.add_connector(
            connector_type,
            begin=begin_sh,
            end=end_sh,
            begin_site=begin_site,
            end_site=end_site,
            left=left,
            top=top,
            width=width,
            height=height,
        )
    emit(
        {
            "ok": True,
            "anchor_id": conn.anchor_id,
            "shapeid": conn.shapeid,
            "connector": conn.to_dict().get("connector"),
            "geometry": conn.geometry(),
        },
        as_text=not ctx.obj["as_json"],
        text=f"connector {conn.shapeid} ({connector_type})",
    )


@shape.command(name="reset-to-layout")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Placeholder to restore (ph:/shape:/shapeid:)."
)
@_deck_command
def shape_reset_to_layout(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Restore a placeholder's geometry + default font size from its slide layout."""
    sh = _resolve_shape(deck, anchor_id)
    with deck.edit(f"CLI: reset to layout {anchor_id}"):
        restored = sh.reset_to_layout()
    emit(
        {"ok": True, "anchor_id": sh.anchor_id, "restored": restored},
        as_text=not ctx.obj["as_json"],
        text=f"reset {sh.anchor_id} to its layout placeholder",
    )


# ---------------------------------------------------------------------------
# text-run-level hyperlinks (v-next): link set / remove / list
# ---------------------------------------------------------------------------


@click.group(name="link")
def link() -> None:
    """Link a word/span inside any text anchor (shape / para / cell / notes)."""


@link.command(name="set")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor (shape:/para:/cell:/notes:)."
)
@click.option("--text", "text", default=None, help="Substring to link (the LLM-friendly form).")
@click.option(
    "--start", "start", type=int, default=None, help="0-based span start (with --length)."
)
@click.option("--length", "length", type=int, default=None, help="Span length (with --start).")
@click.option("--url", "url", default=None, help="External link target (URL / mailto / file path).")
@click.option(
    "--slide", "slide_index", type=int, default=None, help="In-deck jump to this 1-based slide."
)
@click.option("--screen-tip", "screen_tip", default=None, help="Hover tooltip (optional).")
@_deck_command
def link_set(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    text: str | None,
    start: int | None,
    length: int | None,
    url: str | None,
    slide_index: int | None,
    screen_tip: str | None,
) -> None:
    """Link a span of text — by --text substring or --start/--length — to --url or --slide."""
    if (url is None) == (slide_index is None):
        raise click.UsageError("link set requires exactly one of --url or --slide")
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: set link on {anchor_id}"):
        result = anchor.set_link(
            text=text, start=start, length=length, url=url, slide=slide_index, screen_tip=screen_tip
        )
    dest = url if url is not None else f"slide {slide_index}"
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id, "link": result},
        as_text=not ctx.obj["as_json"],
        text=f"linked {result['text']!r} -> {dest}",
    )


@link.command(name="remove")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor (shape:/para:/cell:/notes:)."
)
@click.option("--text", "text", default=None, help="Substring whose link to clear.")
@click.option(
    "--start", "start", type=int, default=None, help="0-based span start (with --length)."
)
@click.option("--length", "length", type=int, default=None, help="Span length (with --start).")
@_deck_command
def link_remove(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    text: str | None,
    start: int | None,
    length: int | None,
) -> None:
    """Remove a span's link, or ALL links in the anchor when no span is given."""
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: remove link on {anchor_id}"):
        removed = anchor.remove_link(text=text, start=start, length=length)
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id, "removed": removed},
        as_text=not ctx.obj["as_json"],
        text=f"removed {removed} link(s) from {anchor.anchor_id}",
    )


@link.command(name="list")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor (shape:/para:/cell:/notes:)."
)
@_deck_command
def link_list(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """List the text-run hyperlinks in a text anchor (text, start, length, address)."""
    anchor = deck.anchor_by_id(anchor_id)
    rows = anchor.links()
    emit(rows, as_text=not ctx.obj["as_json"], text=f"{len(rows)} link(s) in {anchor.anchor_id}")


# ---------------------------------------------------------------------------
# text structure (v0.3): paragraphs, insert, format-paragraph, format-text, list
# ---------------------------------------------------------------------------


@click.command(name="paragraphs")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape whose paragraphs to list (shape:/ph:)."
)
@_deck_command
def paragraphs_cmd(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """List a shape's paragraphs: anchor_id (para:S:N:P), text, indent, bullet."""
    sh = _resolve_shape(deck, anchor_id)
    rows = sh.paragraphs.list()
    emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_paragraphs(rows))


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
@_deck_command
def insert(ctx: click.Context, deck: Presentation, anchor_id: str, text: str, after: bool) -> None:
    """Insert a new paragraph before/after a text anchor (para:/ph:/shape:/notes:)."""
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


@click.command(name="format-paragraph")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to format.")
@click.option("--alignment", type=click.Choice(ALIGNMENT_CHOICES), default=None, help="Alignment.")
@click.option("--space-before", type=float, default=None, help="Space before (points).")
@click.option("--space-after", type=float, default=None, help="Space after (points).")
@click.option(
    "--space-before-lines", type=float, default=None, help="Space before (multiple, e.g. 0.5)."
)
@click.option(
    "--space-after-lines", type=float, default=None, help="Space after (multiple, e.g. 0.5)."
)
@click.option("--line-spacing", type=float, default=None, help="Line spacing (multiple, e.g. 1.5).")
@click.option(
    "--line-spacing-points", type=float, default=None, help="Line spacing (exact points, e.g. 24)."
)
@click.option(
    "--indent-level", type=click.IntRange(1, 5), default=None, help="Outline/bullet level (1-5)."
)
@click.option(
    "--force", is_flag=True, default=False, help="Allow a line-spacing multiple above 5×."
)
@_deck_command
def format_paragraph(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    alignment: str | None,
    space_before: float | None,
    space_after: float | None,
    space_before_lines: float | None,
    space_after_lines: float | None,
    line_spacing: float | None,
    line_spacing_points: float | None,
    indent_level: int | None,
    force: bool,
) -> None:
    """Set paragraph formatting (alignment, spacing, indent level) on a text anchor.

    Line spacing is unit-explicit: `--line-spacing` is a *multiple* (1.5), while
    `--line-spacing-points` is *exact points* (24). Likewise `--space-before/after`
    are points and `--space-before/after-lines` are multiples.
    """
    opts = (
        alignment,
        space_before,
        space_after,
        space_before_lines,
        space_after_lines,
        line_spacing,
        line_spacing_points,
        indent_level,
    )
    if all(v is None for v in opts):
        raise click.UsageError("format-paragraph requires at least one formatting option")
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: format paragraph {anchor_id}"):
        anchor.format_paragraph(
            alignment=alignment,
            space_before=space_before,
            space_after=space_after,
            space_before_lines=space_before_lines,
            space_after_lines=space_after_lines,
            line_spacing=line_spacing,
            line_spacing_points=line_spacing_points,
            indent_level=indent_level,
            force=force,
        )
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id},
        as_text=not ctx.obj["as_json"],
        text=f"formatted {anchor.anchor_id}",
    )


@click.command(name="format-text")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to format.")
@click.option("--bold/--no-bold", "bold", default=None, help="Bold on/off.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italic on/off.")
@click.option("--underline/--no-underline", "underline", default=None, help="Underline on/off.")
@click.option("--size", type=float, default=None, help="Font size (points).")
@click.option("--font", "font", default=None, help="Font name (e.g. 'Arial').")
@click.option("--color", "color", default=None, help="Font color, '#RRGGBB'.")
@_deck_command
def format_text(
    ctx: click.Context,
    deck: Presentation,
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
    if all(v is None for v in (bold, italic, underline, size, font, color)):
        raise click.UsageError("format-text requires at least one formatting option")
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


@click.command(name="reset-format")
@click.option("--anchor-id", "anchor_id", required=True, help="Text anchor to reset.")
@_deck_command
def reset_format(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Reset an anchor's paragraph spacing to clean single-spaced defaults.

    The recovery for a line-spacing spiral (giant spacing, text off the slide).
    Does not reset font size/typeface — use `shape reset-to-layout` for a
    placeholder's geometry + default font, or `format-text` to set a font.
    """
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: reset format {anchor_id}"):
        anchor.reset_format()
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id},
        as_text=not ctx.obj["as_json"],
        text=f"reset paragraph spacing of {anchor.anchor_id}",
    )


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
@_deck_command
def list_apply(
    ctx: click.Context, deck: Presentation, anchor_id: str, list_type: str, char: str | None
) -> None:
    """Turn a text anchor's paragraphs into a bulleted or numbered list."""
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: apply {list_type} list to {anchor_id}"):
        anchor.apply_list(list_type, character=char)
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id, "type": list_type},
        as_text=not ctx.obj["as_json"],
        text=f"applied {list_type} list to {anchor.anchor_id}",
    )


@list_cmd.command(name="remove")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor to strip list formatting."
)
@_deck_command
def list_remove(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Strip bullets / numbering from a text anchor's paragraphs."""
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: remove list from {anchor_id}"):
        anchor.remove_list()
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id},
        as_text=not ctx.obj["as_json"],
        text=f"removed list from {anchor.anchor_id}",
    )


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
@_deck_command
def table_read(ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int) -> None:
    """Read a table as a grid of cells, each carrying its cell:S:N:R:C anchor."""
    grid = _resolve_table(deck, slide_index, shape_index).read()
    emit(grid, as_text=not ctx.obj["as_json"], text=_fmt_table_read(grid))


@table.command(name="add-row")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--values", "values", default=None, help="Optional JSON array of cell values for the new row."
)
@_deck_command
def table_add_row(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, values: str | None
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
    t = _resolve_table(deck, slide_index, shape_index)
    with deck.edit(f"CLI: add row to table shape:{slide_index}:{shape_index}"):
        t.add_row(parsed)
    emit(
        {"ok": True, "anchor_id": t.shape.anchor_id, "rows": t.row_count},
        as_text=not ctx.obj["as_json"],
        text=f"added row to {t.shape.anchor_id} (now {t.row_count} rows)",
    )


@table.command(name="delete-row")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option("--row", "row", type=int, required=True, help="1-based row to delete.")
@_deck_command
def table_delete_row(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, row: int
) -> None:
    """Delete a row from a table (one Ctrl-Z)."""
    t = _resolve_table(deck, slide_index, shape_index)
    with deck.edit(f"CLI: delete row {row} from table shape:{slide_index}:{shape_index}"):
        t.delete_row(row)
    emit(
        {"ok": True, "anchor_id": t.shape.anchor_id, "rows": t.row_count},
        as_text=not ctx.obj["as_json"],
        text=f"deleted row {row} from {t.shape.anchor_id} (now {t.row_count} rows)",
    )


@table.command(name="add-column")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--values", "values", default=None, help="Optional JSON array of cell values (top-to-bottom)."
)
@click.option(
    "--before",
    "before",
    type=int,
    default=None,
    help="1-based column to insert before; omit to append at the right edge.",
)
@_deck_command
def table_add_column(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    shape_index: int,
    values: str | None,
    before: int | None,
) -> None:
    """Add a column to a table — appended, or inserted with --before (one Ctrl-Z)."""
    parsed: list[Any] | None = None
    if values is not None:
        try:
            parsed = json.loads(values)
        except json.JSONDecodeError as e:
            raise click.UsageError(f"--values must be a JSON array: {e}") from e
        if not isinstance(parsed, list):
            raise click.UsageError("--values must be a JSON array")
    t = _resolve_table(deck, slide_index, shape_index)
    with deck.edit(f"CLI: add column to table shape:{slide_index}:{shape_index}"):
        t.add_column(parsed, before=before)
    emit(
        {"ok": True, "anchor_id": t.shape.anchor_id, "columns": t.column_count},
        as_text=not ctx.obj["as_json"],
        text=f"added column to {t.shape.anchor_id} (now {t.column_count} columns)",
    )


@table.command(name="delete-column")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option("--column", "column", type=int, required=True, help="1-based column to delete.")
@_deck_command
def table_delete_column(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, column: int
) -> None:
    """Delete a column from a table (one Ctrl-Z)."""
    t = _resolve_table(deck, slide_index, shape_index)
    with deck.edit(f"CLI: delete column {column} from table shape:{slide_index}:{shape_index}"):
        t.delete_column(column)
    emit(
        {"ok": True, "anchor_id": t.shape.anchor_id, "columns": t.column_count},
        as_text=not ctx.obj["as_json"],
        text=f"deleted column {column} from {t.shape.anchor_id} (now {t.column_count} columns)",
    )


def _parse_axis(value: str | None) -> int | list[int] | None:
    """`--rows`/`--cols`: None (all) | "2" (one) | "1,3" (several) -> a selector."""
    if value is None:
        return None
    try:
        parts = [int(p) for p in str(value).split(",") if p.strip() != ""]
    except ValueError as e:
        raise click.UsageError(f"expected comma-separated integers, got {value!r}") from e
    if not parts:
        raise click.UsageError("selector resolved to no indices")
    return parts[0] if len(parts) == 1 else parts


@table.command(name="set-fill")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--fill", "fill", required=True, help='Color (#RRGGBB / r,g,b / name) or "none" (transparent).'
)
@click.option("--rows", "rows", default=None, help="Row(s): omit=all, '1', or '1,3'.")
@click.option("--cols", "cols", default=None, help="Column(s): omit=all, '2', or '1,3'.")
@click.option(
    "--transparency", "transparency", type=float, default=None, help="Alpha 0.0-1.0 (0 opaque)."
)
@_deck_command
def table_set_fill(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    shape_index: int,
    fill: str,
    rows: str | None,
    cols: str | None,
    transparency: float | None,
) -> None:
    """Shade a region of cells (or clear with --fill none); one Ctrl-Z."""
    t = _resolve_table(deck, slide_index, shape_index)
    with deck.edit(f"CLI: fill cells of table shape:{slide_index}:{shape_index}"):
        n = t.set_fill(
            fill, rows=_parse_axis(rows), cols=_parse_axis(cols), transparency=transparency
        )
    emit(
        {"ok": True, "anchor_id": t.shape.anchor_id, "cells": n},
        as_text=not ctx.obj["as_json"],
        text=f"filled {n} cell(s) of {t.shape.anchor_id} with {fill}",
    )


@table.command(name="set-border")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--color", "color", default=None, help='Edge color (#RRGGBB / name) or "none" (hide).'
)
@click.option("--weight", "weight", type=float, default=None, help="Edge weight in points.")
@click.option("--dash", "dash", default=None, help="Dash style (solid/dash/round_dot/...).")
@click.option(
    "--edges",
    "edges",
    default="all",
    help="Edges: all | top,bottom,left,right,diagonal_down,diagonal_up (comma-separated).",
)
@click.option("--rows", "rows", default=None, help="Row(s): omit=all, '1', or '1,3'.")
@click.option("--cols", "cols", default=None, help="Column(s): omit=all, '2', or '1,3'.")
@click.option("--visible/--hidden", "visible", default=None, help="Force the edge on/off.")
@_deck_command
def table_set_border(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    shape_index: int,
    color: str | None,
    weight: float | None,
    dash: str | None,
    edges: str,
    rows: str | None,
    cols: str | None,
    visible: bool | None,
) -> None:
    """Style cell border(s) across a region; one Ctrl-Z."""
    edge_sel: str | list[str] = (
        [e.strip() for e in edges.split(",") if e.strip()] if "," in edges else edges
    )
    t = _resolve_table(deck, slide_index, shape_index)
    with deck.edit(f"CLI: border cells of table shape:{slide_index}:{shape_index}"):
        n = t.set_border(
            color=color,
            weight=weight,
            dash=dash,
            edges=edge_sel,
            rows=_parse_axis(rows),
            cols=_parse_axis(cols),
            visible=visible,
        )
    emit(
        {"ok": True, "anchor_id": t.shape.anchor_id, "cells": n},
        as_text=not ctx.obj["as_json"],
        text=f"styled border on {n} cell(s) of {t.shape.anchor_id}",
    )


# ---------------------------------------------------------------------------
# chart read | set-type | set-data  (a chart is a shape; data is embedded Excel)
# ---------------------------------------------------------------------------


@click.group(name="chart")
def chart() -> None:
    """Read + edit charts: read, set-type, set-data, recolor-text (a chart is a shape)."""


def _resolve_chart(deck: Presentation, slide_index: int, shape_index: int) -> Any:
    """Resolve the chart on slide S, shape N (z-order). Exit 2 if no such chart."""
    return deck.slides[slide_index].shapes[shape_index].chart


@chart.command(name="read")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@_deck_command
def chart_read(ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int) -> None:
    """Read a chart: type, categories, and series (name + values)."""
    info = _resolve_chart(deck, slide_index, shape_index).read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_chart_read(info))


@chart.command(name="set-type")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--chart-type",
    "chart_type",
    type=click.Choice(CHART_TYPE_CHOICES),
    required=True,
    help="New chart kind.",
)
@_deck_command
def chart_set_type(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, chart_type: str
) -> None:
    """Change a chart's kind (one Ctrl-Z)."""
    ch = _resolve_chart(deck, slide_index, shape_index)
    with deck.edit(f"CLI: set chart type shape:{slide_index}:{shape_index}"):
        ch.set_type(chart_type)
    emit(
        {"ok": True, "anchor_id": ch.shape.anchor_id, "chart_type": ch.chart_type},
        as_text=not ctx.obj["as_json"],
        text=f"set {ch.shape.anchor_id} to {ch.chart_type}",
    )


@chart.command(name="set-data")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--categories",
    "categories",
    required=True,
    help="Category labels: a JSON array or comma-separated list.",
)
@click.option(
    "--series",
    "series",
    required=True,
    help='Series as a JSON object {"name":[values]} or array of [name,[values]].',
)
@_deck_command
def chart_set_data(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    shape_index: int,
    categories: str,
    series: str,
) -> None:
    """Replace a chart's data (categories × series; one Ctrl-Z)."""
    cats = _parse_categories(categories)
    ser = _parse_series(series)
    ch = _resolve_chart(deck, slide_index, shape_index)
    with deck.edit(f"CLI: set chart data shape:{slide_index}:{shape_index}"):
        ch.set_data(cats or [], ser)
    info = ch.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_chart_read(info))


@chart.command(name="recolor-text")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--color",
    "color",
    required=True,
    help='Text color as "#RRGGBB" (or "RRGGBB"). Recolors every shown chart text '
    "element: legend, axis tick labels, title, data labels.",
)
@_deck_command
def chart_recolor_text(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, color: str
) -> None:
    """Recolor ALL of a chart's text (legend/axes/title/data labels; one Ctrl-Z).

    The coarse fix for a chart whose inherited (black) axis/legend text is
    invisible on a custom background — no rebuild from primitives needed.
    """
    ch = _resolve_chart(deck, slide_index, shape_index)
    with deck.edit(f"CLI: recolor chart text shape:{slide_index}:{shape_index}"):
        info = ch.recolor_text(color)
    emit(
        info,
        as_text=not ctx.obj["as_json"],
        text=f"recolored {info['anchor_id']} text -> {info['color']} "
        f"({', '.join(info['recolored'])})",
    )


# ---------------------------------------------------------------------------
# smartart read | set-nodes  (a SmartArt diagram is a shape; content is a tree)
# ---------------------------------------------------------------------------


@click.group(name="smartart")
def smartart() -> None:
    """Read + edit SmartArt: read, set-nodes, recolor-text (a diagram is a shape)."""


def _resolve_smartart(deck: Presentation, slide_index: int, shape_index: int) -> Any:
    """Resolve the SmartArt on slide S, shape N (z-order). Exit 2 if none there."""
    return deck.slides[slide_index].shapes[shape_index].smartart


@smartart.command(name="read")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@_deck_command
def smartart_read(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int
) -> None:
    """Read a SmartArt diagram: layout + the nested node tree (text + level)."""
    info = _resolve_smartart(deck, slide_index, shape_index).read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_smartart_read(info))


@smartart.command(name="set-nodes")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--nodes",
    "nodes",
    required=True,
    help="A JSON array of strings and/or {text, children} objects (flat or nested).",
)
@_deck_command
def smartart_set_nodes(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, nodes: str
) -> None:
    """Replace a SmartArt diagram's nodes (flat list or nested tree; one Ctrl-Z)."""
    parsed = _parse_nodes(nodes)
    sa = _resolve_smartart(deck, slide_index, shape_index)
    with deck.edit(f"CLI: set smartart nodes shape:{slide_index}:{shape_index}"):
        sa.set_nodes(parsed or [])
    info = sa.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_smartart_read(info))


@smartart.command(name="recolor-text")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--color",
    "color",
    required=True,
    help='Text color as "#RRGGBB" (or "RRGGBB"). Recolors every node label.',
)
@_deck_command
def smartart_recolor_text(
    ctx: click.Context, deck: Presentation, slide_index: int, shape_index: int, color: str
) -> None:
    """Recolor ALL of a SmartArt diagram's node text (one Ctrl-Z).

    The coarse fix for a diagram whose inherited (black) node labels are
    invisible on a custom background — no rebuild from primitives needed.
    """
    sa = _resolve_smartart(deck, slide_index, shape_index)
    with deck.edit(f"CLI: recolor smartart text shape:{slide_index}:{shape_index}"):
        info = sa.recolor_text(color)
    emit(
        info,
        as_text=not ctx.obj["as_json"],
        text=f"recolored {info['anchor_id']} text -> {info['color']} "
        f"({info['nodes_recolored']} nodes)",
    )


@smartart.command(name="format-node")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--shape", "shape_index", type=int, required=True, help="1-based shape z-order index."
)
@click.option(
    "--node",
    "node_index",
    type=int,
    required=True,
    help="1-based node_index from `smartart read` (its AllNodes position).",
)
@click.option("--bold/--no-bold", "bold", default=None, help="Bold the node text.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italicize the node text.")
@click.option("--underline/--no-underline", "underline", default=None, help="Underline the text.")
@click.option("--size", "size", type=float, default=None, help="Font size in points.")
@click.option("--font", "font", default=None, help="Font family name.")
@click.option("--color", "color", default=None, help='Text color "#RRGGBB" / "r,g,b".')
@_deck_command
def smartart_format_node(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    shape_index: int,
    node_index: int,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    size: float | None,
    font: str | None,
    color: str | None,
) -> None:
    """Format ONE SmartArt node's text — bold/italic/underline/size/font/color (one Ctrl-Z).

    The per-node companion to `recolor-text`: address a single node by the
    `node_index` that `smartart read` prints for it.
    """
    sa = _resolve_smartart(deck, slide_index, shape_index)
    with deck.edit(f"CLI: format smartart node {node_index} shape:{slide_index}:{shape_index}"):
        info = sa.format_node(
            node_index,
            bold=bold,
            italic=italic,
            underline=underline,
            size=size,
            font=font,
            color=color,
        )
    emit(
        info,
        as_text=not ctx.obj["as_json"],
        text=f"formatted node {node_index} of {info['anchor_id']} ({info['text']!r})",
    )


# ---------------------------------------------------------------------------
# comment — review comments (list / add / reply / delete); threaded, per-slide
# ---------------------------------------------------------------------------


@click.group(name="comment")
def comment() -> None:
    """Read + write review comments: list, add, reply, delete (threaded, per-slide).

    Comments attach to a slide at an (x, y) point and are addressed by
    `--slide S --index N` (1-based, see `comment list`). Adding binds to the
    signed-in Office account; there is no resolve verb (not COM-readable).
    """


@comment.command(name="list")
@click.option(
    "--slide",
    "slide_index",
    type=int,
    default=None,
    help="1-based slide index. Omit for a deck-wide roll-up of every comment.",
)
@_deck_command
def comment_list(ctx: click.Context, deck: Presentation, slide_index: int | None) -> None:
    """List comments on a slide (`--slide S`) or across the whole deck."""
    payload: Any = (
        deck.slides[slide_index].comments.list() if slide_index is not None else deck.comments()
    )
    emit(payload, as_text=not ctx.obj["as_json"], text=_fmt_comment_list(payload))


@comment.command(name="add")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option("--text", "text", required=True, help="The comment body.")
@click.option("--left", "left", type=float, default=None, help="Anchor x, in points (default 12).")
@click.option("--top", "top", type=float, default=None, help="Anchor y, in points (default 12).")
@click.option(
    "--author",
    "author",
    default=None,
    help="Author name (best-effort; modern Office binds to the signed-in account).",
)
@click.option(
    "--initials",
    "initials",
    default=None,
    help="Author initials (best-effort; modern Office binds to the signed-in account).",
)
@_deck_command
def comment_add(
    ctx: click.Context,
    deck: Presentation,
    slide_index: int,
    text: str,
    left: float | None,
    top: float | None,
    author: str | None,
    initials: str | None,
) -> None:
    """Add a comment to slide S (one Ctrl-Z). Binds to the signed-in account."""
    kwargs: dict[str, Any] = {"author": author, "initials": initials}
    if left is not None:
        kwargs["left"] = left
    if top is not None:
        kwargs["top"] = top
    with deck.edit(f"CLI: add comment on slide {slide_index}"):
        c = deck.slides[slide_index].comments.add(text, **kwargs)
        info = c.to_dict()
    info = {"ok": True, "slide": slide_index, "comment": info}
    emit(
        info,
        as_text=not ctx.obj["as_json"],
        text=f"added comment {info['comment']['index']} on slide {slide_index}",
    )


@comment.command(name="reply")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--index", "index", type=int, required=True, help="1-based comment index (see `comment list`)."
)
@click.option("--text", "text", required=True, help="The reply body.")
@_deck_command
def comment_reply(
    ctx: click.Context, deck: Presentation, slide_index: int, index: int, text: str
) -> None:
    """Reply to comment INDEX on slide S (threaded; one Ctrl-Z)."""
    with deck.edit(f"CLI: reply to comment {slide_index}:{index}"):
        rep = deck.slides[slide_index].comments[index].reply(text)
        info = rep.to_dict()
    emit(
        {"ok": True, "slide": slide_index, "parent": index, "reply": info},
        as_text=not ctx.obj["as_json"],
        text=f"replied to comment {index} on slide {slide_index}",
    )


@comment.command(name="delete")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@click.option(
    "--index", "index", type=int, required=True, help="1-based comment index (see `comment list`)."
)
@_deck_command
def comment_delete(ctx: click.Context, deck: Presentation, slide_index: int, index: int) -> None:
    """Delete comment INDEX on slide S (takes its replies; one Ctrl-Z)."""
    with deck.edit(f"CLI: delete comment {slide_index}:{index}"):
        deck.slides[slide_index].comments[index].delete()
    emit(
        {"ok": True, "slide": slide_index, "index": index},
        as_text=not ctx.obj["as_json"],
        text=f"deleted comment {index} on slide {slide_index}",
    )


# ---------------------------------------------------------------------------
# section — deck structure: named slide spans (list / add / rename / delete / move)
# ---------------------------------------------------------------------------


def _fmt_sections(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "(no sections)"
    return "\n".join(
        f"{r['index']}. {r['name']} (slides {r['first_slide']}..+{r['slide_count']})"
        if r["first_slide"] is not None
        else f"{r['index']}. {r['name']} (empty)"
        for r in rows
    )


@click.group(name="section")
def section() -> None:
    """Deck sections — named spans of slides: list, add, rename, delete, move.

    Addressed by 1-based section index (see `section list`). A section starts at a
    slide and runs to the next one; `add --before-slide N` starts a section there.
    """


@section.command(name="list")
@_deck_command
def section_list(ctx: click.Context, deck: Presentation) -> None:
    """List the deck's sections: index, name, first slide, slide count."""
    rows = deck.sections.list()
    emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_sections(rows))


@section.command(name="add")
@click.option("--name", "name", required=True, help="The new section's name.")
@click.option(
    "--before-slide",
    "before_slide",
    type=int,
    default=None,
    help="1-based slide the section starts at. Omit to append an empty trailing section.",
)
@_deck_command
def section_add(
    ctx: click.Context, deck: Presentation, name: str, before_slide: int | None
) -> None:
    """Add a section (one Ctrl-Z); print its resulting row."""
    with deck.edit(f"CLI: add section {name!r}"):
        row = deck.sections.add(name, before_slide=before_slide)
    emit(
        {"ok": True, "section": row},
        as_text=not ctx.obj["as_json"],
        text=f"added section {row['index']}: {row['name']}",
    )


@section.command(name="rename")
@click.option("--section", "index", type=int, required=True, help="1-based section index.")
@click.option("--name", "name", required=True, help="The new name.")
@_deck_command
def section_rename(ctx: click.Context, deck: Presentation, index: int, name: str) -> None:
    """Rename section INDEX (one Ctrl-Z)."""
    with deck.edit(f"CLI: rename section {index}"):
        row = deck.sections.rename(index, name)
    emit(
        {"ok": True, "section": row},
        as_text=not ctx.obj["as_json"],
        text=f"renamed section {index} -> {row['name']}",
    )


@section.command(name="delete")
@click.option("--section", "index", type=int, required=True, help="1-based section index.")
@click.option(
    "--delete-slides",
    "delete_slides",
    is_flag=True,
    default=False,
    help="Also delete the section's slides (default: keep them, drop only the boundary).",
)
@_deck_command
def section_delete(ctx: click.Context, deck: Presentation, index: int, delete_slides: bool) -> None:
    """Delete section INDEX (keeps its slides unless --delete-slides; one Ctrl-Z)."""
    with deck.edit(f"CLI: delete section {index}"):
        out = deck.sections.delete(index, delete_slides=delete_slides)
    emit(
        {"ok": True, **out},
        as_text=not ctx.obj["as_json"],
        text=f"deleted section {index}: {out['name']}"
        + (" (with slides)" if delete_slides else ""),
    )


@section.command(name="move")
@click.option("--section", "index", type=int, required=True, help="1-based section index.")
@click.option("--to", "to", type=int, required=True, help="1-based target position.")
@_deck_command
def section_move(ctx: click.Context, deck: Presentation, index: int, to: int) -> None:
    """Move section INDEX to position TO (carries its slides; one Ctrl-Z)."""
    with deck.edit(f"CLI: move section {index} -> {to}"):
        row = deck.sections.move(index, to)
    emit(
        {"ok": True, "section": row},
        as_text=not ctx.obj["as_json"],
        text=f"moved section to position {row['index']}: {row['name']}",
    )


# ---------------------------------------------------------------------------
# theme — deck-wide palette + typefaces (read / set-color / set-font)
# ---------------------------------------------------------------------------


@click.group(name="theme")
def theme() -> None:
    """Read + edit the deck theme: the 12-slot palette and heading/body fonts.

    Global, anti-polite ops — one change recolors/re-fonts every inheriting slide.
    """


@theme.command(name="read")
@_deck_command
def theme_read(ctx: click.Context, deck: Presentation) -> None:
    """Read the theme palette (12 slots) + the major/minor typefaces."""
    info = deck.theme.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_theme_read(info))


@theme.command(name="set-color")
@click.option(
    "--slot", type=click.Choice(THEME_COLOR_CHOICES), required=True, help="Palette slot to set."
)
@click.option("--color", required=True, help="Color, '#RRGGBB' (or an (r,g,b)/int via the API).")
@_deck_command
def theme_set_color(ctx: click.Context, deck: Presentation, slot: str, color: str) -> None:
    """Set one theme palette slot (e.g. accent1) — recolors the whole deck."""
    with deck.edit(f"CLI: set theme color {slot}"):
        deck.theme.set_color(slot, color)
    info = deck.theme.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_theme_read(info))


@theme.command(name="set-font")
@click.option(
    "--which",
    type=click.Choice(THEME_FONT_SLOTS),
    required=True,
    help="major (headings) or minor (body).",
)
@click.option("--name", "name", required=True, help="Font name (e.g. 'Georgia').")
@click.option(
    "--script",
    type=click.Choice(THEME_FONT_SCRIPT_CHOICES),
    default="latin",
    show_default=True,
    help="Which script sub-typeface to set.",
)
@_deck_command
def theme_set_font(
    ctx: click.Context, deck: Presentation, which: str, name: str, script: str
) -> None:
    """Set the major (headings) or minor (body) theme typeface."""
    with deck.edit(f"CLI: set theme {which} font"):
        deck.theme.set_font(which, name, script=script)
    info = deck.theme.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_theme_read(info))


# ---------------------------------------------------------------------------
# master — deck-wide text styles + background (read / format-* / set-background)
# ---------------------------------------------------------------------------


@click.group(name="master")
def master() -> None:
    """Read + edit master text styles (title/body/default) and the background.

    PowerPoint's nearest 'named style' analog, applied to the primary slide master.
    """


@master.command(name="read")
@_deck_command
def master_read(ctx: click.Context, deck: Presentation) -> None:
    """Read the three master text styles (5 levels each) + the background fill."""
    info = deck.master.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_master_read(info))


@master.command(name="format-text-style")
@click.option(
    "--style", type=click.Choice(TEXT_STYLE_CHOICES), required=True, help="Which text style."
)
@click.option(
    "--level", type=click.IntRange(1, 5), default=1, show_default=True, help="Outline level (1-5)."
)
@click.option("--bold/--no-bold", "bold", default=None, help="Bold on/off.")
@click.option("--italic/--no-italic", "italic", default=None, help="Italic on/off.")
@click.option("--underline/--no-underline", "underline", default=None, help="Underline on/off.")
@click.option("--size", type=float, default=None, help="Font size (points).")
@click.option("--font", "font", default=None, help="Font name (e.g. 'Georgia').")
@click.option("--color", "color", default=None, help="Font color, '#RRGGBB'.")
@_deck_command
def master_format_text_style(
    ctx: click.Context,
    deck: Presentation,
    style: str,
    level: int,
    bold: bool | None,
    italic: bool | None,
    underline: bool | None,
    size: float | None,
    font: str | None,
    color: str | None,
) -> None:
    """Set font formatting on a master text style + level (deck-wide)."""
    if all(v is None for v in (bold, italic, underline, size, font, color)):
        raise click.UsageError("format-text-style requires at least one formatting option")
    with deck.edit(f"CLI: format master {style} L{level}"):
        deck.master.format_text_style(
            style,
            level,
            bold=bold,
            italic=italic,
            underline=underline,
            size=size,
            font=font,
            color=color,
        )
    info = deck.master.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_master_read(info))


@master.command(name="format-paragraph-style")
@click.option(
    "--style", type=click.Choice(TEXT_STYLE_CHOICES), required=True, help="Which text style."
)
@click.option(
    "--level", type=click.IntRange(1, 5), default=1, show_default=True, help="Outline level (1-5)."
)
@click.option("--alignment", type=click.Choice(ALIGNMENT_CHOICES), default=None, help="Alignment.")
@click.option("--space-before", type=float, default=None, help="Space before (points).")
@click.option("--space-after", type=float, default=None, help="Space after (points).")
@click.option("--line-spacing", type=float, default=None, help="Line spacing (multiple, e.g. 1.5).")
@_deck_command
def master_format_paragraph_style(
    ctx: click.Context,
    deck: Presentation,
    style: str,
    level: int,
    alignment: str | None,
    space_before: float | None,
    space_after: float | None,
    line_spacing: float | None,
) -> None:
    """Set paragraph formatting on a master text style + level (deck-wide)."""
    if all(v is None for v in (alignment, space_before, space_after, line_spacing)):
        raise click.UsageError("format-paragraph-style requires at least one option")
    with deck.edit(f"CLI: format master paragraph {style} L{level}"):
        deck.master.format_paragraph_style(
            style,
            level,
            alignment=alignment,
            space_before=space_before,
            space_after=space_after,
            line_spacing=line_spacing,
        )
    info = deck.master.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_master_read(info))


@master.command(name="set-background")
@click.option("--color", required=True, help="Background color, '#RRGGBB' (solid fill).")
@_deck_command
def master_set_background(ctx: click.Context, deck: Presentation, color: str) -> None:
    """Set the master background to a solid color (deck-wide)."""
    with deck.edit("CLI: set master background"):
        deck.master.set_background(color)
    info = deck.master.read()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_master_read(info))


@master.command(name="headers-footers")
@_deck_command
def master_headers_footers(ctx: click.Context, deck: Presentation) -> None:
    """Read the deck-wide footer / slide-number / date defaults (master scope)."""
    hf = deck.master.headers_footers.read()
    emit(hf, as_text=not ctx.obj["as_json"], text=_fmt_hf(hf))


@master.command(name="set-footer")
@click.option(
    "--text", "text", default=None, help="Footer text (setting it auto-shows the footer)."
)
@click.option("--show/--hide", "visible", default=None, help="Show or hide the footer.")
@_deck_command
def master_set_footer(
    ctx: click.Context, deck: Presentation, text: str | None, visible: bool | None
) -> None:
    """Set the deck-wide default footer (master)."""
    if text is None and visible is None:
        raise click.UsageError("master set-footer requires --text and/or --show/--hide")
    with deck.edit("CLI: set master footer"):
        hf = deck.master.headers_footers.set_footer(text=text, visible=visible)
    emit(
        {"ok": True, "scope": "master", "headers_footers": hf},
        as_text=not ctx.obj["as_json"],
        text=_fmt_hf(hf),
    )


@master.command(name="slide-number")
@click.option("--show/--hide", "visible", default=True, help="Show or hide the slide number.")
@_deck_command
def master_slide_number(ctx: click.Context, deck: Presentation, visible: bool) -> None:
    """Show or hide the deck-wide auto slide number (master)."""
    with deck.edit("CLI: master slide number"):
        hf = deck.master.headers_footers.set_slide_number(visible)
    emit(
        {"ok": True, "scope": "master", "headers_footers": hf},
        as_text=not ctx.obj["as_json"],
        text=_fmt_hf(hf),
    )


@master.command(name="set-date")
@click.option("--text", "text", default=None, help="Fixed date text (auto-shows the date).")
@click.option(
    "--format",
    "fmt",
    type=int,
    default=None,
    help="Auto-updating date format (PpDateTimeFormat int).",
)
@click.option("--show/--hide", "visible", default=None, help="Show or hide the date.")
@_deck_command
def master_set_date(
    ctx: click.Context, deck: Presentation, text: str | None, fmt: int | None, visible: bool | None
) -> None:
    """Set the deck-wide default date placeholder (master)."""
    if text is None and fmt is None and visible is None:
        raise click.UsageError("master set-date requires --text, --format, and/or --show/--hide")
    with deck.edit("CLI: set master date"):
        hf = deck.master.headers_footers.set_date(visible=visible, text=text, fmt=fmt)
    emit(
        {"ok": True, "scope": "master", "headers_footers": hf},
        as_text=not ctx.obj["as_json"],
        text=_fmt_hf(hf),
    )


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
@_deck_command
def read_anchor(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Read the text of any text anchor."""
    anchor = deck.anchor_by_id(anchor_id)
    text = anchor.text
    emit(
        {"anchor_id": anchor.anchor_id, "kind": anchor.kind, "text": text},
        as_text=not ctx.obj["as_json"],
        text=text,
    )


@read.command(name="text-frame-status")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Shape to inspect (shape:/ph:/shapeid:)."
)
@_deck_command
def read_text_frame_status(ctx: click.Context, deck: Presentation, anchor_id: str) -> None:
    """Autofit / wrap / margin diagnostics for a shape's text frame (a read)."""
    sh = _resolve_shape(deck, anchor_id)
    status = sh.text_frame_status()
    emit(
        {"anchor_id": sh.anchor_id, **status.to_dict()},
        as_text=not ctx.obj["as_json"],
        text=f"{sh.anchor_id}: autosize={status.autosize} overflow_risk={status.overflow_risk}",
    )


@read.command(name="notes")
@click.option("--slide", "slide_index", type=int, required=True, help="1-based slide index.")
@_deck_command
def read_notes(ctx: click.Context, deck: Presentation, slide_index: int) -> None:
    """Read the speaker notes of slide INDEX."""
    notes = deck.slides[slide_index].notes
    text = notes.text
    emit(
        {"anchor_id": notes.anchor_id, "kind": notes.kind, "text": text},
        as_text=not ctx.obj["as_json"],
        text=text,
    )


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
@click.option(
    "--text",
    "text",
    required=True,
    help="New text (embed \\n or \\r for paragraphs; \\v for a soft line break).",
)
@click.pass_context
def write(ctx: click.Context, anchor_id: str, text: str) -> None:
    """Set the text of a text anchor (preserves the viewed slide; one Ctrl-Z)."""
    _set_text(ctx, anchor_id, text, f"CLI: write {anchor_id}")


@click.command(name="set-paragraphs")
@click.option(
    "--anchor-id", "anchor_id", required=True, help="Text anchor to rewrite (ph:/shape:/notes:)."
)
@click.option(
    "--paragraphs",
    "--json",  # back-compat alias; --paragraphs is preferred (avoids shadowing the global --json)
    "paragraphs_json",
    default=None,
    help='Paragraphs as a JSON array: strings or {"text", "list_type", ...} objects.',
)
@click.option(
    "--file",
    "json_file",
    type=click.Path(exists=True, dir_okay=False),
    default=None,
    help="Read the JSON paragraphs array from a file instead of --paragraphs.",
)
@_deck_command
def set_paragraphs(
    ctx: click.Context,
    deck: Presentation,
    anchor_id: str,
    paragraphs_json: str | None,
    json_file: str | None,
) -> None:
    """Rewrite an anchor as a clean list of paragraphs (no newline inference).

    Each array item is a string or an object with `text` plus optional per-paragraph
    formatting (`list_type`, `indent_level`, `alignment`, `line_spacing` /
    `line_spacing_points`, `size`, `bold`, ...). Each item becomes one addressable
    `para:`.
    """
    if (paragraphs_json is None) == (json_file is None):
        raise click.UsageError("set-paragraphs needs exactly one of --json or --file")
    if paragraphs_json is not None:
        raw = paragraphs_json
    else:
        assert json_file is not None  # narrowed by the xor guard above
        raw = Path(json_file).read_text("utf-8")
    try:
        items = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise click.UsageError(f"invalid JSON for paragraphs: {exc}") from exc
    if not isinstance(items, list):
        raise click.UsageError("paragraphs must be a JSON array")
    anchor = deck.anchor_by_id(anchor_id)
    with deck.edit(f"CLI: set paragraphs {anchor_id}"):
        new_ids = anchor.set_paragraphs(items)
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id, "paragraphs": new_ids},
        as_text=not ctx.obj["as_json"],
        text=f"set {len(new_ids)} paragraphs on {anchor.anchor_id}",
    )


# ---------------------------------------------------------------------------
# exec --script ops.json   (apply a batch script as one Ctrl-Z)
# ---------------------------------------------------------------------------

#: A batch op's error category -> CLI exit code (mirrors main._exit_for, keyed by
#: the string token run_batch reports instead of an exception type).
# The batch-result `error` token -> exit int mapping is the same taxonomy as
# every other failure: reuse the single source of truth in exceptions.
_BATCH_EXIT_FOR = EXIT_CODE_FOR


@click.command(name="exec")
@click.option(
    "--script",
    "script_path",
    type=click.Path(exists=True, dir_okay=False),
    required=True,
    help='JSON batch script: {"label": "...", "ops": [{"op": ..., ...}, ...]}.',
)
@click.option(
    "--continue",
    "continue_on_error",
    is_flag=True,
    default=False,
    help="Keep going after a failing op (default: stop at the first).",
)
@click.option(
    "--no-atomic",
    "no_atomic",
    is_flag=True,
    default=False,
    help="Fence each op as its own undo entry (default: the whole script is one Ctrl-Z).",
)
@click.option(
    "--no-follow-view",
    "no_follow_view",
    is_flag=True,
    default=False,
    help="Snap the view back to the pre-batch slide (default: a script that adds a "
    "slide ends on the last slide it built). Also settable via PPTLIVE_VIEW_FOLLOW=0.",
)
@click.pass_context
def exec_script(
    ctx: click.Context,
    script_path: str,
    continue_on_error: bool,
    no_atomic: bool,
    no_follow_view: bool,
) -> None:
    """Apply a JSON batch script of ops against one connection (the CLI batch verb).

    The script is `{"label": "...", "ops": [{"op": "set_text", "anchor_id": ...,
    "text": ...}, ...]}`. Each op defaults to the `edit` tool (set `"tool"` for a
    read/render/show op); ops run in order and, by default, collapse to a single
    Ctrl-Z. On a failing op the run stops (unless `--continue`) and the exit code
    maps to that op's error category — `shape:S:N` refs resolve live as each op
    runs, so address anything you didn't just create by `ph:S:KIND` or `.Name`.
    """

    def go() -> None:
        raw = Path(script_path).read_text("utf-8")
        try:
            script = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise click.UsageError(f"invalid JSON in script: {exc}") from exc
        if not isinstance(script, dict) or not isinstance(script.get("ops"), list):
            raise click.UsageError('script must be a JSON object with an "ops" array')
        label = str(script.get("label") or "exec script")
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            results = run_batch(
                ppt,
                deck,
                script["ops"],
                doc=ctx.obj["doc_name"],
                atomic=not no_atomic,
                stop_on_error=not continue_on_error,
                follow_view=False if no_follow_view else None,
                label=f"CLI exec: {label}",
            )
        ok = all(r["ok"] for r in results)
        summary = {
            "ok": ok,
            "label": label,
            "atomic": not no_atomic,
            "count": len(results),
            "results": results,
        }
        emit(
            summary,
            as_text=not ctx.obj["as_json"],
            text=f"exec {label!r}: {len(results)} ops, ok={ok}",
        )
        if not ok:
            first_fail = next(r for r in results if not r["ok"])
            sys.exit(_BATCH_EXIT_FOR.get(first_fail.get("error", "error"), EXIT_OTHER))

    _run(ctx, go)


# ---------------------------------------------------------------------------
# find --text "..." [--in SCOPE]
# ---------------------------------------------------------------------------


@click.command(name="find")
@click.option(
    "--text", "text", required=True, help="Text to locate (smart-quote / whitespace fuzzy)."
)
@click.option(
    "--in",
    "in_",
    default=None,
    help="Scope: slide:S, shape:S:N, ph:S:KIND, cell:S:N:R:C, or notes:S (default: whole deck).",
)
@_deck_command
def find(ctx: click.Context, deck: Presentation, text: str, in_: str | None) -> None:
    """Locate every fuzzy occurrence of TEXT (read-only; preserves the view).

    Emits a JSON array of `{anchor_id, start, length, text, context}` hits in
    document order — an empty array (exit 0) when nothing matches.
    """
    matches = deck.find(text, scope=in_)
    emit(matches, as_text=not ctx.obj["as_json"], text=_fmt_find(matches))


# ---------------------------------------------------------------------------
# replace
#   --anchor-id ID --text "..."                                (anchor mode)
#   --find OLD --text NEW [--in SCOPE] [--all|--occurrence N]  (fuzzy mode)
# ---------------------------------------------------------------------------


@click.command(name="replace")
@click.option(
    "--anchor-id", "anchor_id", default=None, help="Replace the entire text at this anchor."
)
@click.option(
    "--find", "find_text", default=None, help="Fuzzy text to locate (alternative to --anchor-id)."
)
@click.option(
    "--text",
    "text",
    required=True,
    help="Replacement text (embed \\n or \\r for paragraphs; \\v for a soft line break).",
)
@click.option(
    "--in", "in_", default=None, help="In fuzzy mode, scope the search (slide:S / an anchor id)."
)
@click.option(
    "--all", "replace_all", is_flag=True, default=False, help="In fuzzy mode, replace every match."
)
@click.option(
    "--occurrence",
    "occurrence",
    type=int,
    default=None,
    help="In fuzzy mode, replace only the Nth match (1-based, document order).",
)
@click.pass_context
def replace(
    ctx: click.Context,
    anchor_id: str | None,
    find_text: str | None,
    text: str,
    in_: str | None,
    replace_all: bool,
    occurrence: int | None,
) -> None:
    """Replace text — either at an anchor (entire text) or via fuzzy find.

    `replace --anchor-id ID --text NEW` overwrites the whole anchor (same effect
    as `write`). `replace --find OLD --text NEW [--in SCOPE] [--all|--occurrence
    N]` fuzzy-locates OLD across the deck and rewrites just the matched spans.
    Either form preserves the viewed slide and collapses to one Ctrl-Z.
    """
    if (anchor_id is None) == (find_text is None):
        raise click.UsageError("provide exactly one of --anchor-id or --find")
    if anchor_id is not None and (in_ or replace_all or occurrence is not None):
        raise click.UsageError("--in / --all / --occurrence are only valid with --find")
    if replace_all and occurrence is not None:
        raise click.UsageError("--all and --occurrence are mutually exclusive")

    if anchor_id is not None:
        _set_text(ctx, anchor_id, text, f"CLI: replace {anchor_id}")
        return

    assert find_text is not None  # guaranteed by the validation above

    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            # An ambiguous match raises AmbiguousMatchError; let it propagate to the
            # _run boundary, which reports it on stderr + exit 5 like every other
            # failure (the CLI contract is: stdout JSON only on success).
            with deck.edit(f"CLI: find/replace {find_text!r}"):
                applied = deck.find_replace(
                    find_text,
                    text,
                    scope=in_,
                    all=replace_all,
                    occurrence=occurrence,
                )
            emit(
                {"ok": True, "count": len(applied), "replacements": applied},
                as_text=not ctx.obj["as_json"],
                text=_fmt_replace_summary(applied),
            )

    _run(ctx, go)


# ---------------------------------------------------------------------------
# selection
# ---------------------------------------------------------------------------


@click.command(name="selection")
@_deck_command
def selection_cmd(ctx: click.Context, deck: Presentation) -> None:
    """Report the user's current selection, resolved to anchors.

    A polite read (it doesn't change the selection): the selected shapes as
    `shape:S:N`, or a text caret as `para:S:N:P`, with the single targetable
    `anchor_id` that `--anchor-id here:` resolves to.
    """
    info = deck.selection().to_dict()
    emit(info, as_text=not ctx.obj["as_json"], text=_fmt_selection(info))


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
@_deck_command
def go_to(ctx: click.Context, deck: Presentation, anchor_id: str, select: bool) -> None:
    """Move the user's view to an anchor's slide (deliberate, opt-in view move)."""
    anchor = deck.anchor_by_id(anchor_id)
    deck.go_to(anchor, select=select)
    emit(
        {"ok": True, "anchor_id": anchor.anchor_id, "kind": anchor.kind},
        as_text=not ctx.obj["as_json"],
        text=f"moved view to {anchor.anchor_id}",
    )


# ---------------------------------------------------------------------------
# Agent self-bootstrapping: llm-help / install-skill / install-mcp
# (all offline — they never touch PowerPoint)
# ---------------------------------------------------------------------------


@click.command(name="llm-help")
@click.option(
    "--python",
    "python",
    is_flag=True,
    default=False,
    help="Print the Python-API guide instead of the CLI guide.",
)
def llm_help_cmd(python: bool) -> None:
    """Print the full pptlive agent guide (the bundled skill) to stdout.

    One-shot orientation for an LLM: the anchor model, every verb, and the
    exit-code taxonomy. `pptlive --help` points here. Defaults to the CLI guide;
    `--python` prints the Python-API guide instead. Output is raw Markdown — not
    JSON, and unaffected by `--json/--text` — so it reads cleanly straight into a
    model's context, exactly like `--help`. Offline: never touches PowerPoint.
    """
    kind = "python" if python else "cli"
    try:
        click.echo(skill_body(kind))
    except (FileNotFoundError, ModuleNotFoundError, OSError, ValueError) as e:
        raise click.ClickException(f"could not read the bundled skill: {e}") from e


@click.command(name="install-skill")
@click.option("--cli", "cli", is_flag=True, default=False, help="Install only the CLI skill.")
@click.option(
    "--python", "python", is_flag=True, default=False, help="Install only the Python-API skill."
)
@click.option(
    "--system",
    "system",
    is_flag=True,
    default=False,
    help="Install to ~/.agents/skills/ instead of the current project's ./.agents/skills/.",
)
@click.option(
    "--force", "force", is_flag=True, default=False, help="Overwrite an existing SKILL.md."
)
@click.pass_context
def install_skill_cmd(
    ctx: click.Context, cli: bool, python: bool, system: bool, force: bool
) -> None:
    """Install pptlive's agent skills (SKILL.md) for LLM coding tools.

    pptlive ships two skills — `pptlive-cli` (the command-line workflow) and
    `pptlive-python` (the `import pptlive as pl` API). By default both are
    written under `.agents/skills/<name>/SKILL.md`; pass `--cli` or `--python`
    for just one. They land under the current directory (default) or your home
    directory (`--system`). Offline — this doesn't touch PowerPoint.
    """
    if cli and not python:
        kinds = ["cli"]
    elif python and not cli:
        kinds = ["python"]
    else:
        kinds = ["cli", "python"]

    base = Path.home() if system else Path.cwd()
    scope = "system" if system else "local"
    dests = [(kind, base / ".agents" / "skills" / skill_name(kind) / "SKILL.md") for kind in kinds]

    # Check every target up front so we never half-write when --force is absent.
    clashes = [str(dest) for _, dest in dests if dest.exists()]
    if clashes and not force:
        raise click.ClickException(
            "already exists (pass --force to overwrite): " + ", ".join(clashes)
        )

    installed = []
    try:
        for kind, dest in dests:
            content = bundled_skill(kind)
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_text(content, encoding="utf-8")
            installed.append(
                {
                    "kind": kind,
                    "name": skill_name(kind),
                    "path": str(dest),
                    "bytes": len(content.encode("utf-8")),
                }
            )
    except (FileNotFoundError, ModuleNotFoundError, OSError, ValueError) as e:
        raise click.ClickException(f"could not install the skill: {e}") from e

    emit(
        {"ok": True, "scope": scope, "installed": installed},
        as_text=not ctx.obj["as_json"],
        text="installed:\n" + "\n".join(f"  {r['name']} → {r['path']}" for r in installed),
    )


def _claude_desktop_config_path() -> Path:
    """Where Claude Desktop keeps `claude_desktop_config.json` on this OS."""
    if sys.platform == "win32":
        appdata = os.environ.get("APPDATA")
        base = Path(appdata) if appdata else Path.home() / "AppData" / "Roaming"
        return base / "Claude" / "claude_desktop_config.json"
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "Claude"
            / "claude_desktop_config.json"
        )
    return Path.home() / ".config" / "Claude" / "claude_desktop_config.json"


def _mcp_server_entry(directory: str | None) -> dict[str, Any]:
    """The `mcpServers` entry that launches the pptlive stdio server.

    Default (repo-less) form runs the published package straight from PyPI with
    `uvx` — `pptlive-mcp` is a console script *inside* `pptlive`, so it needs
    `--from "pptlive[mcp]"` to tell uv which package provides it. With
    `--directory` (a local checkout) pptlive *is* the project, so a plain
    `uv run pptlive-mcp` resolves it without `--from`.
    """
    if directory:
        return {"command": "uv", "args": ["run", "--directory", directory, "pptlive-mcp"]}
    return {"command": "uvx", "args": ["--from", "pptlive[mcp]", "pptlive-mcp"]}


@click.command(name="install-mcp")
@click.option(
    "--client",
    type=click.Choice(["claude-desktop", "claude-code"]),
    default="claude-desktop",
    help="Which MCP client's config to write (default: claude-desktop).",
)
@click.option(
    "--name", "server_name", default="pptlive", help="Server key to register (default: pptlive)."
)
@click.option(
    "--directory",
    "directory",
    default=None,
    help="Register a local checkout via `uv run --directory DIR` (dev), instead of the default `uvx --from pptlive[mcp]`.",
)
@click.option(
    "--config",
    "config_path",
    default=None,
    type=click.Path(),
    help="Write to this config file instead of the client's default location.",
)
@click.option(
    "--print",
    "print_only",
    is_flag=True,
    default=False,
    help="Print the JSON server snippet to stdout instead of writing any file.",
)
@click.option(
    "--force", "force", is_flag=True, default=False, help="Overwrite an existing server entry."
)
@click.pass_context
def install_mcp_cmd(
    ctx: click.Context,
    client: str,
    server_name: str,
    directory: str | None,
    config_path: str | None,
    print_only: bool,
    force: bool,
) -> None:
    """Register the pptlive MCP server in an agent's config.

    Merges an `mcpServers.<name>` entry into Claude Desktop's
    `claude_desktop_config.json` (default) or a Claude Code `.mcp.json`
    (`--client claude-code`, project-local). The entry launches the stdio server
    with `uvx --from "pptlive[mcp]" pptlive-mcp` (no separate install needed), or
    `uv run --directory DIR pptlive-mcp` for a local checkout. Use `--print` to
    just emit the snippet for any client. Offline — never touches PowerPoint;
    restart the client to pick up the change.
    """
    entry = _mcp_server_entry(directory)

    if print_only:
        emit(
            {"ok": True, "server": server_name, "entry": entry, "mcpServers": {server_name: entry}},
            as_text=not ctx.obj["as_json"],
            text=json.dumps({"mcpServers": {server_name: entry}}, indent=2),
        )
        return

    if config_path is not None:
        target = Path(config_path)
    elif client == "claude-desktop":
        target = _claude_desktop_config_path()
    else:  # claude-code: portable, project-local server file
        target = Path.cwd() / ".mcp.json"

    cfg: dict[str, Any] = {}
    if target.exists():
        try:
            raw = target.read_text(encoding="utf-8").strip()
            cfg = json.loads(raw) if raw else {}
        except (OSError, json.JSONDecodeError) as e:
            raise click.ClickException(f"could not read existing config {target}: {e}") from e
        if not isinstance(cfg, dict):
            raise click.ClickException(f"existing config {target} is not a JSON object")

    servers = cfg.setdefault("mcpServers", {})
    if not isinstance(servers, dict):
        raise click.ClickException(f"'mcpServers' in {target} is not a JSON object")
    action = "updated" if server_name in servers else "created"
    if server_name in servers and not force:
        raise click.ClickException(
            f"server '{server_name}' is already in {target}; pass --force to overwrite"
        )
    servers[server_name] = entry

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(cfg, indent=2) + "\n", encoding="utf-8")
    except OSError as e:
        raise click.ClickException(f"could not write {target}: {e}") from e

    emit(
        {
            "ok": True,
            "client": client,
            "path": str(target),
            "server": server_name,
            "action": action,
            "entry": entry,
        },
        as_text=not ctx.obj["as_json"],
        text=f"{action} server '{server_name}' → {target}\n(restart {client} to load it)",
    )

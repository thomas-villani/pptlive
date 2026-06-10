"""Anchor types — semantic, text-bearing handles inside a presentation.

The PowerPoint anchor model is *hierarchical* (slide → shape → paragraph), not a
global character stream, so there is no deck-wide `range:` and offsets are only
meaningful within one shape's text frame (see spec.md §"The anchor model"). An
anchor targets a COM `TextRange`, never the live `Selection`: text is set through
`TextFrame.TextRange.Text` directly, so no edit needs to select anything.

This module holds the abstract `Anchor` base, the `Notes` anchor, and (v0.3) the
`Paragraph` anchor (`para:S:N:P`) over one paragraph of a shape's text frame.
`Shape` — which *is* an `Anchor` when it has a text frame — lives in `_shapes.py`
because it also carries geometry. `Cell` arrives in v0.4.

The text-structure verbs (`format_text`, `format_paragraph`, `apply_list`,
`remove_list`, `insert_paragraph_before/after`) live on the base `Anchor` and act
on `self._text_range()`, so they work on a whole shape's text *and* on a single
`Paragraph` — PowerPoint has no named paragraph styles (the Word `apply_style`
analog), so styling is direct font formatting via `format_text`.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import (
    MsoTriState,
    PpPlaceholderType,
    alignment_for,
    bullet_type_for,
    bullet_type_name,
    color_hex_or_none,
    is_true,
    parse_color,
    tristate_value,
)
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide


def _tristate(value: bool) -> int:
    """Python bool -> `MsoTriState` int (`msoTrue` / `msoFalse`)."""
    return int(MsoTriState.TRUE) if value else int(MsoTriState.FALSE)


#: The character PowerPoint renders as a *soft* line break (`<a:br>`) within one
#: paragraph — a vertical tab. Callers embed it in `set_text` text when they want
#: a line break that does NOT start a new, separately-addressable paragraph.
SOFT_BREAK = "\v"


def normalize_paragraph_breaks(text: str) -> str:
    """Make `\\n` / `\\r\\n` / `\\r` all real paragraph breaks; keep `\\v` soft.

    PowerPoint's COM `TextRange.Text` setter treats `\\r` as a paragraph break
    (`<a:p>`) but `\\n` as a *soft* line break (`<a:br>`) inside one paragraph —
    the opposite of what an LLM building a bullet list expects when it emits
    `"a\\nb\\nc"`. We normalise every newline spelling to `\\r` so each line
    becomes its own addressable `para:` paragraph, and leave `\\v` (`SOFT_BREAK`)
    untouched as the explicit within-paragraph soft break.
    """
    return text.replace("\r\n", "\r").replace("\n", "\r")


#: Per-paragraph keys `set_paragraphs` forwards to `format_paragraph`.
_PARA_FORMAT_KEYS = (
    "alignment",
    "space_before",
    "space_after",
    "space_before_lines",
    "space_after_lines",
    "line_spacing",
    "line_spacing_points",
    "indent_level",
    "force",
)
#: Per-paragraph keys `set_paragraphs` forwards to `format_text`.
_PARA_FONT_KEYS = ("bold", "italic", "underline", "size", "font", "color")


def _as_single_paragraph(text: str) -> str:
    """Collapse any newline inside one `set_paragraphs` item to a soft break.

    `set_paragraphs` is *paragraph-oriented*: each item is exactly one paragraph,
    so a newline embedded in an item must NOT split it into two (the whole point —
    no newline inference). We fold `\\r`/`\\n` to `\\v` (`SOFT_BREAK`), keeping the
    item a single addressable `para:` while preserving an intentional line break.
    """
    return str(text).replace("\r\n", SOFT_BREAK).replace("\n", SOFT_BREAK).replace("\r", SOFT_BREAK)


def _coerce_paragraph_item(item: Any) -> dict[str, Any]:
    """Normalise a `set_paragraphs` item (a `str` or a `{text, ...}` dict)."""
    if isinstance(item, str):
        return {"text": item}
    if isinstance(item, dict):
        if "text" not in item or not isinstance(item["text"], str):
            raise ValueError("each paragraph dict needs a string 'text'")
        return dict(item)
    raise ValueError(f"each paragraph must be a string or a dict, got {type(item).__name__}")


def _bullet_char(character: str | int) -> int:
    """A single-char string or an int code point -> the int `Bullet.Character`."""
    if isinstance(character, str):
        if len(character) != 1:
            raise ValueError(f"bullet character must be a single character, got {character!r}")
        return ord(character)
    if isinstance(character, bool) or not isinstance(character, int):
        raise TypeError("bullet character must be a single-char str or an int code point")
    return int(character)


def apply_font(
    f: Any,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    size: float | None = None,
    font: str | None = None,
    color: str | int | tuple[int, int, int] | None = None,
) -> None:
    """Write font properties onto a COM `Font` object — only the kwargs passed.

    Shared by `Anchor.format_text` (a text range's `.Font`) and the master text
    styles (`Master.format_text_style`, a `TextStyles(t).Levels(n).Font`), so both
    surfaces format fonts identically. `size` is points; `color` is `"#RRGGBB"`,
    an `(r, g, b)` tuple, or a raw RGB int. Caller wraps this in
    `translate_com_errors()`.
    """
    if bold is not None:
        f.Bold = _tristate(bold)
    if italic is not None:
        f.Italic = _tristate(italic)
    if underline is not None:
        f.Underline = _tristate(underline)
    if size is not None:
        f.Size = float(size)
    if font is not None:
        f.Name = str(font)
    if color is not None:
        f.Color.RGB = parse_color(color)


#: Above this, a `line_spacing` *multiple* is almost always a points-vs-multiple
#: confusion (the gpt-5.4 review's `line_spacing: 24` footgun — 24× line height
#: pushes text off the slide). `format_paragraph` rejects it unless `force=True`.
LINE_SPACING_MULTIPLE_MAX = 5.0


def apply_paragraph_format(
    pf: Any,
    *,
    alignment: int | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
    line_spacing_points: float | None = None,
    space_before_lines: float | None = None,
    space_after_lines: float | None = None,
) -> None:
    """Write paragraph properties onto a COM `ParagraphFormat` object.

    Shared by `Anchor.format_paragraph` and `Master.format_paragraph_style`.
    `alignment` is the resolved int (caller coerces a name first).

    PowerPoint stores each spacing value as a **bare number whose unit a companion
    `LineRule*` flag selects** (`msoTrue` ⇒ a multiple/lines, `msoFalse` ⇒ points;
    verified in `scripts/text_model_spike.py`). So each spacing knob comes as a
    points/multiple pair that *also* sets the matching flag, leaving no ambiguity:

    * `line_spacing` (multiple) / `line_spacing_points` (pt) → `SpaceWithin`
      + `LineRuleWithin`
    * `space_before` (pt) / `space_before_lines` (multiple) → `SpaceBefore`
      + `LineRuleBefore`
    * `space_after` (pt) / `space_after_lines` (multiple) → `SpaceAfter`
      + `LineRuleAfter`

    Pass at most one of each pair (the caller validates). Indent level is *not*
    handled here — it lives on the `TextRange`, not `ParagraphFormat`, so
    `Anchor.format_paragraph` sets it separately. Caller wraps this in
    `translate_com_errors()`.
    """
    if alignment is not None:
        pf.Alignment = alignment
    if line_spacing is not None:
        pf.LineRuleWithin = _tristate(True)
        pf.SpaceWithin = float(line_spacing)
    if line_spacing_points is not None:
        pf.LineRuleWithin = _tristate(False)
        pf.SpaceWithin = float(line_spacing_points)
    if space_before is not None:
        pf.LineRuleBefore = _tristate(False)
        pf.SpaceBefore = float(space_before)
    if space_before_lines is not None:
        pf.LineRuleBefore = _tristate(True)
        pf.SpaceBefore = float(space_before_lines)
    if space_after is not None:
        pf.LineRuleAfter = _tristate(False)
        pf.SpaceAfter = float(space_after)
    if space_after_lines is not None:
        pf.LineRuleAfter = _tristate(True)
        pf.SpaceAfter = float(space_after_lines)


class Anchor(ABC):
    """Abstract base for text-bearing handles.

    Concrete subclasses implement `_text_range()` (the COM `TextRange` to read
    and write) and `anchor_id`. `text` / `set_text` are derived and inherited.
    """

    kind: str = "anchor"

    @property
    def com(self) -> Any:
        """Raw COM object for this anchor — the `TextRange` it targets.

        `Shape` overrides this to return the raw `Shape` instead (the more useful
        escape hatch for a shape), exposing its text range via `text`/`set_text`.
        """
        return self._text_range()

    @abstractmethod
    def _text_range(self) -> Any:
        """Return the COM `TextRange` this anchor reads/writes. Must be overridden."""

    @property
    @abstractmethod
    def anchor_id(self) -> str:
        """Stable string identifier (e.g. `notes:3`, `shape:2:1`, `ph:2:body`)."""

    @property
    def name(self) -> str:
        """A display name for this anchor. Defaults to its `anchor_id`."""
        return self.anchor_id

    @property
    def text(self) -> str:
        """The anchor's plain text. PowerPoint separates paragraphs with `\\r`."""
        with _com.translate_com_errors():
            return str(self._text_range().Text or "")

    def set_text(self, text: str) -> None:
        """Replace the anchor's text in place.

        Embed `\\n` (or `\\r`) to start a new paragraph — each line becomes its
        own addressable `para:S:N:P`. For a *soft* line break that stays within
        one paragraph, embed `\\v` (`SOFT_BREAK`). Targets the text range
        directly, never the Selection, so it doesn't move the user's view. Wrap
        in `deck.edit(...)` to preserve the viewed slide and collapse the block
        to one Ctrl-Z (see `EditScope`).
        """
        with _com.translate_com_errors():
            self._text_range().Text = normalize_paragraph_breaks(text)

    def set_paragraphs(self, paragraphs: list[Any]) -> list[str]:
        """Replace this anchor's text with a clean, per-paragraph list.

        The safe alternative to newline inference for list authoring (the gpt-5.4
        review's ask): each item is a plain string **or** a dict
        `{"text": ..., "list_type"?, "indent_level"?, "alignment"?, "line_spacing"?,
        "size"?, ...}` and becomes exactly one addressable `para:` — a newline
        inside an item is folded to a soft break, never a paragraph split. Per-item
        keys are forwarded to `format_paragraph` (spacing/alignment/indent) and
        `format_text` (font); `list_type` applies/removes the bullet
        (`"none"` strips it), resetting list state cleanly. Returns the new
        paragraphs' `anchor_id`s (empty for a text anchor with no paragraph view,
        e.g. notes). Wrap in `deck.edit(...)`.
        """
        items = [_coerce_paragraph_item(p) for p in paragraphs]
        if not items:
            raise ValueError("set_paragraphs needs at least one paragraph")
        joined = "\r".join(_as_single_paragraph(it["text"]) for it in items)
        with _com.translate_com_errors():
            self._text_range().Text = joined
        para_coll = getattr(self, "paragraphs", None)
        if para_coll is None:
            return []  # a text anchor without a paragraph view (notes) — text only
        new_ids: list[str] = []
        for idx, item in enumerate(items, start=1):
            para = para_coll[idx]
            new_ids.append(para.anchor_id)
            self._apply_paragraph_item(para, item)
        return new_ids

    @staticmethod
    def _apply_paragraph_item(para: Anchor, item: dict[str, Any]) -> None:
        """Apply one `set_paragraphs` item's formatting to its built paragraph."""
        list_type = item.get("list_type")
        if list_type == "none":
            para.remove_list()
        elif list_type is not None:
            para.apply_list(list_type, character=item.get("bullet_char"))
        para_fmt = {k: item[k] for k in _PARA_FORMAT_KEYS if k in item}
        if para_fmt:
            para.format_paragraph(**para_fmt)
        font_fmt = {k: item[k] for k in _PARA_FONT_KEYS if k in item}
        if font_fmt:
            para.format_text(**font_fmt)

    # -- text structure (v0.3) -------------------------------------------------
    #
    # These act on `self._text_range()`, so on a whole-shape anchor they apply to
    # all its paragraphs and on a `Paragraph` to just that one. Wrap in
    # `deck.edit(...)` for view preservation + a one-Ctrl-Z fence.

    def paragraph_count(self) -> int:
        """Number of paragraphs in this anchor's text range."""
        with _com.translate_com_errors():
            return int(self._text_range().Paragraphs().Count)

    def format_text(
        self,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        font: str | None = None,
        color: str | int | tuple[int, int, int] | None = None,
    ) -> None:
        """Set font formatting on this anchor's text (PowerPoint's `apply_style`).

        PowerPoint has no named paragraph styles, so styling is direct font
        formatting. Only the kwargs you pass are written. `size` is in points;
        `color` is `"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB int.
        """
        if color is not None:
            parse_color(color)  # validate before any COM
        with _com.translate_com_errors():
            apply_font(
                self._text_range().Font,
                bold=bold,
                italic=italic,
                underline=underline,
                size=size,
                font=font,
                color=color,
            )

    def format_paragraph(
        self,
        *,
        alignment: str | int | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: float | None = None,
        line_spacing_points: float | None = None,
        space_before_lines: float | None = None,
        space_after_lines: float | None = None,
        indent_level: int | None = None,
        force: bool = False,
    ) -> None:
        """Set paragraph formatting on this anchor's paragraphs.

        Only the kwargs you pass are written. `alignment` is a name
        (`"left"`/`"center"`/`"right"`/`"justify"`/`"distribute"`) or int.

        **Spacing is unit-explicit** (PowerPoint stores a bare number whose unit a
        `LineRule*` flag selects, so the same number means wildly different things):

        * `line_spacing` — a **multiple** (`1.0` single, `1.5`, `2.0`).
        * `line_spacing_points` — an **exact point** height (e.g. `24` = 24 pt).
        * `space_before` / `space_after` — **points** before/after the paragraph.
        * `space_before_lines` / `space_after_lines` — the same as a **multiple**.

        Pass at most one of each line-spacing pair (`line_spacing` xor
        `line_spacing_points`, etc.) — passing both raises `ValueError`. A
        `line_spacing` multiple above 5× is rejected unless `force=True` (it's
        almost always a points-vs-multiple mix-up — pass `line_spacing_points`
        instead). `indent_level` is PowerPoint's outline/bullet level, 1-5 (its only
        notion of paragraph indent — there is no points-based left indent).
        """
        align_int = alignment_for(alignment) if alignment is not None else None
        if indent_level is not None and not (1 <= int(indent_level) <= 5):
            raise ValueError(f"indent_level must be between 1 and 5, got {indent_level}")
        if line_spacing is not None and line_spacing_points is not None:
            raise ValueError(
                "pass line_spacing (a multiple) or line_spacing_points (exact pt), not both"
            )
        if space_before is not None and space_before_lines is not None:
            raise ValueError("pass space_before (pt) or space_before_lines (a multiple), not both")
        if space_after is not None and space_after_lines is not None:
            raise ValueError("pass space_after (pt) or space_after_lines (a multiple), not both")
        if (
            line_spacing is not None
            and float(line_spacing) > LINE_SPACING_MULTIPLE_MAX
            and not force
        ):
            raise ValueError(
                f"line_spacing={line_spacing} is a *multiple* — "
                f"{line_spacing}× line height pushes text off the slide. "
                f"Did you mean line_spacing_points={line_spacing} (exact pt)? "
                "Pass force=True to set the multiple anyway."
            )
        with _com.translate_com_errors():
            tr = self._text_range()
            apply_paragraph_format(
                tr.ParagraphFormat,
                alignment=align_int,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
                line_spacing_points=line_spacing_points,
                space_before_lines=space_before_lines,
                space_after_lines=space_after_lines,
            )
            if indent_level is not None:
                tr.IndentLevel = int(indent_level)

    def apply_list(
        self, list_type: str = "bulleted", *, character: str | int | None = None
    ) -> None:
        """Turn this anchor's paragraphs into a bulleted or numbered list.

        `list_type` is `"bulleted"` (default) or `"numbered"`. `character` (a
        single char or int code point) sets a custom bullet glyph — only
        meaningful for a bulleted list. Raises `ValueError` for an unknown
        `list_type`.
        """
        bt = bullet_type_for(list_type)  # ValueError before any COM
        char_int = _bullet_char(character) if character is not None else None
        with _com.translate_com_errors():
            bullet = self._text_range().ParagraphFormat.Bullet
            bullet.Visible = _tristate(True)
            bullet.Type = int(bt)
            if char_int is not None:
                bullet.Character = char_int

    def remove_list(self) -> None:
        """Strip bullets / numbering from this anchor's paragraphs."""
        with _com.translate_com_errors():
            self._text_range().ParagraphFormat.Bullet.Visible = _tristate(False)

    def reset_format(self) -> None:
        """Reset this anchor's paragraph *spacing* to clean single-spaced defaults.

        The recovery verb for the line-spacing / space-before spiral — the gpt-5.4
        review's "giant spacing, text off the slide, unrecoverable without a
        rewrite" case. Sets single line spacing (`LineRuleWithin=msoTrue`,
        `SpaceWithin=1.0`), zero space before/after (in points), and indent level 1.

        **Honest scope (verified in `scripts/text_model_spike.py`):** PowerPoint
        exposes *no* "clear direct formatting" primitive — re-setting the text does
        not drop run overrides, and a run's size / typeface / colour / emphasis have
        no readable "inherited" value to fall back to, so this resets only the
        unambiguous paragraph-spacing knobs (where "single, 0, 0" *is* the clean
        default). To restore a **placeholder's** geometry + default font size from
        its layout, use `Shape.reset_to_layout()`; to set a specific font, use
        `format_text`. A mutation: wrap in `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            tr = self._text_range()
            apply_paragraph_format(
                tr.ParagraphFormat,
                line_spacing=1.0,
                space_before=0.0,
                space_after=0.0,
            )
            tr.IndentLevel = 1

    def insert_paragraph_before(self, text: str) -> None:
        """Insert `text` as a new paragraph immediately before this anchor's range.

        On a whole-shape anchor this prepends a first paragraph; on a `Paragraph`
        it inserts just above that paragraph. Embedded `\\n`/`\\r` in `text` add
        further paragraphs; `\\v` (`SOFT_BREAK`) adds a soft line break.
        """
        text = normalize_paragraph_breaks(text)
        with _com.translate_com_errors():
            tr = self._text_range()
            if str(tr.Text or "") == "":
                tr.Text = text
            else:
                tr.InsertBefore(text + "\r")

    def insert_paragraph_after(self, text: str) -> None:
        """Insert `text` as a new paragraph immediately after this anchor's range.

        On a whole-shape anchor this appends a paragraph (the common "add a
        bullet to the body" case); on a `Paragraph` it inserts just below it. The
        range includes its trailing break for a non-final paragraph, so we detect
        that to land a clean new paragraph either way (verified in the spike).
        Embedded `\\n`/`\\r` in `text` add further paragraphs; `\\v` (`SOFT_BREAK`)
        adds a soft line break.
        """
        text = normalize_paragraph_breaks(text)
        with _com.translate_com_errors():
            tr = self._text_range()
            raw = str(tr.Text or "")
            if raw == "":
                tr.InsertAfter(text)
            elif raw.endswith("\r"):
                tr.InsertAfter(text + "\r")
            else:
                tr.InsertAfter("\r" + text)

    def __repr__(self) -> str:
        return f"<{type(self).__name__} {self.anchor_id!r}>"


class Notes(Anchor):
    """The speaker-notes body of a slide — anchor id `notes:S`.

    Resolves the notes-page **body** placeholder by
    `PlaceholderFormat.Type == ppPlaceholderBody`, not by a hard index, because
    the index varies across templates (spec.md / IMPLEMENTATION.md spike item).
    Reads return `""` for an empty notes body; `set_text` replaces it.
    """

    kind = "notes"

    def __init__(self, slide: Slide) -> None:
        self._slide = slide

    @property
    def slide(self) -> Slide:
        return self._slide

    @property
    def anchor_id(self) -> str:
        return f"notes:{self._slide.index}"

    def _body_placeholder(self) -> Any:
        notes_page = self._slide.com.NotesPage
        for ph in notes_page.Shapes.Placeholders:
            try:
                if int(ph.PlaceholderFormat.Type) == int(PpPlaceholderType.BODY) and is_true(
                    ph.HasTextFrame
                ):
                    return ph
            except Exception:
                continue
        raise AnchorNotFoundError("notes", self.anchor_id)

    def _text_range(self) -> Any:
        return self._body_placeholder().TextFrame.TextRange


# ---------------------------------------------------------------------------
# Paragraphs — para:S:N:P
# ---------------------------------------------------------------------------


def _strip_break(text: str) -> str:
    """Drop the trailing paragraph/line break PowerPoint includes in a range."""
    return text.rstrip("\r\v\n")


_safe = _com.safe_read  # defensive COM-property read (returns a default on failure)


def _font_color_hex(font: Any) -> str | None:
    """`Font.Color` as `"#RRGGBB"`, or `None` for a non-literal (theme/auto) color.

    A theme or automatic color isn't a literal RGB: COM returns the `0x80000000`
    "automatic" sentinel from `.RGB` (which `color_hex` would mis-render as
    `#000000`), and `Color.Type` reports scheme/mixed rather than RGB. We only
    surface a hex when it's a genuine literal RGB (`0..0xFFFFFF`); otherwise the
    color is inherited from the theme and we report `None` (`color_hex_or_none`).
    """
    return color_hex_or_none(font.Color.RGB)


def font_to_dict(text_range: Any) -> dict[str, Any]:
    """The *effective* font of a text range, with tri-state fidelity (PPTLIVE-003).

    Returns the resolved font attributes PowerPoint reports for the range —
    `bold`/`italic`/`underline` as `True`/`False`/`"mixed"` (the `"mixed"` case is
    a range spanning differing runs, the signal `is_true` used to discard),
    `size` (points; `None` when mixed/unset), `font` (typeface name), and `color`
    (`"#RRGGBB"`, or `None` for a theme/automatic color — see `_font_color_hex`).

    Honest scope: these are *effective* values. COM resolves the master/layout
    style cascade before we see them and exposes no general "directly set on the
    run vs inherited" flag (only color carries a usable RGB-vs-theme distinction,
    surfaced here as a literal hex vs `None`). So this is "what is rendered", not
    "what was set on this run".
    """
    font = text_range.Font

    def _size() -> float | None:
        v = float(font.Size)
        return v if v > 0 else None  # PowerPoint returns <=0 for a mixed/unset size

    def _name() -> str | None:
        n = str(font.Name or "")
        return n or None

    return {
        "bold": _safe(lambda: tristate_value(font.Bold), False),
        "italic": _safe(lambda: tristate_value(font.Italic), False),
        "underline": _safe(lambda: tristate_value(font.Underline), False),
        "size": _safe(_size, None),
        "font": _safe(_name, None),
        "color": _safe(lambda: _font_color_hex(font), None),
    }


def _spacing_dict(pf: Any, value_attr: str, rule_attr: str) -> dict[str, Any]:
    """One spacing knob as `{value, mode}` — `mode` reads the paired `LineRule*`.

    PowerPoint stores a bare number whose unit a companion flag selects, so the
    read is meaningless without it: `mode` is `"multiple"` (`LineRule* == msoTrue`)
    or `"points"` (`msoFalse`). Defensive — a value PowerPoint can't supply for the
    range (e.g. a mixed span) degrades to `None`.
    """
    return {
        "value": _safe(lambda: float(getattr(pf, value_attr)), None),
        "mode": _safe(lambda: "multiple" if is_true(getattr(pf, rule_attr)) else "points", None),
    }


def _run_sizes(para_range: Any) -> list[float]:
    """The distinct font sizes across a paragraph's runs, in first-seen order.

    The mixed-run tell the gpt-5.4 review asked for: a stray 5 pt run hiding in an
    otherwise 18 pt paragraph shows up here as `[18.0, 5.0]`. A uniform paragraph
    is a single-element list. Defensive (`[]` if Runs can't be walked).
    """

    def _walk() -> list[float]:
        runs = para_range.Runs()
        count = int(runs.Count)
        seen: list[float] = []
        for i in range(1, count + 1):
            size = float(para_range.Runs(i, 1).Font.Size)
            if size > 0 and size not in seen:  # <=0 is PowerPoint's mixed/unset sentinel
                seen.append(size)
        return seen

    return _safe(_walk, [])


def paragraph_to_dict(para_range: Any, anchor_id: str, index: int) -> dict[str, Any]:
    """Structured snapshot of one paragraph for `shape.paragraphs.list()`.

    Reads are defensive — a property PowerPoint can't supply for this range
    degrades to a sensible default rather than failing the whole listing. The
    `font` sub-dict carries the full effective font (see `font_to_dict`); the
    top-level `bold`/`size` are kept for back-compat. `space_before`/`space_after`/
    `line_spacing` are `{value, mode}` (unit from the paired `LineRule*`), and
    `run_sizes` lists the distinct per-run font sizes so a stray small run is
    visible before it renders.
    """
    pf = para_range.ParagraphFormat
    font = font_to_dict(para_range)
    return {
        "index": index,
        "anchor_id": anchor_id,
        "text": _strip_break(str(para_range.Text or "")),
        "indent_level": _safe(lambda: int(para_range.IndentLevel), 1),
        "alignment": _safe(lambda: int(pf.Alignment), None),
        "bullet": _safe(
            lambda: bullet_type_name(pf.Bullet.Type) if is_true(pf.Bullet.Visible) else "none",
            "none",
        ),
        "space_before": _spacing_dict(pf, "SpaceBefore", "LineRuleBefore"),
        "space_after": _spacing_dict(pf, "SpaceAfter", "LineRuleAfter"),
        "line_spacing": _spacing_dict(pf, "SpaceWithin", "LineRuleWithin"),
        "run_sizes": _run_sizes(para_range),
        "bold": font["bold"],
        "size": font["size"],
        "font": font,
    }


class Paragraph(Anchor):
    """One paragraph of a shape's text frame — anchor id `para:S:N:P`.

    Located by 1-based paragraph index `P` within shape `N` (z-order) on slide
    `S`. Inherits every text verb (`set_text`, `format_text`, `format_paragraph`,
    `apply_list`, `insert_paragraph_before/after`); `_text_range()` is
    `TextFrame.TextRange.Paragraphs(P, 1)`, so those verbs scope to just this
    paragraph. Resolves live on each access (the paragraph count drifts as text
    is inserted/deleted), raising `AnchorNotFoundError` if `P` is out of range or
    `NoTextFrameError` (via the shape) if the shape holds no text.
    """

    kind = "paragraph"

    def __init__(self, shape: Shape, index: int) -> None:
        self._shape = shape
        self._index = int(index)

    @property
    def shape(self) -> Shape:
        return self._shape

    @property
    def slide(self) -> Slide:
        return self._shape.slide

    @property
    def index(self) -> int:
        """1-based paragraph index within the shape's text frame."""
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"para:{self._shape.slide.index}:{self._shape.index}:{self._index}"

    def _text_range(self) -> Any:
        tr = self._shape._text_range()  # NoTextFrameError if the shape has no frame
        count = int(tr.Paragraphs().Count)
        if self._index < 1 or self._index > count:
            raise AnchorNotFoundError("paragraph", self.anchor_id)
        return tr.Paragraphs(self._index, 1)

    @property
    def text(self) -> str:
        """The paragraph's text, without the trailing paragraph break."""
        with _com.translate_com_errors():
            return _strip_break(str(self._text_range().Text or ""))

    @property
    def indent_level(self) -> int:
        """PowerPoint outline/bullet level, 1-5."""
        with _com.translate_com_errors():
            return int(self._text_range().IndentLevel)

    def delete(self) -> None:
        """Delete this paragraph (text + its break). The wrapper is spent."""
        with _com.translate_com_errors():
            self._text_range().Delete()

    def to_dict(self) -> dict[str, Any]:
        with _com.translate_com_errors():
            return paragraph_to_dict(self._text_range(), self.anchor_id, self._index)


class ParagraphCollection:
    """Indexable, iterable view over the paragraphs of a shape's text frame.

    `shape.paragraphs[2]` is the 2nd paragraph (1-based); iteration yields a
    `Paragraph` each; `list()` emits the structured dict used by the
    `paragraphs` CLI command. Raises `NoTextFrameError` (via the shape) if the
    shape holds no text.
    """

    def __init__(self, shape: Shape) -> None:
        self._shape = shape

    def _count(self) -> int:
        with _com.translate_com_errors():
            return int(self._shape._text_range().Paragraphs().Count)

    def __len__(self) -> int:
        return self._count()

    def _anchor_id(self, index: int) -> str:
        return f"para:{self._shape.slide.index}:{self._shape.index}:{index}"

    def __getitem__(self, index: int) -> Paragraph:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"paragraph index must be int, got {type(index).__name__}")
        count = self._count()
        if index < 1 or index > count:
            raise AnchorNotFoundError("paragraph", self._anchor_id(index))
        return Paragraph(self._shape, index)

    def __iter__(self) -> Iterator[Paragraph]:
        for idx in range(1, self._count() + 1):
            yield Paragraph(self._shape, idx)

    def list(self) -> list[dict[str, Any]]:
        """Every paragraph as a structured dict, in order."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            tr = self._shape._text_range()
            count = int(tr.Paragraphs().Count)
            for idx in range(1, count + 1):
                out.append(paragraph_to_dict(tr.Paragraphs(idx, 1), self._anchor_id(idx), idx))
        return out

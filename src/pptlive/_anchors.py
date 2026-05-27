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
    is_true,
    parse_color,
)
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide


def _tristate(value: bool) -> int:
    """Python bool -> `MsoTriState` int (`msoTrue` / `msoFalse`)."""
    return int(MsoTriState.TRUE) if value else int(MsoTriState.FALSE)


def _bullet_char(character: str | int) -> int:
    """A single-char string or an int code point -> the int `Bullet.Character`."""
    if isinstance(character, str):
        if len(character) != 1:
            raise ValueError(f"bullet character must be a single character, got {character!r}")
        return ord(character)
    if isinstance(character, bool) or not isinstance(character, int):
        raise TypeError("bullet character must be a single-char str or an int code point")
    return int(character)


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

        Embed `\\n` (or `\\r`) for multiple paragraphs — PowerPoint treats them
        as paragraph breaks. Targets the text range directly, never the
        Selection, so it doesn't move the user's view. Wrap in `deck.edit(...)`
        to preserve the viewed slide and collapse the block to one Ctrl-Z
        (see `EditScope`).
        """
        with _com.translate_com_errors():
            self._text_range().Text = text

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
        color_long = parse_color(color) if color is not None else None
        with _com.translate_com_errors():
            f = self._text_range().Font
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
            if color_long is not None:
                f.Color.RGB = color_long

    def format_paragraph(
        self,
        *,
        alignment: str | int | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: float | None = None,
        indent_level: int | None = None,
    ) -> None:
        """Set paragraph formatting on this anchor's paragraphs.

        Only the kwargs you pass are written. `alignment` is a name
        (`"left"`/`"center"`/`"right"`/`"justify"`/`"distribute"`) or int.
        `space_before`/`space_after` are in points; `line_spacing` is a multiple
        (`1.0` single, `1.5`, …). `indent_level` is PowerPoint's outline/bullet
        level, 1-5 (its only notion of paragraph indent — there is no points-based
        left indent on `ParagraphFormat`).
        """
        align_int = alignment_for(alignment) if alignment is not None else None
        if indent_level is not None and not (1 <= int(indent_level) <= 5):
            raise ValueError(f"indent_level must be between 1 and 5, got {indent_level}")
        with _com.translate_com_errors():
            tr = self._text_range()
            pf = tr.ParagraphFormat
            if align_int is not None:
                pf.Alignment = align_int
            if space_before is not None:
                pf.SpaceBefore = float(space_before)
            if space_after is not None:
                pf.SpaceAfter = float(space_after)
            if line_spacing is not None:
                pf.SpaceWithin = float(line_spacing)
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

    def insert_paragraph_before(self, text: str) -> None:
        """Insert `text` as a new paragraph immediately before this anchor's range.

        On a whole-shape anchor this prepends a first paragraph; on a `Paragraph`
        it inserts just above that paragraph.
        """
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
        """
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


def paragraph_to_dict(para_range: Any, anchor_id: str, index: int) -> dict[str, Any]:
    """Structured snapshot of one paragraph for `shape.paragraphs.list()`.

    Reads are defensive — a property PowerPoint can't supply for this range
    degrades to a sensible default rather than failing the whole listing.
    """

    def _safe(fn: Any, default: Any) -> Any:
        try:
            return fn()
        except Exception:
            return default

    pf = para_range.ParagraphFormat
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
        "bold": _safe(lambda: is_true(para_range.Font.Bold), False),
        "size": _safe(lambda: float(para_range.Font.Size), None),
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

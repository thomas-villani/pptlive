"""Slides — the ordered containers, and the second level of the anchor hierarchy.

A `Slide` is a *container*, not a text anchor (like a whole table in wordlive):
slide-level verbs live here, while text lives on its shapes/placeholders/notes.
`deck.slides[S]` is 1-based. A slide exposes both its 1-based `index` (what users
say) and its stable `id` (`SlideID`, survives reordering) so listings can be
re-identified after a move.

Slide lifecycle (add/delete/duplicate/move/set-layout) is v0.1 and not built yet.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Notes
from ._shapes import PlaceholderShape, ShapeCollection, is_placeholder
from .constants import is_true, placeholder_types_for
from .exceptions import AnchorNotFoundError, SlideNotFoundError

if TYPE_CHECKING:
    from ._presentation import Presentation


def _paragraphs(text: str) -> list[str]:
    """Split a PowerPoint `TextRange.Text` into non-empty paragraph strings.

    PowerPoint separates paragraphs with `\\r` and soft line breaks with `\\v`;
    `\\n` shows up via COM too. We split on all three and drop blank lines so an
    outline read stays tidy.
    """
    out: list[str] = []
    for chunk in text.replace("\v", "\r").replace("\n", "\r").split("\r"):
        stripped = chunk.strip()
        if stripped:
            out.append(stripped)
    return out


class Slide:
    """Wraps a PowerPoint `Slide` COM object."""

    def __init__(self, deck: Presentation, slide_com: Any) -> None:
        self._deck = deck
        self._slide = slide_com

    @property
    def com(self) -> Any:
        return self._slide

    @property
    def index(self) -> int:
        """1-based position in the deck (`SlideIndex`). Shifts when slides move."""
        with _com.translate_com_errors():
            return int(self._slide.SlideIndex)

    @property
    def id(self) -> int:
        """Stable `SlideID` — survives reordering, unlike `index`."""
        with _com.translate_com_errors():
            return int(self._slide.SlideID)

    @property
    def shapes(self) -> ShapeCollection:
        return ShapeCollection(self)

    @property
    def notes(self) -> Notes:
        """The speaker-notes anchor (`notes:S`)."""
        return Notes(self)

    @property
    def layout_name(self) -> str | None:
        """The slide's custom-layout name (e.g. "Title and Content"), or None."""
        with _com.translate_com_errors():
            try:
                return str(self._slide.CustomLayout.Name)
            except Exception:
                return None

    def _find_placeholder(self, kind: str) -> tuple[Any, int]:
        """Resolve a placeholder KIND to (COM shape, 1-based z-order index).

        Picks the accepted `PpPlaceholderType` of highest preference (see
        `constants._PLACEHOLDER_KINDS`). Raises `AnchorNotFoundError` if no
        matching placeholder exists on the slide, `ValueError` for a bad KIND.
        """
        accepted = placeholder_types_for(kind)  # ValueError on unknown kind
        accepted_ints = [int(t) for t in accepted]
        best_rank: int | None = None
        best: tuple[Any, int] | None = None
        for idx, sh in enumerate(self._slide.Shapes, start=1):
            if not is_placeholder(sh):
                continue
            try:
                ph_type = int(sh.PlaceholderFormat.Type)
            except Exception:
                continue
            if ph_type in accepted_ints:
                rank = accepted_ints.index(ph_type)
                if best_rank is None or rank < best_rank:
                    best_rank, best = rank, (sh, idx)
        if best is None:
            raise AnchorNotFoundError("placeholder", f"ph:{self.index}:{kind.lower()}")
        return best

    def placeholder(self, kind: str) -> PlaceholderShape:
        """Return the `ph:S:KIND` placeholder anchor (resolved live by kind).

        KIND ∈ title, ctrtitle, subtitle, body, footer, date, slidenum. Raises
        `AnchorNotFoundError` if the slide has no such placeholder.
        """
        # Resolve once now so a missing placeholder fails fast with a clean error;
        # the returned anchor still re-resolves live on each use.
        with _com.translate_com_errors():
            self._find_placeholder(kind)
        return PlaceholderShape(self, kind)

    @property
    def title(self) -> str | None:
        """Text of the slide's title placeholder, or None if it has no title."""
        with _com.translate_com_errors():
            try:
                shapes = self._slide.Shapes
                if not is_true(shapes.HasTitle):
                    return None
                title_shape = shapes.Title
                if not is_true(title_shape.HasTextFrame):
                    return None
                return str(title_shape.TextFrame.TextRange.Text or "")
            except Exception:
                return None

    def has_notes(self) -> bool:
        """Whether the slide has non-empty speaker notes."""
        try:
            return bool(self.notes.text.strip())
        except Exception:
            # A missing notes body (AnchorNotFoundError) or any COM hiccup just
            # means "no notes" for the purpose of a listing.
            return False

    def read(self) -> dict[str, Any]:
        """Every shape on the slide plus its metadata — the `slide read S` payload."""
        return {
            "index": self.index,
            "id": self.id,
            "layout": self.layout_name,
            "title": self.title,
            "shapes": self.shapes.list(),
        }

    def __repr__(self) -> str:
        return f"<Slide index={self.index}>"


class SlideCollection:
    """Indexable, iterable view over a presentation's slides (1-based)."""

    def __init__(self, deck: Presentation) -> None:
        self._deck = deck

    @property
    def _com_collection(self) -> Any:
        return self._deck.com.Slides

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def __getitem__(self, index: int) -> Slide:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"slide index must be int, got {type(index).__name__}")
        count = len(self)
        if index < 1 or index > count:
            raise SlideNotFoundError(index)
        with _com.translate_com_errors():
            slide_com = self._com_collection(index)
        return Slide(self._deck, slide_com)

    def __iter__(self) -> Iterator[Slide]:
        with _com.translate_com_errors():
            slides = list(self._com_collection)
        for slide_com in slides:
            yield Slide(self._deck, slide_com)

    def list(self) -> list[dict[str, Any]]:
        """`[{index, id, layout, title, shape_count, has_notes}, ...]`."""
        out: list[dict[str, Any]] = []
        for slide in self:
            out.append(
                {
                    "index": slide.index,
                    "id": slide.id,
                    "layout": slide.layout_name,
                    "title": slide.title,
                    "shape_count": len(slide.shapes),
                    "has_notes": slide.has_notes(),
                }
            )
        return out

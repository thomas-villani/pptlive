"""Slides — the ordered containers, and the second level of the anchor hierarchy.

A `Slide` is a *container*, not a text anchor (like a whole table in wordlive):
slide-level verbs live here, while text lives on its shapes/placeholders/notes.
`deck.slides[S]` is 1-based. A slide exposes both its 1-based `index` (what users
say) and its stable `id` (`SlideID`, survives reordering) so listings can be
re-identified after a move.

Slide lifecycle — `SlideCollection.add()` and `Slide.delete/duplicate/move_to/
set_layout` — is the v0.1 track (the first with no Word analog). These verbs only
mutate; wrap a call in `deck.edit(label)` (as the CLI does) for view preservation
and a one-Ctrl-Z fence. Layout names resolve to a `CustomLayout` via
`Presentation._resolve_layout` (see `constants.match_layout_name`).
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterator
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Notes
from ._comments import CommentCollection
from ._shapes import PlaceholderShape, ShapeCollection, is_placeholder
from .constants import DEFAULT_LEGACY_LAYOUT, image_filter_for, is_true, placeholder_types_for
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    LayoutNotFoundError,
    PowerPointBusyError,
    SlideNotFoundError,
)

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
    def comments(self) -> CommentCollection:
        """The slide's review comments (`comments:S`) — read + add/reply/delete."""
        return CommentCollection(self)

    @property
    def layout_name(self) -> str | None:
        """The slide's custom-layout name (e.g. "Title and Content"), or None."""
        try:
            with _com.translate_com_errors():
                return str(self._slide.CustomLayout.Name)
        except PowerPointBusyError:
            # A transient busy must surface (exit 3, retryable), not masquerade
            # as "this slide has no layout".
            raise
        except Exception:
            return None

    def _find_placeholder(self, kind: str) -> tuple[Any, int]:
        """Resolve a placeholder KIND to (COM shape, 1-based z-order index).

        Picks the accepted `PpPlaceholderType` of highest preference (see
        `constants._PLACEHOLDER_KINDS`); a more-preferred type wins over a
        less-preferred one (so `body` prefers a real BODY over a generic OBJECT).
        But when **two or more** placeholders share the *same* best-preference
        type — e.g. the two OBJECT bodies of a Two Content / Comparison layout —
        the kind is genuinely ambiguous: rather than silently pick the first,
        raise `AmbiguousMatchError` (exit 5) listing the candidate `shape:S:N`
        anchors so the caller targets one explicitly. Raises `AnchorNotFoundError`
        if no matching placeholder exists, `ValueError` for a bad KIND.
        """
        accepted = placeholder_types_for(kind)  # ValueError on unknown kind
        accepted_ints = [int(t) for t in accepted]
        # Collect every accepted placeholder as (rank, idx, com_shape).
        matches: list[tuple[int, int, Any]] = []
        for idx, sh in enumerate(self._slide.Shapes, start=1):
            if not is_placeholder(sh):
                continue
            try:
                ph_type = int(sh.PlaceholderFormat.Type)
            except Exception:
                continue
            if ph_type in accepted_ints:
                matches.append((accepted_ints.index(ph_type), idx, sh))
        if not matches:
            raise AnchorNotFoundError("placeholder", f"ph:{self.index}:{kind.lower()}")
        best_rank = min(rank for rank, _idx, _sh in matches)
        tied = [(idx, sh) for rank, idx, sh in matches if rank == best_rank]
        if len(tied) > 1:
            candidates = [
                {
                    "anchor_id": f"shape:{self.index}:{idx}",
                    "name": str(sh.Name),
                    "id": int(sh.Id),
                    "index": idx,
                }
                for idx, sh in tied
            ]
            raise AmbiguousMatchError.for_placeholder(f"ph:{self.index}:{kind.lower()}", candidates)
        idx, sh = tied[0]
        return sh, idx

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
        try:
            with _com.translate_com_errors():
                shapes = self._slide.Shapes
                if not is_true(shapes.HasTitle):
                    return None
                title_shape = shapes.Title
                if not is_true(title_shape.HasTextFrame):
                    return None
                return str(title_shape.TextFrame.TextRange.Text or "")
        except PowerPointBusyError:
            # Don't let a momentary busy read back as "no title".
            raise
        except Exception:
            return None

    def has_notes(self) -> bool:
        """Whether the slide has non-empty speaker notes."""
        try:
            return bool(self.notes.text.strip())
        except PowerPointBusyError:
            # A transient busy is retryable (exit 3); don't bury it as "no notes".
            raise
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

    # -- render (v0.4; a read — no mutation, polite by nature) -----------------

    def export_image(
        self,
        path: str | os.PathLike[str] | None = None,
        *,
        width: int | None = None,
        height: int | None = None,
        fmt: str = "png",
    ) -> Path:
        """Render the slide to an image file and return its absolute path.

        Wraps `Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)`. The
        export renders the slide's **current in-memory state** — unsaved edits
        included — so an agent can edit over COM and immediately *see* the result;
        and it's polite (it doesn't move the viewed slide or change the Selection).

        `fmt` is a friendly token (`png`/`jpg`/`gif`/`bmp`/`tiff`; see
        `constants.IMAGE_FORMAT_CHOICES`). When `path` is None a temp file is
        created (so export-then-read is one step). `width`/`height` are output
        **pixels**; pass one and the other is filled from the slide's aspect
        ratio, pass neither for the slide's native pixel size. A relative `path`
        is resolved to absolute first — PowerPoint otherwise drops the file in
        its own working directory, not the caller's.
        """
        filter_name, ext = image_filter_for(fmt)  # ValueError before any COM
        if path is None:
            fd, tmp = tempfile.mkstemp(prefix="pptlive_slide_", suffix=f".{ext}")
            os.close(fd)
            os.remove(tmp)  # hand PowerPoint a clean path to write
            abs_path = tmp
        else:
            abs_path = os.path.abspath(os.fspath(path))
        with _com.translate_com_errors():
            w, h = self._export_dims(width, height)
            if w is not None and h is not None:
                self._slide.Export(abs_path, filter_name, int(round(w)), int(round(h)))
            else:
                self._slide.Export(abs_path, filter_name)
        return Path(abs_path)

    def _export_dims(
        self, width: int | None, height: int | None
    ) -> tuple[float | None, float | None]:
        """Resolve requested export pixels, filling a missing dimension from the
        slide's point aspect ratio so a single `width`/`height` keeps proportions."""
        if width is None and height is None:
            return None, None
        if width is not None and height is not None:
            return float(width), float(height)
        ps = self._deck.com.PageSetup
        sw, sh = float(ps.SlideWidth), float(ps.SlideHeight)
        if width is None:
            return float(height) * (sw / sh), float(height)  # type: ignore[arg-type]
        return float(width), float(width) * (sh / sw)

    # -- lifecycle (v0.1; wrap in deck.edit(...) for view + one-Ctrl-Z) --------

    def delete(self) -> None:
        """Delete this slide from the deck (`Slide.Delete`). The wrapper is spent."""
        with _com.translate_com_errors():
            self._slide.Delete()

    def duplicate(self) -> Slide:
        """Duplicate this slide; return the copy (inserted immediately after).

        Wraps `Slide.Duplicate`, which yields a one-item `SlideRange`. The copy
        gets a fresh `SlideID`; everything after the original shifts down by one.
        """
        with _com.translate_com_errors():
            new_range = self._slide.Duplicate()
            new_com = new_range(1)
        return Slide(self._deck, new_com)

    def move_to(self, index: int) -> Slide:
        """Move this slide to 1-based position `index` (`Slide.MoveTo`); return self.

        The wrapper keeps pointing at the same slide, which now reports the new
        `index`. Raises `SlideNotFoundError` if `index` is out of range (1..count).
        """
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"index must be int, got {type(index).__name__}")
        count = len(self._deck.slides)
        if index < 1 or index > count:
            raise SlideNotFoundError(index)
        with _com.translate_com_errors():
            self._slide.MoveTo(index)
        return self

    def set_layout(self, layout: str | int) -> Slide:
        """Re-apply a slide layout by friendly name or 1-based index; return self.

        Resolves `layout` to a `CustomLayout` (see `Presentation._resolve_layout`)
        and assigns `Slide.CustomLayout`. Raises `LayoutNotFoundError` (listing
        the deck's layout names) for an unknown layout.
        """
        custom = self._deck._resolve_layout(layout)
        if custom is None:
            raise LayoutNotFoundError(str(layout), [])
        with _com.translate_com_errors():
            self._slide.CustomLayout = custom
        return self

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

    def add(self, layout: str | int | None = None, index: int | None = None) -> Slide:
        """Insert a new slide and return it (v0.1; wrap in `deck.edit(...)`).

        `layout` is a friendly name or 1-based layout index (default
        `title_and_content`); `index` is the 1-based insertion position
        (default: appended to the end). Prefers the modern
        `Slides.AddSlide(Index, CustomLayout)`, falling back to legacy
        `Slides.Add` only on a deck that exposes no custom layouts. Raises
        `LayoutNotFoundError` for an unknown layout and `SlideNotFoundError`
        for an out-of-range insertion position (1..count+1).
        """
        count = len(self)
        if index is None:
            target = count + 1
        elif isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"index must be int, got {type(index).__name__}")
        elif index < 1 or index > count + 1:
            raise SlideNotFoundError(index)
        else:
            target = index
        with _com.translate_com_errors():
            custom = self._deck._resolve_layout(layout)
            if custom is not None:
                new_com = self._com_collection.AddSlide(target, custom)
            else:
                new_com = self._com_collection.Add(target, int(DEFAULT_LEGACY_LAYOUT))
        return Slide(self._deck, new_com)

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

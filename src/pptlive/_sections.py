"""Sections — the deck's top-level slide grouping (`Presentation.SectionProperties`).

A section is a *named span of consecutive slides* (the bars you collapse in the
slide sorter), the only deck-organization layer above the flat slide list. It owns
no text and moves no view, so — unlike the slide/shape verbs — section edits are
purely structural; still wrap a mutation in `deck.edit(...)` for the one-Ctrl-Z
fence (the CLI/MCP do).

Sections are addressed by a **1-based section index** (the order they appear), not
by slide. The model PowerPoint enforces (verified in `scripts/batch2_spike.py`):

- A section starts at a slide and runs until the next section begins, so
  `add(name, before_slide=N)` (`SectionProperties.AddBeforeSlide`) is the natural
  primitive — "start a section named NAME at slide N". Everything from N onward (to
  the next section) belongs to it.
- Adding the *first* section in front of slide N>1 makes PowerPoint auto-insert a
  leading **"Default Section"** for slides 1..N-1, so one `add` can yield two rows.
- `delete(index)` removes only the section *boundary* and keeps its slides
  (`Delete(index, deleteSlides=msoFalse)`); pass `delete_slides=True` to drop the
  slides too.

Each read row is `{index, name, first_slide, slide_count}` (`first_slide` is -1 for
an empty trailing section).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from .constants import MsoTriState
from .exceptions import AnchorNotFoundError, SlideNotFoundError

if TYPE_CHECKING:
    from ._presentation import Presentation


class SectionCollection:
    """Indexable, mutable view over a deck's sections (1-based section index)."""

    def __init__(self, deck: Presentation) -> None:
        self._deck = deck

    @property
    def _props(self) -> Any:
        return self._deck.com.SectionProperties

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._props.Count)

    def _check_index(self, index: int) -> None:
        """Validate a 1-based section index against the live count (clean error)."""
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"section index must be int, got {type(index).__name__}")
        count = len(self)
        if index < 1 or index > count:
            raise AnchorNotFoundError("section", f"section:{index}")

    def _row(self, props: Any, index: int) -> dict[str, Any]:
        """One section's `{index, name, first_slide, slide_count}` row."""
        first = int(props.FirstSlide(index))
        return {
            "index": index,
            "name": str(props.Name(index)),
            "first_slide": first if first > 0 else None,  # -1 -> empty trailing section
            "slide_count": int(props.SlidesCount(index)),
        }

    def list(self) -> list[dict[str, Any]]:
        """Every section in order — `[{index, name, first_slide, slide_count}, ...]`.

        `first_slide` is the 1-based index of the section's first slide (or `None`
        for an empty trailing section); `slide_count` is how many slides it spans.
        A read — no view move.
        """
        with _com.translate_com_errors():
            props = self._props
            return [self._row(props, i) for i in range(1, int(props.Count) + 1)]

    def add(self, name: str, *, before_slide: int | None = None) -> dict[str, Any]:
        """Create a section named `name`; return its new row.

        `before_slide` (1-based) is the slide the section **starts at** — the
        natural form ("start a 'Results' section at slide 5"); everything from there
        to the next section joins it. Adding the first section in front of slide
        N>1 makes PowerPoint also create a leading "Default Section" for the earlier
        slides. Omit `before_slide` to append an **empty** trailing section (a
        boundary slides can later move into).

        Raises `SlideNotFoundError` for an out-of-range `before_slide`. A mutation:
        wrap in `deck.edit(...)`.
        """
        # Validate before the COM mutation block (mirrors rename/delete/move).
        if before_slide is not None:
            if isinstance(before_slide, bool) or not isinstance(before_slide, int):
                raise TypeError(f"before_slide must be int, got {type(before_slide).__name__}")
            if before_slide < 1 or before_slide > len(self._deck.slides):
                raise SlideNotFoundError(before_slide)
        with _com.translate_com_errors():
            props = self._props
            if before_slide is None:
                new_index = int(props.AddSection(int(props.Count) + 1, str(name)))
            else:
                new_index = int(props.AddBeforeSlide(before_slide, str(name)))
            return self._row(props, new_index)

    def rename(self, index: int, name: str) -> dict[str, Any]:
        """Rename the section at 1-based `index`; return its updated row.

        Raises `AnchorNotFoundError` (exit 2) for an unknown index. A mutation.
        """
        self._check_index(index)
        with _com.translate_com_errors():
            props = self._props
            props.Rename(index, str(name))
            return self._row(props, index)

    def delete(self, index: int, *, delete_slides: bool = False) -> dict[str, Any]:
        """Delete the section at 1-based `index`; return `{deleted, name, ...}`.

        By default only the section **boundary** is removed and its slides stay
        (they merge into the previous section); pass `delete_slides=True` to delete
        the slides too. Raises `AnchorNotFoundError` for an unknown index. A
        mutation: wrap in `deck.edit(...)`.
        """
        self._check_index(index)
        with _com.translate_com_errors():
            props = self._props
            name = str(props.Name(index))
            props.Delete(index, int(MsoTriState.TRUE if delete_slides else MsoTriState.FALSE))
        return {"deleted": True, "index": index, "name": name, "slides_deleted": delete_slides}

    def move(self, index: int, to: int) -> dict[str, Any]:
        """Move the section at 1-based `index` to position `to`; return its new row.

        Reorders the section (and the slides it spans) within the deck. Raises
        `AnchorNotFoundError` for an out-of-range `index` or `to`. A mutation.
        """
        self._check_index(index)
        count = len(self)
        if isinstance(to, bool) or not isinstance(to, int):
            raise TypeError(f"to must be int, got {type(to).__name__}")
        if to < 1 or to > count:
            raise AnchorNotFoundError("section", f"section:{to}")
        with _com.translate_com_errors():
            props = self._props
            props.Move(index, to)
            return self._row(props, to)

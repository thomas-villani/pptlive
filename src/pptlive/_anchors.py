"""Anchor types — semantic, text-bearing handles inside a presentation.

The PowerPoint anchor model is *hierarchical* (slide → shape → paragraph), not a
global character stream, so there is no deck-wide `range:` and offsets are only
meaningful within one shape's text frame (see spec.md §"The anchor model"). An
anchor targets a COM `TextRange`, never the live `Selection`: text is set through
`TextFrame.TextRange.Text` directly, so no edit needs to select anything.

This module holds the abstract `Anchor` base and the `Notes` anchor. `Shape` —
which *is* an `Anchor` when it has a text frame — lives in `_shapes.py` because
it also carries geometry. `Paragraph` and `Cell` arrive in later stages (v0.3 /
v0.4).
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import PpPlaceholderType, is_true
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._slides import Slide


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
        to preserve the viewed slide (note: not atomic undo — see `EditScope`).
        """
        with _com.translate_com_errors():
            self._text_range().Text = text

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

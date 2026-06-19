"""Headers / footers — the footer, slide-number, and date placeholders.

A shared wrapper over a COM `HeadersFooters` object, mounted at two scopes that
mirror the background split (`Slide.headers_footers` for a per-slide override,
`Master.headers_footers` for the deck-wide default). Three elements:

- **footer** — `Footer.Visible` + `Footer.Text`
- **slide_number** — `SlideNumber.Visible` (the auto page number; no text)
- **date** — `DateAndTime.Visible` + a fixed `Text` *or* an auto-updating
  `Format`/`UseFormat`

The batch2 spike pinned the one sharp edge: reading `Footer.Text` / `UseFormat`
while the element is **not visible** raises `"Invalid request"`, so every text read
is guarded (`_com.safe_read`) and degrades to `None` rather than failing. Setting
text auto-shows the element (you don't get a footer by setting hidden text).

A read is polite (no view move); a set is a mutation — wrap in `deck.edit(...)`.
"""

from __future__ import annotations

from typing import Any

from . import _com
from .constants import MsoTriState, is_true

_safe = _com.safe_read


def _tri(value: bool) -> int:
    return int(MsoTriState.TRUE) if value else int(MsoTriState.FALSE)


class HeadersFooters:
    """Wraps a COM `HeadersFooters` (a slide's or the master's)."""

    def __init__(self, com_hf: Any) -> None:
        self._hf = com_hf

    @property
    def com(self) -> Any:
        return self._hf

    def read(self) -> dict[str, Any]:
        """`{footer, slide_number, date, display_on_title_slide}` — guarded.

        `footer` is `{visible, text}`; `slide_number` is `{visible}`; `date` is
        `{visible, text, format, use_format}`. Per the spike, `text`/`use_format`
        are only readable while their element is visible, so they degrade to `None`
        when hidden. `display_on_title_slide` is master-scoped (None on a slide).
        A read — no view move.
        """
        hf = self._hf
        footer = hf.Footer
        date = hf.DateAndTime
        return {
            "footer": {
                "visible": _safe(lambda: is_true(footer.Visible), False),
                "text": _safe(lambda: str(footer.Text), None),
            },
            "slide_number": {
                "visible": _safe(lambda: is_true(hf.SlideNumber.Visible), False),
            },
            "date": {
                "visible": _safe(lambda: is_true(date.Visible), False),
                "text": _safe(lambda: str(date.Text), None),
                "format": _safe(lambda: int(date.Format), None),
                "use_format": _safe(lambda: is_true(date.UseFormat), None),
            },
            "display_on_title_slide": _safe(lambda: is_true(hf.DisplayOnTitleSlide), None),
        }

    def set_footer(self, *, text: str | None = None, visible: bool | None = None) -> dict[str, Any]:
        """Set the footer text and/or visibility; return the resulting read.

        Passing `text` auto-shows the footer (unless `visible=False` is explicit) —
        a hidden footer's text doesn't render, so setting text implies showing it.
        Pass `visible=False` to hide the footer. A mutation: wrap in `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            footer = self._hf.Footer
            show = visible if visible is not None else (True if text is not None else None)
            if show is not None:
                footer.Visible = _tri(show)
            if text is not None:
                footer.Text = str(text)
        return self.read()

    def set_slide_number(self, visible: bool) -> dict[str, Any]:
        """Show or hide the auto slide-number placeholder; return the resulting read.

        A mutation: wrap in `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            self._hf.SlideNumber.Visible = _tri(visible)
        return self.read()

    def set_date(
        self,
        *,
        visible: bool | None = None,
        text: str | None = None,
        fmt: int | None = None,
    ) -> dict[str, Any]:
        """Set the date/time placeholder; return the resulting read.

        Pass `text` for a **fixed** date string (sets `UseFormat=msoFalse`), or
        `fmt` (a raw `PpDateTimeFormat` int) for an **auto-updating** date (sets
        `UseFormat=msoTrue`); they're mutually exclusive. Passing either auto-shows
        the element unless `visible=False` is explicit. A mutation: wrap in
        `deck.edit(...)`.
        """
        if text is not None and fmt is not None:
            raise ValueError("set_date() takes either text (fixed) or fmt (auto), not both")
        with _com.translate_com_errors():
            date = self._hf.DateAndTime
            given = text is not None or fmt is not None
            show = visible if visible is not None else (True if given else None)
            if show is not None:
                date.Visible = _tri(show)
            if text is not None:
                date.UseFormat = _tri(False)
                date.Text = str(text)
            if fmt is not None:
                date.UseFormat = _tri(True)
                date.Format = int(fmt)
        return self.read()

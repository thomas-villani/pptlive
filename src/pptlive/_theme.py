"""Deck-wide styling — the `Theme` and `Master` surfaces (v0.9).

The counterpart to v0.3's per-run `format_text`: instead of styling one anchor,
these restyle the **whole deck** by editing what every slide inherits. Both are
rooted at the primary `Presentation.SlideMaster`, so they share this module:

- **`Theme`** (`deck.theme`) — the template palette + typefaces:
  `SlideMaster.Theme.ThemeColorScheme.Colors(1..12)` and
  `Theme.ThemeFontScheme.MajorFont`/`MinorFont`.
- **`Master`** (`deck.master`) — the master text styles (PowerPoint's nearest
  "named style" analog: `TextStyles(title/body/default).Levels(1..5)`) and the
  master background fill.

These are deliberately **global, anti-polite** ops — one call recolors or
re-fonts every inheriting slide — which is why they live on their own surfaces
rather than folding into `format_text`. They still mutate, so wrap them in
`deck.edit(...)` (as the CLI/MCP do) for the one-Ctrl-Z fence; the user's *view*
doesn't move (the edits never navigate to the master), so view-restore is a
no-op.

Scope (v0.9 first cut): the **primary** `SlideMaster`. Multi-master decks
(`Presentation.Designs`) are deferred — `deck.master` is the primary master, and
a `deck.masters` collection can be added later without breaking this surface.

Findings baked in from the 2026-05-28 feasibility probe + `scripts/master_spike.py`:

1. **Theme color RGB is the same R-low-byte long as `Font.Color.RGB`**, so
   `parse_color` / `color_hex` apply unchanged.
2. **Theme fonts must be reached via `.Item(n)`** — the late-bound `.Latin`
   accessor raises `AttributeError`. `latin` (Item 1) is the default script.
3. **Master text styles drive the same COM `Font` / `ParagraphFormat` objects**
   that `Anchor.format_text` / `format_paragraph` use, so `_anchors.apply_font` /
   `apply_paragraph_format` are reused verbatim.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import apply_font, apply_paragraph_format
from .constants import (
    THEME_COLOR_CHOICES,
    alignment_for,
    color_hex,
    is_true,
    parse_color,
    text_style_for,
    theme_color_for,
    theme_font_script_for,
    theme_font_slot_for,
)

if TYPE_CHECKING:
    from ._presentation import Presentation


def _safe(fn: Any, default: Any) -> Any:
    """Read a COM property defensively — degrade to `default` rather than raise."""
    try:
        return fn()
    except Exception:
        return default


# msoFillType ints -> a friendly name (read-back only; we only *set* solid fills).
_FILL_TYPE_NAMES: dict[int, str] = {
    -2: "mixed",
    1: "solid",
    2: "patterned",
    3: "gradient",
    4: "textured",
    5: "background",
    6: "picture",
}


class Theme:
    """The deck's theme — palette + typefaces — bound to the `Presentation`.

    ```
    deck.theme.read()                              # {colors:{...}, fonts:{...}}
    deck.theme.set_color("accent1", "#C00000")     # recolor every inheriting slide
    deck.theme.set_font("major", "Georgia")        # swap the headings typeface
    ```

    `read()` is side-effect-free. The setters mutate the whole deck; wrap them in
    `deck.edit(...)`. Reaches `SlideMaster.Theme` live every call.
    """

    def __init__(self, deck: Presentation) -> None:
        self._deck = deck

    @property
    def com(self) -> Any:
        """Raw COM `Theme` (`SlideMaster.Theme`), resolved live."""
        with _com.translate_com_errors():
            return self._deck.com.SlideMaster.Theme

    def read(self) -> dict[str, Any]:
        """The full palette + the major/minor (Latin) typefaces.

        `{"colors": {slot: "#RRGGBB", ...}, "fonts": {"major": name, "minor": name}}`.
        Slots are the 12 canonical names in palette order. Side-effect-free.
        """
        with _com.translate_com_errors():
            theme = self.com
            scheme = theme.ThemeColorScheme
            colors = {
                slot: _safe(lambda i=idx: color_hex(scheme.Colors(i).RGB), None)
                for idx, slot in enumerate(THEME_COLOR_CHOICES, start=1)
            }
            fonts_com = theme.ThemeFontScheme
            fonts = {
                "major": _safe(lambda: str(fonts_com.MajorFont.Item(1).Name), None),
                "minor": _safe(lambda: str(fonts_com.MinorFont.Item(1).Name), None),
            }
        return {"colors": colors, "fonts": fonts}

    def set_color(self, slot: str, color: str | int | tuple[int, int, int]) -> None:
        """Set one palette slot (`"accent1"`, `"dark1"`, `"hyperlink"`, …).

        `color` is `"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB int. Unknown
        slot -> `ValueError` (before any COM). Recolors every slide that inherits
        the theme; wrap in `deck.edit(...)`.
        """
        idx = theme_color_for(slot)  # ValueError before any COM
        rgb = parse_color(color)
        with _com.translate_com_errors():
            self.com.ThemeColorScheme.Colors(idx).RGB = rgb

    def set_font(self, which: str, name: str, *, script: str = "latin") -> None:
        """Set the major (headings) or minor (body) typeface.

        `which` is `"major"`/`"minor"` (`"heading"`/`"body"` accepted). `script`
        selects the sub-typeface — `"latin"` (default), `"east_asian"`, or
        `"complex_script"` (the `.Latin` accessor is broken; we go through
        `.Item(n)`). Unknown `which`/`script` -> `ValueError` before any COM.
        """
        slot = theme_font_slot_for(which)  # ValueError before any COM
        script_idx = theme_font_script_for(script)
        with _com.translate_com_errors():
            scheme = self.com.ThemeFontScheme
            font = scheme.MajorFont if slot == "major" else scheme.MinorFont
            font.Item(script_idx).Name = str(name)

    def __repr__(self) -> str:
        return f"<Theme {self._deck.name!r}>"


class Master:
    """The deck's primary slide master — text styles + background.

    ```
    deck.master.read()                                       # {text_styles, background}
    deck.master.format_text_style("body", 1, font="Georgia", size=32)
    deck.master.format_paragraph_style("title", 1, alignment="center")
    deck.master.set_background("#1F1F1F")
    ```

    The text styles drive the same COM `Font` / `ParagraphFormat` objects as
    `Anchor.format_text` / `format_paragraph` (so formatting is identical), but
    deck-wide. Setters mutate; wrap in `deck.edit(...)`. Reaches
    `SlideMaster` live every call.
    """

    _STYLES: tuple[str, ...] = ("title", "body", "default")

    def __init__(self, deck: Presentation) -> None:
        self._deck = deck

    @property
    def com(self) -> Any:
        """Raw COM `SlideMaster`, resolved live."""
        with _com.translate_com_errors():
            return self._deck.com.SlideMaster

    def _level_dict(self, level_com: Any, level: int) -> dict[str, Any]:
        f = level_com.Font
        pf = level_com.ParagraphFormat
        return {
            "level": level,
            "font": _safe(lambda: str(f.Name), None),
            "size": _safe(lambda: float(f.Size), None),
            "bold": _safe(lambda: is_true(f.Bold), None),
            "italic": _safe(lambda: is_true(f.Italic), None),
            "underline": _safe(lambda: is_true(f.Underline), None),
            "color": _safe(lambda: color_hex(f.Color.RGB), None),
            "alignment": _safe(lambda: int(pf.Alignment), None),
        }

    def read(self) -> dict[str, Any]:
        """The three text styles (5 levels each) + the background fill.

        `{"text_styles": {style: {"levels": [{level, font, size, bold, italic,
        underline, color, alignment}, ...]}}, "background": {"type", "color"}}`.
        Reads are defensive — a property the master can't supply degrades to
        `None` rather than failing the whole dump. Side-effect-free.
        """
        with _com.translate_com_errors():
            master = self.com
            text_styles: dict[str, Any] = {}
            for style in self._STYLES:
                ts = master.TextStyles(text_style_for(style))
                levels = [self._level_dict(ts.Levels(lvl), lvl) for lvl in range(1, 6)]
                text_styles[style] = {"levels": levels}
            fill = master.Background.Fill
            fill_type = _safe(lambda: int(fill.Type), None)
            background = {
                "type": _FILL_TYPE_NAMES.get(fill_type, fill_type)
                if fill_type is not None
                else None,
                "color": _safe(lambda: color_hex(fill.ForeColor.RGB), None),
            }
        return {"text_styles": text_styles, "background": background}

    def _level(self, style: str, level: int) -> Any:
        if not (1 <= int(level) <= 5):
            raise ValueError(f"text-style level must be between 1 and 5, got {level}")
        t = text_style_for(style)  # ValueError before any COM
        return self.com.TextStyles(t).Levels(int(level))

    def format_text_style(
        self,
        style: str,
        level: int = 1,
        *,
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
        size: float | None = None,
        font: str | None = None,
        color: str | int | tuple[int, int, int] | None = None,
    ) -> None:
        """Set font formatting on a master text style + outline level (deck-wide).

        `style` is `"title"`/`"body"`/`"default"`; `level` is 1-5 and defaults to
        `1` (the natural choice for `title`, which has a single level). Only the
        kwargs you pass are written (`size` in points; `color` `"#RRGGBB"` / tuple
        / int). Re-renders every slide that inherits the style; wrap in
        `deck.edit(...)`. Unknown style / out-of-range level -> `ValueError`.
        """
        if color is not None:
            parse_color(color)  # validate before any COM
        with _com.translate_com_errors():
            apply_font(
                self._level(style, level).Font,
                bold=bold,
                italic=italic,
                underline=underline,
                size=size,
                font=font,
                color=color,
            )

    def format_paragraph_style(
        self,
        style: str,
        level: int = 1,
        *,
        alignment: str | int | None = None,
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: float | None = None,
    ) -> None:
        """Set paragraph formatting on a master text style + outline level.

        `alignment` is a name (`"left"`/`"center"`/`"right"`/`"justify"`/
        `"distribute"`) or int; `space_before`/`space_after` are points;
        `line_spacing` is the multiple. Only the kwargs you pass are written. Wrap
        in `deck.edit(...)`. Unknown style/alignment / bad level -> `ValueError`.
        """
        align_int = alignment_for(alignment) if alignment is not None else None
        with _com.translate_com_errors():
            apply_paragraph_format(
                self._level(style, level).ParagraphFormat,
                alignment=align_int,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
            )

    def set_background(self, color: str | int | tuple[int, int, int]) -> None:
        """Set the master background to a solid color (deck-wide).

        `color` is `"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB int. Applies a
        solid fill to `SlideMaster.Background`; wrap in `deck.edit(...)`. (v0.9
        ships solid fills only — gradient/picture backgrounds are deferred.)
        """
        rgb = parse_color(color)
        with _com.translate_com_errors():
            fill = self.com.Background.Fill
            fill.Solid()
            fill.ForeColor.RGB = rgb

    def __repr__(self) -> str:
        return f"<Master {self._deck.name!r}>"

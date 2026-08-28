"""`anchor.format_info()` — the format-probe read (`spec-linter.md` §7a) and the
placeholder cascade baseline it compares against (§7b).

The read mirror of `format_text` / `format_paragraph`: an anchor's *effective*
character + paragraph formatting, each field a `{value, baseline, override}` cell.
`value` is what PowerPoint renders. `baseline` is what the anchor **would** show
with no direct formatting — resolved by walking the cascade by hand, because
PowerPoint has no named paragraph styles (spec §2):

    slide placeholder  →  the same-kind placeholder on the slide's `CustomLayout`
                       →  (failing that) `SlideMaster.TextStyles(kind).Levels(n)`

`scripts/lint_cascade_spike.py` (2026-08-27) pinned the two facts this rests on:
an untouched slide placeholder reads **exactly** its layout placeholder's values
(so `override` needs no configuration), and the **layout** — not the master — is
the right rung: a title slide's `ctrtitle` is 60 pt on the layout while the master
title style says 44 pt, so a master-only baseline would flag every title slide.
The layout body placeholder's prompt text carries all five indent levels, which is
where a per-level body baseline comes from.

A free textbox, table cell, or notes body has **no cascade**: `baseline` is
`None`, `override` is `None`, and the linter judges those anchors only by the
peer-mode rules. A field that varies across the anchor's runs is listed in
`mixed` and its `override` is `None` (there is no single value to compare).
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import _safe, font_to_dict
from .constants import PpPlaceholderType, placeholder_kind_name, text_style_for

if TYPE_CHECKING:
    from ._anchors import Anchor

#: The master text style each placeholder kind falls back to when its layout has
#: no matching placeholder (`PpTextStyleType` via `text_style_for`).
_MASTER_STYLE_FOR_TYPE: dict[int, str] = {
    int(PpPlaceholderType.TITLE): "title",
    int(PpPlaceholderType.CENTER_TITLE): "title",
    int(PpPlaceholderType.VERTICAL_TITLE): "title",
    int(PpPlaceholderType.BODY): "body",
    int(PpPlaceholderType.OBJECT): "body",
    int(PpPlaceholderType.SUBTITLE): "body",
    int(PpPlaceholderType.VERTICAL_BODY): "body",
    int(PpPlaceholderType.VERTICAL_OBJECT): "body",
}

_FONT_FIELDS = ("name", "size", "bold", "italic", "underline", "color")
_PARA_FIELDS = ("alignment", "indent_level", "space_before", "space_after")


def layout_placeholder_for(shape_com: Any, slide_com: Any) -> Any | None:
    """The `CustomLayout` placeholder matching `shape_com`'s placeholder type, or
    `None` when the shape isn't a placeholder / the layout has no such placeholder.
    (The same walk `Shape.reset_to_layout` does, without raising.)"""
    try:
        want = int(shape_com.PlaceholderFormat.Type)
    except Exception:
        return None
    try:
        phs = slide_com.CustomLayout.Shapes.Placeholders
        for i in range(1, int(phs.Count) + 1):
            ph = phs(i)
            try:
                if int(ph.PlaceholderFormat.Type) == want:
                    return ph
            except Exception:
                continue
    except Exception:
        return None
    return None


def _font_values(tr: Any) -> dict[str, Any]:
    """The comparable font values of a text range (`font_to_dict` re-keyed to the
    write-verb vocabulary: `font` → `name`)."""
    d = font_to_dict(tr)
    return {
        "name": d["font"],
        "size": d["size"],
        "bold": d["bold"],
        "italic": d["italic"],
        "underline": d["underline"],
        "color": d["color"],
        "color_source": d["color_source"],
        "theme_color": d["theme_color"],
    }


def _para_values(tr: Any) -> dict[str, Any]:
    pf = tr.ParagraphFormat
    return {
        "alignment": _safe(lambda: int(pf.Alignment), None),
        "indent_level": _safe(lambda: int(tr.IndentLevel), None),
        "space_before": _safe(lambda: float(pf.SpaceBefore), None),
        "space_after": _safe(lambda: float(pf.SpaceAfter), None),
    }


def _level_paragraph(layout_tr: Any, level: int) -> Any:
    """The layout placeholder's paragraph at indent `level` (its prompt text carries
    one paragraph per level on a stock layout), or the whole range for level 1 /
    when no paragraph sits at that level."""
    if level <= 1:
        return layout_tr
    try:
        count = int(layout_tr.Paragraphs().Count)
        for p in range(1, count + 1):
            para = layout_tr.Paragraphs(p, 1)
            if int(para.IndentLevel) == level:
                return para
    except Exception:
        pass
    return layout_tr


def _master_level(deck_com: Any, ph_type: int, level: int) -> Any | None:
    style = _MASTER_STYLE_FOR_TYPE.get(ph_type)
    if style is None:
        return None
    try:
        return deck_com.SlideMaster.TextStyles(text_style_for(style)).Levels(max(1, min(level, 5)))
    except Exception:
        return None


def _master_values(level_com: Any) -> tuple[dict[str, Any], dict[str, Any]]:
    f = level_com.Font
    pf = level_com.ParagraphFormat
    font = {
        "name": _safe(lambda: str(f.Name) or None, None),
        "size": _safe(lambda: float(f.Size), None),
        "bold": _safe(lambda: bool(int(f.Bold)), None),
        "italic": _safe(lambda: bool(int(f.Italic)), None),
        "underline": _safe(lambda: bool(int(f.Underline)), None),
        "color": None,  # the master level's colour is a scheme slot; not comparable as hex
    }
    para = {
        "alignment": _safe(lambda: int(pf.Alignment), None),
        "indent_level": None,  # a master level has no IndentLevel; the caller sets it
        "space_before": _safe(lambda: float(pf.SpaceBefore), None),
        "space_after": _safe(lambda: float(pf.SpaceAfter), None),
    }
    return font, para


def resolve_baseline(
    shape_com: Any, slide_com: Any, deck_com: Any, level: int
) -> tuple[str | None, dict[str, Any] | None, dict[str, Any] | None]:
    """`(cascade, font_baseline, paragraph_baseline)` for a placeholder shape at
    indent `level` — `("layout", …)` from the layout placeholder, `("master:<style>",
    …)` from the master text style, or `(None, None, None)` when the shape has no
    cascade (a free textbox, picture, …)."""
    try:
        ph_type = int(shape_com.PlaceholderFormat.Type)
    except Exception:
        return None, None, None
    layout_ph = layout_placeholder_for(shape_com, slide_com)
    if layout_ph is not None:
        try:
            tr = _level_paragraph(layout_ph.TextFrame.TextRange, level)
            return "layout", _font_values(tr), _para_values(tr)
        except Exception:  # a layout placeholder with no text frame (picture, …)
            pass
    lvl = _master_level(deck_com, ph_type, level)
    if lvl is None:
        return None, None, None
    font, para = _master_values(lvl)
    para["indent_level"] = level
    return f"master:{_MASTER_STYLE_FOR_TYPE[ph_type]}", font, para


def _cell(value: Any, baseline: Any, mixed: bool) -> dict[str, Any]:
    override: bool | None
    if mixed or baseline is None or value is None:
        override = None
    else:
        override = _differs(value, baseline)
    return {"value": value, "baseline": baseline, "override": override}


def _differs(a: Any, b: Any) -> bool:
    if isinstance(a, float | int) and isinstance(b, float | int) and not isinstance(a, bool):
        return abs(float(a) - float(b)) > 0.01
    if isinstance(a, str) and isinstance(b, str):
        return a.casefold() != b.casefold()
    return bool(a != b)


def format_info(anchor: Anchor) -> dict[str, Any]:
    """Build the `format_info()` dict for `anchor` — see the module docstring."""
    from ._anchors import Paragraph  # noqa: PLC0415 — avoid an import cycle
    from ._shapes import Shape  # noqa: PLC0415

    with _com.translate_com_errors():
        tr = anchor._text_range()
        font = _font_values(tr)
        para = _para_values(tr)
        text = str(tr.Text or "")

    mixed: list[str] = [f for f in ("bold", "italic", "underline") if font[f] == "mixed"]
    if text.strip():
        if font["size"] is None:
            mixed.append("size")
        if font["name"] is None:
            mixed.append("name")
    if font["color_source"] == "mixed":
        mixed.append("color")

    shape_com: Any | None = None
    placeholder: str | None = None
    if isinstance(anchor, Shape):
        shape_com = anchor.com
    elif isinstance(anchor, Paragraph):
        shape_com = anchor.shape.com
    cascade: str | None = None
    font_base: dict[str, Any] | None = None
    para_base: dict[str, Any] | None = None
    if shape_com is not None:
        try:
            with _com.translate_com_errors():
                placeholder = placeholder_kind_name(shape_com.PlaceholderFormat.Type)
        except Exception:  # not a placeholder -> no PlaceholderFormat
            placeholder = None
        if placeholder is not None:
            level = para["indent_level"] or 1
            with _com.translate_com_errors():
                cascade, font_base, para_base = resolve_baseline(
                    shape_com, anchor.slide.com, anchor.slide.deck.com, int(level)
                )

    out_font: dict[str, Any] = {}
    for f in _FONT_FIELDS:
        cell = _cell(font[f], None if font_base is None else font_base.get(f), f in mixed)
        if f == "color":
            cell["source"] = font["color_source"]
            cell["theme_color"] = font["theme_color"]
        out_font[f] = cell
    out_para: dict[str, Any] = {}
    for f in _PARA_FIELDS:
        base = None if para_base is None else para_base.get(f)
        out_para[f] = _cell(para[f], base, False)
    return {
        "anchor_id": anchor.anchor_id,
        "placeholder": placeholder,
        "cascade": cascade,
        "font": out_font,
        "paragraph": out_para,
        "mixed": mixed,
    }

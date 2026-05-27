"""Typed enums for the PowerPoint magic constants pptlive uses.

Values mirror the official `Mso*` / `Pp*` enumerations exactly. Resist the urge
to pre-populate — add entries only as a feature needs them (the wordlive rule).
Friendly string aliases (`"title"`, `"textbox"`) coerce to the right int the way
wordlive's alignment names do.
"""

from __future__ import annotations

from collections.abc import Sequence
from enum import IntEnum
from typing import Any


class MsoTriState(IntEnum):
    """Office's tri-state boolean — `Shape.HasTextFrame`, `HasTable`, etc.

    COM returns `TRUE` as -1. `MIXED` (-2) only appears for multi-shape
    selections, which pptlive's anchors never hold.
    """

    FALSE = 0
    TRUE = -1
    MIXED = -2
    TOGGLE = -3


def is_true(tristate: Any) -> bool:
    """True iff an MsoTriState-valued COM property is `msoTrue` (-1).

    `bool(shape.HasTextFrame)` already works for the TRUE/FALSE pair, but this
    spells the intent out and ignores the MIXED/TOGGLE sentinels.
    """
    try:
        return int(tristate) == int(MsoTriState.TRUE)
    except (TypeError, ValueError):
        return bool(tristate)


class MsoShapeType(IntEnum):
    """`Shape.Type` values — what kind of object a shape is.

    The subset pptlive reports in `slide.read()`. Emitted as lowercase strings
    via `shape_type_name()` so JSON consumers match `"placeholder"` /
    `"picture"` without importing the enum.
    """

    AUTO_SHAPE = 1
    CALLOUT = 2
    CHART = 3
    COMMENT = 4
    FREEFORM = 5
    GROUP = 6
    EMBEDDED_OLE_OBJECT = 7
    FORM_CONTROL = 8
    LINE = 9
    LINKED_OLE_OBJECT = 10
    LINKED_PICTURE = 11
    OLE_CONTROL_OBJECT = 12
    PICTURE = 13
    PLACEHOLDER = 14
    TEXT_EFFECT = 15
    MEDIA = 16
    TEXT_BOX = 17
    SCRIPT_ANCHOR = 18
    TABLE = 19
    CANVAS = 20
    DIAGRAM = 21
    INK = 22
    INK_COMMENT = 23
    SMART_ART = 24


_SHAPE_TYPE_NAMES: dict[int, str] = {
    MsoShapeType.AUTO_SHAPE: "auto_shape",
    MsoShapeType.CALLOUT: "callout",
    MsoShapeType.CHART: "chart",
    MsoShapeType.COMMENT: "comment",
    MsoShapeType.FREEFORM: "freeform",
    MsoShapeType.GROUP: "group",
    MsoShapeType.EMBEDDED_OLE_OBJECT: "ole_object",
    MsoShapeType.FORM_CONTROL: "form_control",
    MsoShapeType.LINE: "line",
    MsoShapeType.LINKED_OLE_OBJECT: "linked_ole_object",
    MsoShapeType.LINKED_PICTURE: "linked_picture",
    MsoShapeType.OLE_CONTROL_OBJECT: "ole_control",
    MsoShapeType.PICTURE: "picture",
    MsoShapeType.PLACEHOLDER: "placeholder",
    MsoShapeType.TEXT_EFFECT: "text_effect",
    MsoShapeType.MEDIA: "media",
    MsoShapeType.TEXT_BOX: "textbox",
    MsoShapeType.SCRIPT_ANCHOR: "script_anchor",
    MsoShapeType.TABLE: "table",
    MsoShapeType.CANVAS: "canvas",
    MsoShapeType.DIAGRAM: "diagram",
    MsoShapeType.INK: "ink",
    MsoShapeType.INK_COMMENT: "ink_comment",
    MsoShapeType.SMART_ART: "smart_art",
}


def shape_type_name(value: Any) -> str:
    """Friendly lowercase name for a `Shape.Type` int (e.g. 14 -> "placeholder").

    Unknown values render as `"type:<n>"` rather than raising — a read should
    never fail because PowerPoint grew a shape kind we haven't enumerated.
    """
    try:
        return _SHAPE_TYPE_NAMES.get(int(value), f"type:{int(value)}")
    except (TypeError, ValueError):
        return "unknown"


class PpPlaceholderType(IntEnum):
    """`PlaceholderFormat.Type` values — the semantic role of a placeholder.

    `ph:S:KIND` resolves a friendly KIND to one of these (see
    `placeholder_types_for`). Reported by `placeholder_kind_name()` as a
    friendly string in `slide.read()`.
    """

    TITLE = 1
    BODY = 2
    CENTER_TITLE = 3
    SUBTITLE = 4
    VERTICAL_TITLE = 5
    VERTICAL_BODY = 6
    OBJECT = 7
    CHART = 8
    BITMAP = 9
    MEDIA_CLIP = 10
    ORG_CHART = 11
    TABLE = 12
    SLIDE_NUMBER = 13
    HEADER = 14
    FOOTER = 15
    DATE = 16
    VERTICAL_OBJECT = 17
    PICTURE = 18


# Friendly KIND -> the placeholder types it accepts, in preference order.
#
# An LLM asked to "set the title" doesn't know whether the slide is a title
# layout, so `title` also matches a center-title; `body` also matches the
# generic content `object` placeholder (what "Content Placeholder N" usually
# is). The distinct `ctrtitle` kind stays exact for callers that need it.
#
# Spike item (IMPLEMENTATION.md): confirm these type values + the body/object
# overlap against real templates before hardening.
_PLACEHOLDER_KINDS: dict[str, tuple[PpPlaceholderType, ...]] = {
    "title": (PpPlaceholderType.TITLE, PpPlaceholderType.CENTER_TITLE),
    "ctrtitle": (PpPlaceholderType.CENTER_TITLE,),
    "subtitle": (PpPlaceholderType.SUBTITLE,),
    "body": (PpPlaceholderType.BODY, PpPlaceholderType.OBJECT),
    "footer": (PpPlaceholderType.FOOTER,),
    "date": (PpPlaceholderType.DATE,),
    "slidenum": (PpPlaceholderType.SLIDE_NUMBER,),
}

PLACEHOLDER_KINDS: tuple[str, ...] = tuple(_PLACEHOLDER_KINDS)


def placeholder_types_for(kind: str) -> tuple[PpPlaceholderType, ...]:
    """Accepted `PpPlaceholderType`s for a friendly placeholder KIND.

    Raises `ValueError` for an unknown KIND (with the valid set), so callers can
    surface a clean message before touching COM.
    """
    try:
        return _PLACEHOLDER_KINDS[kind.lower()]
    except KeyError:
        raise ValueError(
            f"unknown placeholder kind {kind!r}; expected one of {sorted(_PLACEHOLDER_KINDS)}"
        ) from None


_PLACEHOLDER_TYPE_NAMES: dict[int, str] = {
    PpPlaceholderType.TITLE: "title",
    PpPlaceholderType.BODY: "body",
    PpPlaceholderType.CENTER_TITLE: "ctrtitle",
    PpPlaceholderType.SUBTITLE: "subtitle",
    PpPlaceholderType.VERTICAL_TITLE: "vertical_title",
    PpPlaceholderType.VERTICAL_BODY: "vertical_body",
    PpPlaceholderType.OBJECT: "object",
    PpPlaceholderType.CHART: "chart",
    PpPlaceholderType.BITMAP: "bitmap",
    PpPlaceholderType.MEDIA_CLIP: "media_clip",
    PpPlaceholderType.ORG_CHART: "org_chart",
    PpPlaceholderType.TABLE: "table",
    PpPlaceholderType.SLIDE_NUMBER: "slidenum",
    PpPlaceholderType.HEADER: "header",
    PpPlaceholderType.FOOTER: "footer",
    PpPlaceholderType.DATE: "date",
    PpPlaceholderType.VERTICAL_OBJECT: "vertical_object",
    PpPlaceholderType.PICTURE: "picture",
}


def placeholder_kind_name(value: Any) -> str:
    """Friendly name for a `PlaceholderFormat.Type` int (e.g. 1 -> "title")."""
    try:
        return _PLACEHOLDER_TYPE_NAMES.get(int(value), f"placeholder:{int(value)}")
    except (TypeError, ValueError):
        return "placeholder"


class PpSelectionType(IntEnum):
    """`Selection.Type` — what the user currently has selected in the window.

    Used by the politeness snapshot/restore: a `SHAPES` selection is re-selected
    by name on scope exit; `TEXT` and `NONE` fall back to restoring just the
    viewed slide. (Spike: confirm a shape-range selection round-trips cleanly.)
    """

    NONE = 0
    SLIDES = 1
    SHAPES = 2
    TEXT = 3


class PpViewType(IntEnum):
    """`DocumentWindow.ViewType` — the subset pptlive checks.

    `NORMAL` and `SLIDE` are the views where `View.Slide` (the slide the user is
    looking at) is meaningful; we snapshot/restore it for politeness.
    """

    NORMAL = 9
    SLIDE = 1
    OUTLINE = 6
    NOTES_PAGE = 5
    SLIDE_SORTER = 7


class PpSlideLayout(IntEnum):
    """Legacy `Slides.Add(Index, Layout)` layout codes (the deprecated path).

    Only reached as a fallback when a deck exposes no `CustomLayouts` for the
    modern `AddSlide(Index, CustomLayout)`. Friendly layout names resolve to a
    real `CustomLayout` instead (see `match_layout_name`), so this enum stays
    deliberately tiny — `TEXT` (a title-and-content slide) is the fallback default.
    """

    TITLE = 1
    TEXT = 2
    TWO_COLUMN_TEXT = 4
    TITLE_ONLY = 11
    BLANK = 12
    OBJECT = 16


# Default friendly layout for `slides.add()` when the caller names none, and the
# legacy `Slides.Add` layout used only when a deck has no CustomLayouts at all.
DEFAULT_LAYOUT_ALIAS = "title_and_content"
DEFAULT_LEGACY_LAYOUT = PpSlideLayout.TEXT


def _normalize_name(name: str) -> str:
    """Collapse a friendly name to its case/separator-insensitive comparison key.

    "Title and Content", "title_and_content", and "Title And Content" all map to
    "titleandcontent", so a friendly token matches a real name regardless of
    spacing, casing, or underscores. Used for both layout names and autoshape
    names.
    """
    return "".join(ch for ch in str(name).lower() if ch.isalnum())


# Friendly token (normalized) -> standard Office layout name (normalized), for
# tokens that don't already normalize to the layout's own name. Tokens like
# "two_content" or "blank" need no entry — they normalize straight onto "Two
# Content" / "Blank". This table only bridges the genuinely divergent shorthands.
_LAYOUT_ALIASES: dict[str, str] = {
    "title": "titleslide",
    "content": "titleandcontent",
    "titlecontent": "titleandcontent",
    "titleandbody": "titleandcontent",
    "bullets": "titleandcontent",
    "section": "sectionheader",
    "twocolumn": "twocontent",
    "twocolumntext": "twocontent",
    "caption": "contentwithcaption",
}


def match_layout_name(available: Sequence[str], requested: str) -> int | None:
    """1-based index into `available` of the layout matching `requested`, else None.

    Matches case/separator-insensitively against the deck's *actual* layout names
    first — so any template, including one whose layouts were renamed, resolves
    by its real name — then falls back to a small friendly-alias table for the
    standard Office layouts. Returns None when nothing matches; callers raise
    `LayoutNotFoundError` carrying `available` so an agent can pick a valid name.
    """
    norm_available = [_normalize_name(name) for name in available]
    want = _normalize_name(requested)
    if not want:
        return None
    if want in norm_available:
        return norm_available.index(want) + 1
    canonical = _LAYOUT_ALIASES.get(want)
    if canonical is not None and canonical in norm_available:
        return norm_available.index(canonical) + 1
    return None


class MsoTextOrientation(IntEnum):
    """`Shapes.AddTextbox` orientation. pptlive creates horizontal text boxes."""

    HORIZONTAL = 1
    VERTICAL = 5


class MsoAutoShapeType(IntEnum):
    """The `Shapes.AddShape(Type, …)` autoshape geometries pptlive names.

    A curated common subset of the (large) `MsoAutoShapeType` enumeration —
    added as a feature needs them (the wordlive rule), not pre-populated.
    Friendly names (`"rectangle"`, `"oval"`, `"arrow"`) resolve to these via
    `autoshape_type_for`; a raw int still passes through for the long tail.
    """

    RECTANGLE = 1
    PARALLELOGRAM = 2
    TRAPEZOID = 3
    DIAMOND = 4
    ROUNDED_RECTANGLE = 5
    OCTAGON = 6
    ISOSCELES_TRIANGLE = 7
    RIGHT_TRIANGLE = 8
    OVAL = 9
    HEXAGON = 10
    CROSS = 11
    REGULAR_PENTAGON = 12
    HEART = 21
    RIGHT_ARROW = 33
    LEFT_ARROW = 34
    UP_ARROW = 35
    DOWN_ARROW = 36
    FIVE_POINT_STAR = 92


# Friendly token (normalized) -> autoshape. Several spellings map to one shape
# ("ellipse"/"circle" -> oval, "arrow" -> right_arrow), so an agent needn't know
# Office's exact wording.
_AUTOSHAPE_NAMES: dict[str, MsoAutoShapeType] = {
    "rectangle": MsoAutoShapeType.RECTANGLE,
    "rect": MsoAutoShapeType.RECTANGLE,
    "box": MsoAutoShapeType.RECTANGLE,
    "square": MsoAutoShapeType.RECTANGLE,
    "roundedrectangle": MsoAutoShapeType.ROUNDED_RECTANGLE,
    "roundrect": MsoAutoShapeType.ROUNDED_RECTANGLE,
    "oval": MsoAutoShapeType.OVAL,
    "ellipse": MsoAutoShapeType.OVAL,
    "circle": MsoAutoShapeType.OVAL,
    "diamond": MsoAutoShapeType.DIAMOND,
    "triangle": MsoAutoShapeType.ISOSCELES_TRIANGLE,
    "isoscelestriangle": MsoAutoShapeType.ISOSCELES_TRIANGLE,
    "righttriangle": MsoAutoShapeType.RIGHT_TRIANGLE,
    "parallelogram": MsoAutoShapeType.PARALLELOGRAM,
    "trapezoid": MsoAutoShapeType.TRAPEZOID,
    "octagon": MsoAutoShapeType.OCTAGON,
    "hexagon": MsoAutoShapeType.HEXAGON,
    "pentagon": MsoAutoShapeType.REGULAR_PENTAGON,
    "regularpentagon": MsoAutoShapeType.REGULAR_PENTAGON,
    "cross": MsoAutoShapeType.CROSS,
    "plus": MsoAutoShapeType.CROSS,
    "heart": MsoAutoShapeType.HEART,
    "arrow": MsoAutoShapeType.RIGHT_ARROW,
    "rightarrow": MsoAutoShapeType.RIGHT_ARROW,
    "leftarrow": MsoAutoShapeType.LEFT_ARROW,
    "uparrow": MsoAutoShapeType.UP_ARROW,
    "downarrow": MsoAutoShapeType.DOWN_ARROW,
    "star": MsoAutoShapeType.FIVE_POINT_STAR,
    "fivepointstar": MsoAutoShapeType.FIVE_POINT_STAR,
    "star5": MsoAutoShapeType.FIVE_POINT_STAR,
}

# The canonical, readable names the CLI advertises (a `--shape-type` menu). The
# alias map above accepts more spellings; this is the discoverable shortlist.
AUTOSHAPE_CHOICES: tuple[str, ...] = (
    "rectangle",
    "rounded_rectangle",
    "oval",
    "diamond",
    "triangle",
    "right_triangle",
    "parallelogram",
    "trapezoid",
    "octagon",
    "hexagon",
    "pentagon",
    "cross",
    "heart",
    "arrow",
    "left_arrow",
    "up_arrow",
    "down_arrow",
    "star",
)


def autoshape_type_for(name: str | int) -> int:
    """Friendly autoshape name (or a raw `MsoAutoShapeType` int) -> the int.

    Names match case/separator-insensitively (`"Rounded Rectangle"`,
    `"rounded_rectangle"`, and `"roundrect"` all resolve). A raw int passes
    through unchanged — the escape hatch for autoshapes pptlive hasn't named.
    Raises `ValueError` (listing the friendly names) for an unknown name.
    """
    if isinstance(name, bool):
        raise ValueError(f"invalid autoshape type: {name!r}")
    if isinstance(name, int):
        return int(name)
    found = _AUTOSHAPE_NAMES.get(_normalize_name(name))
    if found is None:
        choices = ", ".join(AUTOSHAPE_CHOICES)
        raise ValueError(f"unknown autoshape {name!r}; expected one of: {choices}")
    return int(found)

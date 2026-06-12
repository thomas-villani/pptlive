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


def tristate_value(tristate: Any) -> bool | str:
    """An MsoTriState font property -> `True` / `False` / `"mixed"`.

    Unlike `is_true` (which collapses MIXED to False), this preserves the
    `msoTriStateMixed` (-2) signal a font property like `Font.Bold` returns when
    a *range* spans both bold and non-bold runs — so a reader can tell "this
    paragraph is uniformly not-bold" apart from "this paragraph mixes bold runs".
    """
    try:
        v = int(tristate)
    except (TypeError, ValueError):
        return bool(tristate)
    if v == int(MsoTriState.MIXED):
        return "mixed"
    return v == int(MsoTriState.TRUE)


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


class MsoAutoSize(IntEnum):
    """`TextFrame2.AutoSize` — how a text frame resizes to fit its text.

    Read off the modern `TextFrame2` (the classic `TextFrame.AutoSize` returns the
    mixed sentinel on current builds — see `scripts/text_model_spike.py`).
    """

    MIXED = -1
    NONE = 0  # neither resizes; text can overflow the frame
    TEXT_TO_FIT_SHAPE = 1  # shrink the text to fit the shape ("shrink text on overflow")
    SHAPE_TO_FIT_TEXT = 2  # grow the shape to fit the text


_AUTOSIZE_NAMES: dict[int, str] = {
    MsoAutoSize.MIXED: "mixed",
    MsoAutoSize.NONE: "none",
    MsoAutoSize.TEXT_TO_FIT_SHAPE: "text_to_fit_shape",
    MsoAutoSize.SHAPE_TO_FIT_TEXT: "shape_to_fit_text",
}


def autosize_name(value: Any) -> str:
    """Friendly name for a `TextFrame2.AutoSize` int (e.g. 2 -> "shape_to_fit_text")."""
    try:
        return _AUTOSIZE_NAMES.get(int(value), f"autosize:{int(value)}")
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


class MsoZOrderCmd(IntEnum):
    """`Shape.ZOrder(cmd)` — how to restack a shape relative to its siblings.

    Only the four that move a shape within the slide's z-stack; the
    in-front-of/behind-text variants aren't named (added as a feature needs them).
    """

    BRING_TO_FRONT = 0
    SEND_TO_BACK = 1
    BRING_FORWARD = 2
    SEND_BACKWARD = 3


# Friendly token (normalized) -> the z-order command. The short forms
# ("front"/"back"/"forward"/"backward") are the discoverable names; the verbose
# Office spellings ("bringtofront", …) resolve too.
_ZORDER_NAMES: dict[str, MsoZOrderCmd] = {
    "front": MsoZOrderCmd.BRING_TO_FRONT,
    "bringtofront": MsoZOrderCmd.BRING_TO_FRONT,
    "tofront": MsoZOrderCmd.BRING_TO_FRONT,
    "back": MsoZOrderCmd.SEND_TO_BACK,
    "sendtoback": MsoZOrderCmd.SEND_TO_BACK,
    "toback": MsoZOrderCmd.SEND_TO_BACK,
    "forward": MsoZOrderCmd.BRING_FORWARD,
    "bringforward": MsoZOrderCmd.BRING_FORWARD,
    "forwards": MsoZOrderCmd.BRING_FORWARD,
    "backward": MsoZOrderCmd.SEND_BACKWARD,
    "sendbackward": MsoZOrderCmd.SEND_BACKWARD,
    "backwards": MsoZOrderCmd.SEND_BACKWARD,
}

# The canonical names the CLI/MCP advertise (a `--to` menu).
ZORDER_CHOICES: tuple[str, ...] = ("front", "back", "forward", "backward")


def zorder_cmd_for(name: str | int) -> int:
    """Friendly z-order name (or a raw `MsoZOrderCmd` int) -> the int.

    `"front"`/`"back"`/`"forward"`/`"backward"` (and the verbose `"bring_to_front"`
    etc.) match case/separator-insensitively. A raw int passes through. Raises
    `ValueError` (listing the friendly names) for an unknown name — symmetric with
    `autoshape_type_for`.
    """
    if isinstance(name, bool):
        raise ValueError(f"invalid z-order command: {name!r}")
    if isinstance(name, int):
        return int(name)
    found = _ZORDER_NAMES.get(_normalize_name(name))
    if found is None:
        choices = ", ".join(ZORDER_CHOICES)
        raise ValueError(f"unknown z-order command {name!r}; expected one of: {choices}")
    return int(found)


# ---------------------------------------------------------------------------
# Text structure (v0.3): paragraph alignment, bullets, font color
# ---------------------------------------------------------------------------


class PpParagraphAlignment(IntEnum):
    """`ParagraphFormat.Alignment` — horizontal alignment of a paragraph."""

    LEFT = 1
    CENTER = 2
    RIGHT = 3
    JUSTIFY = 4
    DISTRIBUTE = 5
    THAI_DISTRIBUTE = 6
    JUSTIFY_LOW = 7


_ALIGNMENT_NAMES: dict[str, PpParagraphAlignment] = {
    "left": PpParagraphAlignment.LEFT,
    "center": PpParagraphAlignment.CENTER,
    "centre": PpParagraphAlignment.CENTER,
    "right": PpParagraphAlignment.RIGHT,
    "justify": PpParagraphAlignment.JUSTIFY,
    "distribute": PpParagraphAlignment.DISTRIBUTE,
}

ALIGNMENT_CHOICES: tuple[str, ...] = ("left", "center", "right", "justify", "distribute")


def alignment_for(value: str | int) -> int:
    """Coerce an alignment name/int to a `PpParagraphAlignment` int.

    Accepts `"left"`/`"center"`/`"right"`/`"justify"`/`"distribute"` (case-
    insensitive, `"centre"` too) or a raw int. Raises `ValueError` for an
    unknown name — symmetric with `autoshape_type_for`.
    """
    if isinstance(value, bool):
        raise ValueError(f"invalid alignment: {value!r}")
    if isinstance(value, int):
        return int(value)
    found = _ALIGNMENT_NAMES.get(_normalize_name(value))
    if found is None:
        choices = ", ".join(ALIGNMENT_CHOICES)
        raise ValueError(f"unknown alignment {value!r}; expected one of: {choices}")
    return int(found)


class PpBulletType(IntEnum):
    """`ParagraphFormat.Bullet.Type` — what kind of bullet a paragraph carries."""

    NONE = 0
    UNNUMBERED = 1
    NUMBERED = 2
    MIXED = -2


# Accepted `list_type` strings -> the bullet type to apply. The two canonical
# names are bulleted / numbered; common variants alias on.
_BULLET_TYPE_FOR: dict[str, PpBulletType] = {
    "bulleted": PpBulletType.UNNUMBERED,
    "bullet": PpBulletType.UNNUMBERED,
    "bullets": PpBulletType.UNNUMBERED,
    "unnumbered": PpBulletType.UNNUMBERED,
    "numbered": PpBulletType.NUMBERED,
    "number": PpBulletType.NUMBERED,
    "numbers": PpBulletType.NUMBERED,
}

LIST_TYPE_CHOICES: tuple[str, ...] = ("bulleted", "numbered")


def bullet_type_for(list_type: str) -> PpBulletType:
    """Resolve a `list_type` string to its `PpBulletType`.

    `"bulleted"` -> unnumbered, `"numbered"` -> numbered. Raises `ValueError`
    for an unknown name.
    """
    found = _BULLET_TYPE_FOR.get(_normalize_name(list_type))
    if found is None:
        choices = ", ".join(LIST_TYPE_CHOICES)
        raise ValueError(f"unknown list type {list_type!r}; expected one of: {choices}")
    return found


_BULLET_TYPE_NAMES: dict[int, str] = {
    int(PpBulletType.NONE): "none",
    int(PpBulletType.UNNUMBERED): "bulleted",
    int(PpBulletType.NUMBERED): "numbered",
    int(PpBulletType.MIXED): "mixed",
}


def bullet_type_name(value: Any) -> str:
    """Friendly name for a `Bullet.Type` int (e.g. 1 -> "bulleted")."""
    try:
        return _BULLET_TYPE_NAMES.get(int(value), f"bullet:{int(value)}")
    except (TypeError, ValueError):
        return "none"


def parse_color(value: str | int | tuple[int, int, int]) -> int:
    """Coerce a color to the long PowerPoint's `Font.Color.RGB` expects.

    Accepts `"#RRGGBB"` / `"RRGGBB"` hex, an `(r, g, b)` tuple (0-255 each), or a
    raw int (passed through — the escape hatch). PowerPoint stores the long in
    R-low-byte order (`red == 0x0000FF`), so `"#FF0000"` -> 255. Raises
    `ValueError` for a malformed hex string or out-of-range channel.
    """
    if isinstance(value, bool):
        raise ValueError(f"invalid color: {value!r}")
    if isinstance(value, int):
        return int(value)
    if isinstance(value, tuple):
        if len(value) != 3 or any(not (0 <= int(c) <= 255) for c in value):
            raise ValueError(f"color tuple must be three 0-255 channels, got {value!r}")
        r, g, b = (int(c) for c in value)
        return r | (g << 8) | (b << 16)
    text = str(value).strip().lstrip("#")
    if len(text) != 6:
        raise ValueError(f"color must be '#RRGGBB' hex, an (r,g,b) tuple, or an int; got {value!r}")
    try:
        r = int(text[0:2], 16)
        g = int(text[2:4], 16)
        b = int(text[4:6], 16)
    except ValueError:
        raise ValueError(f"invalid hex color {value!r}") from None
    return r | (g << 8) | (b << 16)


def color_hex(value: Any) -> str:
    """Render a PowerPoint `Font.Color.RGB` long as `"#RRGGBB"`."""
    n = int(value)
    r, g, b = n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def color_hex_or_none(value: Any) -> str | None:
    """`color_hex`, but `None` for a non-literal (theme/automatic) color.

    A theme or automatic color isn't a literal RGB: COM returns the `0x80000000`
    "automatic" sentinel (which `color_hex` would mis-render as `#000000`), and
    anything outside `0..0xFFFFFF` is likewise not a real RGB. Shared by the font
    (`_anchors._font_color_hex`) and shape fill/line readbacks so a theme-driven
    color reports honestly as `None` rather than a wrong black.
    """
    try:
        rgb = int(value)
    except (TypeError, ValueError):
        return None
    if rgb < 0 or rgb > 0xFFFFFF:
        return None
    return color_hex(rgb)


# ---------------------------------------------------------------------------
# Shape fill / effects (v1.2 advanced cut): gradient / pattern / picture fills,
# shadow / glow / soft-edge / reflection effects. Curated friendly subsets +
# raw-int passthrough + reverse names — spiked 2026-06-11 (fill_advanced_spike /
# effects_spike). Added only as the verbs need them (convention #7).
# ---------------------------------------------------------------------------


class MsoFillType(IntEnum):
    """`Fill.Type` — what kind of fill a shape / background carries (read-back)."""

    MIXED = -2
    SOLID = 1
    PATTERNED = 2
    GRADIENT = 3
    TEXTURED = 4
    BACKGROUND = 5
    PICTURE = 6


# Friendly read-back name for a `Fill.Type` int. Shared by the shape-fill,
# master-background, and slide-background readers (was `_shapes._FILL_TYPE_NAMES`).
FILL_TYPE_NAMES: dict[int, str] = {
    int(MsoFillType.MIXED): "mixed",
    int(MsoFillType.SOLID): "solid",
    int(MsoFillType.PATTERNED): "patterned",
    int(MsoFillType.GRADIENT): "gradient",
    int(MsoFillType.TEXTURED): "textured",
    int(MsoFillType.BACKGROUND): "background",
    int(MsoFillType.PICTURE): "picture",
}


def fill_type_name(value: Any) -> str | int | None:
    """Friendly name for a `Fill.Type` int (`1 -> "solid"`); the int if unknown."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return None
    return FILL_TYPE_NAMES.get(n, n)


class MsoGradientStyle(IntEnum):
    """`Fill.GradientStyle` — the direction a gradient sweeps."""

    MIXED = -2
    HORIZONTAL = 1
    VERTICAL = 2
    DIAGONAL_UP = 3
    DIAGONAL_DOWN = 4
    FROM_CORNER = 5
    FROM_TITLE = 6
    FROM_CENTER = 7


_GRADIENT_STYLES: dict[str, int] = {
    "horizontal": int(MsoGradientStyle.HORIZONTAL),
    "vertical": int(MsoGradientStyle.VERTICAL),
    "diagonal_up": int(MsoGradientStyle.DIAGONAL_UP),
    "diagonal_down": int(MsoGradientStyle.DIAGONAL_DOWN),
    "from_corner": int(MsoGradientStyle.FROM_CORNER),
    "from_title": int(MsoGradientStyle.FROM_TITLE),
    "from_center": int(MsoGradientStyle.FROM_CENTER),
}

GRADIENT_STYLE_CHOICES: tuple[str, ...] = (
    "horizontal",
    "vertical",
    "diagonal_up",
    "diagonal_down",
    "from_corner",
    "from_title",
    "from_center",
)

_GRADIENT_STYLE_NAMES: dict[int, str] = {v: k for k, v in _GRADIENT_STYLES.items()}


def gradient_style_for(style: str | int) -> int:
    """Friendly gradient-style name (or raw `MsoGradientStyle` int) -> the int.

    `"horizontal"`/`"vertical"`/`"diagonal_up"`/… (case- and separator-insensitive)
    or a raw int (passed through). Raises `ValueError` for an unknown name —
    symmetric with `entry_effect_for`.
    """
    if isinstance(style, bool):
        raise ValueError(f"invalid gradient style: {style!r}")
    if isinstance(style, int):
        return int(style)
    found = _GRADIENT_STYLES.get(str(style).strip().lower().replace(" ", "_").replace("-", "_"))
    if found is None:
        choices = ", ".join(GRADIENT_STYLE_CHOICES)
        raise ValueError(f"unknown gradient style {style!r}; expected one of: {choices}")
    return found


def gradient_style_name(value: Any) -> str | int | None:
    """Friendly name for a `Fill.GradientStyle` int (`1 -> "horizontal"`)."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return None
    return _GRADIENT_STYLE_NAMES.get(n, n)


# Curated `MsoPresetGradientType` subset (the 24 named presets; the common ones
# advertised, the rest reachable by raw int). Names match Office's preset menu.
_PRESET_GRADIENTS: dict[str, int] = {
    "early_sunset": 1,
    "late_sunset": 2,
    "nightfall": 3,
    "daybreak": 4,
    "horizon": 5,
    "desert": 6,
    "ocean": 7,
    "calm_water": 8,
    "fire": 9,
    "fog": 10,
    "moss": 11,
    "peacock": 12,
    "wheat": 13,
    "parchment": 14,
    "mahogany": 15,
    "rainbow": 16,
    "rainbow_ii": 17,
    "gold": 18,
    "gold_ii": 19,
    "brass": 20,
    "chrome": 21,
    "chrome_ii": 22,
    "silver": 23,
    "sapphire": 24,
}

PRESET_GRADIENT_CHOICES: tuple[str, ...] = tuple(_PRESET_GRADIENTS)

_PRESET_GRADIENT_NAMES: dict[int, str] = {v: k for k, v in _PRESET_GRADIENTS.items()}


def preset_gradient_for(preset: str | int) -> int:
    """Friendly preset-gradient name (or raw `MsoPresetGradientType` int) -> the int.

    `"ocean"`/`"fire"`/`"rainbow"`/… (case-/separator-insensitive) or a raw int.
    Raises `ValueError` for an unknown name.
    """
    if isinstance(preset, bool):
        raise ValueError(f"invalid preset gradient: {preset!r}")
    if isinstance(preset, int):
        return int(preset)
    found = _PRESET_GRADIENTS.get(str(preset).strip().lower().replace(" ", "_").replace("-", "_"))
    if found is None:
        choices = ", ".join(PRESET_GRADIENT_CHOICES)
        raise ValueError(f"unknown preset gradient {preset!r}; expected one of: {choices}")
    return found


def preset_gradient_name(value: Any) -> str | int | None:
    """Friendly name for a `MsoPresetGradientType` int (`7 -> "ocean"`)."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return None
    return _PRESET_GRADIENT_NAMES.get(n, n)


# Curated `MsoPatternType` subset — the percentage screens (1-12) and the
# confidently-named structural patterns (13-22). Exotic patterns ride the
# raw-int passthrough rather than risking a wrong name->int mapping.
_PATTERNS: dict[str, int] = {
    "percent_5": 1,
    "percent_10": 2,
    "percent_20": 3,
    "percent_25": 4,
    "percent_30": 5,
    "percent_40": 6,
    "percent_50": 7,
    "percent_60": 8,
    "percent_70": 9,
    "percent_75": 10,
    "percent_80": 11,
    "percent_90": 12,
    "dark_horizontal": 13,
    "dark_vertical": 14,
    "dark_downward_diagonal": 15,
    "dark_upward_diagonal": 16,
    "small_checkerboard": 17,
    "trellis": 18,
    "light_horizontal": 19,
    "light_vertical": 20,
    "light_downward_diagonal": 21,
    "light_upward_diagonal": 22,
}

PATTERN_CHOICES: tuple[str, ...] = tuple(_PATTERNS)

_PATTERN_NAMES: dict[int, str] = {v: k for k, v in _PATTERNS.items()}


def pattern_for(pattern: str | int) -> int:
    """Friendly pattern name (or raw `MsoPatternType` int) -> the int.

    `"percent_50"`/`"dark_horizontal"`/`"trellis"`/… (case-/separator-insensitive)
    or a raw int. Raises `ValueError` for an unknown name.
    """
    if isinstance(pattern, bool):
        raise ValueError(f"invalid pattern: {pattern!r}")
    if isinstance(pattern, int):
        return int(pattern)
    found = _PATTERNS.get(str(pattern).strip().lower().replace(" ", "_").replace("-", "_"))
    if found is None:
        choices = ", ".join(PATTERN_CHOICES)
        raise ValueError(f"unknown pattern {pattern!r}; expected one of: {choices}")
    return found


def pattern_name(value: Any) -> str | int | None:
    """Friendly name for a `MsoPatternType` int (`7 -> "percent_50"`)."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return None
    return _PATTERN_NAMES.get(n, n)


class MsoShadowStyle(IntEnum):
    """`Shadow.Style` — inner vs outer shadow (the read-back the spike pinned).

    Setting individual shadow props pushes `Shadow.Type` to the `mixed` (-2)
    sentinel, so `.Style` (not `.Type`) is the reliable read-back.
    """

    MIXED = -2
    INNER = 1
    OUTER = 2


# ---------------------------------------------------------------------------
# Slide render (v0.4): image-export formats
# ---------------------------------------------------------------------------


# Friendly format token -> (graphics-filter name, file extension). `Slide.Export`'s
# FilterName uses PowerPoint's registered export filters; the common raster set is
# reliably present. Added as the feature needs them (the wordlive rule).
_IMAGE_FILTERS: dict[str, tuple[str, str]] = {
    "png": ("PNG", "png"),
    "jpg": ("JPG", "jpg"),
    "jpeg": ("JPG", "jpg"),
    "gif": ("GIF", "gif"),
    "bmp": ("BMP", "bmp"),
    "tif": ("TIF", "tif"),
    "tiff": ("TIF", "tif"),
}

IMAGE_FORMAT_CHOICES: tuple[str, ...] = ("png", "jpg", "gif", "bmp", "tiff")


def image_filter_for(fmt: str) -> tuple[str, str]:
    """Resolve an image-format token to its `(FilterName, extension)` for `Slide.Export`.

    Accepts `"png"`/`"jpg"`/`"jpeg"`/`"gif"`/`"bmp"`/`"tif"`/`"tiff"` (case-
    insensitive, a leading dot tolerated). Raises `ValueError` for an unknown
    format — symmetric with `autoshape_type_for` / `alignment_for`.
    """
    key = str(fmt).strip().lower().lstrip(".")
    found = _IMAGE_FILTERS.get(key)
    if found is None:
        choices = ", ".join(IMAGE_FORMAT_CHOICES)
        raise ValueError(f"unknown image format {fmt!r}; expected one of: {choices}")
    return found


# ---------------------------------------------------------------------------
# Live slide show (v0.6): run state + slide range
# ---------------------------------------------------------------------------


class PpSlideShowState(IntEnum):
    """`SlideShowView.State` — what a running slide show is currently doing.

    `BLACK_SCREEN`/`WHITE_SCREEN` are the presenter "blank the screen" states
    (the B / W keys); setting `State` back to `RUNNING` resumes. pptlive reports
    `DONE` for a deck with no running show (the show window is gone), so the
    `state` read has a value to return without raising.
    """

    RUNNING = 1
    PAUSED = 2
    BLACK_SCREEN = 3
    WHITE_SCREEN = 4
    DONE = 5


_SLIDE_SHOW_STATE_NAMES: dict[int, str] = {
    int(PpSlideShowState.RUNNING): "running",
    int(PpSlideShowState.PAUSED): "paused",
    int(PpSlideShowState.BLACK_SCREEN): "black",
    int(PpSlideShowState.WHITE_SCREEN): "white",
    int(PpSlideShowState.DONE): "done",
}


def slide_show_state_name(value: Any) -> str:
    """Friendly name for a `SlideShowView.State` int (e.g. 3 -> "black")."""
    try:
        return _SLIDE_SHOW_STATE_NAMES.get(int(value), f"state:{int(value)}")
    except (TypeError, ValueError):
        return "unknown"


class PpSlideShowRangeType(IntEnum):
    """`SlideShowSettings.RangeType` — which slides a show runs.

    pptlive sets `SLIDE_RANGE` only to honor `show.start(from_slide=...)`; the
    default `ALL` runs the whole deck.
    """

    ALL = 1
    SLIDE_RANGE = 2
    NAMED_SLIDE_SHOW = 3


# ---------------------------------------------------------------------------
# Per-shape image export (v0.7): Shape.Export filter formats
# ---------------------------------------------------------------------------


class PpShapeFormat(IntEnum):
    """`Shape.Export` filter — the per-shape image-format enum.

    Distinct from `Slide.Export`, whose `FilterName` is a *string* ("PNG"): a
    *shape* export takes this int enum instead. pptlive exposes only the common
    raster set; the vector types (WMF/EMF) stay reachable via the `.com` escape
    hatch.
    """

    GIF = 0
    JPG = 1
    PNG = 2
    BMP = 3


# Friendly format token -> (PpShapeFormat int, file extension), in match order.
_SHAPE_IMAGE_FILTERS: dict[str, tuple[int, str]] = {
    "png": (int(PpShapeFormat.PNG), "png"),
    "jpg": (int(PpShapeFormat.JPG), "jpg"),
    "jpeg": (int(PpShapeFormat.JPG), "jpg"),
    "gif": (int(PpShapeFormat.GIF), "gif"),
    "bmp": (int(PpShapeFormat.BMP), "bmp"),
}

SHAPE_IMAGE_FORMAT_CHOICES: tuple[str, ...] = ("png", "jpg", "gif", "bmp")


def shape_image_filter_for(fmt: str) -> tuple[int, str]:
    """Resolve an image-format token to its `(PpShapeFormat, extension)` for `Shape.Export`.

    Accepts `"png"`/`"jpg"`/`"jpeg"`/`"gif"`/`"bmp"` (case-insensitive, a
    leading dot tolerated). Raises `ValueError` for an unknown format —
    symmetric with `image_filter_for` (the `Slide.Export` resolver). Note the
    raster set is narrower than `Slide.Export`'s (no TIFF — `PpShapeFormat` has
    no TIFF member).
    """
    key = str(fmt).strip().lower().lstrip(".")
    found = _SHAPE_IMAGE_FILTERS.get(key)
    if found is None:
        choices = ", ".join(SHAPE_IMAGE_FORMAT_CHOICES)
        raise ValueError(f"unknown image format {fmt!r}; expected one of: {choices}")
    return found


# ---------------------------------------------------------------------------
# Save / export (v1.1): Presentation.SaveAs file-format enum
# ---------------------------------------------------------------------------


class PpSaveAsFileType(IntEnum):
    """`Presentation.SaveAs(FileFormat=...)` values the save/export verbs expose.

    A deliberately narrow slice of PowerPoint's full `PpSaveAsFileType`: the
    modern Open XML `.pptx` (`OPEN_XML_PRESENTATION`, what `save_as(fmt="pptx")`
    writes) and `PDF` (what `export_pdf` writes). The 2026-06-09 spike found
    `Presentation.ExportAsFixedFormat` won't marshal under pptlive's late-bound
    dispatch (a trailing object-typed param raises `TypeError`), but
    `SaveAs(path, ppSaveAsPDF=32)` produces a faithful PDF *without* rebinding the
    working file or touching its dirty flag — so PDF export rides `SaveAs` too.
    Legacy `.ppt`, image, and slide-show formats are deferred until a use case
    needs them (the wordlive "add only as needed" rule).
    """

    OPEN_XML_PRESENTATION = 24
    PDF = 32


#: Friendly save-format token -> (PpSaveAsFileType int, file extension), match order.
_SAVE_FILE_FORMATS: dict[str, tuple[int, str]] = {
    "pptx": (int(PpSaveAsFileType.OPEN_XML_PRESENTATION), "pptx"),
    "pdf": (int(PpSaveAsFileType.PDF), "pdf"),
}

SAVE_FORMAT_CHOICES: tuple[str, ...] = ("pptx",)


def save_format_for(fmt: str) -> tuple[int, str]:
    """Resolve a `save_as` format token to its `(PpSaveAsFileType, extension)`.

    Accepts `"pptx"` (case-insensitive, a leading dot tolerated). `"pdf"` is
    rejected with a pointer to `export_pdf` — PDF goes through the same `SaveAs`
    COM call but is a *read* (it neither rebinds the working file nor clears the
    dirty flag), so it's a separate verb. Raises `ValueError` for anything else —
    symmetric with `image_filter_for` / `shape_image_filter_for`.
    """
    key = str(fmt).strip().lower().lstrip(".")
    if key == "pdf":
        raise ValueError("save_as does not write PDF; use export_pdf(path) instead")
    found = _SAVE_FILE_FORMATS.get(key)
    if found is None:
        choices = ", ".join(SAVE_FORMAT_CHOICES)
        raise ValueError(
            f"unsupported save format {fmt!r}; supported: {choices} (PDF via export_pdf)"
        )
    return found


# ---------------------------------------------------------------------------
# Charts (v0.7): XlChartType (the chart kind passed to Shapes.AddChart2)
# ---------------------------------------------------------------------------


class XlChartType(IntEnum):
    """`Shapes.AddChart2` / `Chart.ChartType` — the chart kind.

    A small, common subset of Excel's `XlChartType` (the values are shared with
    PowerPoint's chart object model). Added only as needed (the wordlive rule);
    reach for the `.com` escape hatch + a raw int for anything exotic. Note the
    negative members are how Office encodes these specific constants.
    """

    COLUMN_CLUSTERED = 51
    COLUMN_STACKED = 52
    BAR_CLUSTERED = 57
    BAR_STACKED = 58
    LINE = 4
    LINE_MARKERS = 65
    PIE = 5
    DOUGHNUT = -4120
    AREA = 1
    AREA_STACKED = 76
    XY_SCATTER = -4169
    RADAR = -4151


class XlAxisType(IntEnum):
    """`Chart.Axes(type)` — the two axes whose tick labels carry text.

    Excel's `XlAxisType`, shared with PowerPoint's chart object model. Only the
    category and value axes are surfaced (the ones `recolor_text` walks); the
    series axis (3-D charts) isn't needed yet.
    """

    CATEGORY = 1
    VALUE = 2


# Friendly token -> XlChartType int. Short aliases ("column", "bar", "scatter")
# map to the clustered/standard variant; explicit names resolve to themselves.
_CHART_TYPES: dict[str, int] = {
    "column": int(XlChartType.COLUMN_CLUSTERED),
    "column_clustered": int(XlChartType.COLUMN_CLUSTERED),
    "column_stacked": int(XlChartType.COLUMN_STACKED),
    "bar": int(XlChartType.BAR_CLUSTERED),
    "bar_clustered": int(XlChartType.BAR_CLUSTERED),
    "bar_stacked": int(XlChartType.BAR_STACKED),
    "line": int(XlChartType.LINE),
    "line_markers": int(XlChartType.LINE_MARKERS),
    "pie": int(XlChartType.PIE),
    "doughnut": int(XlChartType.DOUGHNUT),
    "area": int(XlChartType.AREA),
    "area_stacked": int(XlChartType.AREA_STACKED),
    "scatter": int(XlChartType.XY_SCATTER),
    "xy_scatter": int(XlChartType.XY_SCATTER),
    "radar": int(XlChartType.RADAR),
}

# The friendly names offered as a CLI choice (canonical spellings, deduped/ordered).
CHART_TYPE_CHOICES: tuple[str, ...] = (
    "column",
    "column_stacked",
    "bar",
    "bar_stacked",
    "line",
    "line_markers",
    "pie",
    "doughnut",
    "area",
    "area_stacked",
    "scatter",
    "radar",
)

# Reverse map (int -> a canonical friendly name) for read-backs.
_CHART_TYPE_NAMES: dict[int, str] = {
    int(XlChartType.COLUMN_CLUSTERED): "column_clustered",
    int(XlChartType.COLUMN_STACKED): "column_stacked",
    int(XlChartType.BAR_CLUSTERED): "bar_clustered",
    int(XlChartType.BAR_STACKED): "bar_stacked",
    int(XlChartType.LINE): "line",
    int(XlChartType.LINE_MARKERS): "line_markers",
    int(XlChartType.PIE): "pie",
    int(XlChartType.DOUGHNUT): "doughnut",
    int(XlChartType.AREA): "area",
    int(XlChartType.AREA_STACKED): "area_stacked",
    int(XlChartType.XY_SCATTER): "xy_scatter",
    int(XlChartType.RADAR): "radar",
}


def chart_type_for(chart_type: str | int) -> int:
    """Resolve a friendly chart-type name (or raw int) to its `XlChartType` int.

    Accepts `"column"`/`"bar"`/`"line"`/`"pie"`/… (case- and separator-
    insensitive: "Line Markers" -> line_markers) or a raw int (passed through, so
    exotic `XlChartType` values still work). Raises `ValueError` for an unknown
    name — symmetric with `autoshape_type_for`.
    """
    if isinstance(chart_type, bool):  # guard: bool is an int subclass
        raise ValueError(f"invalid chart type: {chart_type!r}")
    if isinstance(chart_type, int):
        return int(chart_type)
    key = str(chart_type).strip().lower().replace(" ", "_").replace("-", "_")
    found = _CHART_TYPES.get(key)
    if found is None:
        choices = ", ".join(CHART_TYPE_CHOICES)
        raise ValueError(f"unknown chart type {chart_type!r}; expected one of: {choices}")
    return found


def chart_type_name(value: Any) -> str:
    """Friendly name for an `XlChartType` int (e.g. 51 -> "column_clustered")."""
    try:
        return _CHART_TYPE_NAMES.get(int(value), f"type:{int(value)}")
    except (TypeError, ValueError):
        return "unknown"


# --- SmartArt (v0.8) -------------------------------------------------------
#
# A SmartArt diagram is added from a `SmartArtLayout` pulled from
# `Application.SmartArtLayouts`. The collection *index* drifts between installs,
# but each layout's `.Id` is a stable URN (".../officeart/2005/8/layout/<seg>"),
# so we key friendly names to the trailing URN segment and resolve the live
# layout object by matching `Id.endswith("/" + seg)` (verified in
# scripts/smartart_spike.py: 159 layouts installed, all 7 cores resolved).


class MsoSmartArtNodePosition(IntEnum):
    """`SmartArtNode.AddNode(Position, Type)` — where to add a node.

    The one that matters is `BELOW` (add a *child*): plain
    `SmartArtNodes.Add()` adds a *sibling*, so child nesting must go through
    `AddNode(BELOW, ...)` (verified live).
    """

    DEFAULT = 1
    AFTER = 2
    BEFORE = 3
    ABOVE = 4
    BELOW = 5


# Friendly name -> the trailing segment of the layout's URN `.Id`. The 7 core
# layouts; widen on demand (the wordlive "add only as needed" rule).
_SMARTART_LAYOUTS: dict[str, str] = {
    "list": "list1",
    "process": "process1",
    "cycle": "cycle1",
    "hierarchy": "hierarchy1",
    "orgchart": "orgChart1",
    "org_chart": "orgChart1",
    "pyramid": "pyramid1",
    "venn": "venn1",
}

# The friendly names offered as a CLI choice (canonical spellings, ordered).
SMARTART_CHOICES: tuple[str, ...] = (
    "list",
    "process",
    "cycle",
    "hierarchy",
    "orgchart",
    "pyramid",
    "venn",
)

# Reverse map (URN segment -> a canonical friendly name) for read-backs.
_SMARTART_NAMES: dict[str, str] = {
    "list1": "list",
    "process1": "process",
    "cycle1": "cycle",
    "hierarchy1": "hierarchy",
    "orgChart1": "orgchart",
    "pyramid1": "pyramid",
    "venn1": "venn",
}

# The layouts whose nodes form a tree (one root + children) rather than a flat
# list — `Nodes.Add()` is a no-op at their top level, so set_nodes builds them
# as a single root with `AddNode` children (verified live).
SMARTART_TREE_KINDS: frozenset[str] = frozenset({"hierarchy", "orgchart"})


def smartart_layout_for(kind: str) -> str:
    """Resolve a friendly SmartArt name to its layout URN segment.

    Accepts `"process"`/`"cycle"`/`"orgchart"`/… (case- and separator-
    insensitive). Raises `ValueError` for an unknown name (before any COM) —
    symmetric with `chart_type_for`. Unlike charts there is no raw-int form: a
    layout is a COM object, not an int, so the wrapper resolves the segment
    against `Application.SmartArtLayouts` live.
    """
    key = str(kind).strip().lower().replace(" ", "_").replace("-", "_")
    seg = _SMARTART_LAYOUTS.get(key)
    if seg is None:
        choices = ", ".join(SMARTART_CHOICES)
        raise ValueError(f"unknown SmartArt layout {kind!r}; expected one of: {choices}")
    return seg


def smartart_layout_name(urn: Any) -> str:
    """Friendly name for a SmartArt layout `.Id` URN (e.g. ".../process1" -> "process").

    Falls back to the trailing URN segment (then the raw value) when the layout
    isn't one of the known cores, so a read-back never raises.
    """
    text = str(urn or "")
    seg = text.rsplit("/", 1)[-1] if text else text
    return _SMARTART_NAMES.get(seg, seg or "unknown")


# ---------------------------------------------------------------------------
# Master / theme styling (v0.9): deck-wide text styles, palette, fonts
# ---------------------------------------------------------------------------
#
# The deck-wide counterpart to v0.3's per-run format_text. Feasibility confirmed
# live 2026-05-28 (write+restore round-trips). Three small enums + the usual
# friendly-name resolvers (mirroring smartart_layout_for / chart_type_for), added
# only as v0.9 needs them.


class PpTextStyleType(IntEnum):
    """`SlideMaster.TextStyles(type)` — the master's three named text styles.

    PowerPoint's nearest analog to Word's named paragraph styles: each style has
    5 outline `Levels`, and editing one re-renders every slide that inherits it.
    """

    DEFAULT = 1
    TITLE = 2
    BODY = 3


# Friendly name -> PpTextStyleType int.
_TEXT_STYLES: dict[str, int] = {
    "default": int(PpTextStyleType.DEFAULT),
    "title": int(PpTextStyleType.TITLE),
    "body": int(PpTextStyleType.BODY),
}

# The friendly names offered as a CLI choice (canonical spellings, ordered).
TEXT_STYLE_CHOICES: tuple[str, ...] = ("title", "body", "default")

_TEXT_STYLE_NAMES: dict[int, str] = {v: k for k, v in _TEXT_STYLES.items()}


def text_style_for(style: str) -> int:
    """Resolve a friendly master text-style name to its `PpTextStyleType` int.

    Accepts `"title"`/`"body"`/`"default"` (case-insensitive). Raises
    `ValueError` for an unknown name (before any COM) — symmetric with
    `smartart_layout_for`.
    """
    key = str(style).strip().lower()
    val = _TEXT_STYLES.get(key)
    if val is None:
        choices = ", ".join(TEXT_STYLE_CHOICES)
        raise ValueError(f"unknown text style {style!r}; expected one of: {choices}")
    return val


def text_style_name(value: Any) -> str:
    """Friendly name for a `PpTextStyleType` int (e.g. 3 -> "body")."""
    try:
        return _TEXT_STYLE_NAMES.get(int(value), f"style:{int(value)}")
    except (TypeError, ValueError):
        return "default"


class MsoThemeColorSchemeIndex(IntEnum):
    """`Theme.ThemeColorScheme.Colors(index)` — the 12 theme palette slots.

    The slot ints PowerPoint uses; `.RGB` on each is the same R-low-byte long as
    `Font.Color.RGB` (so `parse_color` / `color_hex` apply unchanged).
    """

    DARK1 = 1
    LIGHT1 = 2
    DARK2 = 3
    LIGHT2 = 4
    ACCENT1 = 5
    ACCENT2 = 6
    ACCENT3 = 7
    ACCENT4 = 8
    ACCENT5 = 9
    ACCENT6 = 10
    HYPERLINK = 11
    FOLLOWED_HYPERLINK = 12


# Friendly slot name -> palette index. Includes a couple of common aliases.
_THEME_COLORS: dict[str, int] = {
    "dark1": int(MsoThemeColorSchemeIndex.DARK1),
    "light1": int(MsoThemeColorSchemeIndex.LIGHT1),
    "dark2": int(MsoThemeColorSchemeIndex.DARK2),
    "light2": int(MsoThemeColorSchemeIndex.LIGHT2),
    "accent1": int(MsoThemeColorSchemeIndex.ACCENT1),
    "accent2": int(MsoThemeColorSchemeIndex.ACCENT2),
    "accent3": int(MsoThemeColorSchemeIndex.ACCENT3),
    "accent4": int(MsoThemeColorSchemeIndex.ACCENT4),
    "accent5": int(MsoThemeColorSchemeIndex.ACCENT5),
    "accent6": int(MsoThemeColorSchemeIndex.ACCENT6),
    "hyperlink": int(MsoThemeColorSchemeIndex.HYPERLINK),
    "hlink": int(MsoThemeColorSchemeIndex.HYPERLINK),
    "followed_hyperlink": int(MsoThemeColorSchemeIndex.FOLLOWED_HYPERLINK),
    "folhlink": int(MsoThemeColorSchemeIndex.FOLLOWED_HYPERLINK),
}

# The canonical slot names, in palette order — used as CLI choices *and* as the
# ordered key set when reading the whole scheme back.
THEME_COLOR_CHOICES: tuple[str, ...] = (
    "dark1",
    "light1",
    "dark2",
    "light2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hyperlink",
    "followed_hyperlink",
)


def theme_color_for(slot: str) -> int:
    """Resolve a friendly theme-color slot name to its palette index (1-12).

    Accepts `"accent1"`/`"dark1"`/`"hyperlink"`/… (case- and separator-
    insensitive; `"hlink"`/`"folhlink"` aliases too). Raises `ValueError` for an
    unknown name (before any COM).
    """
    key = str(slot).strip().lower().replace(" ", "").replace("-", "")
    # Match against keys with their own separators stripped, so "accent 6",
    # "accent6", and "followed-hyperlink" all resolve.
    idx = _THEME_COLORS.get(key) or {k.replace("_", ""): v for k, v in _THEME_COLORS.items()}.get(
        key
    )
    if idx is None:
        choices = ", ".join(THEME_COLOR_CHOICES)
        raise ValueError(f"unknown theme color slot {slot!r}; expected one of: {choices}")
    return idx


# Theme font scheme: the two typeface roles and the per-script sub-index.
# `Theme.ThemeFontScheme.MajorFont`/`MinorFont` are accessed by
# `.Item(1=Latin / 2=EastAsian / 3=ComplexScript).Name` — the late-bound `.Latin`
# accessor raises AttributeError, so `.Item(n)` is the only reliable path.
THEME_FONT_SLOTS: tuple[str, ...] = ("major", "minor")

_THEME_FONT_SCRIPTS: dict[str, int] = {
    "latin": 1,
    "east_asian": 2,
    "complex_script": 3,
}

THEME_FONT_SCRIPT_CHOICES: tuple[str, ...] = ("latin", "east_asian", "complex_script")


def theme_font_slot_for(which: str) -> str:
    """Normalize the typeface role to `"major"` or `"minor"`.

    `"major"` is the headings font, `"minor"` the body font; `"heading"`/`"body"`
    are accepted aliases. Raises `ValueError` for anything else (before any COM).
    """
    key = str(which).strip().lower()
    if key in ("major", "heading", "headings"):
        return "major"
    if key in ("minor", "body"):
        return "minor"
    choices = ", ".join(THEME_FONT_SLOTS)
    raise ValueError(f"unknown theme font {which!r}; expected one of: {choices}")


def theme_font_script_for(script: str) -> int:
    """Resolve a font script name to its `.Item(n)` index (latin=1/…).

    Raises `ValueError` for an unknown name (before any COM).
    """
    key = str(script).strip().lower().replace(" ", "_").replace("-", "_")
    idx = _THEME_FONT_SCRIPTS.get(key)
    if idx is None:
        choices = ", ".join(THEME_FONT_SCRIPT_CHOICES)
        raise ValueError(f"unknown font script {script!r}; expected one of: {choices}")
    return idx


# ---------------------------------------------------------------------------
# Navigation — hyperlinks / actions (v0.4.0, the v1.4 cut)
# ---------------------------------------------------------------------------
#
# A shape-level hyperlink lives on `Shape.ActionSettings(ppMouseClick).Hyperlink`
# (`.Address` for a URL/file, `.SubAddress` for a slide jump). Spike findings
# (scripts/hyperlink_spike.py): setting `.Address` auto-flips `.Action` to
# `ppActionHyperlink`; `Hyperlink.Delete()` reverts `.Action` to `ppActionNone`
# and `.Address` to `""`; the slide-jump `SubAddress` form is `"<SlideID>,<index>,<title>"`.


class PpMouseActivation(IntEnum):
    """`Shape.ActionSettings(activation)` — which mouse event the action fires on.

    Only `MOUSE_CLICK` is used (the common "click to follow the link"); `MOUSE_OVER`
    is named for the `.com` escape hatch but not wired into a verb yet.
    """

    MOUSE_CLICK = 1
    MOUSE_OVER = 2


class PpActionType(IntEnum):
    """`ActionSetting.Action` — what the click does.

    Only the two pptlive sets/reads: `NONE` (no action — what `Hyperlink.Delete()`
    leaves behind) and `HYPERLINK` (follow `.Hyperlink.Address`/`.SubAddress`, which
    PowerPoint sets implicitly when an address is assigned). Widen on demand.
    """

    NONE = 0
    HYPERLINK = 7


# ---------------------------------------------------------------------------
# Motion — slide transitions (v0.4.0, the v1.5 cut)
# ---------------------------------------------------------------------------
#
# `Slide.SlideShowTransition.EntryEffect` takes a `PpEntryEffect` int. The enum is
# huge; we expose a **curated** subset of the common, well-documented families and
# fall back to a raw-int passthrough for anything exotic (the chart_type_for rule).
# Every value below is round-trip-verified on a live build (scripts/transition_spike.py)
# — PowerPoint *validates* EntryEffect (the "wipe" family 3329-3332 is rejected as
# "not valid for transitions"), so only verified families are named here.


class PpEntryEffect(IntEnum):
    """`Slide.SlideShowTransition.EntryEffect` — the slide's entrance transition.

    A curated subset of the documented `PpEntryEffect` enum (the families this
    build accepts): cut, blinds, checkerboard, cover, dissolve, fade, uncover.
    `NONE` is no transition. Pass a raw int to reach any value pptlive hasn't named.
    """

    NONE = 0
    CUT = 257
    CUT_THROUGH_BLACK = 258
    RANDOM = 513
    BLINDS_HORIZONTAL = 769
    BLINDS_VERTICAL = 770
    CHECKERBOARD_ACROSS = 1025
    CHECKERBOARD_DOWN = 1026
    COVER_LEFT = 1281
    COVER_UP = 1282
    COVER_RIGHT = 1283
    COVER_DOWN = 1284
    DISSOLVE = 1537
    FADE = 1793
    UNCOVER_LEFT = 2049
    UNCOVER_UP = 2050
    UNCOVER_RIGHT = 2051
    UNCOVER_DOWN = 2052


# Friendly token -> PpEntryEffect int. Short aliases ("blinds", "checkerboard",
# "cover", "uncover") map to a sensible default direction; explicit names resolve
# to themselves.
_ENTRY_EFFECTS: dict[str, int] = {
    "none": int(PpEntryEffect.NONE),
    "cut": int(PpEntryEffect.CUT),
    "cut_through_black": int(PpEntryEffect.CUT_THROUGH_BLACK),
    "random": int(PpEntryEffect.RANDOM),
    "blinds": int(PpEntryEffect.BLINDS_HORIZONTAL),
    "blinds_horizontal": int(PpEntryEffect.BLINDS_HORIZONTAL),
    "blinds_vertical": int(PpEntryEffect.BLINDS_VERTICAL),
    "checkerboard": int(PpEntryEffect.CHECKERBOARD_ACROSS),
    "checkerboard_across": int(PpEntryEffect.CHECKERBOARD_ACROSS),
    "checkerboard_down": int(PpEntryEffect.CHECKERBOARD_DOWN),
    "cover": int(PpEntryEffect.COVER_LEFT),
    "cover_left": int(PpEntryEffect.COVER_LEFT),
    "cover_up": int(PpEntryEffect.COVER_UP),
    "cover_right": int(PpEntryEffect.COVER_RIGHT),
    "cover_down": int(PpEntryEffect.COVER_DOWN),
    "dissolve": int(PpEntryEffect.DISSOLVE),
    "fade": int(PpEntryEffect.FADE),
    "uncover": int(PpEntryEffect.UNCOVER_LEFT),
    "uncover_left": int(PpEntryEffect.UNCOVER_LEFT),
    "uncover_up": int(PpEntryEffect.UNCOVER_UP),
    "uncover_right": int(PpEntryEffect.UNCOVER_RIGHT),
    "uncover_down": int(PpEntryEffect.UNCOVER_DOWN),
}

# The friendly names offered as a CLI choice (canonical spellings, ordered).
ENTRY_EFFECT_CHOICES: tuple[str, ...] = (
    "none",
    "fade",
    "cut",
    "dissolve",
    "random",
    "blinds_horizontal",
    "blinds_vertical",
    "checkerboard_across",
    "checkerboard_down",
    "cover_left",
    "cover_up",
    "cover_right",
    "cover_down",
    "uncover_left",
    "uncover_up",
    "uncover_right",
    "uncover_down",
)

# Reverse map (int -> a canonical friendly name) for read-backs.
_ENTRY_EFFECT_NAMES: dict[int, str] = {
    int(PpEntryEffect.NONE): "none",
    int(PpEntryEffect.CUT): "cut",
    int(PpEntryEffect.CUT_THROUGH_BLACK): "cut_through_black",
    int(PpEntryEffect.RANDOM): "random",
    int(PpEntryEffect.BLINDS_HORIZONTAL): "blinds_horizontal",
    int(PpEntryEffect.BLINDS_VERTICAL): "blinds_vertical",
    int(PpEntryEffect.CHECKERBOARD_ACROSS): "checkerboard_across",
    int(PpEntryEffect.CHECKERBOARD_DOWN): "checkerboard_down",
    int(PpEntryEffect.COVER_LEFT): "cover_left",
    int(PpEntryEffect.COVER_UP): "cover_up",
    int(PpEntryEffect.COVER_RIGHT): "cover_right",
    int(PpEntryEffect.COVER_DOWN): "cover_down",
    int(PpEntryEffect.DISSOLVE): "dissolve",
    int(PpEntryEffect.FADE): "fade",
    int(PpEntryEffect.UNCOVER_LEFT): "uncover_left",
    int(PpEntryEffect.UNCOVER_UP): "uncover_up",
    int(PpEntryEffect.UNCOVER_RIGHT): "uncover_right",
    int(PpEntryEffect.UNCOVER_DOWN): "uncover_down",
}


def entry_effect_for(effect: str | int) -> int:
    """Resolve a friendly transition name (or raw int) to its `PpEntryEffect` int.

    Accepts `"fade"`/`"cut"`/`"cover_left"`/… (case- and separator-insensitive) or
    a raw int (passed through, so exotic `PpEntryEffect` values still work). Raises
    `ValueError` for an unknown name — symmetric with `chart_type_for`.
    """
    if isinstance(effect, bool):  # guard: bool is an int subclass
        raise ValueError(f"invalid transition effect: {effect!r}")
    if isinstance(effect, int):
        return int(effect)
    key = str(effect).strip().lower().replace(" ", "_").replace("-", "_")
    found = _ENTRY_EFFECTS.get(key)
    if found is None:
        choices = ", ".join(ENTRY_EFFECT_CHOICES)
        raise ValueError(f"unknown transition effect {effect!r}; expected one of: {choices}")
    return found


def entry_effect_name(value: Any) -> str:
    """Friendly name for a `PpEntryEffect` int (e.g. 1793 -> "fade")."""
    try:
        return _ENTRY_EFFECT_NAMES.get(int(value), f"effect:{int(value)}")
    except (TypeError, ValueError):
        return "unknown"

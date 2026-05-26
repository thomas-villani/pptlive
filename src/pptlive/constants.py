"""Typed enums for the PowerPoint magic constants pptlive uses.

Values mirror the official `Mso*` / `Pp*` enumerations exactly. Resist the urge
to pre-populate — add entries only as a feature needs them (the wordlive rule).
Friendly string aliases (`"title"`, `"textbox"`) coerce to the right int the way
wordlive's alignment names do.
"""

from __future__ import annotations

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

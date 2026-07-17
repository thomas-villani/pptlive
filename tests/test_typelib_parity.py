"""Pin every hand-transcribed constant against the live Office type libraries.

`constants.py` is 36 hand-transcribed `IntEnum`s, and the fake COM mirrors whatever
they say — so a wrong value is invisible to the unit suite *and* to a round-trip
spike (the wrong id round-trips perfectly). That blind spot shipped four real bugs:

* `MsoAlignCmd` 1-based instead of 0-based — `align "left"` aligned centers,
  `"bottom"` sent an out-of-range 6.
* `MsoAnimEffect` — 7 of 10 curated effects wrong (`"zoom"` sent 31, which is
  PowerPoint's `GrowAndTurn`).
* `MsoAutoSize` — `SHAPE_TO_FIT_TEXT`/`TEXT_TO_FIT_SHAPE` swapped, so
  `text_frame_status` reported the opposite autofit mode.
* `PpViewType.NOTES_PAGE` / `PpSlideLayout.TWO_COLUMN_TEXT` — latent, pointing at
  `ppViewNotesMaster` / `ppLayoutTable`.

This reads the **type library**, not a running app, so it needs Office installed but
no PowerPoint window and no deck — seconds, not minutes. Marked `smoke` only because
of the Office dependency.

Two things worth knowing before editing the map below:

1. **The name is not derivable from ours.** `MsoZOrderCmd.SEND_TO_BACK` is Office's
   `msoSendToBack`, *not* `msoZOrderSendToBack` (which doesn't exist). Hence an
   explicit table rather than a clever prefix rule.
2. **The typelib matters.** `msoAnimEffect*` lives in the **PowerPoint** typelib
   despite the `mso` prefix; looking it up in the Office one silently finds nothing.
   All three are loaded and merged here, mirroring how `constants` resolves.

All 36 enums are mapped (a 2026-07-17 audit promoted the last 17 "curated subset"
enums into `EXPECTED` after confirming every member — mapping a partial enum
asserts only that the members we transcribed are right, not that we have all of
them). `EXPECT_ABSENT` is the escape hatch for a future genuinely-unmappable enum
(a synthetic or derived value) and is empty for now; a member with no Office
counterpart opts out by omission, and the companion test makes sure "unmapped and
unexcused" can never silently read as "checked".
"""

from __future__ import annotations

from typing import Any

import pytest

from pptlive import constants as K

pytestmark = pytest.mark.smoke

#: (GUID, major, minor) for the three type libraries our constants come from.
TYPELIBS = [
    ("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", 2, 8),  # Office  — Mso*
    ("{91493440-5A91-11CF-8700-00AA0060263B}", 2, 12),  # PowerPoint — Pp*, msoAnim*
    ("{00020813-0000-0000-C000-000000000046}", 1, 9),  # Excel — Xl*
]

#: our enum member -> the Office/PowerPoint/Excel constant it must equal.
#: Deliberately explicit: the Office name is not mechanically derivable from ours.
EXPECTED: dict[str, dict[str, str]] = {
    "MsoAlignCmd": {
        "LEFTS": "msoAlignLefts",
        "CENTERS": "msoAlignCenters",
        "RIGHTS": "msoAlignRights",
        "TOPS": "msoAlignTops",
        "MIDDLES": "msoAlignMiddles",
        "BOTTOMS": "msoAlignBottoms",
    },
    "MsoDistributeCmd": {
        "HORIZONTALLY": "msoDistributeHorizontally",
        "VERTICALLY": "msoDistributeVertically",
    },
    "MsoConnectorType": {
        "STRAIGHT": "msoConnectorStraight",
        "ELBOW": "msoConnectorElbow",
        "CURVED": "msoConnectorCurve",
    },
    "MsoZOrderCmd": {
        "BRING_TO_FRONT": "msoBringToFront",
        "SEND_TO_BACK": "msoSendToBack",
        "BRING_FORWARD": "msoBringForward",
        "SEND_BACKWARD": "msoSendBackward",
    },
    "MsoAutoSize": {
        "MIXED": "msoAutoSizeMixed",
        "NONE": "msoAutoSizeNone",
        "SHAPE_TO_FIT_TEXT": "msoAutoSizeShapeToFitText",
        "TEXT_TO_FIT_SHAPE": "msoAutoSizeTextToFitShape",
    },
    "MsoAnimEffect": {
        "APPEAR": "msoAnimEffectAppear",
        "FLY_IN": "msoAnimEffectFly",
        "FADE": "msoAnimEffectFade",
        "SPLIT": "msoAnimEffectSplit",
        "SWIVEL": "msoAnimEffectSwivel",
        "WHEEL": "msoAnimEffectWheel",
        "WIPE": "msoAnimEffectWipe",
        "ZOOM": "msoAnimEffectZoom",
        "FLOAT_IN": "msoAnimEffectFloat",
        "GROW_TURN": "msoAnimEffectGrowAndTurn",
    },
    "MsoAnimTriggerType": {
        "ON_CLICK": "msoAnimTriggerOnPageClick",
        "WITH_PREVIOUS": "msoAnimTriggerWithPrevious",
        "AFTER_PREVIOUS": "msoAnimTriggerAfterPrevious",
    },
    "MsoTriState": {
        "TRUE": "msoTrue",
        "FALSE": "msoFalse",
        "MIXED": "msoTriStateMixed",
        "TOGGLE": "msoTriStateToggle",
    },
    "MsoColorType": {
        "MIXED": "msoColorTypeMixed",
        "RGB": "msoColorTypeRGB",
        "SCHEME": "msoColorTypeScheme",
        "CMYK": "msoColorTypeCMYK",
        "CMS": "msoColorTypeCMS",
        "INK": "msoColorTypeInk",
    },
    "MsoFillType": {
        "MIXED": "msoFillMixed",
        "SOLID": "msoFillSolid",
        "PATTERNED": "msoFillPatterned",
        "GRADIENT": "msoFillGradient",
        "TEXTURED": "msoFillTextured",
        "BACKGROUND": "msoFillBackground",
        "PICTURE": "msoFillPicture",
    },
    "MsoTextOrientation": {
        "HORIZONTAL": "msoTextOrientationHorizontal",
        "VERTICAL": "msoTextOrientationVertical",
    },
    "MsoTextUnderlineType": {
        "NONE": "msoNoUnderline",
        "SINGLE_LINE": "msoUnderlineSingleLine",
    },
    "PpViewType": {
        "NORMAL": "ppViewNormal",
        "SLIDE": "ppViewSlide",
        "OUTLINE": "ppViewOutline",
        "NOTES_PAGE": "ppViewNotesPage",
        "SLIDE_SORTER": "ppViewSlideSorter",
    },
    "PpSlideLayout": {
        "TITLE": "ppLayoutTitle",
        "TEXT": "ppLayoutText",
        "TWO_COLUMN_TEXT": "ppLayoutTwoColumnText",
        "TITLE_ONLY": "ppLayoutTitleOnly",
        "BLANK": "ppLayoutBlank",
        "OBJECT": "ppLayoutObject",
    },
    "PpSaveAsFileType": {
        "OPEN_XML_PRESENTATION": "ppSaveAsOpenXMLPresentation",
        "PDF": "ppSaveAsPDF",
    },
    "PpMediaType": {
        "SOUND": "ppMediaTypeSound",
        "MOVIE": "ppMediaTypeMovie",
        "OTHER": "ppMediaTypeOther",
    },
    "PpTextStyleType": {
        "TITLE": "ppTitleStyle",
        "BODY": "ppBodyStyle",
        "DEFAULT": "ppDefaultStyle",
    },
    "PpSelectionType": {
        "NONE": "ppSelectionNone",
        "SLIDES": "ppSelectionSlides",
        "SHAPES": "ppSelectionShapes",
        "TEXT": "ppSelectionText",
    },
    "XlAxisType": {"CATEGORY": "xlCategory", "VALUE": "xlValue"},
    # --- the "curated subset" enums, promoted from EXPECT_ABSENT after a full
    #     audit (2026-07-17). Every one was correct — but mapping a curated subset
    #     asserts only that the members we *did* transcribe are right, not that we
    #     have all of them, so a partial enum is still fully checked here. Names
    #     were machine-generated off the typelib and pasted, not hand-typed.
    "MsoAutoShapeType": {
        "RECTANGLE": "msoShapeRectangle",
        "PARALLELOGRAM": "msoShapeParallelogram",
        "TRAPEZOID": "msoShapeTrapezoid",
        "DIAMOND": "msoShapeDiamond",
        "ROUNDED_RECTANGLE": "msoShapeRoundedRectangle",
        "OCTAGON": "msoShapeOctagon",
        "ISOSCELES_TRIANGLE": "msoShapeIsoscelesTriangle",
        "RIGHT_TRIANGLE": "msoShapeRightTriangle",
        "OVAL": "msoShapeOval",
        "HEXAGON": "msoShapeHexagon",
        "CROSS": "msoShapeCross",
        "REGULAR_PENTAGON": "msoShapeRegularPentagon",
        "HEART": "msoShapeHeart",
        "RIGHT_ARROW": "msoShapeRightArrow",
        "LEFT_ARROW": "msoShapeLeftArrow",
        "UP_ARROW": "msoShapeUpArrow",
        "DOWN_ARROW": "msoShapeDownArrow",
        "FIVE_POINT_STAR": "msoShape5pointStar",
    },
    "MsoGradientStyle": {
        "MIXED": "msoGradientMixed",
        "HORIZONTAL": "msoGradientHorizontal",
        "VERTICAL": "msoGradientVertical",
        "DIAGONAL_UP": "msoGradientDiagonalUp",
        "DIAGONAL_DOWN": "msoGradientDiagonalDown",
        "FROM_CORNER": "msoGradientFromCorner",
        "FROM_TITLE": "msoGradientFromTitle",
        "FROM_CENTER": "msoGradientFromCenter",
    },
    "MsoShadowStyle": {
        "MIXED": "msoShadowStyleMixed",
        "INNER": "msoShadowStyleInnerShadow",
        "OUTER": "msoShadowStyleOuterShadow",
    },
    "MsoShapeType": {
        "AUTO_SHAPE": "msoAutoShape",
        "CALLOUT": "msoCallout",
        "CHART": "msoChart",
        "COMMENT": "msoComment",
        "FREEFORM": "msoFreeform",
        "GROUP": "msoGroup",
        "EMBEDDED_OLE_OBJECT": "msoEmbeddedOLEObject",
        "FORM_CONTROL": "msoFormControl",
        "LINE": "msoLine",
        "LINKED_OLE_OBJECT": "msoLinkedOLEObject",
        "LINKED_PICTURE": "msoLinkedPicture",
        "OLE_CONTROL_OBJECT": "msoOLEControlObject",
        "PICTURE": "msoPicture",
        "PLACEHOLDER": "msoPlaceholder",
        "TEXT_EFFECT": "msoTextEffect",
        "MEDIA": "msoMedia",
        "TEXT_BOX": "msoTextBox",
        "SCRIPT_ANCHOR": "msoScriptAnchor",
        "TABLE": "msoTable",
        "CANVAS": "msoCanvas",
        "DIAGRAM": "msoDiagram",
        "INK": "msoInk",
        "INK_COMMENT": "msoInkComment",
        "SMART_ART": "msoSmartArt",
    },
    "MsoThemeColorSchemeIndex": {
        "DARK1": "msoThemeDark1",
        "LIGHT1": "msoThemeLight1",
        "DARK2": "msoThemeDark2",
        "LIGHT2": "msoThemeLight2",
        "ACCENT1": "msoThemeAccent1",
        "ACCENT2": "msoThemeAccent2",
        "ACCENT3": "msoThemeAccent3",
        "ACCENT4": "msoThemeAccent4",
        "ACCENT5": "msoThemeAccent5",
        "ACCENT6": "msoThemeAccent6",
        "HYPERLINK": "msoThemeHyperlink",
        "FOLLOWED_HYPERLINK": "msoThemeFollowedHyperlink",
    },
    "PpActionType": {
        "NONE": "ppActionNone",
        "HYPERLINK": "ppActionHyperlink",
    },
    "PpBulletType": {
        "NONE": "ppBulletNone",
        "UNNUMBERED": "ppBulletUnnumbered",
        "NUMBERED": "ppBulletNumbered",
        "MIXED": "ppBulletMixed",
    },
    "PpEntryEffect": {
        "NONE": "ppEffectNone",
        "CUT": "ppEffectCut",
        "CUT_THROUGH_BLACK": "ppEffectCutThroughBlack",
        "RANDOM": "ppEffectRandom",
        "BLINDS_HORIZONTAL": "ppEffectBlindsHorizontal",
        "BLINDS_VERTICAL": "ppEffectBlindsVertical",
        "CHECKERBOARD_ACROSS": "ppEffectCheckerboardAcross",
        "CHECKERBOARD_DOWN": "ppEffectCheckerboardDown",
        "COVER_LEFT": "ppEffectCoverLeft",
        "COVER_UP": "ppEffectCoverUp",
        "COVER_RIGHT": "ppEffectCoverRight",
        "COVER_DOWN": "ppEffectCoverDown",
        "DISSOLVE": "ppEffectDissolve",
        "FADE": "ppEffectFade",
        "UNCOVER_LEFT": "ppEffectUncoverLeft",
        "UNCOVER_UP": "ppEffectUncoverUp",
        "UNCOVER_RIGHT": "ppEffectUncoverRight",
        "UNCOVER_DOWN": "ppEffectUncoverDown",
    },
    "PpParagraphAlignment": {
        "LEFT": "ppAlignLeft",
        "CENTER": "ppAlignCenter",
        "RIGHT": "ppAlignRight",
        "JUSTIFY": "ppAlignJustify",
        "DISTRIBUTE": "ppAlignDistribute",
        "THAI_DISTRIBUTE": "ppAlignThaiDistribute",
        "JUSTIFY_LOW": "ppAlignJustifyLow",
    },
    "PpPlaceholderType": {
        "TITLE": "ppPlaceholderTitle",
        "BODY": "ppPlaceholderBody",
        "CENTER_TITLE": "ppPlaceholderCenterTitle",
        "SUBTITLE": "ppPlaceholderSubtitle",
        "VERTICAL_TITLE": "ppPlaceholderVerticalTitle",
        "VERTICAL_BODY": "ppPlaceholderVerticalBody",
        "OBJECT": "ppPlaceholderObject",
        "CHART": "ppPlaceholderChart",
        "BITMAP": "ppPlaceholderBitmap",
        "MEDIA_CLIP": "ppPlaceholderMediaClip",
        "ORG_CHART": "ppPlaceholderOrgChart",
        "TABLE": "ppPlaceholderTable",
        "SLIDE_NUMBER": "ppPlaceholderSlideNumber",
        "HEADER": "ppPlaceholderHeader",
        "FOOTER": "ppPlaceholderFooter",
        "DATE": "ppPlaceholderDate",
        "VERTICAL_OBJECT": "ppPlaceholderVerticalObject",
        "PICTURE": "ppPlaceholderPicture",
    },
    "PpShapeFormat": {
        "GIF": "ppShapeFormatGIF",
        "JPG": "ppShapeFormatJPG",
        "PNG": "ppShapeFormatPNG",
        "BMP": "ppShapeFormatBMP",
    },
    "PpSlideShowState": {
        "RUNNING": "ppSlideShowRunning",
        "PAUSED": "ppSlideShowPaused",
        "BLACK_SCREEN": "ppSlideShowBlackScreen",
        "WHITE_SCREEN": "ppSlideShowWhiteScreen",
        "DONE": "ppSlideShowDone",
    },
    "PpSlideShowRangeType": {
        "ALL": "ppShowAll",
        "SLIDE_RANGE": "ppShowSlideRange",
        "NAMED_SLIDE_SHOW": "ppShowNamedSlideShow",
    },
    "PpMouseActivation": {
        "MOUSE_CLICK": "ppMouseClick",
        "MOUSE_OVER": "ppMouseOver",
    },
    "PpMediaTaskStatus": {
        "NONE": "ppMediaTaskStatusNone",
        "IN_PROGRESS": "ppMediaTaskStatusInProgress",
        "QUEUED": "ppMediaTaskStatusQueued",
        "DONE": "ppMediaTaskStatusDone",
        "FAILED": "ppMediaTaskStatusFailed",
    },
    "MsoSmartArtNodePosition": {
        "DEFAULT": "msoSmartArtNodeDefault",
        "AFTER": "msoSmartArtNodeAfter",
        "BEFORE": "msoSmartArtNodeBefore",
        "ABOVE": "msoSmartArtNodeAbove",
        "BELOW": "msoSmartArtNodeBelow",
    },
    "XlChartType": {
        "COLUMN_CLUSTERED": "xlColumnClustered",
        "COLUMN_STACKED": "xlColumnStacked",
        "BAR_CLUSTERED": "xlBarClustered",
        "BAR_STACKED": "xlBarStacked",
        "LINE": "xlLine",
        "LINE_MARKERS": "xlLineMarkers",
        "PIE": "xlPie",
        "DOUGHNUT": "xlDoughnut",
        "AREA": "xlArea",
        "AREA_STACKED": "xlAreaStacked",
        "XY_SCATTER": "xlXYScatter",
        "RADAR": "xlRadar",
    },
}

#: Enums deliberately NOT mapped, with the reason — so an unmapped enum is a
#: recorded decision rather than an oversight. Empty since the 2026-07-17 audit
#: promoted every curated enum into EXPECTED; keep it here so a future
#: genuinely-unmappable enum (a synthetic/derived one) has a home with a reason.
EXPECT_ABSENT: dict[str, str] = {}


@pytest.fixture(scope="module")
def typelib_constants() -> dict[str, Any]:
    """Every constant from the Office/PowerPoint/Excel typelibs, merged.

    Reads the type libraries only — no Application object, so nothing launches.
    """
    try:
        from win32com.client import constants, gencache
    except ImportError:  # pragma: no cover - non-Windows
        pytest.skip("pywin32 not available")
    loaded = 0
    for guid, major, minor in TYPELIBS:
        try:
            gencache.EnsureModule(guid, 0, major, minor)
            loaded += 1
        except Exception:  # noqa: BLE001 - a missing app (e.g. no Excel) isn't fatal
            continue
    if not loaded:
        pytest.skip("no Office type libraries available")
    merged: dict[str, Any] = {}
    for d in constants.__dicts__:
        merged.update(d)
    if not merged:
        pytest.skip("Office type libraries exposed no constants")
    return merged


@pytest.mark.parametrize("enum_name", sorted(EXPECTED))
def test_enum_matches_the_office_typelib(enum_name: str, typelib_constants: dict[str, Any]) -> None:
    """Every mapped member equals the constant Office/PowerPoint/Excel defines."""
    enum_cls = getattr(K, enum_name)
    mismatches: list[str] = []
    for member_name, office_name in EXPECTED[enum_name].items():
        ours = int(getattr(enum_cls, member_name))
        if office_name not in typelib_constants:
            # A name we expect but the typelib lacks means the map is stale — that's
            # a real failure, not a skip: it's how a silent rename would slip past.
            mismatches.append(f"{enum_name}.{member_name}: {office_name!r} not in any typelib")
            continue
        theirs = int(typelib_constants[office_name])
        if ours != theirs:
            mismatches.append(f"{enum_name}.{member_name} = {ours}, but {office_name} = {theirs}")
    assert not mismatches, "constants disagree with the live Office typelib:\n  " + "\n  ".join(
        mismatches
    )


def test_every_enum_is_either_mapped_or_explicitly_excused() -> None:
    """No enum may be silently unchecked — it's mapped, or excused with a reason.

    Without this, adding a new enum and forgetting to map it would read as "passing"
    (the parametrize just wouldn't cover it) — the same silence that hid the four
    bugs above. Needs no typelib, so it runs anywhere.
    """
    import inspect
    from enum import IntEnum

    declared = {
        name
        for name, obj in vars(K).items()
        if inspect.isclass(obj) and issubclass(obj, IntEnum) and obj is not IntEnum
    }
    accounted = set(EXPECTED) | set(EXPECT_ABSENT)
    unaccounted = declared - accounted
    assert not unaccounted, (
        "these IntEnums are neither pinned to the typelib nor excused in "
        f"EXPECT_ABSENT: {sorted(unaccounted)}"
    )
    stale = set(EXPECT_ABSENT) - declared - set(EXPECTED)
    assert not stale, f"EXPECT_ABSENT names enums that no longer exist: {sorted(stale)}"

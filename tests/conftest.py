"""Test fixtures.

`fake_powerpoint` builds a connected object graph that quacks like a
PowerPoint.Application COM object — Application → Presentations → Slides →
Shapes → TextFrame.TextRange, plus placeholders and a notes page — enough to
exercise the politeness logic, anchors, and CLI shape without a real PowerPoint
install. Modeled on wordlive's `fake_word`, redesigned for the 2-D object model.

`no_powerpoint` simulates PowerPoint not running. `real_powerpoint` is for the
@pytest.mark.smoke suite and skips if PowerPoint can't be reached.

The fake uses plain Python classes (not bare MagicMocks) so writes round-trip
deterministically: setting `TextRange.Text` and reading it back goes through the
same object, the way the real COM property does.
"""

from __future__ import annotations

import os
from datetime import UTC, datetime
from types import SimpleNamespace
from typing import Any

import pytest

# PpPlaceholderType values used below (kept literal so the fake doesn't depend on
# the module under test for its constants).
_PH_TITLE = 1
_PH_BODY = 2
_PH_CENTER_TITLE = 3
_PH_SUBTITLE = 4
_PH_OBJECT = 7  # generic content placeholder (what `body` also matches)

# MsoShapeType values used below.
_MSO_CHART = 3
_MSO_AUTO_SHAPE = 1
_MSO_LINE = 9
_MSO_PICTURE = 13
_MSO_PLACEHOLDER = 14
_MSO_TEXT_BOX = 17
_MSO_TABLE = 19
_MSO_SMARTART = 24

# MsoTriState
_MSO_TRUE = -1
_MSO_FALSE = 0

# PpSelectionType
_SEL_NONE = 0
_SEL_SLIDES = 1
_SEL_SHAPES = 2
_SEL_TEXT = 3

# PpSlideShowState / PpSlideShowRangeType
_SHOW_RUNNING = 1
_SHOW_BLACK = 3
_SHOW_WHITE = 4
_SHOW_DONE = 5
_RANGE_SLIDE = 2


def _minimal_png(width: int, height: int) -> bytes:
    """A 24-byte stub PNG (signature + IHDR) — enough that a reader can recover
    the dimensions, so `Slide.Export` round-trips in unit tests without rendering."""
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = (
        (13).to_bytes(4, "big")
        + b"IHDR"
        + int(width).to_bytes(4, "big")
        + int(height).to_bytes(4, "big")
    )
    return sig + ihdr


# The standard Office layout names a deck's SlideMaster.CustomLayouts exposes;
# the fake offers these so layout-name resolution can be exercised end to end.
_STANDARD_LAYOUTS = (
    "Title Slide",
    "Title and Content",
    "Section Header",
    "Two Content",
    "Comparison",
    "Title Only",
    "Blank",
    "Content with Caption",
    "Picture with Caption",
)


# ---------------------------------------------------------------------------
# Shape and its text frame
# ---------------------------------------------------------------------------


# A paragraph-aware text model (v0.3). The source of truth is a list of
# `_FakePara` objects on the frame; a `_FakeTextRange` is a view over a 1-based
# contiguous span of them. Setting/inserting text reconstructs the paragraph list
# exactly as PowerPoint's char-spliced TextRange does (verified in text_spike.py),
# so paragraph addressing, the trailing-CR behavior, and insert all round-trip.


class _FakeFont:
    def __init__(self) -> None:
        self.Bold = 0
        self.Italic = 0
        self.Underline = 0
        self.Size = 18.0
        self.Name = "Calibri"
        self.Color = SimpleNamespace(RGB=0)


class _FakeBullet:
    def __init__(self) -> None:
        self.Visible = 0
        self.Type = 0
        self.Character = 0


class _FakeParagraphFormat:
    def __init__(self) -> None:
        self.Alignment = 1
        self.SpaceBefore = 0.0
        self.SpaceAfter = 0.0
        self.SpaceWithin = 1.0
        # LineRule* select each Space*'s unit: msoTrue (-1) = multiple/lines,
        # msoFalse (0) = points. Default within is a multiple (SpaceWithin 1.0 =
        # single); before/after default to points (0 pt). Mirrors live COM.
        self.LineRuleWithin = -1
        self.LineRuleBefore = 0
        self.LineRuleAfter = 0
        self.Bullet = _FakeBullet()


class _FakePara:
    def __init__(self, text: str = "") -> None:
        self.text = text
        self.Font = _FakeFont()
        self.ParagraphFormat = _FakeParagraphFormat()
        self.IndentLevel = 1


class _SpanNestedProxy:
    """Forwards `.attr.sub.<name>` get/set across a range's paragraph span
    (e.g. `Font.Color`, `ParagraphFormat.Bullet`): reads the first, writes all."""

    def __init__(self, rng: _FakeTextRange, attr: str, sub: str) -> None:
        object.__setattr__(self, "_rng", rng)
        object.__setattr__(self, "_attr", attr)
        object.__setattr__(self, "_sub", sub)

    def __getattr__(self, name: str) -> Any:
        p0 = self._rng._paras_in_span()[0]
        return getattr(getattr(getattr(p0, self._attr), self._sub), name)

    def __setattr__(self, name: str, value: Any) -> None:
        for p in self._rng._paras_in_span():
            setattr(getattr(getattr(p, self._attr), self._sub), name, value)


class _SpanAttrProxy:
    """Forwards a per-paragraph sub-object's attrs across a range's span: reads
    the first paragraph's value, writes to every paragraph (`Font`, `ParagraphFormat`)."""

    def __init__(self, rng: _FakeTextRange, attr: str) -> None:
        object.__setattr__(self, "_rng", rng)
        object.__setattr__(self, "_attr", attr)

    def __getattr__(self, name: str) -> Any:
        if name in ("Color", "Bullet"):
            return _SpanNestedProxy(self._rng, self._attr, name)
        return getattr(getattr(self._rng._paras_in_span()[0], self._attr), name)

    def __setattr__(self, name: str, value: Any) -> None:
        for p in self._rng._paras_in_span():
            setattr(getattr(p, self._attr), name, value)


class _FakeTextRange:
    """A view over paragraphs [start .. start+length-1] (1-based) of a frame."""

    def __init__(self, frame: _FakeTextFrame, start: int = 1, length: int = -1) -> None:
        self._frame = frame
        self._start = start
        self._length = length

    def _span(self) -> tuple[int, int]:
        total = len(self._frame._paras)
        start = max(1, self._start)
        end = total if self._length == -1 else min(total, start + self._length - 1)
        return start, end

    def _paras_in_span(self) -> list[_FakePara]:
        s, e = self._span()
        return self._frame._paras[s - 1 : e]

    @property
    def Text(self) -> str:
        s, e = self._span()
        text = "\r".join(p.text for p in self._frame._paras[s - 1 : e])
        if e < len(self._frame._paras):  # a non-final paragraph carries its break
            text += "\r"
        return text

    @Text.setter
    def Text(self, value: str) -> None:
        s, e = self._span()
        parts = str(value).split("\r")
        if len(parts) == 1 and (e - s + 1) == 1:
            self._frame._paras[s - 1].text = parts[0]  # keep the paragraph's formatting
            return
        replacement = [_FakePara(p) for p in parts]
        self._frame._paras = self._frame._paras[: s - 1] + replacement + self._frame._paras[e:]
        if not self._frame._paras:
            self._frame._paras = [_FakePara("")]

    @property
    def Count(self) -> int:
        s, e = self._span()
        return e - s + 1

    def Paragraphs(self, start: int = -1, length: int = -1) -> _FakeTextRange:
        s, _e = self._span()
        if start == -1:
            return _FakeTextRange(self._frame, s, self._length)
        return _FakeTextRange(self._frame, s + start - 1, length)

    def Runs(self, start: int = -1, length: int = -1) -> _FakeTextRange:
        # The fake models formatting at paragraph granularity (one Font per
        # paragraph), so it exposes one run per spanned paragraph — enough to
        # exercise the run-walk; true sub-paragraph mixed runs live in smoke.
        return self.Paragraphs(start, length)

    @property
    def IndentLevel(self) -> int:
        return self._paras_in_span()[0].IndentLevel

    @IndentLevel.setter
    def IndentLevel(self, value: int) -> None:
        for p in self._paras_in_span():
            p.IndentLevel = int(value)

    @property
    def Font(self) -> _SpanAttrProxy:
        return _SpanAttrProxy(self, "Font")

    @property
    def ParagraphFormat(self) -> _SpanAttrProxy:
        return _SpanAttrProxy(self, "ParagraphFormat")

    def _full_text(self) -> str:
        return "\r".join(p.text for p in self._frame._paras)

    def _char_start(self) -> int:
        s, _e = self._span()
        return sum(len(p.text) + 1 for p in self._frame._paras[: s - 1])

    def InsertBefore(self, text: str) -> _FakeTextRange:
        full = self._full_text()
        off = self._char_start()
        self._frame._set_full_text(full[:off] + str(text) + full[off:])
        return self

    def InsertAfter(self, text: str) -> _FakeTextRange:
        full = self._full_text()
        off = self._char_start() + len(self.Text)
        self._frame._set_full_text(full[:off] + str(text) + full[off:])
        return self

    def Delete(self) -> None:
        s, e = self._span()
        self._frame._paras = self._frame._paras[: s - 1] + self._frame._paras[e:]
        if not self._frame._paras:
            self._frame._paras = [_FakePara("")]

    def Characters(self, start: int = 1, length: int = -1) -> _FakeCharRange:
        """`TextRange.Characters(Start, Length)` — a 1-based char sub-range view
        (relative to this range), the handle find/replace writes a span through."""
        return _FakeCharRange(self, int(start), int(length))


class _FakeCharRange:
    """A view over `Characters(start, length)` of a parent `_FakeTextRange`.

    1-based `start`, relative to the parent range; setting `.Text` char-splices
    the underlying frame (exactly as PowerPoint's char-addressed TextRange does),
    so only the matched span changes and the rest of the frame is untouched."""

    def __init__(self, parent: _FakeTextRange, start: int, length: int) -> None:
        self._parent = parent
        self._start = int(start)
        self._length = int(length)

    def _bounds(self) -> tuple[int, int]:
        ptext = self._parent.Text
        s = max(1, self._start) - 1
        e = len(ptext) if self._length == -1 else min(len(ptext), s + self._length)
        return s, max(s, e)

    @property
    def Text(self) -> str:
        s, e = self._bounds()
        return self._parent.Text[s:e]

    @Text.setter
    def Text(self, value: str) -> None:
        s, e = self._bounds()
        base = self._parent._char_start()
        full = self._parent._full_text()
        self._parent._frame._set_full_text(full[: base + s] + str(value) + full[base + e :])

    @property
    def Count(self) -> int:
        s, e = self._bounds()
        return e - s


class _FakeTextFrame:
    def __init__(self, text: str = "") -> None:
        self._paras = [_FakePara(p) for p in str(text).split("\r")]
        # Autofit container props (text_frame_status). Classic AutoSize reads mixed
        # on real builds (-2); the wrapper prefers TextFrame2.AutoSize. Margins in
        # points mirror PowerPoint's defaults.
        self.AutoSize = -2
        self.WordWrap = _MSO_TRUE
        self.MarginLeft = 7.2
        self.MarginRight = 7.2
        self.MarginTop = 3.6
        self.MarginBottom = 3.6

    @property
    def TextRange(self) -> _FakeTextRange:
        return _FakeTextRange(self, 1, -1)

    def _set_full_text(self, full: str) -> None:
        self._paras = [_FakePara(p) for p in str(full).split("\r")]


class _FakeGradientStop:
    """One `Fill.GradientStops(i)` — Position + Transparency + Color.RGB."""

    def __init__(self, rgb: int, position: float, transparency: float = 0.0) -> None:
        self.Position = float(position)
        self.Transparency = float(transparency)
        self.Color = SimpleNamespace(RGB=int(rgb))


class _FakeGradientStops:
    """`Fill.GradientStops` — Count + `(i)` access + legacy `Insert(rgb, position)`.

    Mirrors the spike: `Insert` *appends* a stop (the wrapper / reader sorts by
    position on read); `Insert2` is the one that won't marshal, so it's absent here.
    """

    def __init__(self) -> None:
        self._stops: list[_FakeGradientStop] = []

    @property
    def Count(self) -> int:
        return len(self._stops)

    def __call__(self, index: int) -> _FakeGradientStop:
        return self._stops[int(index) - 1]  # 1-based

    def Insert(self, rgb: int, position: float) -> None:
        self._stops.append(_FakeGradientStop(rgb, position))


class _FakeShapeFill:
    """`Shape.Fill` — solid / gradient / pattern / picture.

    Defaults to a visible *theme* fill (the `0x80000000` automatic sentinel, so a
    readback reports `color: None` until a literal RGB is set), matching how a
    fresh autoshape inherits the theme accent. The gradient / pattern / picture
    methods set `Type` and the type-specific read-back fields the wrappers expect
    (spike-verified in scripts/fill_advanced_spike.py).
    """

    def __init__(self) -> None:
        self.Visible = _MSO_TRUE
        self.Type = 5  # msoFillBackground (theme/inherited) until a fill verb runs
        self.ForeColor = SimpleNamespace(RGB=0x80000000)
        self.BackColor = SimpleNamespace(RGB=0x80000000)
        self.Transparency = 0.0  # opaque until a partial-alpha is set
        self.GradientStyle = -2  # msoGradientMixed until a gradient is set
        self.GradientVariant = 0
        self.GradientColorType = 0
        self.GradientDegree = 0.0
        self.Pattern = -2  # msoPatternMixed until a pattern is set
        self.TextureType = 0
        self.GradientStops = _FakeGradientStops()
        self.picture_path: str | None = None

    def Solid(self) -> None:
        self.Type = 1  # msoFillSolid

    def _two_default_stops(self) -> None:
        # The two endpoint stops track ForeColor/BackColor (the wrapper sets those
        # AFTER TwoColorGradient, mirroring real COM), so they share the namespaces.
        self.GradientStops = _FakeGradientStops()
        s0 = _FakeGradientStop(0, 0.0)
        s0.Color = self.ForeColor
        s1 = _FakeGradientStop(0, 1.0)
        s1.Color = self.BackColor
        self.GradientStops._stops = [s0, s1]

    def TwoColorGradient(self, style: int, variant: int) -> None:
        self.Type = 3  # msoFillGradient
        self.GradientColorType = 2
        self.GradientStyle = int(style)
        self.GradientVariant = int(variant)
        self._two_default_stops()

    def OneColorGradient(self, style: int, variant: int, degree: float) -> None:
        self.Type = 3
        self.GradientColorType = 1
        self.GradientStyle = int(style)
        self.GradientVariant = int(variant)
        self.GradientDegree = float(degree)

    def PresetGradient(self, style: int, variant: int, preset: int) -> None:
        self.Type = 3
        self.GradientColorType = 3
        self.GradientStyle = int(style)
        self.GradientVariant = int(variant)

    def Patterned(self, pattern: int) -> None:
        self.Type = 2  # msoFillPatterned
        self.Pattern = int(pattern)

    def UserPicture(self, path: str) -> None:
        self.Type = 6  # msoFillPicture
        self.TextureType = 2
        self.picture_path = str(path)


class _FakeShapeLine:
    """`Shape.Line` — Visible + Weight + ForeColor.RGB + dash/arrowheads (the border)."""

    def __init__(self) -> None:
        self.Visible = _MSO_TRUE
        self.Weight = 1.0
        self.ForeColor = SimpleNamespace(RGB=0x80000000)
        self.Transparency = 0.0  # opaque until a partial-alpha is set
        self.DashStyle = 1  # msoLineSolid
        self.BeginArrowheadStyle = 1  # msoArrowheadNone
        self.EndArrowheadStyle = 1
        self.BeginArrowheadLength = 2  # msoArrowheadLengthMedium
        self.BeginArrowheadWidth = 2
        self.EndArrowheadLength = 2
        self.EndArrowheadWidth = 2


class _FakeHyperlink:
    """`ActionSetting.Hyperlink` — Address / SubAddress / ScreenTip + Delete().

    Mirrors the live finding (scripts/hyperlink_spike.py): assigning a non-empty
    `Address` or `SubAddress` flips the owning action to `ppActionHyperlink` (7);
    `Delete()` reverts the action to `ppActionNone` (0) and clears the address.
    """

    def __init__(self, action: _FakeActionSetting) -> None:
        self._action = action
        self._address = ""
        self._sub_address = ""
        self.ScreenTip = ""

    @property
    def Address(self) -> str:
        return self._address

    @Address.setter
    def Address(self, value: str) -> None:
        self._address = str(value or "")
        self._action._refresh()

    @property
    def SubAddress(self) -> str:
        return self._sub_address

    @SubAddress.setter
    def SubAddress(self, value: str) -> None:
        self._sub_address = str(value or "")
        self._action._refresh()

    def Delete(self) -> None:
        self._address = ""
        self._sub_address = ""
        self.ScreenTip = ""
        self._action.Action = 0  # ppActionNone


class _FakeActionSetting:
    """`Shape.ActionSettings(ppMouseClick)` — `.Action` + `.Hyperlink`."""

    def __init__(self) -> None:
        self.Action = 0  # ppActionNone until a link is assigned
        self.Hyperlink = _FakeHyperlink(self)

    def _refresh(self) -> None:
        linked = bool(self.Hyperlink._address or self.Hyperlink._sub_address)
        self.Action = 7 if linked else 0  # ppActionHyperlink / ppActionNone


class _FakeShape:
    """A shape. `text=None` means no text frame (picture/line)."""

    def __init__(
        self,
        *,
        name: str,
        shape_id: int,
        shape_type: int,
        text: str | None = None,
        placeholder_type: int | None = None,
        table: _FakeTable | None = None,
        left: float = 10.0,
        top: float = 20.0,
        width: float = 100.0,
        height: float = 50.0,
        rotation: float = 0.0,
    ) -> None:
        self.Name = name
        self.Id = shape_id
        self.Type = shape_type
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.Rotation = rotation
        self.AlternativeText = ""
        self.Fill = _FakeShapeFill()
        self.Line = _FakeShapeLine()
        # Effects (v1.2). Defaults are "off": shadow invisible, glow radius 0,
        # soft-edge / reflection preset 0. Transparency defaults to the unset
        # sentinel (-2147483648) the wrapper maps to None on read.
        self.Shadow = SimpleNamespace(
            Visible=_MSO_FALSE,
            Type=1,
            Style=1,
            Transparency=-2147483648,
            Blur=0.0,
            Size=100.0,
            OffsetX=0.0,
            OffsetY=0.0,
            ForeColor=SimpleNamespace(RGB=0x000000),
        )
        self.Glow = SimpleNamespace(
            Radius=0.0, Transparency=-2147483648, Color=SimpleNamespace(RGB=0x000000)
        )
        self.SoftEdge = SimpleNamespace(Type=0, Radius=0.0)
        self.Reflection = SimpleNamespace(Type=0)
        self._placeholder_type = placeholder_type
        self._table = table
        self._chart: _FakeChart | None = None
        self._smartart: _FakeSmartArt | None = None
        self._text_frame = _FakeTextFrame(text) if text is not None else None
        # Modern TextFrame2 — the wrapper reads AutoSize off this (the classic
        # TextFrame.AutoSize is the mixed sentinel on real builds). 1 =
        # msoAutoSizeTextToFitShape, a content placeholder's "shrink on overflow".
        self._text_frame2 = (
            SimpleNamespace(AutoSize=1, WordWrap=_MSO_TRUE, HasText=_MSO_TRUE)
            if text is not None
            else None
        )
        self.selected = False
        self.last_export: dict[str, Any] | None = None
        self._collection: _FakeShapes | None = None  # set when adopted by _FakeShapes
        self._action_settings: dict[int, _FakeActionSetting] = {}

    def ActionSettings(self, activation: int) -> _FakeActionSetting:
        """`Shape.ActionSettings(ppMouseClick=1)` — created on first access."""
        key = int(activation)
        if key not in self._action_settings:
            self._action_settings[key] = _FakeActionSetting()
        return self._action_settings[key]

    def Delete(self) -> None:
        assert self._collection is not None
        self._collection._shapes.remove(self)

    @property
    def ZOrderPosition(self) -> int:
        """1-based slot in the collection (index 1 = back, Count = front)."""
        assert self._collection is not None
        return self._collection._shapes.index(self) + 1

    def ZOrder(self, cmd: int) -> None:
        """Restack within the collection. `cmd` is the `MsoZOrderCmd` int
        (0=BringToFront, 1=SendToBack, 2=BringForward, 3=SendBackward); the list
        is back→front, so front == last slot."""
        assert self._collection is not None
        lst = self._collection._shapes
        i = lst.index(self)
        lst.remove(self)
        if cmd == 0:  # BringToFront
            lst.append(self)
        elif cmd == 1:  # SendToBack
            lst.insert(0, self)
        elif cmd == 2:  # BringForward (one toward front)
            lst.insert(min(i + 1, len(lst)), self)
        elif cmd == 3:  # SendBackward (one toward back)
            lst.insert(max(i - 1, 0), self)
        else:
            lst.insert(i, self)

    @property
    def HasTextFrame(self) -> int:
        return _MSO_TRUE if self._text_frame is not None else _MSO_FALSE

    @property
    def TextFrame(self) -> _FakeTextFrame:
        if self._text_frame is None:
            raise AttributeError("shape has no text frame")
        return self._text_frame

    @property
    def TextFrame2(self) -> Any:
        if self._text_frame2 is None:
            raise AttributeError("shape has no text frame")
        return self._text_frame2

    @property
    def HasTable(self) -> int:
        return _MSO_TRUE if self._table is not None else _MSO_FALSE

    @property
    def Table(self) -> _FakeTable:
        if self._table is None:
            raise AttributeError("shape has no table")
        return self._table

    @property
    def HasChart(self) -> int:
        return _MSO_TRUE if self._chart is not None else _MSO_FALSE

    @property
    def Chart(self) -> _FakeChart:
        if self._chart is None:
            raise AttributeError("shape has no chart")
        return self._chart

    @property
    def HasSmartArt(self) -> int:
        return _MSO_TRUE if self._smartart is not None else _MSO_FALSE

    @property
    def SmartArt(self) -> _FakeSmartArt:
        if self._smartart is None:
            raise AttributeError("shape has no smartart")
        return self._smartart

    @property
    def PlaceholderFormat(self) -> Any:
        if self._placeholder_type is None:
            raise AttributeError("shape is not a placeholder")
        return SimpleNamespace(Type=self._placeholder_type)

    def Select(self, *args: Any, **kwargs: Any) -> None:
        self.selected = True

    def Export(
        self,
        PathName: str,
        Filter: int,
        ScaleWidth: int | None = None,
        ScaleHeight: int | None = None,
    ) -> None:
        """Render this shape to a stub PNG. `Filter` is a PpShapeFormat int (not a
        FilterName string). Honors ScaleWidth/Height; else the shape's native
        pixel size at 96 DPI (so a 100x50 pt shape -> ~133x67 px)."""
        if ScaleWidth and ScaleHeight:
            w, h = int(ScaleWidth), int(ScaleHeight)
        else:
            w = int(round(float(self.Width) * 96 / 72))
            h = int(round(float(self.Height) * 96 / 72))
        self.last_export = {"PathName": PathName, "Filter": int(Filter), "Width": w, "Height": h}
        with open(PathName, "wb") as fh:
            fh.write(_minimal_png(w, h))

    def _clone(self) -> _FakeShape:
        """A duplicate-slide copy — same name/id/type/geometry/text."""
        text = self._text_frame.TextRange.Text if self._text_frame is not None else None
        return _FakeShape(
            name=self.Name,
            shape_id=self.Id,
            shape_type=self.Type,
            text=text,
            placeholder_type=self._placeholder_type,
            left=self.Left,
            top=self.Top,
            width=self.Width,
            height=self.Height,
            rotation=self.Rotation,
        )


# ---------------------------------------------------------------------------
# Table (v0.5): a shape with HasTable; cells carry their own text-framed Shape
# ---------------------------------------------------------------------------


class _FakeCell:
    """A table cell — its text lives in `Cell.Shape.TextFrame.TextRange`, exactly
    like real PowerPoint (a normal text frame, paragraphs split by `\\r`)."""

    def __init__(self, text: str = "") -> None:
        self.Shape = _FakeShape(name="Table Cell", shape_id=0, shape_type=_MSO_TEXT_BOX, text=text)


class _FakeRows:
    """`Table.Rows` — 1-based callable + `.Add()` (appends) + `.Count`."""

    def __init__(self, table: _FakeTable) -> None:
        self._table = table

    @property
    def Count(self) -> int:
        return len(self._table._cells)

    def Add(self, before_row: int | None = None) -> None:
        self._table._cells.append([_FakeCell("") for _ in range(self._table._ncols)])

    def __call__(self, index: int) -> _FakeRow:
        return _FakeRow(self._table, int(index))


class _FakeRow:
    def __init__(self, table: _FakeTable, index: int) -> None:
        self._table = table
        self._index = index

    def Delete(self) -> None:
        del self._table._cells[self._index - 1]


class _FakeColumns:
    def __init__(self, table: _FakeTable) -> None:
        self._table = table

    @property
    def Count(self) -> int:
        return self._table._ncols


class _FakeTable:
    """A table grid: rows × columns of `_FakeCell`s, 1-based `Cell(r, c)`."""

    def __init__(self, rows: int, cols: int) -> None:
        self._ncols = int(cols)
        self._cells = [[_FakeCell("") for _ in range(int(cols))] for _ in range(int(rows))]

    @property
    def Rows(self) -> _FakeRows:
        return _FakeRows(self)

    @property
    def Columns(self) -> _FakeColumns:
        return _FakeColumns(self)

    def Cell(self, row: int, col: int) -> _FakeCell:
        # 1-based; real COM raises for out-of-range, so do we (the wrapper
        # bounds-checks first, but this keeps the fake honest).
        if row < 1 or row > len(self._cells) or col < 1 or col > self._ncols:
            raise IndexError((row, col))
        return self._cells[row - 1][col - 1]


# ---------------------------------------------------------------------------
# Chart (v0.7): a shape with HasChart; data lives in an embedded Excel workbook
# ---------------------------------------------------------------------------
#
# The fake models the workbook faithfully enough to exercise the real write
# sequence (UsedRange.ClearContents -> Cells writes -> SetSourceData(string)):
# the worksheet is a {(row, col): value} dict, SetSourceData records the plotted
# range as a string, and SeriesCollection parses that range against the cells —
# exactly how a freshly-read chart recovers its names/categories/values.


class _FakeCellRef:
    """`ws.Cells(r, c)` — a settable `.Value`/`.NumberFormat` view over the ws dict."""

    def __init__(self, ws: _FakeWorksheet, row: int, col: int) -> None:
        self._ws = ws
        self._key = (int(row), int(col))

    @property
    def Value(self) -> Any:
        return self._ws._cells.get(self._key)

    @Value.setter
    def Value(self, v: Any) -> None:
        self._ws._cells[self._key] = v

    @property
    def NumberFormat(self) -> Any:
        return self._ws._formats.get(self._key, "General")

    @NumberFormat.setter
    def NumberFormat(self, fmt: Any) -> None:
        self._ws._formats[self._key] = fmt


class _FakeUsedRange:
    def __init__(self, ws: _FakeWorksheet) -> None:
        self._ws = ws

    def ClearContents(self) -> None:
        self._ws._cells.clear()
        self._ws._formats.clear()


class _FakeWorksheet:
    def __init__(self) -> None:
        self._cells: dict[tuple[int, int], Any] = {}
        self._formats: dict[tuple[int, int], Any] = {}
        # Excel's first sheet is named "Sheet1" only on English Office; the real
        # code now references it by `.Name`, so the fake carries one too.
        self.Name = "Sheet1"

    def Cells(self, row: int, col: int) -> _FakeCellRef:
        return _FakeCellRef(self, row, col)

    @property
    def UsedRange(self) -> _FakeUsedRange:
        return _FakeUsedRange(self)


class _FakeExcelApp:
    def __init__(self) -> None:
        self.quit_called = False

    def Quit(self) -> None:
        self.quit_called = True


class _FakeWorkbook:
    def __init__(self, ws: _FakeWorksheet) -> None:
        self._ws = ws
        self.Application = _FakeExcelApp()
        self.closed = False

    def Worksheets(self, index: int) -> _FakeWorksheet:
        return self._ws

    def Close(self) -> None:
        self.closed = True


class _FakeChartData:
    def __init__(self, chart: _FakeChart) -> None:
        self._chart = chart

    def Activate(self) -> None:
        self._chart._activated = True

    @property
    def Workbook(self) -> _FakeWorkbook:
        return self._chart._wb


class _FakeSeries:
    def __init__(
        self,
        name: Any,
        x_values: list[Any],
        values: list[Any],
        chart: _FakeChart | None = None,
        index: int | None = None,
    ) -> None:
        self.Name = name
        self.XValues = x_values
        self.Values = values
        self._chart = chart
        self._index = index

    @property
    def HasDataLabels(self) -> bool:
        return bool(self._chart._series_labels_on) if self._chart is not None else False

    @HasDataLabels.setter
    def HasDataLabels(self, value: bool) -> None:
        if self._chart is not None:
            self._chart._series_labels_on = bool(value)

    def DataLabels(self) -> Any:
        # A method (not a property), like the live COM — recolor reaches the font
        # via `DataLabels().Font.Color`. Persist the font on the chart so a test
        # can read the color back after the per-call series object is gone.
        assert self._chart is not None and self._index is not None
        return SimpleNamespace(Font=self._chart._label_font(self._index))


class _FakeSeriesCollection:
    def __init__(self, series: list[_FakeSeries]) -> None:
        self._series = series

    @property
    def Count(self) -> int:
        return len(self._series)

    def __call__(self, index: int) -> _FakeSeries:
        return self._series[int(index) - 1]


class _FakeChart:
    """A chart backed by an embedded `_FakeWorkbook`. `SetSourceData(string)`
    records the plotted range; `SeriesCollection()` reads it from the cells."""

    def __init__(self, chart_type: int) -> None:
        self.ChartType = int(chart_type)
        self._ws = _FakeWorksheet()
        self._wb = _FakeWorkbook(self._ws)
        self._activated = False
        # Source range as (nrows, ncols); seed PowerPoint's default placeholder
        # data: 3 series x 4 categories in $A$1:$D$5.
        self._seed_default_data()
        self._nrows, self._ncols = 5, 4
        # Text elements `recolor_text` touches. ChartArea's color lives on the
        # modern TextFrame2; legend/title/axis fonts use the classic chart-Font
        # model where `Font.Color` is the RGB long directly.
        self.HasLegend = True
        self.HasTitle = False
        self.Legend = SimpleNamespace(Font=SimpleNamespace(Color=0))
        self.ChartTitle = SimpleNamespace(Font=SimpleNamespace(Color=0), Text="")
        self.ChartArea = SimpleNamespace(Format=SimpleNamespace(TextFrame2=_fake_textframe2()))
        self._axes_present = True  # a pie chart sets this False -> Axes() raises
        self._axis_fonts: dict[int, Any] = {
            1: SimpleNamespace(Color=0),  # category
            2: SimpleNamespace(Color=0),  # value
        }
        self._series_labels_on = False
        self._label_fonts: dict[int, Any] = {}

    def Axes(self, axis_type: int, group: int | None = None) -> Any:
        if not self._axes_present:
            from pptlive.exceptions import ComError

            raise ComError("this chart has no axes")
        return SimpleNamespace(TickLabels=SimpleNamespace(Font=self._axis_fonts[int(axis_type)]))

    def _label_font(self, index: int) -> Any:
        return self._label_fonts.setdefault(int(index), SimpleNamespace(Color=0))

    def _seed_default_data(self) -> None:
        c = self._ws._cells
        c[(1, 1)] = " "
        for col in range(2, 5):
            c[(1, col)] = f"Series {col - 1}"
        for row in range(2, 6):
            c[(row, 1)] = f"Category {row - 1}"
            for col in range(2, 5):
                c[(row, col)] = float(row - 1)

    @property
    def ChartData(self) -> _FakeChartData:
        return _FakeChartData(self)

    def SetSourceData(self, source: str) -> None:
        # source like "'Sheet1'!$A$1:$C$4" -> parse the bottom-right corner.
        self._last_source = source  # captured so tests can assert the sheet ref
        ref = source.split("!", 1)[-1].replace("$", "")
        _tl, _, br = ref.partition(":")
        col_letters = "".join(ch for ch in br if ch.isalpha())
        row_digits = "".join(ch for ch in br if ch.isdigit())
        ncols = 0
        for ch in col_letters:
            ncols = ncols * 26 + (ord(ch.upper()) - 64)
        self._ncols, self._nrows = ncols, int(row_digits)

    def SeriesCollection(self) -> _FakeSeriesCollection:
        cells = self._ws._cells
        cats = [cells.get((r, 1)) for r in range(2, self._nrows + 1)]
        series = []
        for col in range(2, self._ncols + 1):
            name = cells.get((1, col))
            values = [cells.get((r, col)) for r in range(2, self._nrows + 1)]
            series.append(_FakeSeries(name, list(cats), values, chart=self, index=col - 1))
        return _FakeSeriesCollection(series)


# ---------------------------------------------------------------------------
# SmartArt (v0.8): a shape with HasSmartArt; content is a tree of nodes whose
# text lives on TextFrame2. The fake mirrors the live findings the wrapper relies
# on: node `Type` is always default (assistant doesn't round-trip),
# `SmartArtNode.AddNode(BELOW)` adds a child while `Nodes.Add()` adds a top-level
# sibling, and tree layouts (hierarchy/orgChart) seed a skeleton (1 root + empty
# children) that set_nodes must clear.
# ---------------------------------------------------------------------------

_SMARTART_CATALOG: tuple[tuple[str, str], ...] = (
    ("Vertical Box List", "urn:microsoft.com/office/officeart/2005/8/layout/list1"),
    ("Basic Process", "urn:microsoft.com/office/officeart/2005/8/layout/process1"),
    ("Text Cycle", "urn:microsoft.com/office/officeart/2005/8/layout/cycle1"),
    ("Hierarchy", "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1"),
    ("Organization Chart", "urn:microsoft.com/office/officeart/2005/8/layout/orgChart1"),
    ("Basic Pyramid", "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1"),
    ("Basic Venn", "urn:microsoft.com/office/officeart/2005/8/layout/venn1"),
)


class _FakeSmartArtLayouts:
    """`Application.SmartArtLayouts` — the installed catalog, keyed by stable URN."""

    def __init__(self) -> None:
        self._items = [SimpleNamespace(Name=name, Id=urn) for name, urn in _SMARTART_CATALOG]

    @property
    def Count(self) -> int:
        return len(self._items)

    def Item(self, index: int) -> Any:
        return self._items[int(index) - 1]


def _fake_textframe2() -> SimpleNamespace:
    """A TextFrame2 whose text-color lives on `TextRange.Font.Fill.ForeColor.RGB`.

    SmartArt nodes (and a chart's ChartArea) carry color there, not on
    `Font.Color` — mirrors `SmartArt.recolor_text` / `Chart.recolor_text`. The RGB
    seeds to the `0x80000000` automatic sentinel, like a real theme-driven color.
    """
    fore = SimpleNamespace(RGB=0x80000000)
    font = SimpleNamespace(Fill=SimpleNamespace(ForeColor=fore))
    return SimpleNamespace(TextRange=SimpleNamespace(Text="", Font=font))


class _FakeSmartArtNode:
    def __init__(self, level: int, parent_list: _FakeSmartArtNodes) -> None:
        self._level = int(level)
        self._parent_list = parent_list
        self.TextFrame2 = _fake_textframe2()
        self._children = _FakeSmartArtNodes(level=self._level + 1)

    @property
    def Level(self) -> int:
        return self._level

    @property
    def Type(self) -> int:
        return 1  # always msoSmartArtNodeTypeDefault — assistant doesn't round-trip

    @property
    def Nodes(self) -> _FakeSmartArtNodes:
        return self._children

    def AddNode(self, position: int = 5, node_type: int = 1) -> _FakeSmartArtNode:
        """Add a *child* (position is BELOW=5 in practice) and return it."""
        return self._children._append()

    def Delete(self) -> None:
        self._parent_list._remove(self)


class _FakeSmartArtNodes:
    def __init__(self, level: int) -> None:
        self._level = int(level)
        self._nodes: list[_FakeSmartArtNode] = []

    def _append(self) -> _FakeSmartArtNode:
        node = _FakeSmartArtNode(self._level, self)
        self._nodes.append(node)
        return node

    def _remove(self, node: _FakeSmartArtNode) -> None:
        self._nodes.remove(node)

    @property
    def Count(self) -> int:
        return len(self._nodes)

    def Item(self, index: int) -> _FakeSmartArtNode:
        return self._nodes[int(index) - 1]

    def Add(self) -> _FakeSmartArtNode:
        """Add a top-level sibling. (The fake always grows; tree-layout caps are a
        live-COM behavior the wrapper guards against via its kind pre-check.)"""
        return self._append()


class _FakeSmartArt:
    """A SmartArt diagram backed by a node tree; `Layout.Id` is the stable URN."""

    def __init__(self, layout_id: str) -> None:
        self.Layout = SimpleNamespace(Id=str(layout_id))
        self._nodes = _FakeSmartArtNodes(level=1)
        self._seed_default()

    def _seed_default(self) -> None:
        seg = self.Layout.Id.rsplit("/", 1)[-1]
        if seg in ("hierarchy1", "orgChart1"):  # tree: 1 root + 2 empty children
            root = self._nodes._append()
            root.AddNode()
            root.AddNode()
        else:  # flat: 3 empty top-level nodes
            for _ in range(3):
                self._nodes._append()

    @property
    def Nodes(self) -> _FakeSmartArtNodes:
        return self._nodes

    @property
    def AllNodes(self) -> Any:
        flat: list[_FakeSmartArtNode] = []

        def walk(coll: _FakeSmartArtNodes) -> None:
            for i in range(1, coll.Count + 1):
                node = coll.Item(i)
                flat.append(node)
                walk(node.Nodes)

        walk(self._nodes)
        return SimpleNamespace(Count=len(flat), Item=lambda i: flat[int(i) - 1])


class _FakeShapeRange:
    def __init__(self, app: _FakeApplication, shapes: list[_FakeShape]) -> None:
        self._app = app
        self._shapes = shapes

    def __iter__(self) -> Any:
        return iter(self._shapes)

    def __call__(self, index: int) -> _FakeShape:
        return self._shapes[index - 1]

    @property
    def Count(self) -> int:
        return len(self._shapes)

    def Select(self, *args: Any, **kwargs: Any) -> None:
        for sh in self._shapes:
            sh.selected = True
        self._app._selection_type = _SEL_SHAPES
        self._app._selected_names = tuple(sh.Name for sh in self._shapes)


class _FakeShapes:
    def __init__(self, shapes: list[_FakeShape]) -> None:
        self._shapes = shapes
        self._app: _FakeApplication | None = None
        self._id_counter = max((s.Id for s in shapes), default=1)
        for sh in shapes:
            sh._collection = self

    def _next_id(self) -> int:
        self._id_counter += 1
        return self._id_counter

    def _adopt(self, sh: _FakeShape) -> _FakeShape:
        """Append a new shape at the top of the z-order (last slot), like COM."""
        sh._collection = self
        self._shapes.append(sh)
        return sh

    def AddTextbox(
        self, orientation: int, left: float, top: float, width: float, height: float
    ) -> _FakeShape:
        sid = self._next_id()
        return self._adopt(
            _FakeShape(
                name=f"TextBox {sid}",
                shape_id=sid,
                shape_type=_MSO_TEXT_BOX,
                text="",
                left=left,
                top=top,
                width=width,
                height=height,
            )
        )

    def AddShape(
        self, shape_type: int, left: float, top: float, width: float, height: float
    ) -> _FakeShape:
        sid = self._next_id()
        sh = _FakeShape(
            name=f"Shape {sid}",
            shape_id=sid,
            shape_type=_MSO_AUTO_SHAPE,
            text="",
            left=left,
            top=top,
            width=width,
            height=height,
        )
        sh.AutoShapeType = int(shape_type)
        return self._adopt(sh)

    def AddPicture(
        self,
        filename: str,
        link_to_file: int,
        save_with_document: int,
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> _FakeShape:
        sid = self._next_id()
        # -1 means "native size"; the fake just substitutes a nominal size.
        w = 120.0 if width == -1 else width
        h = 90.0 if height == -1 else height
        return self._adopt(
            _FakeShape(
                name=f"Picture {sid}",
                shape_id=sid,
                shape_type=_MSO_PICTURE,
                text=None,
                left=left,
                top=top,
                width=w,
                height=h,
            )
        )

    def AddTable(
        self,
        num_rows: int,
        num_cols: int,
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> _FakeShape:
        sid = self._next_id()
        return self._adopt(
            _FakeShape(
                name=f"Table {sid}",
                shape_id=sid,
                shape_type=_MSO_TABLE,
                text=None,  # the table shape itself has no text frame; cells do
                table=_FakeTable(int(num_rows), int(num_cols)),
                left=left,
                top=top,
                width=width,
                height=height,
            )
        )

    def AddChart2(
        self,
        style: int,
        chart_type: int,
        left: float,
        top: float,
        width: float,
        height: float,
    ) -> _FakeShape:
        sid = self._next_id()
        sh = _FakeShape(
            name=f"Chart {sid}",
            shape_id=sid,
            shape_type=_MSO_CHART,
            text=None,  # a chart shape has no text frame
            left=left,
            top=top,
            width=width,
            height=height,
        )
        sh._chart = _FakeChart(int(chart_type))
        return self._adopt(sh)

    def AddSmartArt(
        self, layout: Any, left: float, top: float, width: float, height: float
    ) -> _FakeShape:
        sid = self._next_id()
        sh = _FakeShape(
            name=f"Diagram {sid}",
            shape_id=sid,
            shape_type=_MSO_SMARTART,
            text=None,  # a SmartArt shape has no text frame
            left=left,
            top=top,
            width=width,
            height=height,
        )
        sh._smartart = _FakeSmartArt(layout.Id)
        return self._adopt(sh)

    @property
    def Application(self) -> _FakeApplication | None:
        return self._app

    @property
    def Count(self) -> int:
        return len(self._shapes)

    def __call__(self, index: int) -> _FakeShape:
        if index < 1 or index > len(self._shapes):
            raise IndexError(index)
        return self._shapes[index - 1]

    def __iter__(self) -> Any:
        return iter(self._shapes)

    @staticmethod
    def _is_title(sh: _FakeShape) -> bool:
        return sh._placeholder_type in (_PH_TITLE, _PH_CENTER_TITLE)

    @property
    def HasTitle(self) -> int:
        return _MSO_TRUE if any(self._is_title(s) for s in self._shapes) else _MSO_FALSE

    @property
    def Title(self) -> _FakeShape:
        for s in self._shapes:
            if self._is_title(s):
                return s
        raise AttributeError("slide has no title")

    def Range(self, names: Any) -> _FakeShapeRange:
        wanted = list(names) if isinstance(names, (list, tuple)) else [names]
        selected = [s for s in self._shapes if s.Name in wanted]
        assert self._app is not None
        return _FakeShapeRange(self._app, selected)


# ---------------------------------------------------------------------------
# Comments (v1.3): threaded, slide-attached, identity-bound (Add2 / legacy Add)
# ---------------------------------------------------------------------------
#
# The fake mirrors the live findings the wrapper relies on: a comment carries
# Author/AuthorInitials/Text/DateTime/Left/Top plus a ProviderID/UserID identity,
# and a `.Replies` collection whose `Add2` appends a threaded child. `Add2` binds
# to the passed identity (the wrapper sources it off an existing comment); the
# legacy identity-free `Add` is the fallback on a deck with no comment to source
# from. `Delete()` removes the comment (and, being a list slice, its replies).

_SEED_DT = datetime(2026, 6, 7, 10, 30, tzinfo=UTC)


class _FakeComment:
    """One review comment (or reply). Lives in a parent Python list so `Delete`
    can remove it; `Replies` wraps its own child list."""

    def __init__(
        self,
        *,
        container: list[_FakeComment],
        text: str,
        author: str = "",
        initials: str = "",
        provider: str = "",
        user: str = "",
        left: float = 12.0,
        top: float = 12.0,
        dt: datetime | None = None,
        is_reply: bool = False,
    ) -> None:
        self._container = container
        self.Text = text
        self.Author = author
        self.AuthorInitials = initials
        self.ProviderID = provider
        self.UserID = user
        self.Left = float(left)
        self.Top = float(top)
        self.DateTime = dt if dt is not None else _SEED_DT
        self._replies: list[_FakeComment] = []
        self._is_reply = is_reply

    @property
    def Replies(self) -> _FakeReplies:
        # A top-level comment exposes its own reply list; a *reply* exposes the
        # sibling list it belongs to (which includes itself) — the self-referential
        # live behavior that makes a naive thread walk recurse forever.
        return _FakeReplies(self._container if self._is_reply else self._replies)

    def Delete(self) -> None:
        self._container.remove(self)


class _FakeReplies:
    """A comment's `.Replies` — 1-based `Item`/callable + `Count` + `Add2`."""

    def __init__(self, replies: list[_FakeComment]) -> None:
        self._replies = replies

    @property
    def Count(self) -> int:
        return len(self._replies)

    def Item(self, index: int) -> _FakeComment:
        return self._replies[int(index) - 1]

    def __call__(self, index: int) -> _FakeComment:
        return self._replies[int(index) - 1]

    def Add2(
        self,
        left: float,
        top: float,
        author: str,
        initials: str,
        text: str,
        provider: str = "",
        user: str = "",
    ) -> _FakeComment:
        rep = _FakeComment(
            container=self._replies,
            text=text,
            author=author,
            initials=initials,
            provider=provider,
            user=user,
            left=left,
            top=top,
            is_reply=True,
        )
        self._replies.append(rep)
        return rep


class _FakeCommentCollection:
    """`Slide.Comments` — 1-based `Item`/callable + `Count` + modern `Add2` /
    legacy `Add`. Records `last_add_method` so a test can assert which path ran."""

    def __init__(self, comments: list[_FakeComment] | None = None) -> None:
        self._comments: list[_FakeComment] = comments if comments is not None else []
        self.last_add_method: str | None = None

    @property
    def Count(self) -> int:
        return len(self._comments)

    def Item(self, index: int) -> _FakeComment:
        return self._comments[int(index) - 1]

    def __call__(self, index: int) -> _FakeComment:
        return self._comments[int(index) - 1]

    def Add2(
        self,
        left: float,
        top: float,
        author: str,
        initials: str,
        text: str,
        provider: str = "",
        user: str = "",
    ) -> _FakeComment:
        self.last_add_method = "Add2"
        c = _FakeComment(
            container=self._comments,
            text=text,
            author=author,
            initials=initials,
            provider=provider,
            user=user,
            left=left,
            top=top,
        )
        self._comments.append(c)
        return c

    def Add(self, left: float, top: float, author: str, initials: str, text: str) -> _FakeComment:
        self.last_add_method = "Add"
        c = _FakeComment(
            container=self._comments,
            text=text,
            author=author,
            initials=initials,
            left=left,
            top=top,
        )
        self._comments.append(c)
        return c


def _seeded_comments() -> list[_FakeComment]:
    """One parent comment with a threaded reply — both carrying a real identity,
    so reads/threads and identity-sourced Add2 are exercised on the default deck."""
    comments: list[_FakeComment] = []
    parent = _FakeComment(
        container=comments,
        text="Tighten this headline.",
        author="Thomas Villani",
        initials="TV",
        provider="AD",
        user="S::tom@example.com::abc-123",
    )
    parent._replies.append(
        _FakeComment(
            container=parent._replies,
            text="Agreed — will do.",
            author="Thomas Villani",
            initials="TV",
            provider="AD",
            user="S::tom@example.com::abc-123",
            is_reply=True,
        )
    )
    comments.append(parent)
    return comments


# ---------------------------------------------------------------------------
# Slide, notes page, layouts
# ---------------------------------------------------------------------------


class _FakeCustomLayout:
    """A `CustomLayout` — a name plus (for reset_to_layout) its placeholders."""

    def __init__(self, name: str, placeholders: list[_FakeShape] | None = None) -> None:
        self.Name = name
        self.Shapes = SimpleNamespace(Placeholders=_FakePlaceholders(placeholders or []))


class _FakeCustomLayouts:
    """`SlideMaster.CustomLayouts` — iterable + 1-based callable."""

    def __init__(self, names: tuple[str, ...]) -> None:
        self._layouts = [_FakeCustomLayout(n) for n in names]

    @property
    def Count(self) -> int:
        return len(self._layouts)

    def __call__(self, index: int) -> _FakeCustomLayout:
        return self._layouts[index - 1]

    def __iter__(self) -> Any:
        return iter(self._layouts)


class _FakeSlideRange:
    """What `Slide.Duplicate()` returns — a 1-based callable over the new slides."""

    def __init__(self, slides: list[_FakeSlide]) -> None:
        self._slides = slides

    @property
    def Count(self) -> int:
        return len(self._slides)

    def __call__(self, index: int) -> _FakeSlide:
        return self._slides[index - 1]

    def __iter__(self) -> Any:
        return iter(self._slides)


class _FakeNotesPage:
    def __init__(self, body_text: str | None) -> None:
        placeholders: list[_FakeShape] = []
        if body_text is not None:
            placeholders.append(
                _FakeShape(
                    name="Notes Placeholder 2",
                    shape_id=2,
                    shape_type=_MSO_PLACEHOLDER,
                    text=body_text,
                    placeholder_type=_PH_BODY,
                )
            )
        self.Shapes = SimpleNamespace(Placeholders=_FakePlaceholders(placeholders))


class _FakePlaceholders:
    def __init__(self, items: list[_FakeShape]) -> None:
        self._items = items

    def __iter__(self) -> Any:
        return iter(self._items)

    @property
    def Count(self) -> int:
        return len(self._items)

    def __call__(self, index: int) -> _FakeShape:
        return self._items[index - 1]


class _FakeSlideTransition:
    """`Slide.SlideShowTransition` — entry effect + duration + advance model.

    Defaults match the live spike (scripts/transition_spike.py): no effect,
    click-to-advance on, timed-advance off.
    """

    def __init__(self) -> None:
        self.EntryEffect = 0  # ppEffectNone
        self.Duration = 0.0
        self.AdvanceOnClick = _MSO_TRUE
        self.AdvanceOnTime = _MSO_FALSE
        self.AdvanceTime = 0.0


class _FakeSlide:
    """A slide. `SlideIndex` is derived from list position so add/move/delete
    shift indices the way real PowerPoint does; `SlideID` stays stable."""

    def __init__(
        self,
        *,
        slide_id: int,
        layout_name: str,
        shapes: list[_FakeShape],
        notes_text: str | None = "",
        comments: list[_FakeComment] | None = None,
        layout_placeholders: list[_FakeShape] | None = None,
    ) -> None:
        self.SlideID = slide_id
        self.Shapes = _FakeShapes(shapes)
        self.CustomLayout = _FakeCustomLayout(layout_name, layout_placeholders)
        self._notes_text = notes_text
        self.NotesPage = _FakeNotesPage(notes_text)
        self.Comments = _FakeCommentCollection(comments)
        self.SlideShowTransition = _FakeSlideTransition()
        # Per-slide background: inherits the master by default (msoTrue), with its
        # own solid-fill object for when an override is set.
        self.FollowMasterBackground = _MSO_TRUE
        self.Background = SimpleNamespace(Fill=_FakeFillFormat())
        self._collection: _FakeSlides | None = None
        self._app: _FakeApplication | None = None

    @property
    def SlideIndex(self) -> int:
        assert self._collection is not None
        return self._collection._slides.index(self) + 1

    def Delete(self) -> None:
        assert self._collection is not None
        self._collection._slides.remove(self)

    def MoveTo(self, to_pos: int) -> None:
        assert self._collection is not None
        slides = self._collection._slides
        slides.remove(self)
        slides.insert(int(to_pos) - 1, self)

    def Duplicate(self) -> _FakeSlideRange:
        assert self._collection is not None
        clone = _FakeSlide(
            slide_id=self._collection._next_id(),
            layout_name=self.CustomLayout.Name,
            shapes=[sh._clone() for sh in self.Shapes._shapes],
            notes_text=self._notes_text,
        )
        slides = self._collection._slides
        slides.insert(slides.index(self) + 1, clone)
        self._collection._adopt(clone)
        return _FakeSlideRange([clone])

    def Export(
        self,
        FileName: str,
        FilterName: str,
        ScaleWidth: int | None = None,
        ScaleHeight: int | None = None,
    ) -> None:
        """Render to a stub PNG. Honors ScaleWidth/Height; else the slide's native
        pixel size at 96 DPI (so a 960x540 pt slide -> 1280x720, matching real COM)."""
        if ScaleWidth and ScaleHeight:
            w, h = int(ScaleWidth), int(ScaleHeight)
        else:
            w, h = 1280, 720
            try:
                ps = self._app.ActivePresentation.PageSetup  # type: ignore[union-attr]
                w = int(round(float(ps.SlideWidth) * 96 / 72))
                h = int(round(float(ps.SlideHeight) * 96 / 72))
            except Exception:
                pass
        self.last_export = {"FileName": FileName, "FilterName": FilterName, "Width": w, "Height": h}
        with open(FileName, "wb") as fh:
            fh.write(_minimal_png(w, h))


class _FakeSlides:
    def __init__(self, slides: list[_FakeSlide]) -> None:
        self._slides = slides
        self._app: _FakeApplication | None = None
        self._id_counter = max((s.SlideID for s in slides), default=255)
        for s in slides:
            self._adopt(s)

    def _adopt(self, slide: _FakeSlide) -> None:
        """Wire a slide's back-refs so it knows its collection + app."""
        slide._collection = self
        slide._app = self._app
        slide.Shapes._app = self._app

    def _next_id(self) -> int:
        self._id_counter += 1
        return self._id_counter

    @property
    def Count(self) -> int:
        return len(self._slides)

    def __call__(self, index: int) -> _FakeSlide:
        if index < 1 or index > len(self._slides):
            raise IndexError(index)
        return self._slides[index - 1]

    def __iter__(self) -> Any:
        return iter(self._slides)

    def AddSlide(self, index: int, custom_layout: _FakeCustomLayout) -> _FakeSlide:
        slide = _FakeSlide(
            slide_id=self._next_id(),
            layout_name=str(custom_layout.Name),
            shapes=[],
            notes_text="",
        )
        self._slides.insert(int(index) - 1, slide)
        self._adopt(slide)
        return slide

    def Add(self, index: int, pp_layout: int) -> _FakeSlide:
        """Legacy `Slides.Add` — only reached on a deck with no CustomLayouts."""
        slide = _FakeSlide(
            slide_id=self._next_id(),
            layout_name=f"Layout {int(pp_layout)}",
            shapes=[],
            notes_text="",
        )
        self._slides.insert(int(index) - 1, slide)
        self._adopt(slide)
        return slide


# ---------------------------------------------------------------------------
# Live slide show (v0.6): SlideShowSettings -> SlideShowWindow.View
# ---------------------------------------------------------------------------


class _FakeSlideShowView:
    """`SlideShowView` — the running show's controller. `State` is read/write;
    Next past the last slide ends the show, as PowerPoint does."""

    def __init__(self, pres: _FakePresentation, start_pos: int) -> None:
        self._pres = pres
        self._pos = max(1, int(start_pos))
        self.State = _SHOW_RUNNING

    @property
    def CurrentShowPosition(self) -> int:
        return self._pos

    @property
    def Slide(self) -> _FakeSlide:
        return self._pres.Slides(self._pos)

    def Next(self) -> None:
        if self._pos >= self._pres.Slides.Count:
            self._pres._end_show()  # advancing past the last slide ends the show
        else:
            self._pos += 1

    def Previous(self) -> None:
        if self._pos > 1:
            self._pos -= 1

    def GotoSlide(self, index: int, reset: Any = None) -> None:
        self._pos = int(index)

    def First(self) -> None:
        self._pos = 1

    def Last(self) -> None:
        self._pos = self._pres.Slides.Count

    def Exit(self) -> None:
        self._pres._end_show()


class _FakeSlideShowWindow:
    def __init__(self, pres: _FakePresentation, start_pos: int) -> None:
        self._pres = pres
        self.View = _FakeSlideShowView(pres, start_pos)

    @property
    def Presentation(self) -> _FakePresentation:
        return self._pres


class _FakeSlideShowSettings:
    """`SlideShowSettings` — `.Run()` starts the show at StartingSlide when the
    RangeType is a slide range, else from the top."""

    def __init__(self, pres: _FakePresentation) -> None:
        self._pres = pres
        self.StartingSlide = 1
        self.EndingSlide = pres.Slides.Count
        self.RangeType = 1  # ppShowAll

    def Run(self) -> _FakeSlideShowWindow:
        start = int(self.StartingSlide) if self.RangeType == _RANGE_SLIDE else 1
        return self._pres._start_show(start)


# ---------------------------------------------------------------------------
# Slide master: text styles + theme (palette/fonts) + background  (v0.9)
# ---------------------------------------------------------------------------


class _FakeThemeColorScheme:
    """`Theme.ThemeColorScheme` — `Colors(1..12)` each carry a writable `.RGB`."""

    def __init__(self) -> None:
        # Seed 12 distinct longs so reads are deterministic and writes round-trip.
        self._colors = [SimpleNamespace(RGB=(i * 0x111111)) for i in range(1, 13)]

    def Colors(self, index: int) -> Any:
        return self._colors[int(index) - 1]


class _FakeThemeFont:
    """`MajorFont`/`MinorFont` — `.Item(1=Latin/2=EastAsian/3=ComplexScript).Name`."""

    def __init__(self, latin: str) -> None:
        self._scripts = [
            SimpleNamespace(Name=latin),
            SimpleNamespace(Name=""),
            SimpleNamespace(Name=""),
        ]

    def Item(self, index: int) -> Any:
        return self._scripts[int(index) - 1]


class _FakeThemeFontScheme:
    def __init__(self) -> None:
        self.MajorFont = _FakeThemeFont("Calibri Light")
        self.MinorFont = _FakeThemeFont("Calibri")


class _FakeTheme:
    def __init__(self) -> None:
        self.ThemeColorScheme = _FakeThemeColorScheme()
        self.ThemeFontScheme = _FakeThemeFontScheme()


class _FakeTextStyleLevel:
    def __init__(self) -> None:
        self.Font = _FakeFont()
        self.ParagraphFormat = _FakeParagraphFormat()


class _FakeTextStyle:
    """One `TextStyles(t)` — `.Levels(1..5)`, each a Font + ParagraphFormat."""

    def __init__(self) -> None:
        self._levels = [_FakeTextStyleLevel() for _ in range(5)]

    def Levels(self, index: int) -> _FakeTextStyleLevel:
        return self._levels[int(index) - 1]


class _FakeFillFormat:
    """`Background.Fill` — solid-fill subset (Type + ForeColor.RGB + Solid())."""

    def __init__(self) -> None:
        self.Type = 5  # msoFillBackground (inherits) until Solid() is called
        self.ForeColor = SimpleNamespace(RGB=0)

    def Solid(self) -> None:
        self.Type = 1  # msoFillSolid


class _FakeSlideMaster:
    """`Presentation.SlideMaster` — layouts + text styles + theme + background."""

    def __init__(self, layout_names: tuple[str, ...]) -> None:
        self.CustomLayouts = _FakeCustomLayouts(layout_names)
        # 1=default, 2=title, 3=body (PpTextStyleType).
        self._text_styles = {1: _FakeTextStyle(), 2: _FakeTextStyle(), 3: _FakeTextStyle()}
        self.Theme = _FakeTheme()
        self.Background = SimpleNamespace(Fill=_FakeFillFormat())

    def TextStyles(self, style_type: int) -> _FakeTextStyle:
        return self._text_styles[int(style_type)]


class _FakePresentation:
    def __init__(
        self,
        *,
        name: str,
        full_name: str,
        slides: list[_FakeSlide],
        slide_width: float = 960.0,
        slide_height: float = 540.0,
        layout_names: tuple[str, ...] = _STANDARD_LAYOUTS,
    ) -> None:
        self.Name = name
        self.FullName = full_name
        # MsoTriState: -1 (msoTrue) = no unsaved changes, 0 (msoFalse) = dirty.
        # A freshly-opened deck is clean; tests dirty it via `pres.Saved = 0`.
        self.Saved = -1
        self.Slides = _FakeSlides(slides)
        self.PageSetup = SimpleNamespace(SlideWidth=slide_width, SlideHeight=slide_height)
        self.SlideMaster = _FakeSlideMaster(layout_names)
        self._show_settings = _FakeSlideShowSettings(self)
        self._show_window: _FakeSlideShowWindow | None = None

    @property
    def Path(self) -> str:
        """The deck's folder — `""` for a never-saved deck (bare `FullName`).

        Mirrors real COM: `Presentation.Path` is the directory of `FullName`, or
        empty when the deck has never been saved (its `FullName` is just a name
        like ``"Presentation1"``). `save()` reads this to refuse a path-less save.
        """
        full = str(self.FullName)
        return os.path.dirname(full) if (os.sep in full or "/" in full) else ""

    def Save(self) -> None:
        """Persist to the existing file — here, just clear the dirty flag.

        The wrapper guards on an empty `Path` before calling this, so a fake
        never-saved `Save()` is never exercised through `deck.save()`.
        """
        self.Saved = -1

    def SaveAs(self, FileName: str, FileFormat: int = 24) -> None:  # noqa: N803 (COM name)
        """Write a stand-in file at `FileName`; rebind for pptx, export for PDF.

        Mirrors the 2026-06-09 spike's verified behavior: SaveAs to the Open XML
        format (24) rebinds the working file (`FullName`/`Path` follow) and marks
        the deck clean; SaveAs to PDF (32) writes the PDF but leaves `FullName`,
        `Path`, and the dirty flag untouched (a pure export).
        """
        fmt = int(FileFormat)
        if fmt == 32:  # ppSaveAsPDF — export, no rebind, dirty flag preserved
            with open(FileName, "wb") as fh:
                fh.write(b"%PDF-1.4\n%fake pptlive pdf\n")
            return
        with open(FileName, "wb") as fh:  # pptx-family — a zip signature
            fh.write(b"PK\x03\x04fake pptlive pptx")
        self.FullName = str(FileName)
        self.Saved = -1

    @property
    def SlideShowSettings(self) -> _FakeSlideShowSettings:
        return self._show_settings

    @property
    def SlideShowWindow(self) -> _FakeSlideShowWindow:
        # Real COM raises when no show is running; the wrapper treats any failure
        # here as "not running", so raising keeps the fake honest.
        if self._show_window is None:
            raise RuntimeError("no slide show is running")
        return self._show_window

    def _start_show(self, start_pos: int) -> _FakeSlideShowWindow:
        self._show_window = _FakeSlideShowWindow(self, start_pos)
        return self._show_window

    def _end_show(self) -> None:
        self._show_window = None


# ---------------------------------------------------------------------------
# Application + window/view/selection
# ---------------------------------------------------------------------------


class _FakeView:
    def __init__(self, app: _FakeApplication) -> None:
        self._app = app

    @property
    def Slide(self) -> _FakeSlide:
        pres = self._app.ActivePresentation
        return pres.Slides(self._app._viewed)

    def GotoSlide(self, index: int) -> None:
        self._app._viewed = int(index)


class _FakeSelection:
    def __init__(self, app: _FakeApplication) -> None:
        self._app = app

    @property
    def Type(self) -> int:
        return self._app._selection_type

    def _slide(self) -> _FakeSlide:
        pres = self._app.ActivePresentation
        return pres.Slides(self._app._viewed)

    @property
    def ShapeRange(self) -> _FakeShapeRange:
        slide = self._slide()
        if self._app._selection_type == _SEL_TEXT and self._app._text_shape is not None:
            names: tuple[str, ...] = (self._app._text_shape,)
        else:
            names = self._app._selected_names
        shapes = [s for s in slide.Shapes if s.Name in names]
        return _FakeShapeRange(self._app, shapes)

    @property
    def TextRange(self) -> Any:
        """The selected text run — a paragraph caret. Exposes `Start` (1-based
        char offset) and `Text`, the two fields read_selection reads."""
        slide = self._slide()
        host = next(s for s in slide.Shapes if s.Name == self._app._text_shape)
        frame = host.TextFrame
        para = self._app._text_para
        start = sum(len(p.text) + 1 for p in frame._paras[: para - 1]) + 1
        return SimpleNamespace(Start=start, Text=frame.TextRange.Paragraphs(para, 1).Text)

    def SlideRange(self, index: int = 1) -> _FakeSlide:
        return self._slide()

    def Unselect(self) -> None:
        self._app._selection_type = _SEL_NONE
        self._app._selected_names = ()
        self._app._text_shape = None


class _FakeWindow:
    def __init__(self, app: _FakeApplication) -> None:
        self.View = _FakeView(app)
        self.Selection = _FakeSelection(app)


class _FakePresentations:
    def __init__(self, presentations: list[_FakePresentation]) -> None:
        self._presentations = presentations

    @property
    def Count(self) -> int:
        return len(self._presentations)

    def __iter__(self) -> Any:
        return iter(self._presentations)

    def __call__(self, index: int) -> _FakePresentation:
        return self._presentations[index - 1]


class _FakeApplication:
    def __init__(self, presentations: list[_FakePresentation]) -> None:
        self._presentations = presentations
        self._viewed = 1
        self._selection_type = _SEL_NONE
        self._selected_names: tuple[str, ...] = ()
        self._text_shape: str | None = None  # host shape name for a TEXT selection
        self._text_para = 1  # 1-based paragraph the text caret is in
        self.Visible = True
        self._window = _FakeWindow(self)
        self._undo_entries = 0  # count of StartNewUndoEntry() calls (edit() fences)
        # Wire back-refs now the app exists: Shapes.Range(...).Select() updates the
        # selection, and each slide knows its collection + app for lifecycle verbs.
        for pres in presentations:
            pres.Slides._app = self
            for slide in list(pres.Slides._slides):
                pres.Slides._adopt(slide)

    def StartNewUndoEntry(self) -> None:
        """Mirror PowerPoint's boundary primitive; edit() calls this on entry."""
        self._undo_entries += 1

    # -- test helpers: drive the live Selection (what read_selection reads) -----

    def _select_shapes(self, *names: str) -> None:
        self._selection_type = _SEL_SHAPES
        self._selected_names = tuple(names)
        self._text_shape = None

    def _select_text(self, shape_name: str, paragraph: int = 1) -> None:
        self._selection_type = _SEL_TEXT
        self._text_shape = shape_name
        self._text_para = int(paragraph)

    def _select_slide(self) -> None:
        self._selection_type = _SEL_SLIDES
        self._selected_names = ()
        self._text_shape = None

    @property
    def Presentations(self) -> _FakePresentations:
        return _FakePresentations(self._presentations)

    @property
    def ActivePresentation(self) -> _FakePresentation:
        if not self._presentations:
            raise RuntimeError("no active presentation")
        return self._presentations[0]

    @property
    def ActiveWindow(self) -> _FakeWindow:
        return self._window

    @property
    def SmartArtLayouts(self) -> _FakeSmartArtLayouts:
        return _FakeSmartArtLayouts()


# ---------------------------------------------------------------------------
# Builders
# ---------------------------------------------------------------------------


def _default_deck() -> _FakePresentation:
    """A 3-slide deck exercising the v0 surface.

    Slide 1 (Title Slide): center-title "Welcome" + subtitle, with notes.
    Slide 2 (Title and Content): title "Agenda", body with 3 bullets, a picture
        (no text frame), no notes.
    Slide 3 (Blank): a text box and a line (no text frame), no title/body.
    """
    slide1 = _FakeSlide(
        slide_id=256,
        layout_name="Title Slide",
        notes_text="Lead with the vision.",
        comments=_seeded_comments(),
        shapes=[
            _FakeShape(
                name="Title 1",
                shape_id=2,
                shape_type=_MSO_PLACEHOLDER,
                text="Welcome",
                placeholder_type=_PH_CENTER_TITLE,
            ),
            _FakeShape(
                name="Subtitle 2",
                shape_id=3,
                shape_type=_MSO_PLACEHOLDER,
                text="A demo deck",
                placeholder_type=_PH_SUBTITLE,
            ),
        ],
    )
    # The layout's body placeholder carries the geometry + default font size
    # reset_to_layout restores from (the spike's live values for this layout).
    _layout_body = _FakeShape(
        name="Content Placeholder 2",
        shape_id=2,
        shape_type=_MSO_PLACEHOLDER,
        text="",
        placeholder_type=_PH_BODY,
        left=66.0,
        top=143.75,
        width=828.0,
        height=342.625,
    )
    _layout_body.TextFrame.TextRange.Font.Size = 28.0
    slide2 = _FakeSlide(
        slide_id=257,
        layout_name="Title and Content",
        notes_text="",
        layout_placeholders=[_layout_body],
        shapes=[
            _FakeShape(
                name="Title 1",
                shape_id=2,
                shape_type=_MSO_PLACEHOLDER,
                text="Agenda",
                placeholder_type=_PH_TITLE,
            ),
            _FakeShape(
                name="Content Placeholder 2",
                shape_id=3,
                shape_type=_MSO_PLACEHOLDER,
                text="Intro\rDemo\rQ&A",
                placeholder_type=_PH_BODY,
            ),
            _FakeShape(
                name="Picture 3",
                shape_id=4,
                shape_type=_MSO_PICTURE,
                text=None,
                left=400.0,
                top=120.0,
                width=300.0,
                height=200.0,
            ),
        ],
    )
    slide3 = _FakeSlide(
        slide_id=258,
        layout_name="Blank",
        notes_text=None,  # no notes body placeholder at all
        shapes=[
            _FakeShape(
                name="TextBox 1",
                shape_id=2,
                shape_type=_MSO_TEXT_BOX,
                text="Free text",
            ),
            _FakeShape(
                name="Line 2",
                shape_id=3,
                shape_type=_MSO_LINE,
                text=None,
            ),
        ],
    )
    return _FakePresentation(
        name="Pitch.pptx",
        full_name=r"C:\decks\Pitch.pptx",
        slides=[slide1, slide2, slide3],
    )


@pytest.fixture
def fake_powerpoint(monkeypatch: pytest.MonkeyPatch) -> _FakeApplication:
    """A fake PowerPoint.Application with the default 3-slide deck open."""
    app = _FakeApplication([_default_deck()])

    from pptlive import _com

    monkeypatch.setattr(_com, "get_active_powerpoint", lambda: app)
    monkeypatch.setattr(_com, "launch_powerpoint", lambda: app)
    return app


@pytest.fixture
def fake_powerpoint_same_named_decks(monkeypatch: pytest.MonkeyPatch) -> _FakeApplication:
    """Two open decks sharing a display `Name` but with distinct `FullName`s.

    Models the real collision the `--doc` selector must disambiguate: two files
    called `Deck.pptx` opened from different folders. The first is the active deck.
    """
    first = _default_deck()
    first.Name = "Deck.pptx"
    first.FullName = r"C:\\a\\Deck.pptx"
    second = _default_deck()
    second.Name = "Deck.pptx"
    second.FullName = r"C:\\b\\Deck.pptx"
    app = _FakeApplication([first, second])

    from pptlive import _com

    monkeypatch.setattr(_com, "get_active_powerpoint", lambda: app)
    monkeypatch.setattr(_com, "launch_powerpoint", lambda: app)
    return app


@pytest.fixture
def ppt(fake_powerpoint: _FakeApplication):  # type: ignore[no-untyped-def]
    """A `pptlive.PowerPoint` handle attached to the fake app (context held open)."""
    import pptlive

    with pptlive.attach() as handle:
        yield handle


@pytest.fixture
def deck(ppt: Any):  # type: ignore[no-untyped-def]
    """The active `Presentation` of the fake deck."""
    return ppt.presentations.active


@pytest.fixture
def no_powerpoint(monkeypatch: pytest.MonkeyPatch) -> None:
    """Simulate PowerPoint not running."""
    from pptlive import _com
    from pptlive.exceptions import PowerPointNotRunningError

    def _raise(*args: Any, **kwargs: Any) -> None:
        raise PowerPointNotRunningError("no running Microsoft PowerPoint instance found")

    monkeypatch.setattr(_com, "get_active_powerpoint", _raise)
    monkeypatch.setattr(_com, "launch_powerpoint", _raise)


@pytest.fixture
def real_powerpoint():  # type: ignore[no-untyped-def]
    """Smoke fixture: yields a pptlive.PowerPoint, or skips if PowerPoint isn't reachable."""
    import pptlive
    from pptlive.exceptions import PowerPointNotRunningError

    try:
        ctx = pptlive.attach()
        ppt = ctx.__enter__()
    except PowerPointNotRunningError as e:
        pytest.skip(f"PowerPoint not running: {e}")
        return
    try:
        yield ppt
    finally:
        ctx.__exit__(None, None, None)

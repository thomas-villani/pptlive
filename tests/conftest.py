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

from types import SimpleNamespace
from typing import Any

import pytest

# PpPlaceholderType values used below (kept literal so the fake doesn't depend on
# the module under test for its constants).
_PH_TITLE = 1
_PH_BODY = 2
_PH_CENTER_TITLE = 3
_PH_SUBTITLE = 4

# MsoShapeType values used below.
_MSO_AUTO_SHAPE = 1
_MSO_LINE = 9
_MSO_PICTURE = 13
_MSO_PLACEHOLDER = 14
_MSO_TEXT_BOX = 17

# MsoTriState
_MSO_TRUE = -1
_MSO_FALSE = 0

# PpSelectionType
_SEL_NONE = 0
_SEL_SHAPES = 2

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


class _FakeTextRange:
    def __init__(self, text: str = "") -> None:
        self.Text = text


class _FakeTextFrame:
    def __init__(self, text: str = "") -> None:
        self.TextRange = _FakeTextRange(text)


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
        self._placeholder_type = placeholder_type
        self._text_frame = _FakeTextFrame(text) if text is not None else None
        self.selected = False

    @property
    def HasTextFrame(self) -> int:
        return _MSO_TRUE if self._text_frame is not None else _MSO_FALSE

    @property
    def TextFrame(self) -> _FakeTextFrame:
        if self._text_frame is None:
            raise AttributeError("shape has no text frame")
        return self._text_frame

    @property
    def PlaceholderFormat(self) -> Any:
        if self._placeholder_type is None:
            raise AttributeError("shape is not a placeholder")
        return SimpleNamespace(Type=self._placeholder_type)

    def Select(self, *args: Any, **kwargs: Any) -> None:
        self.selected = True

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


class _FakeShapeRange:
    def __init__(self, app: _FakeApplication, shapes: list[_FakeShape]) -> None:
        self._app = app
        self._shapes = shapes

    def __iter__(self) -> Any:
        return iter(self._shapes)

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
# Slide, notes page, layouts
# ---------------------------------------------------------------------------


class _FakeCustomLayout:
    """A `CustomLayout` — just a name, as far as resolution cares."""

    def __init__(self, name: str) -> None:
        self.Name = name


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
    ) -> None:
        self.SlideID = slide_id
        self.Shapes = _FakeShapes(shapes)
        self.CustomLayout = _FakeCustomLayout(layout_name)
        self._notes_text = notes_text
        self.NotesPage = _FakeNotesPage(notes_text)
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
        self.Slides = _FakeSlides(slides)
        self.PageSetup = SimpleNamespace(SlideWidth=slide_width, SlideHeight=slide_height)
        self.SlideMaster = SimpleNamespace(CustomLayouts=_FakeCustomLayouts(layout_names))


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

    @property
    def ShapeRange(self) -> list[_FakeShape]:
        pres = self._app.ActivePresentation
        slide = pres.Slides(self._app._viewed)
        return [s for s in slide.Shapes if s.Name in self._app._selected_names]

    def Unselect(self) -> None:
        self._app._selection_type = _SEL_NONE
        self._app._selected_names = ()


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
    slide2 = _FakeSlide(
        slide_id=257,
        layout_name="Title and Content",
        notes_text="",
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

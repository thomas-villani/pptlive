"""Slide-level reads: list, read, outline, page setup, title, has_notes."""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from pptlive._slides import Slide
from pptlive.exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    PowerPointBusyError,
    SlideNotFoundError,
)

_MSO_PLACEHOLDER = 14
_PH_TITLE, _PH_BODY, _PH_OBJECT = 1, 2, 7


def _ph_shape(name: str, shape_id: int, ph_type: int) -> SimpleNamespace:
    """A minimal duck-typed placeholder COM shape for `_find_placeholder`."""
    return SimpleNamespace(
        Name=name,
        Id=shape_id,
        Type=_MSO_PLACEHOLDER,
        PlaceholderFormat=SimpleNamespace(Type=ph_type),
    )


def _slide_with(*shapes: SimpleNamespace, index: int = 5) -> Slide:
    return Slide(None, SimpleNamespace(Shapes=list(shapes), SlideIndex=index))  # type: ignore[arg-type]


class _Boom:
    """Every attribute access (and call) raises a transient busy error.

    Stands in for a COM object PowerPoint momentarily refuses to serve, so a read
    that touches it has the chance to either re-raise (correct) or swallow the
    busy as a soft default (the bug these tests guard against).
    """

    def __getattr__(self, _name: str) -> object:
        raise PowerPointBusyError(hresult=0x80010001)

    def __call__(self, *_args: object, **_kw: object) -> object:
        raise PowerPointBusyError(hresult=0x80010001)


def test_slides_list_shape(deck) -> None:  # type: ignore[no-untyped-def]
    rows = deck.slides.list()
    assert [r["index"] for r in rows] == [1, 2, 3]
    assert [r["id"] for r in rows] == [256, 257, 258]
    by_index = {r["index"]: r for r in rows}
    assert by_index[1]["title"] == "Welcome"
    assert by_index[1]["layout"] == "Title Slide"
    assert by_index[1]["has_notes"] is True
    assert by_index[2]["has_notes"] is False
    assert by_index[2]["shape_count"] == 3
    assert by_index[3]["title"] is None


def test_title_surfaces_busy_instead_of_none(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    slide.com.Shapes = _Boom()
    with pytest.raises(PowerPointBusyError):
        _ = slide.title


def test_layout_name_surfaces_busy_instead_of_none(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    slide.com.CustomLayout = _Boom()
    with pytest.raises(PowerPointBusyError):
        _ = slide.layout_name


# -- placeholder resolution + ambiguity guard (PPTLIVE-004) -------------------


def test_find_placeholder_single_object_resolves() -> None:
    slide = _slide_with(
        _ph_shape("Title 1", 2, _PH_TITLE),
        _ph_shape("Content Placeholder 2", 3, _PH_OBJECT),
    )
    sh, idx = slide._find_placeholder("body")
    assert (sh.Name, idx) == ("Content Placeholder 2", 2)


def test_find_placeholder_prefers_body_over_object_by_rank() -> None:
    # A real BODY (rank 0) wins over a generic OBJECT (rank 1) — different ranks,
    # so it is NOT ambiguous.
    slide = _slide_with(
        _ph_shape("Body 1", 2, _PH_BODY),
        _ph_shape("Content Placeholder 2", 3, _PH_OBJECT),
    )
    sh, idx = slide._find_placeholder("body")
    assert (sh.Name, idx) == ("Body 1", 1)


def test_find_placeholder_two_objects_is_ambiguous() -> None:
    # Two Content layout: two OBJECT placeholders share the best rank → ambiguous.
    slide = _slide_with(
        _ph_shape("Title 1", 2, _PH_TITLE),
        _ph_shape("Content Placeholder 2", 3, _PH_OBJECT),
        _ph_shape("Content Placeholder 3", 4, _PH_OBJECT),
    )
    with pytest.raises(AmbiguousMatchError) as exc:
        slide._find_placeholder("body")
    # The error lists the candidate shape anchors so the caller can pick one.
    anchors = {m["anchor_id"] for m in exc.value.matches}
    assert anchors == {"shape:5:2", "shape:5:3"}
    assert "shape:5:2" in str(exc.value) and "shape:5:3" in str(exc.value)


def test_find_placeholder_missing_raises_not_found() -> None:
    slide = _slide_with(_ph_shape("Title 1", 2, _PH_TITLE))
    with pytest.raises(AnchorNotFoundError):
        slide._find_placeholder("body")


def test_has_notes_surfaces_busy_instead_of_false(deck) -> None:  # type: ignore[no-untyped-def]
    slide = deck.slides[1]
    slide.com.NotesPage = _Boom()
    with pytest.raises(PowerPointBusyError):
        slide.has_notes()


def test_find_placeholder_surfaces_busy_on_type_read(deck) -> None:  # type: ignore[no-untyped-def]
    # A placeholder whose PlaceholderFormat.Type read goes busy must surface as
    # PowerPointBusyError, not be swallowed by the per-shape skip (which would
    # mis-report the placeholder as absent → AnchorNotFoundError).
    slide = deck.slides[2]
    boom_ph = SimpleNamespace(Type=_MSO_PLACEHOLDER, PlaceholderFormat=_Boom(), Name="Boom", Id=999)
    slide.com.Shapes = [boom_ph]
    with pytest.raises(PowerPointBusyError):
        slide.placeholder("title")


def test_slide_indexing_is_one_based(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[1].index == 1
    assert len(deck.slides) == 3


def test_slide_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.slides[0]
    with pytest.raises(SlideNotFoundError):
        deck.slides[4]


def test_slide_read_payload(deck) -> None:  # type: ignore[no-untyped-def]
    grid = deck.slides[2].read()
    assert grid["index"] == 2
    assert grid["id"] == 257
    assert grid["layout"] == "Title and Content"
    assert grid["title"] == "Agenda"
    shapes = grid["shapes"]
    assert [s["anchor_id"] for s in shapes] == ["shape:2:1", "shape:2:2", "shape:2:3"]
    assert shapes[0]["placeholder"] == "title"
    assert shapes[1]["placeholder"] == "body"
    assert shapes[2]["type"] == "picture"
    assert shapes[2]["has_text_frame"] is False
    assert shapes[2]["text"] is None
    # geometry is in points
    assert shapes[2]["geometry"] == {
        "left": 400.0,
        "top": 120.0,
        "width": 300.0,
        "height": 200.0,
        "rotation": 0.0,
    }


def test_outline(deck) -> None:  # type: ignore[no-untyped-def]
    items = deck.outline()
    assert items[1] == {"slide": 2, "title": "Agenda", "bullets": ["Intro", "Demo", "Q&A"]}
    # Title slide has no body placeholder -> no bullets, but keeps its title.
    assert items[0]["title"] == "Welcome"
    assert items[0]["bullets"] == []
    assert items[2]["title"] is None


def test_outline_tolerates_textless_body_placeholder(deck) -> None:  # type: ignore[no-untyped-def]
    # A body placeholder filled with a chart/table/picture has no text frame, so
    # reading its `.text` raises NoTextFrameError. outline() must skip it and keep
    # going, not crash the whole deck overview (regression: it used to exit 6).
    body = deck.slides[2].shapes[2].com  # the "Content Placeholder 2" (body)
    body._text_frame = None  # simulate the placeholder now holding a chart
    items = deck.outline()
    assert items[1]["title"] == "Agenda"
    assert items[1]["bullets"] == []  # no bullets, but no crash either
    assert items[0]["title"] == "Welcome"  # other slides unaffected


def test_page_setup_points(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.page_setup() == {"width": 960.0, "height": 540.0}


def test_iteration_yields_slides_in_order(deck) -> None:  # type: ignore[no-untyped-def]
    assert [s.index for s in deck.slides] == [1, 2, 3]

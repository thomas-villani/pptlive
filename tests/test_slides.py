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


def test_find_placeholder_ordinal_picks_nth_in_z_order() -> None:
    # ph:S:body:1 / ph:S:body:2 — the two columns of a Two Content slide, no
    # ambiguity error, no slide.read() + name-filter dance.
    slide = _slide_with(
        _ph_shape("Title 1", 2, _PH_TITLE),
        _ph_shape("Content Placeholder 2", 3, _PH_OBJECT),
        _ph_shape("Content Placeholder 3", 4, _PH_OBJECT),
    )
    assert slide._find_placeholder("body", ordinal=1)[0].Name == "Content Placeholder 2"
    sh, idx = slide._find_placeholder("body", ordinal=2)
    assert (sh.Name, idx) == ("Content Placeholder 3", 3)
    with pytest.raises(AnchorNotFoundError):
        slide._find_placeholder("body", ordinal=3)
    with pytest.raises(AnchorNotFoundError):
        slide._find_placeholder("body", ordinal=0)


def test_find_placeholder_ordinal_ranks_preferred_type_first() -> None:
    # Comparison layout: BODY (rank 0) placeholders come before OBJECT (rank 1)
    # ones regardless of z-order, then z-order within a rank.
    slide = _slide_with(
        _ph_shape("Title 1", 2, _PH_TITLE),
        _ph_shape("Content Placeholder 2", 3, _PH_OBJECT),
        _ph_shape("Text Placeholder 3", 4, _PH_BODY),
    )
    assert slide._find_placeholder("body", ordinal=1)[0].Name == "Text Placeholder 3"
    assert slide._find_placeholder("body", ordinal=2)[0].Name == "Content Placeholder 2"


def test_ambiguous_placeholder_error_names_the_ordinal_form() -> None:
    slide = _slide_with(
        _ph_shape("Content Placeholder 2", 3, _PH_OBJECT),
        _ph_shape("Content Placeholder 3", 4, _PH_OBJECT),
    )
    with pytest.raises(AmbiguousMatchError) as exc:
        slide._find_placeholder("body")
    assert "ph:5:body:1" in str(exc.value) and "ph:5:body:2" in str(exc.value)


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


# ---------------------------------------------------------------------------
# The one-op slide: slides.add(title=/body=/content=/notes=)
# ---------------------------------------------------------------------------


def test_add_with_content_fills_placeholders_and_notes(deck) -> None:  # type: ignore[no-untyped-def]
    new = deck.slides.add(
        layout="two_content",
        title="Q3 results",
        content={"body:1": ["Revenue up 12%", "Churn flat"], "body:2": "Right\ncolumn"},
        notes="Speaker notes",
    )
    s = new.index
    assert deck.anchor_by_id(f"ph:{s}:title").text == "Q3 results"
    assert deck.anchor_by_id(f"ph:{s}:body:1").text == "Revenue up 12%\rChurn flat"
    assert deck.anchor_by_id(f"ph:{s}:body:2").text == "Right\rcolumn"
    assert new.notes.text == "Speaker notes"


def test_add_body_shorthand_and_paragraph_items(deck) -> None:  # type: ignore[no-untyped-def]
    new = deck.slides.add(
        layout="title_and_content",
        title="Agenda",
        body=[{"text": "Intro", "bold": True}, "Demo"],
    )
    assert deck.anchor_by_id(f"ph:{new.index}:body").text == "Intro\rDemo"


def test_add_content_validates_before_any_com(deck) -> None:  # type: ignore[no-untyped-def]
    before = len(deck.slides)
    bad = [
        dict(body=[]),
        dict(body=[{"bold": True}]),
        dict(title=3),
        dict(content={"bogus": "x"}),
        dict(content={"body:0": "x"}),
        dict(content={"body:1:1": "x"}),
        dict(content={"body": "x"}, body="y"),
        dict(notes=1),
    ]
    for kwargs in bad:
        with pytest.raises(ValueError):
            deck.slides.add(layout="title_and_content", **kwargs)
    assert len(deck.slides) == before  # nothing was added


def test_add_content_ambiguous_kind_leaves_new_slide_untouched(deck) -> None:  # type: ignore[no-untyped-def]
    # `body` on a Two Content slide is ambiguous; the resolve-all-first rule means
    # the title is NOT written either, even though it comes first.
    with pytest.raises(AmbiguousMatchError):
        deck.slides.add(layout="two_content", title="T", body="ambiguous")
    new = deck.slides[len(deck.slides)]
    assert new.layout_name == "Two Content"
    assert deck.anchor_by_id(f"ph:{new.index}:title").text == ""


def test_add_content_unknown_on_layout_raises_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides.add(layout="title_only", body="no body here")

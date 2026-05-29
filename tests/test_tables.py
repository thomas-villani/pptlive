"""Tables (v0.5): add_table, the Table wrapper, and Cell anchors (cell:S:N:R:C).

Exercised against the fake COM graph — a table is a shape that satisfies
`HasTable`, and a cell's text lives in its own text frame
(`Table.Cell(r,c).Shape.TextFrame.TextRange`).
"""

from __future__ import annotations

import pytest

from pptlive import Cell, Shape
from pptlive.exceptions import AnchorNotFoundError


def _add_table(deck, rows=2, cols=3):  # type: ignore[no-untyped-def]
    """Add a `rows`×`cols` table to slide 3 (Blank) and return the Shape."""
    with deck.edit("test: add table"):
        return deck.slides[3].shapes.add_table(rows, cols)


def test_add_table_returns_table_shape(deck) -> None:  # type: ignore[no-untyped-def]
    sh = _add_table(deck, 2, 3)
    assert isinstance(sh, Shape)
    assert sh.has_table is True
    # The default-deck slide 3 had 2 shapes, so the table lands at z-order 3.
    assert sh.anchor_id == "shape:3:3"


def test_add_table_rejects_nonpositive(deck) -> None:  # type: ignore[no-untyped-def]
    # ValueError is raised before any COM.
    with pytest.raises(ValueError):
        deck.slides[3].shapes.add_table(0, 3)


def test_table_dimensions(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 3).table
    assert table.row_count == 2
    assert table.column_count == 3


def test_cell_text_round_trips(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 3).table
    with deck.edit("test: fill cell"):
        table.cell(1, 2).set_text("hello")
    assert table.cell(1, 2).text == "hello"
    assert isinstance(table.cell(1, 2), Cell)


def test_cell_anchor_id(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 3).table
    assert table.cell(2, 3).anchor_id == "cell:3:3:2:3"


def test_cell_out_of_range_is_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 3).table
    with pytest.raises(AnchorNotFoundError):
        table.cell(3, 1)
    with pytest.raises(AnchorNotFoundError):
        table.cell(1, 4)


def test_grid(deck) -> None:  # type: ignore[no-untyped-def]
    sh = _add_table(deck, 2, 2)
    table = sh.table
    with deck.edit("test: fill grid"):
        table.cell(1, 1).set_text("a")
        table.cell(1, 2).set_text("b")
        table.cell(2, 1).set_text("c")
        table.cell(2, 2).set_text("d")
    assert table.grid() == [["a", "b"], ["c", "d"]]


def test_read_emits_cells_with_anchors(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 2).table
    with deck.edit("test: fill"):
        table.cell(1, 1).set_text("R1C1")
    grid = table.read()
    assert grid["slide"] == 3
    assert grid["shape"] == 3
    assert grid["anchor_id"] == "shape:3:3"
    assert grid["rows"] == 2 and grid["columns"] == 2
    assert grid["cells"][0][0] == {
        "row": 1,
        "col": 1,
        "text": "R1C1",
        "anchor_id": "cell:3:3:1:1",
    }


def test_add_row_appends_and_fills(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 1, 2).table
    with deck.edit("test: add row"):
        table.add_row(["x", "y"])
    assert table.row_count == 2
    assert table.cell(2, 1).text == "x"
    assert table.cell(2, 2).text == "y"


def test_add_row_extra_values_ignored(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 1, 2).table
    with deck.edit("test: add row"):
        table.add_row(["x", "y", "z-ignored"])
    assert table.grid()[-1] == ["x", "y"]


def test_delete_row(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 3, 1).table
    with deck.edit("test: fill rows"):
        table.cell(1, 1).set_text("one")
        table.cell(2, 1).set_text("two")
        table.cell(3, 1).set_text("three")
    with deck.edit("test: delete row 2"):
        table.delete_row(2)
    assert table.row_count == 2
    assert table.grid() == [["one"], ["three"]]


def test_delete_row_out_of_range(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 1).table
    with pytest.raises(AnchorNotFoundError):
        table.delete_row(5)


def test_delete_last_row_is_rejected(deck) -> None:  # type: ignore[no-untyped-def]
    # PowerPoint has no zero-row table; deleting the only row would corrupt the
    # shape, so a one-row table refuses delete_row with a clear ValueError.
    table = _add_table(deck, 1, 2).table
    with pytest.raises(ValueError, match="last row"):
        table.delete_row(1)
    assert table.row_count == 1


def test_anchor_by_id_resolves_cell(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 2).table
    with deck.edit("test: fill"):
        table.cell(2, 2).set_text("corner")
    cell = deck.anchor_by_id("cell:3:3:2:2")
    assert isinstance(cell, Cell)
    assert cell.text == "corner"


def test_anchor_by_id_writes_through_cell(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 2, 2).table
    with deck.edit("test: write cell"):
        deck.anchor_by_id("cell:3:3:1:1").set_text("written")
    assert table.cell(1, 1).text == "written"


def test_cell_inherits_format_text(deck) -> None:  # type: ignore[no-untyped-def]
    table = _add_table(deck, 1, 1).table
    with deck.edit("test: bold cell"):
        table.cell(1, 1).set_text("bold me")
        table.cell(1, 1).format_text(bold=True)
    # The cell's text-frame font carries the formatting (read via the COM seam).
    assert table.cell(1, 1).com.Font.Bold == -1  # msoTrue


def test_anchor_by_id_cell_no_table_is_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    # shape:3:1 is the TextBox on slide 3 — it has no table.
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id("cell:3:1:1:1")


def test_shape_table_raises_without_table(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[3].shapes[1].table  # the TextBox, no table


def test_shape_read_reports_has_table(deck) -> None:  # type: ignore[no-untyped-def]
    _add_table(deck, 2, 2)
    rows = deck.slides[3].shapes.list()
    assert rows[-1]["has_table"] is True
    assert rows[0]["has_table"] is False  # the TextBox


def test_bad_cell_anchor_id_shapes(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.anchor_by_id("cell:3:3")  # too few parts

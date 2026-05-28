"""Tables — the `Table` wrapper over a shape, and the `Cell` anchor.

Unlike wordlive (where tables are a document-scoped collection), a PowerPoint
table is a **shape on a slide**: the shape satisfies `Shape.HasTable` and exposes
the grid via `Shape.Table`. So there is no deck-wide `tables` collection — a table
is reached through its shape (`slide.shapes[N].table`, or the `cell:S:N:R:C`
anchor), and the `Table` here is bound to a `Shape`, re-resolving its COM object
live so z-order drift is handled exactly as it is for the shape.

A `Cell` *is* an `Anchor`, exactly like wordlive: it targets the cell's own text
frame (`Table.Cell(r, c).Shape.TextFrame.TextRange`), so the inherited
`set_text` / `format_text` / `format_paragraph` / `apply_list` /
`insert_paragraph_*` machinery works on a cell with no special-casing, and
`anchor_by_id("cell:S:N:R:C")` resolves through `Presentation.anchor_by_id` like
any other anchor. PowerPoint cell text is a plain text frame (paragraphs split by
`\\r`, no Word end-of-cell markers), so `Cell.text` needs no extra stripping.

The anchor-id scheme is `cell:S:N:R:C` (slide S, shape N by z-order, 1-based row
R, column C). `shape:S:N` addresses the table shape itself (geometry, delete);
the bare slide/shape forms are not cell anchors.

Spike (verified live 2026-05-28, `scripts/table_spike.py`, net-zero):
`Shapes.AddTable` returns a shape whose `Type` may report **placeholder** (14),
not table (19), when it fills a content placeholder — so `HasTable` is the only
reliable gate. `Table.Cell(r, c).Shape.TextFrame.TextRange.Text` round-trips
(multi-paragraph `\\r` preserved); `Rows.Add()` appends one row (the new last row
is addressable); `Rows(n).Delete()` removes it; an out-of-range `Cell` raises, so
cell access is bounds-checked here for a clean `AnchorNotFoundError`.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide


class Cell(Anchor):
    """A single table cell, addressed by 1-based (row, column) — `cell:S:N:R:C`.

    Subclasses `Anchor`, so it inherits every text verb (`set_text`,
    `format_text`, `format_paragraph`, `apply_list`, `insert_paragraph_*`)
    unchanged. `_text_range()` is the cell's own text frame, re-resolved live (the
    grid can drift as rows are added/deleted), raising `AnchorNotFoundError` if
    (row, col) falls outside the current grid.
    """

    kind = "cell"

    def __init__(self, table: Table, row: int, col: int) -> None:
        self._table = table
        self._row = int(row)
        self._col = int(col)

    @property
    def table(self) -> Table:
        return self._table

    @property
    def slide(self) -> Slide:
        return self._table.shape.slide

    @property
    def row(self) -> int:
        return self._row

    @property
    def column(self) -> int:
        return self._col

    @property
    def anchor_id(self) -> str:
        return f"cell:{self._table.shape.slide.index}:{self._table.shape.index}:{self._row}:{self._col}"

    def _cell_com(self) -> Any:
        """The COM `Cell` at (row, col), bounds-checked against the live grid."""
        com_table = self._table.com
        rows = int(com_table.Rows.Count)
        cols = int(com_table.Columns.Count)
        if not (1 <= self._row <= rows and 1 <= self._col <= cols):
            raise AnchorNotFoundError("table cell", self.anchor_id)
        return com_table.Cell(self._row, self._col)

    def _text_range(self) -> Any:
        return self._cell_com().Shape.TextFrame.TextRange

    def to_dict(self) -> dict[str, Any]:
        with _com.translate_com_errors():
            return {
                "row": self._row,
                "col": self._col,
                "text": str(self._text_range().Text or ""),
                "anchor_id": self.anchor_id,
            }


class Table:
    """A table on a slide, bound to its `Shape` (over `Shape.Table`).

    Re-resolves the COM table live through the shape, so it survives z-order
    drift. A table has no anchor of its own — `shape:S:N` addresses the table
    shape (geometry/delete) and `cell:S:N:R:C` addresses a cell — so `Table`
    exposes structural reads + row edits, while text edits go through `Cell`.
    """

    def __init__(self, shape: Shape) -> None:
        self._shape = shape

    @property
    def com(self) -> Any:
        """Raw COM `Table` (`Shape.Table`), resolved live."""
        with _com.translate_com_errors():
            return self._shape.com.Table

    @property
    def shape(self) -> Shape:
        return self._shape

    @property
    def row_count(self) -> int:
        with _com.translate_com_errors():
            return int(self._shape.com.Table.Rows.Count)

    @property
    def column_count(self) -> int:
        with _com.translate_com_errors():
            return int(self._shape.com.Table.Columns.Count)

    def cell(self, row: int, col: int) -> Cell:
        """Return the `Cell` at 1-based (row, col).

        Raises `AnchorNotFoundError` (kind `"table cell"`) if the coordinates fall
        outside the table's grid.
        """
        rows, cols = self.row_count, self.column_count
        if not (1 <= int(row) <= rows and 1 <= int(col) <= cols):
            raise AnchorNotFoundError(
                "table cell", f"cell:{self._shape.slide.index}:{self._shape.index}:{row}:{col}"
            )
        return Cell(self, int(row), int(col))

    def grid(self) -> list[list[str]]:
        """All cell text as a row-major `list[list[str]]`."""
        rows, cols = self.row_count, self.column_count
        return [[self.cell(r, c).text for c in range(1, cols + 1)] for r in range(1, rows + 1)]

    def read(self) -> dict[str, Any]:
        """Structured dump: dimensions plus every cell with its addressable id.

        Each cell carries its `anchor_id` (`cell:S:N:R:C`) so a caller can feed it
        straight back into `write` / `format-text` / `format-paragraph`.
        """
        rows, cols = self.row_count, self.column_count
        cells = [
            [self.cell(r, c).to_dict() for c in range(1, cols + 1)] for r in range(1, rows + 1)
        ]
        return {
            "slide": self._shape.slide.index,
            "shape": self._shape.index,
            "anchor_id": self._shape.anchor_id,
            "rows": rows,
            "columns": cols,
            "cells": cells,
        }

    def add_row(self, values: list[Any] | None = None) -> None:
        """Append a row at the end of the table, optionally filling its cells.

        `values` are matched to columns left-to-right; extras past the column
        count are ignored, short lists leave trailing cells empty. Wrap in
        `deck.edit(...)` for view preservation + a one-Ctrl-Z fence.
        """
        with _com.translate_com_errors():
            com_table = self._shape.com.Table
            com_table.Rows.Add()
            if values:
                last = int(com_table.Rows.Count)
                cols = int(com_table.Columns.Count)
                for c, val in enumerate(values, start=1):
                    if c > cols:
                        break
                    com_table.Cell(last, c).Shape.TextFrame.TextRange.Text = str(val)

    def delete_row(self, index: int) -> None:
        """Delete the 1-based row `index`.

        Raises `AnchorNotFoundError` (kind `"table row"`) if out of range.
        """
        rows = self.row_count
        if not (1 <= int(index) <= rows):
            raise AnchorNotFoundError(
                "table row", f"cell:{self._shape.slide.index}:{self._shape.index}:row:{index}"
            )
        with _com.translate_com_errors():
            self._shape.com.Table.Rows(int(index)).Delete()

    def __iter__(self) -> Iterator[Cell]:
        """Iterate cells row-major."""
        rows, cols = self.row_count, self.column_count
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                yield Cell(self, r, c)

    def __repr__(self) -> str:
        return f"<Table {self._shape.anchor_id} {self.row_count}x{self.column_count}>"

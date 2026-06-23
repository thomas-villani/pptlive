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

from collections.abc import Iterator, Sequence
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor
from ._shapes import _is_none_token, apply_shape_fill
from .constants import (
    MsoTriState,
    border_edges_for,
    color_hex_or_none,
    dash_style_for,
    parse_color,
)
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide

# A cell index selector: None (the whole axis), one 1-based index, or a list of them.
AxisSelector = int | Sequence[int] | None


def _cell_fill_hex(com_cell: Any) -> str | None:
    """The cell's **effective** solid-fill color as `#RRGGBB`, or `None`.

    `None` means a hidden fill (`Fill.Visible == msoFalse`) or the theme sentinel
    (`0x80000000`) — never a wrong `#000000`, matching the `color_hex_or_none` guard
    used for shape fills. Caveat (live-verified): a default PowerPoint **table
    style** writes a real per-cell banded RGB, so a fresh, never-touched cell reads
    back that *style* color, not `None`. There is no COM flag distinguishing a fill
    set on the cell from one cascaded by the table style (it's OOXML-only), so this
    reports what is actually rendered — direct or style-inherited alike.
    """
    fill = com_cell.Shape.Fill
    if int(fill.Visible) == int(MsoTriState.FALSE):
        return None
    return color_hex_or_none(fill.ForeColor.RGB)


def _apply_cell_border(
    border: Any,
    *,
    rgb: int | None,
    hide: bool,
    weight: float | None,
    dash_int: int | None,
    visible: bool | None,
) -> None:
    """Write one resolved `Borders(index)` edge — caller wraps in `translate_com_errors`.

    `hide` (`color="none"`) takes the edge invisible; an explicit `rgb` turns it on
    and colors it. `weight`/`dash_int` set the geometry; `visible` forces the on/off
    state last (so it wins over the color-implied visibility when both are given).
    """
    if hide:
        border.Visible = int(MsoTriState.FALSE)
    elif rgb is not None:
        border.Visible = int(MsoTriState.TRUE)
        border.ForeColor.RGB = rgb
    if weight is not None:
        border.Weight = float(weight)
    if dash_int is not None:
        border.DashStyle = dash_int
    if visible is not None:
        border.Visible = int(MsoTriState.TRUE if visible else MsoTriState.FALSE)


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
            com_cell = self._cell_com()
            return {
                "row": self._row,
                "col": self._col,
                "text": str(com_cell.Shape.TextFrame.TextRange.Text or ""),
                "fill": _cell_fill_hex(com_cell),
                "anchor_id": self.anchor_id,
            }

    def set_fill(
        self,
        fill: str | int | tuple[int, int, int],
        *,
        transparency: float | None = None,
    ) -> int:
        """Solid-fill this cell (cell shading) — or clear it with `fill="none"`.

        Thin per-cell wrapper over `Table.set_fill` (which is where the row/column
        bulk form lives). `fill` is a color (`"#RRGGBB"` / `(r, g, b)` / raw RGB int)
        or the string `"none"` (transparent — inherits the table style's shading);
        `transparency` is a `0.0..1.0` alpha fraction. A mutation: wrap in
        `deck.edit(...)`. Returns the number of cells filled (always 1).
        """
        return self._table.set_fill(fill, rows=self._row, cols=self._col, transparency=transparency)

    def set_border(
        self,
        *,
        color: str | int | tuple[int, int, int] | None = None,
        weight: float | None = None,
        dash: str | int | None = None,
        edges: str | int | Sequence[str | int] = "all",
        visible: bool | None = None,
    ) -> int:
        """Style this cell's border(s) — color / weight / dash / visibility.

        Thin per-cell wrapper over `Table.set_border`. `edges` selects which edges
        (`"all"` = the four sides, or `"top"`/`"bottom"`/`"left"`/`"right"`/
        `"diagonal_down"`/`"diagonal_up"`, one or a list); `color` is a color or
        `"none"` (hide that edge). A mutation: wrap in `deck.edit(...)`. Returns the
        number of cells touched (always 1).
        """
        return self._table.set_border(
            rows=self._row,
            cols=self._col,
            color=color,
            weight=weight,
            dash=dash,
            edges=edges,
            visible=visible,
        )


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

        Merged cells are **not** flagged: PowerPoint's COM `Cell` exposes no
        merge-state read property (it lives only in the OOXML, off the automation
        surface), so the grid is reported by its raw row×column geometry. A merged
        region still occupies every covered coordinate, and writing to a covered
        cell lands on the merge origin's text — address merged regions by their
        top-left (origin) cell to avoid surprises.
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

        Raises `AnchorNotFoundError` (kind `"table row"`) if out of range, and
        `ValueError` if it would empty the table — PowerPoint has no zero-row
        table, and `Rows(1).Delete()` on a one-row grid corrupts the shape rather
        than failing cleanly. Delete the whole table shape (`shape:S:N`) instead.
        """
        rows = self.row_count
        if not (1 <= int(index) <= rows):
            raise AnchorNotFoundError(
                "table row", f"cell:{self._shape.slide.index}:{self._shape.index}:row:{index}"
            )
        if rows <= 1:
            raise ValueError(
                "cannot delete the last row of a table; delete the table shape "
                f"({self._shape.anchor_id}) instead"
            )
        with _com.translate_com_errors():
            self._shape.com.Table.Rows(int(index)).Delete()

    def add_column(self, values: list[Any] | None = None, *, before: int | None = None) -> None:
        """Add a column, optionally filling its cells — appended by default.

        `before` (1-based) inserts the new column *before* that existing column;
        omitting it appends at the right edge (verified live: `Columns.Add()`
        appends, `Columns.Add(n)` inserts before column `n`). `values` are matched
        to rows top-to-bottom; extras past the row count are ignored, short lists
        leave trailing cells empty. Raises `AnchorNotFoundError` (kind
        `"table column"`) for an out-of-range `before`. Wrap in `deck.edit(...)`
        for view preservation + a one-Ctrl-Z fence.
        """
        if before is not None:
            cols = self.column_count
            if not (1 <= int(before) <= cols):
                raise AnchorNotFoundError(
                    "table column",
                    f"cell:{self._shape.slide.index}:{self._shape.index}:column:{before}",
                )
        with _com.translate_com_errors():
            com_table = self._shape.com.Table
            if before is None:
                com_table.Columns.Add()
                target = int(com_table.Columns.Count)
            else:
                com_table.Columns.Add(int(before))
                target = int(before)
            if values:
                rows = int(com_table.Rows.Count)
                for r, val in enumerate(values, start=1):
                    if r > rows:
                        break
                    com_table.Cell(r, target).Shape.TextFrame.TextRange.Text = str(val)

    def delete_column(self, index: int) -> None:
        """Delete the 1-based column `index`.

        Raises `AnchorNotFoundError` (kind `"table column"`) if out of range, and
        `ValueError` if it would empty the table — PowerPoint has no zero-column
        table, and deleting the only column corrupts the shape rather than failing
        cleanly. Delete the whole table shape (`shape:S:N`) instead.
        """
        cols = self.column_count
        if not (1 <= int(index) <= cols):
            raise AnchorNotFoundError(
                "table column",
                f"cell:{self._shape.slide.index}:{self._shape.index}:column:{index}",
            )
        if cols <= 1:
            raise ValueError(
                "cannot delete the last column of a table; delete the table shape "
                f"({self._shape.anchor_id}) instead"
            )
        with _com.translate_com_errors():
            self._shape.com.Table.Columns(int(index)).Delete()

    def _resolve_axis(self, sel: AxisSelector, count: int, axis: str) -> list[int]:
        """Normalize a row/column selector to a validated list of 1-based indices.

        `None` selects the whole axis; an int selects one; a sequence selects each.
        Raises `AnchorNotFoundError` (kind `"table row"`/`"table column"`) for an
        out-of-range index — before any COM mutation.
        """
        if sel is None:
            return list(range(1, count + 1))
        # bool is an int subclass; reject it explicitly (True->1 would silently
        # select row/col 1) for consistency with parse_color / dash_style_for /
        # border_edges_for, which all guard the same footgun.
        if isinstance(sel, bool):
            raise ValueError(f"{axis} selector must be an int or list of ints, not a bool")
        raw: Sequence[int]
        if isinstance(sel, int):
            raw = [int(sel)]
        else:
            raw = [int(x) for x in sel]
        for i in raw:
            if not (1 <= i <= count):
                raise AnchorNotFoundError(
                    f"table {axis}",
                    f"cell:{self._shape.slide.index}:{self._shape.index}:{axis}:{i}",
                )
        return list(raw)

    def set_fill(
        self,
        fill: str | int | tuple[int, int, int],
        *,
        rows: AxisSelector = None,
        cols: AxisSelector = None,
        transparency: float | None = None,
    ) -> int:
        """Solid-fill (cell shading) a region of cells — or clear it with `fill="none"`.

        `rows`/`cols` select the region: `None` is the whole axis, an int one
        index, a list several. Their **intersection** is filled — so `rows=1` shades
        the header row, `cols=2` a column, `rows=1, cols=1` a single cell, and both
        `None` the whole table. `fill` is a color (`"#RRGGBB"` / `(r, g, b)` / raw
        RGB int) or `"none"` (transparent — falls back to the table style's shading).
        `transparency` is a `0.0..1.0` alpha fraction. Colors / indices are validated
        before any COM mutation. A mutation: wrap in `deck.edit(...)`. Returns the
        number of cells filled.
        """
        row_idx = self._resolve_axis(rows, self.row_count, "row")
        col_idx = self._resolve_axis(cols, self.column_count, "column")
        if not _is_none_token(fill):
            parse_color(fill)  # validate the color once, before any COM (like set_border)
        with _com.translate_com_errors():
            com_table = self._shape.com.Table
            count = 0
            for r in row_idx:
                for c in col_idx:
                    apply_shape_fill(
                        com_table.Cell(r, c).Shape, fill=fill, fill_transparency=transparency
                    )
                    count += 1
        return count

    def set_border(
        self,
        *,
        color: str | int | tuple[int, int, int] | None = None,
        weight: float | None = None,
        dash: str | int | None = None,
        edges: str | int | Sequence[str | int] = "all",
        rows: AxisSelector = None,
        cols: AxisSelector = None,
        visible: bool | None = None,
    ) -> int:
        """Style cell border(s) across a region — color / weight / dash / visibility.

        `rows`/`cols` select the region exactly like `set_fill` (intersection;
        `None` = whole axis). `edges` picks which edges of each cell: `"all"` (the
        four sides) or one/several of `"top"`/`"bottom"`/`"left"`/`"right"`/
        `"diagonal_down"`/`"diagonal_up"`. `color` is a color (turns the edge on) or
        `"none"` (hides it); `weight` is points; `dash` a friendly `MsoLineDashStyle`
        name; `visible` forces the edge on/off. At least one of
        `color`/`weight`/`dash`/`visible` is required. All names / colors / indices
        are validated before any COM mutation. A mutation: wrap in `deck.edit(...)`.
        Returns the number of cells touched.
        """
        if color is None and weight is None and dash is None and visible is None:
            raise ValueError(
                "set_border() requires at least one of color=, weight=, dash=, visible="
            )
        if weight is not None and weight < 0:  # validate before any COM
            raise ValueError(f"weight must be >= 0 points, got {weight}")
        edge_idx = border_edges_for(edges)  # ValueError before any COM
        hide = color is not None and _is_none_token(color)
        rgb = None if (color is None or hide) else parse_color(color)  # ValueError before COM
        dash_int = dash_style_for(dash) if dash is not None else None
        row_idx = self._resolve_axis(rows, self.row_count, "row")
        col_idx = self._resolve_axis(cols, self.column_count, "column")
        with _com.translate_com_errors():
            com_table = self._shape.com.Table
            count = 0
            for r in row_idx:
                for c in col_idx:
                    com_cell = com_table.Cell(r, c)
                    for idx in edge_idx:
                        _apply_cell_border(
                            com_cell.Borders(idx),
                            rgb=rgb,
                            hide=hide,
                            weight=weight,
                            dash_int=dash_int,
                            visible=visible,
                        )
                    count += 1
        return count

    def __iter__(self) -> Iterator[Cell]:
        """Iterate cells row-major."""
        rows, cols = self.row_count, self.column_count
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                yield Cell(self, r, c)

    def __repr__(self) -> str:
        return f"<Table {self._shape.anchor_id} {self.row_count}x{self.column_count}>"

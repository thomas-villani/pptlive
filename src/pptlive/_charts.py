"""Charts — the `Chart` wrapper over a chart shape and its embedded-Excel data.

Like a table, a PowerPoint chart is a **shape on a slide**: the shape satisfies
`Shape.HasChart` and exposes the chart via `Shape.Chart`. So there is no deck-wide
`charts` collection — a chart is reached through its shape (`slide.shapes[N].chart`
or `anchor_by_id("shape:S:N").chart`), and the `Chart` here is bound to a `Shape`,
re-resolving its COM object live so z-order drift is handled exactly as for the
shape. There is no new anchor scheme: `shape:S:N` addresses the chart shape (for
geometry / delete / export), and `.chart` reaches the data.

A chart's data lives in an **embedded Excel workbook**, not in the slide. So the
data verbs drive that workbook over COM. The sequence is the one the 2026-05-28
spike verified (`scripts/chart_spike.py`, net-zero) and it has two non-obvious
findings baked in:

1. **`SetSourceData` takes a STRING range, not a `Range` object.** Passing a
   `Range` raised `E_FAIL`; the string form (`"Sheet1!$A$1:$C$4"`) works.
2. **`SetSourceData` dissolves the default Excel Table (ListObject).** So we never
   rely on the ListObject to resize data — `ClearContents` wipes stale values and
   `SetSourceData` defines exactly the plotted range, which is the only thing the
   chart reads. This makes re-writes (shrink *and* grow) clean, verified live.

Editing chart data calls `ChartData.Activate()`, which briefly opens the embedded
workbook (a momentary Excel window) and is closed again here — that flicker is
inherent to PowerPoint chart automation, not a politeness regression.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from typing import TYPE_CHECKING, Any

from . import _com
from .constants import chart_type_for, chart_type_name

if TYPE_CHECKING:
    from ._shapes import Shape
    from ._slides import Slide

# A series spec accepted by set_data / add_chart: either a name->values mapping
# (the common case; insertion-ordered) or an ordered sequence of (name, values)
# pairs (when names repeat or order must be explicit).
SeriesInput = Mapping[str, Sequence[float]] | Sequence[tuple[str, Sequence[float]]]


def _col_letter(n: int) -> str:
    """1 -> A, 2 -> B, ... 27 -> AA (Excel column letters)."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _normalize_series(series: SeriesInput) -> list[tuple[str, list[float]]]:
    """Coerce a SeriesInput into an ordered `list[(name, [float, ...])]`."""
    items: list[tuple[str, Sequence[float]]]
    if isinstance(series, Mapping):
        items = list(series.items())
    else:
        items = [(name, values) for name, values in series]
    return [(str(name), [float(v) for v in values]) for name, values in items]


class Chart:
    """A chart on a slide, bound to its `Shape` — reached via `shape.chart`.

    ```
    chart = deck.slides[2].shapes[3].chart
    chart.read()                       # {chart_type, categories, series:[...]}
    chart.set_type("line")             # change the chart kind
    chart.set_data(["Q1", "Q2"], {"Revenue": [10, 20]})   # rewrite the data
    ```

    `read()` is side-effect-free. `set_type` / `set_data` mutate; wrap them in
    `deck.edit(...)` (as the CLI/MCP do) for view preservation + a one-Ctrl-Z
    fence. The COM chart is resolved live every call, so z-order drift on the
    host shape is handled.
    """

    def __init__(self, shape: Shape) -> None:
        self._shape = shape

    @property
    def com(self) -> Any:
        """Raw COM `Chart` (`Shape.Chart`), resolved live."""
        with _com.translate_com_errors():
            return self._shape.com.Chart

    @property
    def shape(self) -> Shape:
        return self._shape

    @property
    def slide(self) -> Slide:
        return self._shape.slide

    @property
    def chart_type(self) -> str:
        """Friendly chart-type name (e.g. "column_clustered", "line", "pie")."""
        with _com.translate_com_errors():
            return chart_type_name(self.com.ChartType)

    def set_type(self, chart_type: str | int) -> None:
        """Change the chart kind (`Chart.ChartType`).

        `chart_type` is a friendly name (`"column"`, `"bar"`, `"line"`, `"pie"`,
        …; see `constants.CHART_TYPE_CHOICES`) or a raw `XlChartType` int. Raises
        `ValueError` for an unknown name (before any COM). Wrap in `deck.edit(...)`.
        """
        type_int = chart_type_for(chart_type)  # ValueError before COM
        with _com.translate_com_errors():
            self.com.ChartType = type_int

    def categories(self) -> list[str]:
        """The category (X-axis) labels, as strings."""
        with _com.translate_com_errors():
            sc = self.com.SeriesCollection()
            if int(sc.Count) < 1:
                return []
            return [_as_label(v) for v in sc(1).XValues]

    def series(self) -> list[dict[str, Any]]:
        """Each series as `{"name": str, "values": [float, ...]}`, in plot order."""
        with _com.translate_com_errors():
            sc = self.com.SeriesCollection()
            out: list[dict[str, Any]] = []
            for i in range(1, int(sc.Count) + 1):
                s = sc(i)
                out.append({"name": str(s.Name), "values": [_as_number(v) for v in s.Values]})
            return out

    def read(self) -> dict[str, Any]:
        """Structured dump: chart kind + categories + series.

        `{slide, shape, anchor_id, chart_type, chart_type_code, categories,
        series:[{name, values}]}`. Side-effect-free.
        """
        with _com.translate_com_errors():
            chart = self.com
            type_code = int(chart.ChartType)
            sc = chart.SeriesCollection()
            count = int(sc.Count)
            cats = [_as_label(v) for v in sc(1).XValues] if count >= 1 else []
            series = [
                {"name": str(sc(i).Name), "values": [_as_number(v) for v in sc(i).Values]}
                for i in range(1, count + 1)
            ]
        return {
            "slide": self._shape.slide.index,
            "shape": self._shape.index,
            "anchor_id": self._shape.anchor_id,
            "chart_type": chart_type_name(type_code),
            "chart_type_code": type_code,
            "categories": cats,
            "series": series,
        }

    def set_data(self, categories: Sequence[str], series: SeriesInput) -> None:
        """Replace the chart's data with `categories` × `series`.

        `categories` are the X-axis labels; `series` is a name->values mapping
        (e.g. `{"Revenue": [10, 20, 30]}`) or an ordered sequence of
        `(name, values)` pairs. Every series must have exactly `len(categories)`
        values. Raises `ValueError` for empty inputs or a length mismatch (before
        any COM). Wrap in `deck.edit(...)` for the one-Ctrl-Z fence.

        Drives the chart's embedded Excel workbook: activate it, clear stale
        values, write the corner / series names (row 1) / categories (column A) /
        values, point `SetSourceData` at the exact range (string form), and close
        the workbook. See the module docstring for why no Excel Table is used.
        """
        cats = [str(c) for c in categories]
        norm = _normalize_series(series)
        if not cats:
            raise ValueError("set_data requires at least one category")
        if not norm:
            raise ValueError("set_data requires at least one series")
        for name, values in norm:
            if len(values) != len(cats):
                raise ValueError(
                    f"series {name!r} has {len(values)} values but there are {len(cats)} categories"
                )

        ncols = 1 + len(norm)
        nrows = 1 + len(cats)

        def _write() -> None:
            with _com.translate_com_errors():
                chart = self.com
                cdata = chart.ChartData
                cdata.Activate()
                wb = cdata.Workbook
                try:
                    ws = wb.Worksheets(1)
                    ws.UsedRange.ClearContents()
                    ws.Cells(1, 1).Value = ""  # corner
                    for c, (name, _values) in enumerate(norm, start=2):
                        ws.Cells(1, c).Value = name
                    for r, cat in enumerate(cats, start=2):
                        # Force the category column to Text, else Excel coerces a
                        # numeric-looking label ("2026") to a number and it reads
                        # back as "2026.0". Categories are labels, never values.
                        cell = ws.Cells(r, 1)
                        cell.NumberFormat = "@"
                        cell.Value = cat
                    for c, (_name, values) in enumerate(norm, start=2):
                        for r, v in enumerate(values, start=2):
                            ws.Cells(r, c).Value = v
                    target = f"Sheet1!$A$1:${_col_letter(ncols)}${nrows}"
                    chart.SetSourceData(target)
                finally:
                    wb.Close()

        # The embedded workbook can be momentarily unavailable right after the
        # chart is created (RPC_S_UNKNOWN_IF); the write is a clean rewrite, so
        # retrying the whole sequence is safe. See `_com.retry_on_busy`.
        _com.retry_on_busy(_write)

    def __repr__(self) -> str:
        return f"<Chart {self._shape.anchor_id} type={self.chart_type!r}>"


def _as_number(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_label(value: Any) -> str:
    return "" if value is None else str(value)

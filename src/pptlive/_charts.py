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
from .constants import XlAxisType, chart_type_for, chart_type_name, color_hex, parse_color
from .exceptions import PowerPointBusyError, PptliveError

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


def _sheet_ref(name: str) -> str:
    """Quote a worksheet name for an A1-style range reference.

    The first sheet's name is localized (English "Sheet1", but "Feuil1",
    "Tabelle1", "Hoja1", … on non-English Office), so we must reference it by its
    actual `.Name` rather than the literal "Sheet1". Quoting and doubling any
    embedded apostrophe keeps names with spaces or quotes valid.
    """
    return "'" + name.replace("'", "''") + "'"


def _normalize_series(series: SeriesInput) -> list[tuple[str, list[float]]]:
    """Coerce a SeriesInput into an ordered `list[(name, [float, ...])]`.

    Raises a `ValueError` naming the offending series and value when a value is
    not numeric (before any COM), instead of a bare `float()` message that
    doesn't say which series it came from.
    """
    items: list[tuple[str, Sequence[float]]]
    if isinstance(series, Mapping):
        items = list(series.items())
    else:
        items = [(name, values) for name, values in series]
    out: list[tuple[str, list[float]]] = []
    for name, values in items:
        sname = str(name)
        floats: list[float] = []
        for v in values:
            try:
                floats.append(float(v))
            except (TypeError, ValueError):
                raise ValueError(f"series {sname!r} value {v!r} is not a number") from None
        out.append((sname, floats))
    return out


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
        `(name, values)` pairs. Series are written — and plotted — in insertion
        order (note bar charts render series bottom-to-top by Excel/PowerPoint
        convention, so the first series sits at the bottom; this is a render
        order, not a reorder of the data). Every series must have exactly
        `len(categories)` values. Raises `ValueError` for empty inputs or a length
        mismatch (before any COM). Wrap in `deck.edit(...)` for the one-Ctrl-Z fence.

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
                    target = f"{_sheet_ref(ws.Name)}!$A$1:${_col_letter(ncols)}${nrows}"
                    chart.SetSourceData(target)
                finally:
                    wb.Close()
            # The embedded-Excel commit can SILENTLY race: a first-chance
            # RPC_S_CALL_FAILED (0x800706BE) during the workbook teardown leaves
            # SetSourceData bound to an empty range with *no* Python exception, so
            # the chart reads back blank (SeriesCollection empty). Verify the write
            # actually landed; if not, raise a retryable busy so retry_on_busy
            # re-runs the whole idempotent sequence (a fresh ChartData.Activate
            # re-establishes the workbook). Observed live ~50% of the time on the
            # chart smoke test before this guard.
            if not self._reflects_data(len(norm), len(cats)):
                raise PowerPointBusyError(
                    "chart data did not commit (embedded-Excel teardown raced); retrying"
                )

        # The embedded workbook can be momentarily unavailable right after the
        # chart is created (RPC_S_UNKNOWN_IF), and the commit above can race; the
        # write is a clean rewrite, so retrying the whole sequence is safe. See
        # `_com.retry_on_busy`.
        _com.retry_on_busy(_write, attempts=5)

    def _reflects_data(self, nseries: int, ncats: int) -> bool:
        """True if the plotted chart now reflects a just-written shape (count check).

        The minimal read-back `set_data` uses to detect the silent embedded-Excel
        commit race — the write no-ops with no exception, leaving `SeriesCollection`
        empty. Compares only counts (series count and the first series' category
        count), not values: the observed failure is *total* (an empty collection),
        so counts catch it without false-retrying a valid write whose labels happen
        to round-trip oddly. Any COM hiccup during the check counts as "not yet
        reflected", so the idempotent write is simply retried.
        """
        try:
            with _com.translate_com_errors():
                sc = self.com.SeriesCollection()
                if int(sc.Count) != nseries:
                    return False
                return len(list(sc(1).XValues)) == ncats
        except (PptliveError, TypeError, ValueError):
            # Any COM hiccup or a degenerate read-back (e.g. XValues coming back
            # as a non-iterable scalar) counts as "not yet reflected" — retry the
            # idempotent write rather than letting it escape the retry loop.
            return False

    def recolor_text(self, color: str | int | tuple[int, int, int]) -> dict[str, Any]:
        """Set the font color of **every shown** chart text element (one Ctrl-Z).

        A chart has no addressable text frame of its own — its text is split across
        the legend, axis tick labels, title, and data labels, each its own COM
        element — so there is no per-anchor `format_text` path. This is the coarse
        fix PPTLIVE-009 asked for: recolor all chart text at once, the move a dark-
        (or any custom-background) theme needs when the inherited black axis/legend
        text goes invisible.

        `color` is a `"#RRGGBB"` / `"RRGGBB"` hex string, an `(r, g, b)` tuple, or a
        raw RGB int (same forms as `format_text`'s `color`). Raises `ValueError` for
        a malformed color (before any COM). Recolors only the elements that are
        actually **present**: the legend/title are guarded by `HasLegend` /
        `HasTitle`; the axis tick labels and per-series data labels are best-effort
        (an absent axis — e.g. a pie chart has none — is simply skipped, not an
        error), so it never adds a legend, title, or labels the deck didn't show.
        Always sets the `ChartArea` global text default (the modern
        `Format.TextFrame2` path) so any text shown later inherits the color too.
        Wrap in `deck.edit(...)` for view preservation + the one-Ctrl-Z fence.
        Returns `{ok, slide, shape, anchor_id, color, recolored, series_data_labels}`
        where `recolored` lists the element kinds that were touched.

        The chart `Font` model differs from a text frame's: a chart element's
        `Font.Color` is *itself* the RGB long (set/read directly), not a
        `ColorFormat` with `.RGB`; and `Series.DataLabels` is a **method**.
        `Chart.HasAxis` is an Excel-ism PowerPoint's chart COM rejects, so axes are
        probed by attempting the set rather than asking first.

        The core recolor (chart area + legend/title/data labels) runs under
        `retry_on_busy`: every set is idempotent (recoloring to the same RGB twice
        is a no-op), so a transient busy mid-sequence retries the whole block rather
        than leaving a half-recolored chart — the same safety `set_data` uses. The
        axes stay best-effort outside that fence (they're already probe-and-skip).
        """
        rgb = parse_color(color)  # ValueError before any COM

        def _recolor_core() -> tuple[list[str], int]:
            recolored: list[str] = []
            labeled = 0
            with _com.translate_com_errors():
                chart = self.com
                # Global default for all chart text (modern TextFrame2 path);
                # reliably present, and covers text elements shown later.
                chart.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb
                recolored.append("chart_area")
                if bool(chart.HasLegend):
                    chart.Legend.Font.Color = rgb
                    recolored.append("legend")
                if bool(chart.HasTitle):
                    chart.ChartTitle.Font.Color = rgb
                    recolored.append("title")
                sc = chart.SeriesCollection()
                for i in range(1, int(sc.Count) + 1):
                    s = sc(i)
                    if bool(s.HasDataLabels):
                        s.DataLabels().Font.Color = rgb
                        labeled += 1
                if labeled:
                    recolored.append("data_labels")
            return recolored, labeled

        recolored, labeled = _com.retry_on_busy(_recolor_core)

        def _attempt_axis(axis_type: int) -> bool:
            """Set an axis's tick-label color; skip (False) if that axis is absent.

            `Chart.HasAxis` is unreliable on PowerPoint's chart COM, so we attempt
            the set and treat a COM failure (e.g. a pie chart has no axes) as
            "not present" rather than an error.
            """
            try:
                with _com.translate_com_errors():
                    self.com.Axes(axis_type).TickLabels.Font.Color = rgb
                return True
            except PowerPointBusyError:
                raise  # a transient busy is retryable — don't mask it as "axis absent"
            except PptliveError:
                return False

        for axis_type, name in (
            (XlAxisType.CATEGORY, "category_axis"),
            (XlAxisType.VALUE, "value_axis"),
        ):
            if _attempt_axis(axis_type):
                recolored.append(name)
        return {
            "ok": True,
            "slide": self._shape.slide.index,
            "shape": self._shape.index,
            "anchor_id": self._shape.anchor_id,
            "color": color_hex(rgb),
            "recolored": recolored,
            "series_data_labels": labeled,
        }

    def __repr__(self) -> str:
        return f"<Chart {self._shape.anchor_id} type={self.chart_type!r}>"


def _as_number(value: Any) -> float | None:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_label(value: Any) -> str:
    return "" if value is None else str(value)

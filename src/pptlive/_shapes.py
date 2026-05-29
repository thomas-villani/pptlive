"""Shapes — the 2-D objects on a slide, and the thing a `Shape` anchor IS.

A `Shape` is an `Anchor` (inherits `text`/`set_text`) *when it has a text frame*,
plus geometry verbs (`geometry`/`move`/`resize`/`delete`) that have no Word
analog. A text op on a frameless shape (picture, line) raises `NoTextFrameError`
(exit 6). New shapes come from `ShapeCollection.add_textbox`/`add_shape`/
`add_picture` (v0.2); like the slide-lifecycle verbs they only mutate, so wrap a
call in `deck.edit(label)` for view preservation + a one-Ctrl-Z fence.

z-order ids drift: `shape:S:N` is the 1-based z-order index, which shifts when
shapes are added or removed, so a `Shape` resolves its COM object **live** on
every access and never caches it (spec.md / Resolved Open Q #3). Listings emit
`name` (`Shape.Name`) and `id` (`Shape.Id`, stable across reorder) so an agent
can re-identify after drift; `PlaceholderShape` (`ph:S:KIND`) re-resolves by
semantic kind, the drift-proof form.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterator, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor, Paragraph, ParagraphCollection
from .constants import (
    MsoShapeType,
    MsoTextOrientation,
    MsoTriState,
    autoshape_type_for,
    chart_type_for,
    is_true,
    placeholder_kind_name,
    placeholder_types_for,
    shape_image_filter_for,
    shape_type_name,
)
from .exceptions import AnchorNotFoundError, NoTextFrameError

if TYPE_CHECKING:
    from ._charts import Chart, SeriesInput
    from ._slides import Slide
    from ._tables import Table

# Default geometry (points) for added shapes when the caller names none. A wide,
# short box for text; a square for autoshapes; a picture keeps its native size.
_DEFAULT_LEFT = 72.0
_DEFAULT_TOP = 72.0
_DEFAULT_TEXTBOX_WIDTH = 288.0  # 4 in
_DEFAULT_TEXTBOX_HEIGHT = 72.0  # 1 in
_DEFAULT_SHAPE_WIDTH = 144.0  # 2 in
_DEFAULT_SHAPE_HEIGHT = 144.0  # 2 in
_DEFAULT_TABLE_WIDTH = 480.0  # ~6.7 in
_DEFAULT_ROW_HEIGHT = 30.0  # advisory; PowerPoint auto-fits rows to content
_DEFAULT_CHART_WIDTH = 480.0  # ~6.7 in
_DEFAULT_CHART_HEIGHT = 300.0  # ~4.2 in
_CHART_DEFAULT_STYLE = -1  # AddChart2 Style: -1 = the type's default style


# ---------------------------------------------------------------------------
# COM-level helpers (operate on a raw Shape dispatch object)
# ---------------------------------------------------------------------------


def has_text_frame(com_shape: Any) -> bool:
    """True iff the shape can hold text (`Shape.HasTextFrame == msoTrue`)."""
    try:
        return is_true(com_shape.HasTextFrame)
    except Exception:
        return False


def is_placeholder(com_shape: Any) -> bool:
    """True iff the shape is a layout placeholder (`Shape.Type == msoPlaceholder`)."""
    try:
        return int(com_shape.Type) == int(MsoShapeType.PLACEHOLDER)
    except Exception:
        return False


def has_table(com_shape: Any) -> bool:
    """True iff the shape holds a table (`Shape.HasTable == msoTrue`).

    The reliable gate — a table that fills a content placeholder reports
    `Shape.Type == placeholder` (14), not table (19), so `Type` can't be trusted
    (verified live in `scripts/table_spike.py`).
    """
    try:
        return is_true(com_shape.HasTable)
    except Exception:
        return False


def has_chart(com_shape: Any) -> bool:
    """True iff the shape holds a chart (`Shape.HasChart == msoTrue`).

    The reliable gate — like a table, a chart that fills a content placeholder can
    report a placeholder `Shape.Type`, so `Type` can't be trusted.
    """
    try:
        return is_true(com_shape.HasChart)
    except Exception:
        return False


def _alt_text_of(com_shape: Any) -> str:
    """The shape's `AlternativeText` (accessibility text), or "" if unreadable."""
    try:
        return str(com_shape.AlternativeText or "")
    except Exception:
        return ""


def _geometry_of(com_shape: Any) -> dict[str, float] | None:
    try:
        return {
            "left": float(com_shape.Left),
            "top": float(com_shape.Top),
            "width": float(com_shape.Width),
            "height": float(com_shape.Height),
            "rotation": float(com_shape.Rotation),
        }
    except Exception:
        return None


def shape_to_dict(com_shape: Any, slide_index: int, z_index: int) -> dict[str, Any]:
    """Structured snapshot of one shape for `slide.read()` / `shapes.list()`.

    Emits the canonical `anchor_id` (`shape:S:N`) plus the drift-proof `name`
    and `id`, the friendly shape `type`, `geometry` (points), placeholder kind
    (or None), and `text` when the shape has a text frame.
    """
    d: dict[str, Any] = {
        "anchor_id": f"shape:{slide_index}:{z_index}",
        "index": z_index,
        "name": str(com_shape.Name),
        "id": int(com_shape.Id),
        "type": shape_type_name(com_shape.Type),
        "alt_text": _alt_text_of(com_shape),
        "geometry": _geometry_of(com_shape),
    }
    if is_placeholder(com_shape):
        try:
            d["placeholder"] = placeholder_kind_name(com_shape.PlaceholderFormat.Type)
        except Exception:
            d["placeholder"] = None
    else:
        d["placeholder"] = None

    d["has_table"] = has_table(com_shape)
    d["has_chart"] = has_chart(com_shape)
    if has_text_frame(com_shape):
        d["has_text_frame"] = True
        try:
            d["text"] = str(com_shape.TextFrame.TextRange.Text or "")
        except Exception:
            d["text"] = None
    else:
        d["has_text_frame"] = False
        d["text"] = None
    return d


# ---------------------------------------------------------------------------
# Shape
# ---------------------------------------------------------------------------


class Shape(Anchor):
    """A shape on a slide, addressed by 1-based z-order index — `shape:S:N`.

    Resolves its COM object live on every access (z-order drifts). Inherits
    `text` / `set_text` from `Anchor` (raising `NoTextFrameError` if the shape
    has no text frame), and adds geometry verbs.
    """

    kind = "shape"

    def __init__(self, slide: Slide, index: int) -> None:
        self._slide = slide
        self._index = int(index)

    @property
    def slide(self) -> Slide:
        return self._slide

    @property
    def index(self) -> int:
        """1-based z-order index this shape was addressed by."""
        return self._index

    @property
    def anchor_id(self) -> str:
        return f"shape:{self._slide.index}:{self._index}"

    def _com_shape(self) -> Any:
        """Resolve the COM `Shape` live by z-order index. Never cached."""
        shapes = self._slide.com.Shapes
        count = int(shapes.Count)
        if self._index < 1 or self._index > count:
            raise AnchorNotFoundError("shape", self.anchor_id)
        return shapes(self._index)

    @property
    def com(self) -> Any:
        """Raw COM `Shape` (overrides `Anchor.com`, which would give a text range)."""
        with _com.translate_com_errors():
            return self._com_shape()

    @property
    def name(self) -> str:
        """The shape's `.Name` (e.g. "Title 1") — drift-proof, unique per slide."""
        try:
            with _com.translate_com_errors():
                return str(self._com_shape().Name)
        except Exception:
            return self.anchor_id

    @property
    def shape_id(self) -> int:
        """`Shape.Id` — stable across z-order reordering."""
        with _com.translate_com_errors():
            return int(self._com_shape().Id)

    @property
    def shape_type(self) -> str:
        """Friendly shape-type name (e.g. "placeholder", "textbox", "picture")."""
        with _com.translate_com_errors():
            return shape_type_name(self._com_shape().Type)

    @property
    def alt_text(self) -> str:
        """The shape's alternative (accessibility) text — `Shape.AlternativeText`.

        Doubles as a stable, **LLM-readable re-identification handle**: tag a
        picture/diagram with a descriptive alt text and find it again after
        z-order drift (it shows up in every shape listing) without relying on the
        volatile `shape:S:N` index. Empty string when unset. Set it with
        `set_alt_text`.
        """
        with _com.translate_com_errors():
            return _alt_text_of(self._com_shape())

    @property
    def has_text_frame(self) -> bool:
        with _com.translate_com_errors():
            return has_text_frame(self._com_shape())

    @property
    def has_table(self) -> bool:
        """Whether this shape holds a table (`Shape.HasTable`)."""
        with _com.translate_com_errors():
            return has_table(self._com_shape())

    @property
    def table(self) -> Table:
        """The shape's `Table` (cells are `cell:S:N:R:C` anchors).

        Raises `AnchorNotFoundError` (kind `"table"`, exit 2) if the shape holds
        no table.
        """
        from ._tables import Table

        with _com.translate_com_errors():
            if not has_table(self._com_shape()):
                raise AnchorNotFoundError("table", self.anchor_id)
        return Table(self)

    @property
    def has_chart(self) -> bool:
        """Whether this shape holds a chart (`Shape.HasChart`)."""
        with _com.translate_com_errors():
            return has_chart(self._com_shape())

    @property
    def chart(self) -> Chart:
        """The shape's `Chart` (data lives in an embedded Excel workbook).

        Raises `AnchorNotFoundError` (kind `"chart"`, exit 2) if the shape holds
        no chart.
        """
        from ._charts import Chart

        with _com.translate_com_errors():
            if not has_chart(self._com_shape()):
                raise AnchorNotFoundError("chart", self.anchor_id)
        return Chart(self)

    def _text_range(self) -> Any:
        sh = self._com_shape()
        if not has_text_frame(sh):
            raise NoTextFrameError(self.anchor_id)
        return sh.TextFrame.TextRange

    # -- paragraphs (v0.3) -------------------------------------------------

    @property
    def paragraphs(self) -> ParagraphCollection:
        """The shape's paragraphs (`para:S:N:P`); raises `NoTextFrameError` if none."""
        return ParagraphCollection(self)

    def paragraph(self, index: int) -> Paragraph:
        """The `index`-th paragraph (1-based) of this shape's text frame."""
        return ParagraphCollection(self)[index]

    # -- geometry (points throughout, never EMUs) --------------------------

    def geometry(self) -> dict[str, float]:
        """`{left, top, width, height, rotation}` in points."""
        with _com.translate_com_errors():
            geo = _geometry_of(self._com_shape())
        if geo is None:
            raise AnchorNotFoundError("shape", self.anchor_id)
        return geo

    def move(self, *, left: float | None = None, top: float | None = None) -> None:
        """Set the shape's absolute position in points. Pass `left` and/or `top`."""
        if left is None and top is None:
            raise ValueError("move() requires at least one of left= or top=")
        with _com.translate_com_errors():
            sh = self._com_shape()
            if left is not None:
                sh.Left = float(left)
            if top is not None:
                sh.Top = float(top)

    def resize(self, *, width: float | None = None, height: float | None = None) -> None:
        """Set the shape's size in points. Pass `width` and/or `height`."""
        if width is None and height is None:
            raise ValueError("resize() requires at least one of width= or height=")
        with _com.translate_com_errors():
            sh = self._com_shape()
            if width is not None:
                sh.Width = float(width)
            if height is not None:
                sh.Height = float(height)

    def set_alt_text(self, value: str) -> None:
        """Set the shape's alternative (accessibility) text (`Shape.AlternativeText`).

        The drift-proof re-identification handle (see `alt_text`). A mutation:
        wrap in `deck.edit(...)` for view preservation + a one-Ctrl-Z fence.
        """
        with _com.translate_com_errors():
            self._com_shape().AlternativeText = str(value)

    def delete(self) -> None:
        """Delete this shape from its slide (`Shape.Delete`).

        The wrapper is spent afterwards; later shapes' z-order indices shift down
        by one. Wrap in `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        with _com.translate_com_errors():
            self._com_shape().Delete()

    # -- render (v0.7; a read — no mutation, polite by nature) -----------------

    def export_image(
        self,
        path: str | os.PathLike[str] | None = None,
        *,
        fmt: str = "png",
    ) -> Path:
        """Render *just this shape* to an image file and return its absolute path.

        The per-shape complement to `Slide.export_image` (v0.4): lets a vision
        model see one picture / chart / diagram in isolation, cropped to the
        shape's (rendered) bounds. Wraps `Shape.Export(PathName, Filter)` —
        `Filter` is the `PpShapeFormat` int enum, **not** `Slide.Export`'s string
        FilterName.

        `fmt` is a friendly token (`png`/`jpg`/`gif`/`bmp`; see
        `constants.SHAPE_IMAGE_FORMAT_CHOICES` — narrower than the slide set, no
        TIFF). When `path` is None a temp file is created (export-then-`Read` in
        one step). A relative `path` is resolved to absolute first (PowerPoint
        otherwise writes to its own working directory, not the caller's). It's a
        read — no mutation, and polite (it doesn't move the viewed slide or the
        Selection).

        The image is the shape's **native pixel size** (the slide's 96-DPI scale,
        e.g. a 720 pt-wide shape on a 960 pt slide → 960 px). Unlike
        `Slide.export_image`, there is **no** output-size override: the 2026-05-28
        live spike found `Shape.Export`'s ScaleWidth/ScaleHeight do *not* map to
        output pixels the way `Slide.Export`'s do (requesting 400×300 gave
        399×241 — width roughly tracked, height didn't, aspect wasn't preserved),
        so pptlive only exposes the reliable native export. Reach for the `.com`
        escape hatch (`shape.com.Export(path, filter, w, h, mode)`) if you need to
        experiment with scaling.
        """
        filter_int, ext = shape_image_filter_for(fmt)  # ValueError before any COM
        if path is None:
            fd, tmp = tempfile.mkstemp(prefix="pptlive_shape_", suffix=f".{ext}")
            os.close(fd)
            os.remove(tmp)  # hand PowerPoint a clean path to write
            abs_path = tmp
        else:
            abs_path = os.path.abspath(os.fspath(path))
        with _com.translate_com_errors():
            self._com_shape().Export(abs_path, int(filter_int))
        return Path(abs_path)

    def to_dict(self) -> dict[str, Any]:
        with _com.translate_com_errors():
            return shape_to_dict(self._com_shape(), self._slide.index, self._index)


class PlaceholderShape(Shape):
    """A placeholder addressed by semantic kind — `ph:S:KIND`.

    The LLM-preferred, drift-proof form: "the title of slide 3" without caring
    about z-order. Re-resolves its COM shape by `PlaceholderFormat.Type` on every
    access (via `Slide._find_placeholder`), so it survives shape reordering. It
    *is* a `Shape`, so all geometry and text verbs work; only the resolution
    strategy and `anchor_id` differ.
    """

    kind = "placeholder"

    def __init__(self, slide: Slide, ph_kind: str) -> None:
        # Validate the kind eagerly so a typo fails before any COM work.
        placeholder_types_for(ph_kind)
        super().__init__(slide, index=0)
        self._ph_kind = ph_kind.lower()

    @property
    def placeholder_kind(self) -> str:
        return self._ph_kind

    @property
    def anchor_id(self) -> str:
        return f"ph:{self._slide.index}:{self._ph_kind}"

    @property
    def index(self) -> int:
        """Current 1-based z-order index of the resolved placeholder."""
        with _com.translate_com_errors():
            _shape, idx = self._slide._find_placeholder(self._ph_kind)
        return idx

    def _com_shape(self) -> Any:
        shape, _idx = self._slide._find_placeholder(self._ph_kind)
        return shape

    def to_dict(self) -> dict[str, Any]:
        with _com.translate_com_errors():
            shape, idx = self._slide._find_placeholder(self._ph_kind)
            return shape_to_dict(shape, self._slide.index, idx)


# ---------------------------------------------------------------------------
# ShapeCollection
# ---------------------------------------------------------------------------


class ShapeCollection:
    """Indexable, iterable view over a slide's shapes.

    Index by 1-based z-order (`slide.shapes[2]`) or by name
    (`slide.shapes["Title 1"]`). Iteration yields a `Shape` per shape in
    z-order. `list()` emits the structured dict used by `slide read`.
    """

    def __init__(self, slide: Slide) -> None:
        self._slide = slide

    @property
    def _com_collection(self) -> Any:
        return self._slide.com.Shapes

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def __getitem__(self, key: int | str) -> Shape:
        if isinstance(key, bool):
            raise TypeError(f"shape key must be int or str, got {type(key).__name__}")
        if isinstance(key, int):
            count = len(self)
            if key < 1 or key > count:
                raise AnchorNotFoundError("shape", f"shape:{self._slide.index}:{key}")
            return Shape(self._slide, key)
        if isinstance(key, str):
            with _com.translate_com_errors():
                for idx, sh in enumerate(self._com_collection, start=1):
                    if str(sh.Name) == key:
                        return Shape(self._slide, idx)
            raise AnchorNotFoundError("shape", key)
        raise TypeError(f"shape key must be int or str, got {type(key).__name__}")

    def __contains__(self, key: object) -> bool:
        if isinstance(key, bool) or not isinstance(key, (int, str)):
            return False
        try:
            self[key]
            return True
        except AnchorNotFoundError:
            return False

    def __iter__(self) -> Iterator[Shape]:
        count = len(self)
        for idx in range(1, count + 1):
            yield Shape(self._slide, idx)

    # -- creators (v0.2; wrap in deck.edit(...) for view + one-Ctrl-Z) ---------
    #
    # PowerPoint adds a new shape at the top of the z-order, i.e. the last slot
    # of the Shapes collection, so the new shape's 1-based z-order index is the
    # post-add Count. We return a live `Shape` addressed by that index (verified
    # against real PowerPoint in scripts/shape_spike.py).

    def _added(self) -> Shape:
        return Shape(self._slide, int(self._com_collection.Count))

    def add_textbox(
        self,
        text: str = "",
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> Shape:
        """Add a horizontal text box and return it (`Shapes.AddTextbox`).

        Geometry is in points; omitted values default to a 4×1 in box near the
        top-left. `text`, if given, is written into the new frame.
        """
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = _DEFAULT_TEXTBOX_WIDTH if width is None else float(width)
        height = _DEFAULT_TEXTBOX_HEIGHT if height is None else float(height)
        with _com.translate_com_errors():
            com_shape = self._com_collection.AddTextbox(
                int(MsoTextOrientation.HORIZONTAL), left, top, width, height
            )
            if text:
                com_shape.TextFrame.TextRange.Text = text
            return self._added()

    def add_shape(
        self,
        shape_type: str | int,
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> Shape:
        """Add an autoshape and return it (`Shapes.AddShape`).

        `shape_type` is a friendly name (`"rectangle"`, `"oval"`, `"arrow"`, …;
        see `constants.AUTOSHAPE_CHOICES`) or a raw `MsoAutoShapeType` int.
        Geometry is in points; omitted values default to a 2×2 in box near the
        top-left. Raises `ValueError` for an unknown shape name (before any COM).
        """
        type_int = autoshape_type_for(shape_type)  # ValueError before COM
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = _DEFAULT_SHAPE_WIDTH if width is None else float(width)
        height = _DEFAULT_SHAPE_HEIGHT if height is None else float(height)
        with _com.translate_com_errors():
            self._com_collection.AddShape(type_int, left, top, width, height)
            return self._added()

    def add_picture(
        self,
        path: str | os.PathLike[str],
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        alt_text: str | None = None,
    ) -> Shape:
        """Embed a picture from a local file and return it (`Shapes.AddPicture`).

        The image is **embedded**, never linked (so the deck stays portable).
        `left`/`top` default to the top-left; omitted `width`/`height` keep the
        image's native size. `alt_text`, if given, sets the picture's
        alternative text — a drift-proof, LLM-readable re-identification handle
        (see `Shape.alt_text`). Raises `FileNotFoundError` if `path` doesn't exist.
        """
        fs_path = os.fspath(path)
        if not os.path.isfile(fs_path):
            raise FileNotFoundError(f"picture not found: {fs_path}")
        abs_path = os.path.abspath(fs_path)
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        with _com.translate_com_errors():
            com_shape = self._com_collection.AddPicture(
                abs_path,
                int(MsoTriState.FALSE),  # LinkToFile: no
                int(MsoTriState.TRUE),  # SaveWithDocument: yes (embed)
                left,
                top,
                -1.0 if width is None else float(width),  # -1 = native size
                -1.0 if height is None else float(height),
            )
            if alt_text is not None:
                com_shape.AlternativeText = str(alt_text)
            return self._added()

    def add_table(
        self,
        rows: int,
        columns: int,
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> Shape:
        """Add a `rows`×`columns` table and return its `Shape` (`Shapes.AddTable`).

        Geometry is in points; omitted values default to a wide grid near the
        top-left (height is advisory — PowerPoint auto-fits rows to content).
        Address cells through the returned shape's `.table` or the `cell:S:N:R:C`
        anchor; the shape's `.has_table` is True. Raises `ValueError` for
        non-positive `rows`/`columns` (before any COM).
        """
        if int(rows) < 1 or int(columns) < 1:
            raise ValueError(f"table needs >=1 row and >=1 column, got {rows}x{columns}")
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = _DEFAULT_TABLE_WIDTH if width is None else float(width)
        height = _DEFAULT_ROW_HEIGHT * int(rows) if height is None else float(height)
        with _com.translate_com_errors():
            self._com_collection.AddTable(int(rows), int(columns), left, top, width, height)
            return self._added()

    def add_chart(
        self,
        chart_type: str | int = "column",
        categories: Sequence[str] | None = None,
        series: SeriesInput | None = None,
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> Shape:
        """Add a chart and return its `Shape` (`Shapes.AddChart2`).

        `chart_type` is a friendly name (`"column"`, `"bar"`, `"line"`, `"pie"`,
        …; see `constants.CHART_TYPE_CHOICES`) or a raw `XlChartType` int.
        Geometry is in points; omitted values default to a ~6.7×4.2 in chart near
        the top-left. Address the chart's data through the returned shape's
        `.chart` (or `anchor_id`); the shape's `.has_chart` is True.

        If both `categories` and `series` are given, the chart's embedded-Excel
        data is replaced with them (via `Chart.set_data`); otherwise the chart
        keeps PowerPoint's default placeholder data. Pass one without the other
        and it's a `ValueError`. Raises `ValueError` for an unknown `chart_type`
        (before any COM). Wrap in `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        type_int = chart_type_for(chart_type)  # ValueError before any COM
        if (categories is None) != (series is None):
            raise ValueError("pass both categories and series, or neither")
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = _DEFAULT_CHART_WIDTH if width is None else float(width)
        height = _DEFAULT_CHART_HEIGHT if height is None else float(height)
        with _com.translate_com_errors():
            self._com_collection.AddChart2(_CHART_DEFAULT_STYLE, type_int, left, top, width, height)
            shape = self._added()
        if categories is not None and series is not None:
            shape.chart.set_data(categories, series)
        return shape

    def list(self) -> list[dict[str, Any]]:
        """Every shape as a structured dict, in z-order."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, sh in enumerate(self._com_collection, start=1):
                out.append(shape_to_dict(sh, self._slide.index, idx))
        return out

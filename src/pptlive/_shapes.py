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
from dataclasses import dataclass
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor, Paragraph, ParagraphCollection, links_in_range, slide_jump_subaddress
from .constants import (
    MsoShadowStyle,
    MsoShapeType,
    MsoTextOrientation,
    MsoTriState,
    PpActionType,
    PpMouseActivation,
    align_cmd_for,
    anim_effect_for,
    anim_effect_name,
    anim_trigger_for,
    anim_trigger_name,
    arrowhead_size_for,
    arrowhead_style_for,
    arrowhead_style_name,
    autoshape_type_for,
    autosize_name,
    chart_type_for,
    color_hex_or_none,
    connector_type_for,
    connector_type_name,
    dash_style_for,
    dash_style_name,
    distribute_cmd_for,
    fill_type_name,
    gradient_style_for,
    gradient_style_name,
    is_true,
    media_type_name,
    parse_color,
    pattern_for,
    pattern_name,
    placeholder_kind_name,
    placeholder_types_for,
    preset_gradient_for,
    relative_to_for,
    shape_image_filter_for,
    shape_type_name,
    smartart_layout_for,
    zorder_cmd_for,
)
from .exceptions import AmbiguousMatchError, AnchorNotFoundError, NoTextFrameError

if TYPE_CHECKING:
    from ._charts import Chart, SeriesInput
    from ._slides import Slide
    from ._smartart import NodeInput, SmartArt
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
_DEFAULT_SMARTART_WIDTH = 480.0  # ~6.7 in
_DEFAULT_SMARTART_HEIGHT = 300.0  # ~4.2 in
_DEFAULT_AUDIO_WIDTH = 60.0  # a small speaker-icon box
_DEFAULT_AUDIO_HEIGHT = 60.0
_DEFAULT_VIDEO_WIDTH = 480.0  # 16:9 video frame
_DEFAULT_VIDEO_HEIGHT = 270.0

_safe = _com.safe_read  # defensive COM-property read (returns a default on failure)


# ---------------------------------------------------------------------------
# COM-level helpers (operate on a raw Shape dispatch object)
# ---------------------------------------------------------------------------


def find_shape_by_id(slide_com: Any, shape_id: int) -> tuple[int, Any] | None:
    """Find a shape on a slide by its stable `Shape.Id`.

    Returns `(1-based z-order index, COM Shape)`, or `None` when no shape on the
    slide carries that id. The single home for the "scan `Shapes` for a matching
    `.Id`" loop shared by `ShapeById._com_shape` (resolving the `shapeid:S:ID`
    anchor) and `_selection._zorder_index` (mapping a live Selection back to its
    volatile `shape:S:N` z-order index).
    """
    shapes = slide_com.Shapes
    for idx in range(1, int(shapes.Count) + 1):
        sh = shapes(idx)
        if int(sh.Id) == int(shape_id):
            return idx, sh
    return None


def has_text_frame(com_shape: Any) -> bool:
    """True iff the shape can hold text (`Shape.HasTextFrame == msoTrue`)."""
    try:
        return is_true(com_shape.HasTextFrame)
    except Exception:
        return False


def _layout_default_size(layout_placeholder: Any) -> float | None:
    """A layout placeholder's default font size in points, or `None` if unreadable."""
    try:
        size = float(layout_placeholder.TextFrame.TextRange.Font.Size)
    except Exception:
        return None
    return size if size > 0 else None


@dataclass(frozen=True)
class TextFrameStatus:
    """A shape's text-frame container state — the autofit/overflow diagnostics read.

    Surfaces the state that makes a "formatting spiral" visible *before* it bites
    (the gpt-5.4 review's ask). `autosize` is the friendly `TextFrame2.AutoSize`
    mode; `word_wrap` is on/off; `margins` is the four inner margins in points;
    `overflow_risk` is a **coarse, mode-derived** flag (PowerPoint exposes no
    shrink-% on this build, so it reads the autofit *mode*, not a measured extent):
    `"possible"` (autosize off — text can clip), `"low"` (an autofit mode is
    active), or `"unknown"`.
    """

    autosize: str
    word_wrap: bool
    margins: dict[str, float]
    overflow_risk: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "autosize": self.autosize,
            "word_wrap": self.word_wrap,
            "margins": self.margins,
            "overflow_risk": self.overflow_risk,
        }


def _overflow_risk(autosize: str) -> str:
    """Coarse overflow risk from the autofit mode (no measured extent available)."""
    if autosize == "none":
        return "possible"  # nothing shrinks the text or grows the shape — it can clip
    if autosize in ("text_to_fit_shape", "shape_to_fit_text"):
        return "low"  # an autofit mode keeps text and frame reconciled
    return "unknown"


def _autosize_of(com_shape: Any) -> str:
    """The shape's autofit mode name, preferring the clean `TextFrame2` reading."""
    try:
        return autosize_name(com_shape.TextFrame2.AutoSize)
    except Exception:
        return _safe_autosize_classic(com_shape)


def _safe_autosize_classic(com_shape: Any) -> str:
    try:
        return autosize_name(com_shape.TextFrame.AutoSize)
    except Exception:
        return "unknown"


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


def has_smartart(com_shape: Any) -> bool:
    """True iff the shape holds a SmartArt diagram (`Shape.HasSmartArt == msoTrue`).

    The reliable gate — like a table/chart, `Shape.Type` reports msoSmartArt here
    but can't be trusted in general, so gate on `Has*` (verified live in
    `scripts/smartart_spike.py`).
    """
    try:
        return is_true(com_shape.HasSmartArt)
    except Exception:
        return False


def is_picture(com_shape: Any) -> bool:
    """True iff the shape is an embedded or linked picture (`msoPicture`/`msoLinkedPicture`).

    The gate for `Shape.set_picture` — `Shape.Type`, unlike the table/chart case,
    *is* reliable for a picture (a picture is never a placeholder masquerade).
    """
    try:
        return int(com_shape.Type) in (
            int(MsoShapeType.PICTURE),
            int(MsoShapeType.LINKED_PICTURE),
        )
    except Exception:
        return False


def is_media(com_shape: Any) -> bool:
    """True iff the shape is an embedded/linked media clip (`msoMedia`).

    The gate for the media reads — `Shape.Type == msoMedia (16)` is reliable for
    audio/video (a media clip is never a placeholder masquerade), like `is_picture`.
    """
    try:
        return int(com_shape.Type) == int(MsoShapeType.MEDIA)
    except Exception:
        return False


def _media_to_dict(com_shape: Any) -> dict[str, Any]:
    """The media sub-dict for a media shape: kind + duration + playback state.

    `length_s` is the clip duration in **seconds** (`MediaFormat.Length` is ms);
    `autoplay` reflects `AnimationSettings.PlaySettings.PlayOnEntry`. Every field is
    read defensively (`_safe`) — some props raise on certain clips — so a media
    shape never fails the whole structured read.
    """
    length_ms = _safe(lambda: float(com_shape.MediaFormat.Length), None)
    return {
        "type": media_type_name(_safe(lambda: int(com_shape.MediaType), None)),
        "length_s": round(length_ms / 1000.0, 3) if length_ms else None,
        "muted": _safe(lambda: is_true(com_shape.MediaFormat.Muted), None),
        "volume": _safe(lambda: round(float(com_shape.MediaFormat.Volume), 3), None),
        "autoplay": _safe(
            lambda: is_true(com_shape.AnimationSettings.PlaySettings.PlayOnEntry), None
        ),
    }


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


def _is_none_token(value: Any) -> bool:
    """True iff `value` is the string `"none"` (the transparent-fill / no-line token)."""
    return isinstance(value, str) and value.strip().lower() == "none"


def _precheck_fill(fill: Any, line: Any) -> None:
    """Validate fill/line color args **before any COM** (the wordlive rule).

    A `None` (unchanged) or the `"none"` token (transparent / no border) needs no
    color; anything else must `parse_color` cleanly, so a bad hex raises a clean
    `ValueError` before a shape is created or mutated.
    """
    if fill is not None and not _is_none_token(fill):
        parse_color(fill)
    if line is not None and not _is_none_token(line):
        parse_color(line)


def _check_transparency(value: float | None, label: str) -> None:
    """A transparency must be a `0.0..1.0` fraction (ValueError before any COM)."""
    if value is None:
        return
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{label} must be a number in [0, 1], got {value!r}")
    if not 0.0 <= float(value) <= 1.0:
        raise ValueError(
            f"{label} must be in [0, 1] (0 opaque, 1 fully transparent), got {value!r}"
        )


def apply_shape_fill(
    com_shape: Any,
    *,
    fill: str | int | tuple[int, int, int] | None = None,
    line: str | int | tuple[int, int, int] | None = None,
    line_width: float | None = None,
    fill_transparency: float | None = None,
    line_transparency: float | None = None,
) -> None:
    """Write fill / line (border) formatting onto a COM `Shape` — only the kwargs passed.

    `fill`/`line` take a color (`"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB
    int) for a solid fill / line of that color, or the string `"none"` to make
    the fill transparent / remove the border. `line_width` is the border weight
    in points. `fill_transparency`/`line_transparency` are `0.0..1.0` alpha
    fractions (0 opaque, 1 fully transparent). Colors + transparencies are
    validated up front (so a bad value raises before any COM mutation). Caller
    wraps this in `translate_com_errors()`.
    """
    _precheck_fill(fill, line)  # ValueError before any COM mutation
    _check_transparency(fill_transparency, "fill_transparency")
    _check_transparency(line_transparency, "line_transparency")
    if fill is not None:
        if _is_none_token(fill):
            com_shape.Fill.Visible = int(MsoTriState.FALSE)
        else:
            com_shape.Fill.Visible = int(MsoTriState.TRUE)
            com_shape.Fill.Solid()
            com_shape.Fill.ForeColor.RGB = parse_color(fill)
    if line is not None:
        if _is_none_token(line):
            com_shape.Line.Visible = int(MsoTriState.FALSE)
        else:
            com_shape.Line.Visible = int(MsoTriState.TRUE)
            com_shape.Line.ForeColor.RGB = parse_color(line)
    if line_width is not None:
        com_shape.Line.Weight = float(line_width)
    if fill_transparency is not None:
        com_shape.Fill.Transparency = float(fill_transparency)
    if line_transparency is not None:
        com_shape.Line.Transparency = float(line_transparency)


def apply_line_style(
    com_shape: Any,
    *,
    dash: str | int | None = None,
    begin_arrow: str | int | None = None,
    end_arrow: str | int | None = None,
    begin_arrow_size: str | int | None = None,
    end_arrow_size: str | int | None = None,
) -> None:
    """Write line dash + arrowhead styling onto a COM `Shape` — only the kwargs passed.

    `dash` is a friendly `MsoLineDashStyle` name (`"solid"`/`"dash"`/`"round_dot"`/…)
    or raw int. `begin_arrow`/`end_arrow` are `MsoArrowheadStyle` names
    (`"none"`/`"triangle"`/`"open"`/`"stealth"`/`"diamond"`/`"oval"`) or raw ints;
    `begin_arrow_size`/`end_arrow_size` are `"small"`/`"medium"`/`"large"` (set both
    arrowhead length + width). All names are resolved up front (ValueError before any
    COM). **Arrowheads only apply to lines/connectors** — PowerPoint raises on a
    closed shape (rectangle/oval/…). Caller wraps this in `translate_com_errors()`.
    """
    # Resolve every name first so a typo raises before any COM mutation.
    dash_int = dash_style_for(dash) if dash is not None else None
    begin_int = arrowhead_style_for(begin_arrow) if begin_arrow is not None else None
    end_int = arrowhead_style_for(end_arrow) if end_arrow is not None else None
    begin_size_int = arrowhead_size_for(begin_arrow_size) if begin_arrow_size is not None else None
    end_size_int = arrowhead_size_for(end_arrow_size) if end_arrow_size is not None else None
    if (
        dash_int is None
        and begin_int is None
        and end_int is None
        and begin_size_int is None
        and end_size_int is None
    ):
        raise ValueError(
            "apply_line_style() needs at least one of dash=, begin_arrow=, end_arrow=, "
            "begin_arrow_size=, end_arrow_size="
        )
    line = com_shape.Line
    if dash_int is not None:
        line.DashStyle = dash_int
    if begin_int is not None:
        line.BeginArrowheadStyle = begin_int
    if end_int is not None:
        line.EndArrowheadStyle = end_int
    if begin_size_int is not None:
        line.BeginArrowheadLength = begin_size_int
        line.BeginArrowheadWidth = begin_size_int
    if end_size_int is not None:
        line.EndArrowheadLength = end_size_int
        line.EndArrowheadWidth = end_size_int


# Advanced fills (v1.2): gradient / picture / pattern. The spike
# (scripts/fill_advanced_spike.py) pinned each COM recipe; the wrappers below
# encode those findings (legacy `Insert` for stops, abs-path picture, …).

_GRADIENT_DEFAULT_DEGREE = 0.5  # OneColorGradient brightness when none is given


def _gradient_stop_args(
    colors: Sequence[Any], positions: Sequence[float] | None
) -> tuple[list[int], list[float | None]]:
    """Validate + normalize gradient colors/positions to RGB ints + positions.

    All colors `parse_color` cleanly (ValueError before any COM). `positions`, when
    given, must match `colors` in length and each lie in `[0, 1]`. Returns the RGB
    ints and a parallel positions list (the raw floats or all-`None`).
    """
    rgbs = [parse_color(c) for c in colors]  # ValueError before any COM
    if not rgbs:
        raise ValueError("set_gradient_fill() needs at least one color")
    if positions is None:
        return rgbs, [None] * len(rgbs)
    if len(positions) != len(rgbs):
        raise ValueError("positions= must have the same length as colors=")
    norm: list[float | None] = []
    for p in positions:
        fp = float(p)
        if not (0.0 <= fp <= 1.0):
            raise ValueError(f"gradient position {p!r} must be between 0.0 and 1.0")
        norm.append(fp)
    return rgbs, norm


def apply_gradient_fill(
    com_shape: Any,
    *,
    colors: Sequence[Any] | None = None,
    positions: Sequence[float] | None = None,
    style: str | int = "horizontal",
    variant: int = 1,
    degree: float | None = None,
    preset: str | int | None = None,
) -> None:
    """Write a gradient fill onto a COM `Shape`. Caller wraps in `translate_com_errors()`.

    Three shapes, matching the spike's verified COM recipes:
    - `preset=` (e.g. `"ocean"`) → `Fill.PresetGradient` (a named multi-stop ramp).
    - one color → `Fill.OneColorGradient` (`degree` is the 0..1 brightness, default 0.5).
    - two+ colors → `Fill.TwoColorGradient` with the first/last colors at stops 0.0/1.0;
      interior colors become extra stops via the legacy `GradientStops.Insert` (since
      `Insert2` won't marshal). `positions` places the *interior* stops (endpoints stay
      at 0.0/1.0); omitted, interior stops space evenly.

    All colors are validated up front (ValueError before any COM mutation).
    """
    style_int = gradient_style_for(style)  # ValueError before any COM
    if preset is not None:
        preset_int = preset_gradient_for(preset)
        com_shape.Fill.PresetGradient(style_int, int(variant), preset_int)
        return
    if colors is None:
        raise ValueError("set_gradient_fill() requires colors= or preset=")
    rgbs, norm = _gradient_stop_args(colors, positions)
    fill = com_shape.Fill
    if len(rgbs) == 1:
        deg = _GRADIENT_DEFAULT_DEGREE if degree is None else float(degree)
        fill.OneColorGradient(style_int, int(variant), deg)
        fill.ForeColor.RGB = rgbs[0]
        return
    fill.TwoColorGradient(style_int, int(variant))
    fill.ForeColor.RGB = rgbs[0]
    fill.BackColor.RGB = rgbs[-1]
    interior = rgbs[1:-1]
    if not interior:
        return
    n = len(rgbs)
    for i, rgb in enumerate(interior, start=1):
        pos = norm[i]
        if pos is None:
            pos = i / (n - 1)  # evenly spaced between the 0.0/1.0 endpoints
        fill.GradientStops.Insert(rgb, float(pos))  # legacy Insert (Insert2 won't marshal)


def apply_picture_fill(com_shape: Any, path: str | Path) -> None:
    """Fill a COM `Shape` with an image (`Fill.UserPicture`). Caller wraps in errors.

    The path is resolved to an **absolute** path first — `UserPicture` raises
    `ERROR_FILE_NOT_FOUND` on a relative path (the `Slide.Export` footgun, confirmed
    in the spike). Raises `FileNotFoundError` (before any COM) if the file is missing.
    """
    abspath = os.path.abspath(os.fspath(path))
    if not os.path.isfile(abspath):
        raise FileNotFoundError(f"picture fill image not found: {abspath}")
    com_shape.Fill.UserPicture(abspath)


def replace_picture(
    slide_com: Any,
    com_old: Any,
    abs_path: str,
    slide_index: int,
    *,
    alt_text: str | None = None,
) -> int:
    """Re-source a picture shape in place: add the new image at the old picture's
    box / name / alt text / z-order, then delete the old. Returns the **new**
    `Shape.Id`. Caller wraps this in `translate_com_errors()`.

    PowerPoint's COM exposes no in-place image swap for a picture shape
    (`Fill.UserPicture` only sets a fill *behind* the unchanged raster — confirmed
    in `scripts/set_picture_spike.py`), so this is a delete + re-insert that
    preserves everything addressable. The spike pinned the three pitfalls handled
    below: pictures default to **locked aspect** (so the copied box must be applied
    with `LockAspectRatio` off, else width/height snap to the new image's ratio);
    the old delete **drifts z-order indices** (so the new shape is re-resolved by
    its stable `Shape.Id`, never an index); and the old **z-order slot** is
    restored by send-to-back-then-step-forward.
    """
    # Snapshot the old picture's identity + geometry before touching anything.
    left = float(com_old.Left)
    top = float(com_old.Top)
    width = float(com_old.Width)
    height = float(com_old.Height)
    rotation = float(com_old.Rotation)
    name = str(com_old.Name)
    carried_alt = str(com_old.AlternativeText or "")
    z = int(com_old.ZOrderPosition)

    new_com = slide_com.Shapes.AddPicture(
        abs_path,
        int(MsoTriState.FALSE),  # LinkToFile: no
        int(MsoTriState.TRUE),  # SaveWithDocument: yes (embed)
        left,
        top,
        -1.0,  # native size — the old box is forced on below
        -1.0,
    )
    new_id = int(new_com.Id)

    def _force_box() -> None:
        # Clearing LockAspectRatio lets the copied box win; if the assignment is
        # rejected (some builds/shape states), setting Width snaps Height to the
        # new image's ratio (and vice versa) — the exact drift this re-source
        # guards against. Best-effort clear, then set the box.
        try:
            new_com.LockAspectRatio = int(MsoTriState.FALSE)
        except Exception:  # noqa: BLE001 — best-effort; verified/retried below
            pass
        new_com.Left = left
        new_com.Top = top
        new_com.Width = width
        new_com.Height = height

    _force_box()
    # Verify the box stuck; if a wedged aspect-lock snapped it, retry once. A
    # genuinely stuck lock can't be cleared in software, but the retry recovers
    # the common transient case instead of silently shipping a wrong-sized picture.
    if abs(float(new_com.Width) - width) > 0.5 or abs(float(new_com.Height) - height) > 0.5:
        _force_box()
    new_com.Rotation = rotation
    new_com.AlternativeText = carried_alt if alt_text is None else str(alt_text)

    com_old.Delete()

    # The delete shifted z-order indices — re-resolve the new shape by its Id.
    found = find_shape_by_id(slide_com, new_id)
    if found is None:  # defensive — we just added it
        raise AnchorNotFoundError("shape", f"shapeid:{slide_index}:{new_id}")
    new_com = found[1]
    new_com.Name = name  # safe now the old one is gone (no name clash)

    # Restore the old z-order slot: to the back, then forward to position `z`.
    new_com.ZOrder(zorder_cmd_for("back"))
    for _ in range(max(0, z - 1)):
        new_com.ZOrder(zorder_cmd_for("forward"))
    return new_id


def apply_pattern_fill(
    com_shape: Any,
    *,
    pattern: str | int,
    fore: Any,
    back: Any | None = None,
) -> None:
    """Write a two-color pattern fill (`Fill.Patterned`). Caller wraps in errors.

    `pattern` is a friendly `MsoPatternType` name (`"percent_50"`, `"trellis"`, …)
    or a raw int; `fore`/`back` are the pattern's foreground/background colors.
    Colors validated up front (ValueError before any COM).
    """
    pattern_int = pattern_for(pattern)  # ValueError before any COM
    fore_rgb = parse_color(fore)
    back_rgb = parse_color(back) if back is not None else None
    fill = com_shape.Fill
    fill.Patterned(pattern_int)
    fill.ForeColor.RGB = fore_rgb
    if back_rgb is not None:
        fill.BackColor.RGB = back_rgb


# Shape effects (v1.2): shadow / glow / soft-edge / reflection. The spike
# (scripts/effects_spike.py) confirmed all four round-trip; the props below are
# the verified-clean ones.

_TRANSPARENCY_SENTINEL = -2147483648  # default/unset Transparency reads this


def _is_none_effect(value: Any) -> bool:
    return isinstance(value, str) and value.strip().lower() == "none"


def _apply_shadow(com_shape: Any, spec: Any) -> None:
    shadow = com_shape.Shadow
    if _is_none_effect(spec):
        shadow.Visible = int(MsoTriState.FALSE)
        return
    if not isinstance(spec, dict):
        raise ValueError('shadow= must be a dict of properties or "none"')
    color = spec.get("color")
    rgb = parse_color(color) if color is not None else None  # ValueError before COM
    shadow.Visible = int(MsoTriState.TRUE)
    shadow.Style = int(MsoShadowStyle.OUTER)
    if rgb is not None:
        shadow.ForeColor.RGB = rgb
    for key, attr in (
        ("transparency", "Transparency"),
        ("blur", "Blur"),
        ("size", "Size"),
        ("offset_x", "OffsetX"),
        ("offset_y", "OffsetY"),
    ):
        if spec.get(key) is not None:
            setattr(shadow, attr, float(spec[key]))


def _apply_glow(com_shape: Any, spec: Any) -> None:
    glow = com_shape.Glow
    if _is_none_effect(spec):
        glow.Radius = 0.0
        return
    if not isinstance(spec, dict):
        raise ValueError('glow= must be a dict of properties or "none"')
    color = spec.get("color")
    rgb = parse_color(color) if color is not None else None  # ValueError before COM
    if rgb is not None:
        glow.Color.RGB = rgb
    glow.Radius = float(spec.get("radius", 8.0))
    if spec.get("transparency") is not None:
        glow.Transparency = float(spec["transparency"])


def _apply_soft_edge(com_shape: Any, spec: Any) -> None:
    preset = 0 if _is_none_effect(spec) else int(spec)
    com_shape.SoftEdge.Type = preset


def _apply_reflection(com_shape: Any, spec: Any) -> None:
    preset = 0 if _is_none_effect(spec) else int(spec)
    com_shape.Reflection.Type = preset


def apply_effect(
    com_shape: Any,
    *,
    shadow: Any | None = None,
    glow: Any | None = None,
    soft_edge: Any | None = None,
    reflection: Any | None = None,
) -> None:
    """Write shape effects (shadow / glow / soft-edge / reflection). Only kwargs passed.

    Each kwarg is the friendly form the wrapper documents (a dict for shadow/glow, an
    int preset for soft_edge/reflection) or the string `"none"` to turn it off. Colors
    validated up front. Caller wraps in `translate_com_errors()`.
    """
    if shadow is not None:
        _apply_shadow(com_shape, shadow)
    if glow is not None:
        _apply_glow(com_shape, glow)
    if soft_edge is not None:
        _apply_soft_edge(com_shape, soft_edge)
    if reflection is not None:
        _apply_reflection(com_shape, reflection)


def _gradient_stops(fill: Any) -> list[dict[str, Any]] | None:
    """`[{color, position}]` for a gradient fill's stops, sorted by position; None if unreadable."""
    try:
        stops = fill.GradientStops
        count = int(stops.Count)
    except Exception:
        return None
    out: list[dict[str, Any]] = []
    for i in range(1, count + 1):
        try:
            s = stops(i)
            out.append(
                {
                    "color": color_hex_or_none(s.Color.RGB),
                    "position": round(float(s.Position), 4),
                }
            )
        except Exception:
            continue
    out.sort(key=lambda d: d["position"] if d["position"] is not None else 0.0)
    return out


def _fill_to_dict(com_shape: Any) -> dict[str, Any] | None:
    """`{type, color, visible}` for the shape's fill (+ detail), or None when absent.

    `type` is the friendly `Fill.Type` (`solid`/`gradient`/`pattern`/`picture`/…),
    `color` the literal `"#RRGGBB"` foreground (or `None` for a theme/automatic fill —
    the `0x80000000` sentinel guard, `color_hex_or_none`). A gradient adds `stops`
    (`[{color, position}]`) and `gradient_style`; a pattern adds `pattern` (the friendly
    name) and `back_color`.
    """
    try:
        fill = com_shape.Fill
    except Exception:
        return None
    out: dict[str, Any] = {
        "type": fill_type_name(_safe(lambda: int(fill.Type), None)),
        "color": _safe(lambda: color_hex_or_none(fill.ForeColor.RGB), None),
        "visible": _safe(lambda: is_true(fill.Visible), None),
        "transparency": _safe(lambda: round(float(fill.Transparency), 3), None),
    }
    if out["type"] == "gradient":
        out["gradient_style"] = gradient_style_name(_safe(lambda: int(fill.GradientStyle), None))
        out["stops"] = _gradient_stops(fill)
    elif out["type"] == "patterned":
        out["pattern"] = pattern_name(_safe(lambda: int(fill.Pattern), None))
        out["back_color"] = _safe(lambda: color_hex_or_none(fill.BackColor.RGB), None)
    return out


def _line_to_dict(com_shape: Any) -> dict[str, Any] | None:
    """`{color, weight, visible, transparency, dash}` for the shape's line (border).

    Adds `begin_arrow`/`end_arrow` (friendly `MsoArrowheadStyle` names) only when an
    arrowhead is actually set (not `none`/unset) — keeps closed shapes' line dicts
    lean. Returns None when the shape exposes no `.Line`.
    """
    try:
        line = com_shape.Line
    except Exception:
        return None
    try:
        weight: float | None = float(line.Weight)
    except Exception:
        weight = None
    out: dict[str, Any] = {
        "color": color_hex_or_none(line.ForeColor.RGB),
        "weight": weight,
        "visible": is_true(line.Visible),
        "transparency": _safe(lambda: round(float(line.Transparency), 3), None),
        "dash": dash_style_name(_safe(lambda: int(line.DashStyle), None)),
    }
    begin = arrowhead_style_name(_safe(lambda: int(line.BeginArrowheadStyle), None))
    if begin is not None and begin != "none":
        out["begin_arrow"] = begin
    end = arrowhead_style_name(_safe(lambda: int(line.EndArrowheadStyle), None))
    if end is not None and end != "none":
        out["end_arrow"] = end
    return out


def background_to_dict(com_with_background: Any) -> dict[str, Any]:
    """`{type, color}` for any object that exposes a `.Background.Fill` (slide / master).

    `type` is the friendly `msoFillType` name (solid/gradient/picture/…) and `color`
    is the literal `"#RRGGBB"`, or `None` for a theme/automatic fill (the same honest
    `color_hex_or_none` guard as font / shape fill). Best-effort: a property that
    won't read returns `None` rather than raising, so a read never blows up.
    """
    fill = _safe(lambda: com_with_background.Background.Fill, None)
    if fill is None:
        return {"type": None, "color": None}
    return {
        "type": fill_type_name(_safe(lambda: int(fill.Type), None)),
        "color": _safe(lambda: color_hex_or_none(fill.ForeColor.RGB), None),
    }


def _effects_to_dict(com_shape: Any) -> dict[str, Any] | None:
    """Active shape effects (`{shadow?, glow?, soft_edge?, reflection?}`) or None.

    Best-effort: reads the cheap "is it on" gate for each effect and only expands the
    active ones (so a plain shape returns `None`, not four empty sub-dicts). The
    `Transparency` sentinel (`-2147483648`, an unset value) reports as `None`.
    Mirrors the spike: read `Shadow.Style` (not `.Type`, which goes mixed), `Glow.Radius`,
    `SoftEdge.Type`, `Reflection.Type`.
    """
    out: dict[str, Any] = {}

    shadow = _safe(lambda: com_shape.Shadow, None)
    if shadow is not None and _safe(lambda: is_true(shadow.Visible), False):
        out["shadow"] = {
            "color": _safe(lambda: color_hex_or_none(shadow.ForeColor.RGB), None),
            "transparency": _transparency(shadow),
            "blur": _safe(lambda: round(float(shadow.Blur), 3), None),
            "size": _safe(lambda: round(float(shadow.Size), 3), None),
            "offset_x": _safe(lambda: round(float(shadow.OffsetX), 3), None),
            "offset_y": _safe(lambda: round(float(shadow.OffsetY), 3), None),
        }

    glow = _safe(lambda: com_shape.Glow, None)
    glow_radius = _safe(lambda: float(glow.Radius), 0.0) if glow is not None else 0.0
    if glow is not None and glow_radius and glow_radius > 0:
        out["glow"] = {
            "color": _safe(lambda: color_hex_or_none(glow.Color.RGB), None),
            "radius": round(glow_radius, 3),
            "transparency": _transparency(glow),
        }

    soft = _safe(lambda: com_shape.SoftEdge, None)
    soft_type = _safe(lambda: int(soft.Type), 0) if soft is not None else 0
    if soft_type and soft_type > 0:
        out["soft_edge"] = {
            "type": soft_type,
            "radius": _safe(lambda: round(float(soft.Radius), 3), None),
        }

    refl = _safe(lambda: com_shape.Reflection, None)
    refl_type = _safe(lambda: int(refl.Type), 0) if refl is not None else 0
    if refl_type and refl_type > 0:
        out["reflection"] = {"type": refl_type}

    return out or None


def _transparency(effect: Any) -> float | None:
    """Read an effect's `.Transparency`, mapping the unset sentinel to None."""
    raw = _safe(lambda: float(effect.Transparency), None)
    if raw is None or int(raw) == _TRANSPARENCY_SENTINEL:
        return None
    return round(raw, 3)


def _hyperlink_to_dict(com_shape: Any) -> dict[str, Any] | None:
    """`{address, sub_address}` for the shape's mouse-click hyperlink, or None when unset.

    Reads `ActionSettings(ppMouseClick)`: a shape carries a link only when its
    `.Action` is `ppActionHyperlink`. `address` is an external URL/file target;
    `sub_address` is the in-deck jump (`"<SlideID>,<index>,<title>"`). Both empty
    strings collapse to `None`. Guarded — a frameless / action-less shape returns None.
    """
    try:
        acts = com_shape.ActionSettings(int(PpMouseActivation.MOUSE_CLICK))
        if int(acts.Action) != int(PpActionType.HYPERLINK):
            return None
        link = acts.Hyperlink
        address = str(_safe(lambda: link.Address, "") or "")
        sub_address = str(_safe(lambda: link.SubAddress, "") or "")
    except Exception:
        return None
    if not address and not sub_address:
        return None
    return {"address": address or None, "sub_address": sub_address or None}


def is_connector(com_shape: Any) -> bool:
    """True iff the shape is a connector line (`Shape.Connector` is msoTrue)."""
    return _safe(lambda: is_true(com_shape.Connector), False)


def is_group(com_shape: Any) -> bool:
    """True iff the shape is a group (`Shape.Type == msoGroup`)."""
    return _safe(lambda: int(com_shape.Type) == int(MsoShapeType.GROUP), False)


def _group_item_ids(com_shape: Any) -> list[int]:
    """The stable `Shape.Id`s of a group's members (the spike verified they survive
    grouping), so a read of a group exposes its children by drift-proof id."""

    def _walk() -> list[int]:
        items = com_shape.GroupItems
        return [int(items(i).Id) for i in range(1, int(items.Count) + 1)]

    return _safe(_walk, [])


def _center_of(com_shape: Any) -> tuple[float, float]:
    """The center point (x, y) of a shape's bounding box, in points."""
    return (
        float(com_shape.Left) + float(com_shape.Width) / 2.0,
        float(com_shape.Top) + float(com_shape.Height) / 2.0,
    )


def _check_site(com_shape: Any, site: int, label: str) -> None:
    """Validate a 1-based connection-site index against `Shape.ConnectionSiteCount`."""
    count = _safe(lambda: int(com_shape.ConnectionSiteCount), 0)
    if int(site) < 1 or int(site) > count:
        raise ValueError(
            f"{label}={site} out of range; {shape_type_name(com_shape.Type)} has "
            f"{count} connection site(s)"
        )


def _connector_to_dict(com_shape: Any) -> dict[str, Any] | None:
    """`{type, begin_shape_id, end_shape_id}` for a connector, or None for a non-connector.

    A connector reports its line geometry (`ConnectorFormat.Type`) and the stable
    `Shape.Id` of whatever shape each end is glued to (or `None` when that end
    floats free). Per the spike, `RerouteConnections()` may have re-chosen the
    actual connection sites, so the glued-shape ids — not the requested sites — are
    what reads back meaningfully.
    """
    if not is_connector(com_shape):
        return None
    cf = _safe(lambda: com_shape.ConnectorFormat, None)
    if cf is None:
        return None

    def _end_id(connected_attr: str, shape_attr: str) -> int | None:
        if not _safe(lambda: is_true(getattr(cf, connected_attr)), False):
            return None
        return _safe(lambda: int(getattr(cf, shape_attr).Id), None)

    return {
        "type": _safe(lambda: connector_type_name(int(cf.Type)), None),
        "begin_shape_id": _end_id("BeginConnected", "BeginConnectedShape"),
        "end_shape_id": _end_id("EndConnected", "EndConnectedShape"),
    }


def effect_to_dict(eff: Any, slide_index: int) -> dict[str, Any]:
    """One animation `Effect` -> `{shapeid, shape, effect, exit, trigger, duration,
    delay}` (the row `slide.animations()` and `Shape.animate` echo).

    Maps the effect back to its target shape via the drift-proof
    `shapeid:S:ID` (`Effect.Shape.Id`) plus the shape `name`, and decodes the
    `EffectType`/`TriggerType` ints to friendly names. `exit` is True for an exit
    (animate-out) effect. `duration` is the animation length in seconds and `delay`
    the `after`/`with` start delay — both read off `Effect.Timing`.

    Every field is read defensively (`_safe`): one unreadable property on a single
    effect (an exotic `Timing`, a motion-path `EffectType`) degrades to `None`
    rather than failing the whole `slide.animations()` listing — matching the
    rest of the read-to-dict helpers. A genuine busy still propagates.
    """
    shape_id = _safe(lambda: int(eff.Shape.Id), None)
    return {
        "shapeid": None if shape_id is None else f"shapeid:{slide_index}:{shape_id}",
        "shape": _safe(lambda: str(eff.Shape.Name), None),
        "effect": _safe(lambda: anim_effect_name(int(eff.EffectType)), None),
        "exit": _safe(lambda: is_true(eff.Exit), None),
        "trigger": _safe(lambda: anim_trigger_name(int(eff.Timing.TriggerType)), None),
        "duration": _safe(lambda: float(eff.Timing.Duration), None),
        "delay": _safe(lambda: float(eff.Timing.TriggerDelayTime), None),
    }


def shape_to_dict(com_shape: Any, slide_index: int, z_index: int) -> dict[str, Any]:
    """Structured snapshot of one shape for `slide.read()` / `shapes.list()`.

    Emits the canonical `anchor_id` (`shape:S:N`) plus the drift-proof `shapeid`
    (`shapeid:S:ID`), `name`, and `id`, the friendly shape `type`, `geometry`
    (points), placeholder kind (or None), and `text` when the shape has a text
    frame.
    """
    d: dict[str, Any] = {
        "anchor_id": f"shape:{slide_index}:{z_index}",
        "shapeid": f"shapeid:{slide_index}:{int(com_shape.Id)}",
        "index": z_index,
        "name": str(com_shape.Name),
        "id": int(com_shape.Id),
        "type": shape_type_name(com_shape.Type),
        "alt_text": _alt_text_of(com_shape),
        "geometry": _geometry_of(com_shape),
        "fill": _fill_to_dict(com_shape),
        "line": _line_to_dict(com_shape),
        "effects": _effects_to_dict(com_shape),
        "hyperlink": _hyperlink_to_dict(com_shape),
        "connector": _connector_to_dict(com_shape),
    }
    if is_group(com_shape):
        d["group_item_ids"] = _group_item_ids(com_shape)
    if is_placeholder(com_shape):
        try:
            d["placeholder"] = placeholder_kind_name(com_shape.PlaceholderFormat.Type)
        except Exception:
            d["placeholder"] = None
    else:
        d["placeholder"] = None

    d["has_table"] = has_table(com_shape)
    d["has_chart"] = has_chart(com_shape)
    d["has_smartart"] = has_smartart(com_shape)
    d["has_media"] = is_media(com_shape)
    if d["has_media"]:
        d["media"] = _media_to_dict(com_shape)
    if has_text_frame(com_shape):
        d["has_text_frame"] = True
        try:
            d["text"] = str(com_shape.TextFrame.TextRange.Text or "")
        except Exception:
            d["text"] = None
        d["links"] = _safe(lambda: links_in_range(com_shape.TextFrame.TextRange), [])
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
        """The shape's `.Name` (e.g. "Title 1") — drift-proof, unique per slide.

        Propagates a missing-shape / busy error like `shape_id` and `shape_type`
        do — it must never fabricate an `anchor_id`-shaped string, which would
        collide with the `shape:S:N` anchor format and mislead a caller into
        treating a failed lookup as a real shape name.
        """
        with _com.translate_com_errors():
            return str(self._com_shape().Name)

    @property
    def shape_id(self) -> int:
        """`Shape.Id` — stable across z-order reordering."""
        with _com.translate_com_errors():
            return int(self._com_shape().Id)

    @property
    def shapeid(self) -> str:
        """The restack-proof anchor for this shape — `shapeid:S:ID`.

        Built live from `Shape.Id`, so it survives the z-order drift that shifts
        `shape:S:N` when shapes are added / deleted / reordered. Every shape read
        and mutation echoes this alongside `anchor_id`, so an agent can chain
        edits on a shape across a restack without re-reading the slide.
        """
        with _com.translate_com_errors():
            return f"shapeid:{self._slide.index}:{int(self._com_shape().Id)}"

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

    @property
    def has_smartart(self) -> bool:
        """Whether this shape holds a SmartArt diagram (`Shape.HasSmartArt`)."""
        with _com.translate_com_errors():
            return has_smartart(self._com_shape())

    @property
    def smartart(self) -> SmartArt:
        """The shape's `SmartArt` diagram (its node tree).

        Raises `AnchorNotFoundError` (kind `"smartart"`, exit 2) if the shape
        holds no SmartArt.
        """
        from ._smartart import SmartArt

        with _com.translate_com_errors():
            if not has_smartart(self._com_shape()):
                raise AnchorNotFoundError("smartart", self.anchor_id)
        return SmartArt(self)

    @property
    def has_media(self) -> bool:
        """Whether this shape is an audio/video clip (`Shape.Type == msoMedia`)."""
        with _com.translate_com_errors():
            return is_media(self._com_shape())

    @property
    def media(self) -> dict[str, Any]:
        """The media clip's `{type, length_s, muted, volume, autoplay}` read.

        Raises `AnchorNotFoundError` (kind `"media"`, exit 2) if the shape holds
        no media.
        """
        with _com.translate_com_errors():
            sh = self._com_shape()
            if not is_media(sh):
                raise AnchorNotFoundError("media", self.anchor_id)
            return _media_to_dict(sh)

    def _text_range(self) -> Any:
        sh = self._com_shape()
        if not has_text_frame(sh):
            raise NoTextFrameError(self.anchor_id)
        return sh.TextFrame.TextRange

    def text_frame_status(self) -> TextFrameStatus:
        """Autofit / wrap / margin diagnostics for this shape's text frame.

        A **read** (no view move, no edit fence) that exposes the state behind a
        "formatting spiral": the autofit mode, word-wrap, inner margins (points),
        and a coarse `overflow_risk` (see `TextFrameStatus`). Raises
        `NoTextFrameError` if the shape holds no text frame.
        """
        with _com.translate_com_errors():
            sh = self._com_shape()
            if not has_text_frame(sh):
                raise NoTextFrameError(self.anchor_id)
            tf = sh.TextFrame
            autosize = _autosize_of(sh)
            margins = {
                "left": _safe(lambda: float(tf.MarginLeft), 0.0),
                "right": _safe(lambda: float(tf.MarginRight), 0.0),
                "top": _safe(lambda: float(tf.MarginTop), 0.0),
                "bottom": _safe(lambda: float(tf.MarginBottom), 0.0),
            }
            word_wrap = _safe(lambda: is_true(tf.WordWrap), True)
        return TextFrameStatus(
            autosize=autosize,
            word_wrap=word_wrap,
            margins=margins,
            overflow_risk=_overflow_risk(autosize),
        )

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

    def reset_to_layout(self) -> dict[str, float]:
        """Restore this *placeholder's* geometry (+ default font size) from its layout.

        The recovery verb for a placeholder that's been manually moved/resized or
        shrunk to an unreadable font (the gpt-5.4 review's "5 pt font, overflow off
        the slide" case). Matches this shape to its slide's `CustomLayout`
        placeholder by `PlaceholderFormat.Type`, then copies that placeholder's
        `Left`/`Top`/`Width`/`Height` (and, best-effort, its default font size onto
        the live text) — the layout's reading is the source of truth (verified in
        `scripts/text_model_spike.py`). Returns the restored `{left, top, width,
        height[, font_size]}` in points.

        Raises `ValueError` if this shape isn't a placeholder, or
        `AnchorNotFoundError` if the layout has no placeholder of the same kind.
        Pairs with `reset_format` (which clears the paragraph-spacing spiral). A
        mutation: wrap in `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            sh = self._com_shape()
            try:
                want_type = int(sh.PlaceholderFormat.Type)
            except Exception as exc:  # not a placeholder -> no PlaceholderFormat
                raise ValueError(
                    f"reset_to_layout() needs a placeholder shape, got {self.anchor_id}"
                ) from exc
            match = self._layout_placeholder(want_type)
            sh.Left = float(match.Left)
            sh.Top = float(match.Top)
            sh.Width = float(match.Width)
            sh.Height = float(match.Height)
            restored: dict[str, float] = {
                "left": float(sh.Left),
                "top": float(sh.Top),
                "width": float(sh.Width),
                "height": float(sh.Height),
            }
            size = _layout_default_size(match)
            if size is not None and has_text_frame(sh):
                sh.TextFrame.TextRange.Font.Size = size
                restored["font_size"] = size
        return restored

    def _layout_placeholder(self, placeholder_type: int) -> Any:
        """The `CustomLayout` placeholder matching `placeholder_type` (live COM)."""
        phs = self._slide.com.CustomLayout.Shapes.Placeholders
        for i in range(1, int(phs.Count) + 1):
            ph = phs(i)
            try:
                if int(ph.PlaceholderFormat.Type) == placeholder_type:
                    return ph
            except Exception:
                continue
        raise AnchorNotFoundError("layout placeholder", self.anchor_id)

    def set_alt_text(self, value: str) -> None:
        """Set the shape's alternative (accessibility) text (`Shape.AlternativeText`).

        The drift-proof re-identification handle (see `alt_text`). A mutation:
        wrap in `deck.edit(...)` for view preservation + a one-Ctrl-Z fence.
        """
        with _com.translate_com_errors():
            self._com_shape().AlternativeText = str(value)

    def set_fill(
        self,
        *,
        fill: str | int | tuple[int, int, int] | None = None,
        line: str | int | tuple[int, int, int] | None = None,
        line_width: float | None = None,
        fill_transparency: float | None = None,
        line_transparency: float | None = None,
    ) -> None:
        """Set the shape's **fill** and/or **line** (border). Only the kwargs passed.

        `fill`/`line` take a color (`"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB
        int) for a solid fill / line of that color, or the string `"none"` to make
        the fill transparent / remove the border. `line_width` is the border weight
        in points. `fill_transparency`/`line_transparency` are `0.0..1.0` alpha
        fractions (0 opaque, 1 fully transparent) — the partial-alpha knob, distinct
        from `"none"` (which hides the fill/line entirely). Distinct from
        `format_text`'s `color`, which is *font* color. Raises `ValueError` for a bad
        color / out-of-range transparency (before any COM) or if nothing is passed. A
        mutation: wrap in `deck.edit(...)`.
        """
        if (
            fill is None
            and line is None
            and line_width is None
            and fill_transparency is None
            and line_transparency is None
        ):
            raise ValueError(
                "set_fill() requires at least one of fill=, line=, line_width=, "
                "fill_transparency=, or line_transparency="
            )
        with _com.translate_com_errors():
            apply_shape_fill(
                self._com_shape(),
                fill=fill,
                line=line,
                line_width=line_width,
                fill_transparency=fill_transparency,
                line_transparency=line_transparency,
            )

    def set_line_style(
        self,
        *,
        dash: str | int | None = None,
        begin_arrow: str | int | None = None,
        end_arrow: str | int | None = None,
        begin_arrow_size: str | int | None = None,
        end_arrow_size: str | int | None = None,
    ) -> None:
        """Set the shape's line **dash** pattern and/or **arrowheads**. Only the kwargs passed.

        `dash` is a friendly `MsoLineDashStyle` (`"solid"`/`"dash"`/`"round_dot"`/
        `"dash_dot"`/`"long_dash"`/…) or raw int. `begin_arrow`/`end_arrow` are
        `MsoArrowheadStyle` names (`"none"`/`"triangle"`/`"open"`/`"stealth"`/
        `"diamond"`/`"oval"`) or raw ints; `begin_arrow_size`/`end_arrow_size` are
        `"small"`/`"medium"`/`"large"` (set both arrowhead length + width). Names are
        validated up front (`ValueError` before any COM). **Arrowheads apply to
        lines/connectors only** — PowerPoint raises on a closed shape (use `dash` for
        those). Border color/weight stay on `set_fill`. A mutation: wrap in
        `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            apply_line_style(
                self._com_shape(),
                dash=dash,
                begin_arrow=begin_arrow,
                end_arrow=end_arrow,
                begin_arrow_size=begin_arrow_size,
                end_arrow_size=end_arrow_size,
            )

    def set_gradient_fill(
        self,
        colors: Sequence[Any] | None = None,
        *,
        positions: Sequence[float] | None = None,
        style: str | int = "horizontal",
        variant: int = 1,
        degree: float | None = None,
        preset: str | int | None = None,
    ) -> None:
        """Give this shape a **gradient** fill. A mutation: wrap in `deck.edit(...)`.

        Pass `colors` (a list of `"#RRGGBB"` / `(r,g,b)` / int colors) **or** `preset`
        (a named ramp like `"ocean"` / `"fire"` / `"rainbow"`):
        - one color → a one-color gradient (`degree` is the 0..1 brightness, default 0.5);
        - two colors → a two-color gradient (first at stop 0.0, second at 1.0);
        - three+ colors → a multi-stop gradient; `positions` (floats 0..1, same length as
          `colors`) places the *interior* stops, which otherwise space evenly (the
          endpoints stay at 0.0/1.0).

        `style` is a friendly `MsoGradientStyle` (`"horizontal"`/`"vertical"`/
        `"diagonal_up"`/…) and `variant` (1-4) picks the shading variant. Raises
        `ValueError` (before any COM) for a bad color / style / preset, or if neither
        `colors` nor `preset` is given.
        """
        with _com.translate_com_errors():
            apply_gradient_fill(
                self._com_shape(),
                colors=colors,
                positions=positions,
                style=style,
                variant=variant,
                degree=degree,
                preset=preset,
            )

    def set_picture_fill(self, path: str | Path) -> None:
        """Fill this shape with an **image** (`Fill.UserPicture`). A mutation: `deck.edit(...)`.

        `path` is resolved to an absolute path (a relative path raises
        `ERROR_FILE_NOT_FOUND` in PowerPoint); a missing file raises `FileNotFoundError`
        (before any COM).

        For an actual **picture** shape this is the wrong verb — it sets a fill
        *behind* the unchanged picture raster (so the image doesn't visibly change);
        use `set_picture` to re-source a picture in place.
        """
        with _com.translate_com_errors():
            apply_picture_fill(self._com_shape(), path)

    def set_picture(
        self,
        path: str | os.PathLike[str],
        *,
        alt_text: str | None = None,
    ) -> Shape:
        """Re-source this **picture** in place — swap its image without delete-and-recreate.

        The post-creation edit for a picture: replaces the displayed image while
        **preserving the picture's position, size, rotation, name, alt text, and
        z-order slot**, so an agent can update a logo / screenshot / chart export
        without re-deriving geometry (the wordlive delete-then-recreate habit). The
        new image is **embedded** (never linked); `alt_text`, if given, overrides
        the carried-over alt text.

        Under the hood this *is* a delete + re-insert — PowerPoint's COM exposes no
        in-place image swap for a picture shape (`Fill.UserPicture` only sets a
        fill *behind* the unchanged raster, confirmed in
        `scripts/set_picture_spike.py`). Two honest consequences: the picture gets
        a **new `Shape.Id`** (so this returns a fresh handle — the old wrapper is
        spent, like after `delete()`), and anything bound to the old picture object
        — **animations, hyperlinks, crop, and picture adjustments**
        (brightness / contrast / recolor) — is **not** carried over. Position, size,
        rotation, name, alt text, and z-order are.

        Raises `FileNotFoundError` if `path` is missing, or `ValueError` if this
        shape isn't a picture (both before any COM mutation — use `set_picture_fill`
        to put an image into a non-picture shape's fill). Returns a `shapeid:S:ID`
        handle to the new picture. A mutation: wrap in `deck.edit(...)`.
        """
        fs_path = os.fspath(path)
        if not os.path.isfile(fs_path):
            raise FileNotFoundError(f"picture not found: {fs_path}")
        abs_path = os.path.abspath(fs_path)
        with _com.translate_com_errors():
            com_old = self._com_shape()
            if not is_picture(com_old):
                raise ValueError(
                    f"set_picture() needs a picture shape, got "
                    f"{shape_type_name(com_old.Type)} ({self.anchor_id}); use "
                    f"set_picture_fill() to put an image into a non-picture shape's fill"
                )
            new_id = replace_picture(
                self._slide.com, com_old, abs_path, self._slide.index, alt_text=alt_text
            )
        return ShapeById(self._slide, new_id)

    def set_pattern_fill(
        self,
        pattern: str | int,
        *,
        fore: Any,
        back: Any | None = None,
    ) -> None:
        """Give this shape a two-color **pattern** fill. A mutation: wrap in `deck.edit(...)`.

        `pattern` is a friendly `MsoPatternType` name (`"percent_50"`, `"trellis"`,
        `"dark_horizontal"`, …) or a raw int; `fore` is the pattern color and `back`
        the (optional) background color. Raises `ValueError` (before any COM) for a bad
        pattern name or color.
        """
        with _com.translate_com_errors():
            apply_pattern_fill(self._com_shape(), pattern=pattern, fore=fore, back=back)

    def set_effect(
        self,
        *,
        shadow: Any | None = None,
        glow: Any | None = None,
        soft_edge: Any | None = None,
        reflection: Any | None = None,
    ) -> None:
        """Set shape **effects** — shadow / glow / soft-edge / reflection. Only kwargs passed.

        Each takes its friendly form or the string `"none"` to turn it off:
        - `shadow` — a dict `{color?, transparency?, blur?, size?, offset_x?, offset_y?}`
          (an outer drop shadow);
        - `glow` — a dict `{color?, radius?, transparency?}` (radius 0 = off);
        - `soft_edge` — an int preset `0`-`6` (0 = off);
        - `reflection` — an int preset `0`-`9` (0 = off).

        Raises `ValueError` (before any COM) for a bad color or a non-dict shadow/glow.
        A mutation: wrap in `deck.edit(...)`.
        """
        if shadow is None and glow is None and soft_edge is None and reflection is None:
            raise ValueError(
                "set_effect() requires at least one of shadow=, glow=, soft_edge=, or reflection="
            )
        with _com.translate_com_errors():
            apply_effect(
                self._com_shape(),
                shadow=shadow,
                glow=glow,
                soft_edge=soft_edge,
                reflection=reflection,
            )

    def reorder(self, to: str | int) -> int:
        """Restack this shape in the slide's z-order; return its new 1-based position.

        `to` is `"front"` / `"back"` / `"forward"` / `"backward"` (or a raw
        `MsoZOrderCmd` int) — bring to front, send to back, or step one level
        (`constants.ZORDER_CHOICES`). Lets a freshly added background panel slide
        *behind* existing content (otherwise it always lands on top). Note this
        shifts the `shape:S:N` indices of the shapes it passes — re-read after, or
        address by `shapeid:S:ID` / `.Name`. Raises `ValueError` for an unknown
        command (before any COM). A mutation: wrap in `deck.edit(...)`.
        """
        cmd = zorder_cmd_for(to)  # ValueError before any COM
        with _com.translate_com_errors():
            sh = self._com_shape()  # raw ref tracks the shape across the restack
            shape_id = int(sh.Id)
            sh.ZOrder(cmd)
            # Report the Shapes-collection index (the basis shape:S:N resolves by),
            # not ZOrderPosition — they coincide on a flat slide but can diverge,
            # and the returned position should be usable as shape:S:N.
            found = find_shape_by_id(self._slide.com, shape_id)
            return found[0] if found is not None else int(sh.ZOrderPosition)

    def ungroup(self) -> list[ShapeById]:
        """Ungroup this group shape; return drift-proof handles to the freed members.

        Reverses `ShapeCollection.group` — `GroupShape.Ungroup()` dissolves the
        group and frees its members back onto the slide. Per the spike
        (`scripts/arrangement_spike.py`), the freed children **keep their original
        `Shape.Id`s**, so this returns a `ShapeById` per member (resolvable as
        `shapeid:S:ID`). The group's own wrapper is spent afterwards (its shape no
        longer exists), like after `delete()`.

        Raises `ValueError` (before any COM) if this shape is not a group. A
        mutation: wrap in `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        with _com.translate_com_errors():
            com = self._com_shape()
            if not is_group(com):
                raise ValueError(
                    f"ungroup() needs a group shape, got "
                    f"{shape_type_name(com.Type)} ({self.anchor_id})"
                )
            freed = com.Ungroup()
            child_ids = [int(freed(i).Id) for i in range(1, int(freed.Count) + 1)]
        return [ShapeById(self._slide, cid) for cid in child_ids]

    def animate(
        self,
        effect: str | int = "fade",
        *,
        trigger: str | int = "on_click",
        duration: float | None = None,
        delay: float | None = None,
        exit: bool = False,
    ) -> dict[str, Any]:
        """Give this shape an entrance (or exit) animation; return the effect dict.

        Appends an effect to the slide's main animation sequence
        (`Slide.TimeLine.MainSequence.AddEffect`). `effect` is a friendly
        `MsoAnimEffect` name (`"fade"`/`"appear"`/`"fly_in"`/… — see
        `constants.ANIM_EFFECT_CHOICES`) or a raw int. `trigger` is when it fires —
        `"on_click"` (default), `"with_previous"`, or `"after_previous"`.
        `duration` is the animation length in seconds and `delay` the start delay in
        seconds (both optional; PowerPoint's per-effect default applies when None).
        Pass `exit=True` to make the shape animate **out** instead of in (the
        "disappear" case) — the same effect ids serve both.

        A shape can carry several effects; each `animate()` call adds one (clear
        them with `clear_animations()`). Raises `ValueError` (before any COM) for an
        unknown effect/trigger name. A mutation: wrap in `deck.edit(...)`.
        """
        effect_int = anim_effect_for(effect)  # ValueError before any COM
        trigger_int = anim_trigger_for(trigger)
        with _com.translate_com_errors():
            seq = self._slide.com.TimeLine.MainSequence
            eff = seq.AddEffect(self._com_shape(), effect_int, 0, trigger_int)
            if exit:
                eff.Exit = int(MsoTriState.TRUE)
            if duration is not None:
                eff.Timing.Duration = float(duration)
            if delay is not None:
                eff.Timing.TriggerDelayTime = float(delay)
            return effect_to_dict(eff, self.slide.index)

    def clear_animations(self) -> int:
        """Remove every animation effect targeting **this shape**; return the count.

        Walks the slide's main sequence and deletes each effect whose target is this
        shape (matched by stable `Shape.Id`, so a restack doesn't matter), leaving
        other shapes' animations intact. Use `Slide.clear_animations()` to wipe the
        whole slide. A no-op (returns 0) if the shape has no animations. A mutation:
        wrap in `deck.edit(...)`.
        """
        return self._slide.clear_animations(anchor=self)

    def set_hyperlink(
        self,
        *,
        url: str | None = None,
        slide: int | None = None,
        screen_tip: str | None = None,
    ) -> dict[str, Any]:
        """Make this shape a clickable hyperlink; return the resulting link dict.

        Pass **exactly one** destination: `url` for an external link (a URL,
        `mailto:`, or file path) or `slide` for an in-deck jump to a 1-based slide
        index ("back to TOC" / agenda navigation). `screen_tip` is the optional
        hover tooltip. The link fires on mouse click (`ppMouseClick`); setting it
        implicitly flips the shape's action to `ppActionHyperlink`. A shape needs
        no text frame to carry a link (it's a shape-level action).

        Raises `ValueError` (before any COM) if neither or both destinations are
        given, `url` is blank, or `slide` is out of range. A mutation: wrap in
        `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        if (url is None) == (slide is None):
            raise ValueError("set_hyperlink() requires exactly one of url= or slide=")
        sub_address: str | None = None
        if url is not None:
            if not str(url).strip():
                raise ValueError("set_hyperlink(url=) must be a non-empty string")
            url = str(url)
        else:
            assert slide is not None  # narrowed by the exactly-one check above
            sub_address = self._slide_jump_subaddress(int(slide))  # ValueError if out of range
        with _com.translate_com_errors():
            sh = self._com_shape()  # resolve once; reuse for mutation + readback
            acts = sh.ActionSettings(int(PpMouseActivation.MOUSE_CLICK))
            link = acts.Hyperlink
            if url is not None:
                link.Address = url
            else:
                # SubAddress alone makes it an in-deck jump; clear any stale Address.
                link.Address = ""
                link.SubAddress = sub_address
            if screen_tip is not None:
                link.ScreenTip = str(screen_tip)
            return _hyperlink_to_dict(sh) or {}

    def _slide_jump_subaddress(self, slide_index: int) -> str:
        """The `"<SlideID>,<index>,<title>"` SubAddress for an in-deck jump.

        Delegates to the shared `_anchors.slide_jump_subaddress` (also used by the
        text-run `Anchor.set_link`).
        """
        return slide_jump_subaddress(self._slide, slide_index)

    def remove_hyperlink(self) -> None:
        """Remove this shape's mouse-click hyperlink (`Hyperlink.Delete`).

        Reverts the shape's action to `ppActionNone` and clears the address. A
        no-op if the shape has no link. A mutation: wrap in `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            acts = self._com_shape().ActionSettings(int(PpMouseActivation.MOUSE_CLICK))
            acts.Hyperlink.Delete()

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


class ShapeById(Shape):
    """A shape addressed by its stable `Shape.Id` — `shapeid:S:ID`.

    The **delete-proof** handle (PPTLIVE-010): `Shape.Id` is assigned once and is
    never renumbered or reused, so unlike `shape:S:N` — a z-order index that
    shifts down when a lower shape is deleted or restacked — a `shapeid` keeps
    pointing at the same shape across structural edits. Resolves by scanning the
    slide's shapes for the matching `.Id` on every access (so z-order drift on the
    host slide is handled). The `id` emitted in every shape listing *is* this
    value, so an agent can build `shapeid:S:ID` straight from a read.
    """

    kind = "shape"

    def __init__(self, slide: Slide, shape_id: int) -> None:
        super().__init__(slide, index=0)
        self._shape_id = int(shape_id)

    @property
    def target_id(self) -> int:
        """The `Shape.Id` this handle resolves by."""
        return self._shape_id

    @property
    def anchor_id(self) -> str:
        return f"shapeid:{self._slide.index}:{self._shape_id}"

    def _resolve(self) -> tuple[int, Any]:
        """Resolve to `(1-based Shapes-collection index, COM Shape)` live by `.Id`.

        Uses the collection index from `find_shape_by_id`, *not* `ZOrderPosition`:
        `shape:S:N` resolves via `Shapes(N)`, so the emitted `shape:S:N` anchor
        must use the same basis or it could point at a different shape than the
        `shapeid` it was read from (the two coincide on a flat slide but can
        diverge with grouped/placeholder orderings).
        """
        found = find_shape_by_id(self._slide.com, self._shape_id)
        if found is None:
            raise AnchorNotFoundError("shape", self.anchor_id)
        return found

    def _com_shape(self) -> Any:
        """Resolve the COM `Shape` live by stable `.Id`. Never cached."""
        return self._resolve()[1]

    @property
    def index(self) -> int:
        """Current 1-based `Shapes`-collection index of the resolved shape."""
        with _com.translate_com_errors():
            return self._resolve()[0]

    def to_dict(self) -> dict[str, Any]:
        with _com.translate_com_errors():
            idx, sh = self._resolve()
            return shape_to_dict(sh, self._slide.index, idx)


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
                matches: list[dict[str, Any]] = [
                    {
                        "anchor_id": f"shape:{self._slide.index}:{idx}",
                        "name": key,
                        "id": int(sh.Id),
                        "index": idx,
                    }
                    for idx, sh in enumerate(self._com_collection, start=1)
                    if str(sh.Name) == key
                ]
            if not matches:
                raise AnchorNotFoundError("shape", key)
            if len(matches) > 1:
                # PowerPoint allows duplicate shape names (paste/duplicate). Rather
                # than silently pick the first, surface the ambiguity with the
                # candidate shape:S:N anchors (same convention as _find_placeholder).
                anchors = ", ".join(str(c["anchor_id"]) for c in matches)
                raise AmbiguousMatchError(
                    key,
                    matches,
                    message=(
                        f"{len(matches)} shapes named {key!r} on slide "
                        f"{self._slide.index} ({anchors}); target one by shape:S:N"
                    ),
                )
            return Shape(self._slide, int(matches[0]["index"]))
        raise TypeError(f"shape key must be int or str, got {type(key).__name__}")

    def __contains__(self, key: object) -> bool:
        if isinstance(key, bool) or not isinstance(key, (int, str)):
            return False
        try:
            self[key]
            return True
        except AnchorNotFoundError:
            return False
        except AmbiguousMatchError:
            return True  # the name exists (more than once) — membership is still True

    def __iter__(self) -> Iterator[Shape]:
        count = len(self)
        for idx in range(1, count + 1):
            yield Shape(self._slide, idx)

    def by_id(self, shape_id: int) -> ShapeById:
        """A shape addressed by its stable `Shape.Id` (`shapeid:S:ID`) — delete-proof.

        Unlike `slide.shapes[N]` (z-order index) or `slide.shapes["Name"]`, the id
        survives a delete/restack that renumbers indices (PPTLIVE-010). Verifies
        the id exists now (raising `AnchorNotFoundError` if not), then returns a
        `ShapeById` that re-resolves live on each use.
        """
        handle = ShapeById(self._slide, shape_id)
        with _com.translate_com_errors():
            handle._com_shape()  # eager existence check (clean exit-2 if absent)
        return handle

    # -- arrangement (v-next; wrap in deck.edit(...) for view + one-Ctrl-Z) ----

    def _range_indices(self, shapes: Sequence[Shape]) -> list[int]:
        """Resolve a list of `Shape` handles to their *current* z-order indices.

        Uses the stable `Shape.Id` of each handle to look up its live collection
        index (so the result is robust against duplicate shape names and the
        z-order an upstream op may have drifted). Raises `AnchorNotFoundError` if a
        handle no longer resolves on this slide. The caller passes the int list to
        `Shapes.Range(...)`.
        """
        slide_com = self._slide.com
        indices: list[int] = []
        for sh in shapes:
            sid = int(sh._com_shape().Id)
            found = find_shape_by_id(slide_com, sid)
            if found is None:
                raise AnchorNotFoundError("shape", sh.anchor_id)
            indices.append(found[0])
        return indices

    def group(self, shapes: Sequence[Shape]) -> ShapeById:
        """Group two or more shapes into a single group shape; return its handle.

        `Shapes.Range([...]).Group()` combines the shapes; the new group gets a
        **fresh `Shape.Id`** (so this returns a `ShapeById` for it), while the
        members keep their own ids inside `group.GroupItems` (read back in the
        group's `group_item_ids`) — verified in `scripts/arrangement_spike.py`.
        Reverse with `Shape.ungroup`.

        Raises `ValueError` (before any COM) for fewer than two shapes; an unknown
        member raises `AnchorNotFoundError`. A mutation: wrap in `deck.edit(...)`.
        """
        members = list(shapes)
        if len(members) < 2:
            raise ValueError(f"group() needs at least 2 shapes, got {len(members)}")
        with _com.translate_com_errors():
            indices = self._range_indices(members)
            group = self._com_collection.Range(indices).Group()
            group_id = int(group.Id)
        return ShapeById(self._slide, group_id)

    def align(
        self, shapes: Sequence[Shape], how: str | int, *, relative_to: str | int | bool = "slide"
    ) -> None:
        """Align a set of shapes to a common edge / center.

        `how` is `"left"`/`"center"`/`"right"` (horizontal) or `"top"`/`"middle"`/
        `"bottom"` (vertical) — see `constants.ALIGN_CHOICES`. `relative_to` is
        `"slide"` (align against the slide, the default) or `"selection"` (align
        the shapes to one another). `Shapes.Range([...]).Align(cmd, RelativeTo)`.

        Raises `ValueError` (before any COM) for an unknown `how`/`relative_to`, an
        empty set, or a selection-relative align of fewer than two shapes. A
        mutation: wrap in `deck.edit(...)`.
        """
        cmd = align_cmd_for(how)  # ValueError before any COM
        rel = relative_to_for(relative_to)
        members = list(shapes)
        if not members:
            raise ValueError("align() needs at least one shape")
        if rel == int(MsoTriState.FALSE) and len(members) < 2:
            raise ValueError("aligning relative to the selection needs at least 2 shapes")
        with _com.translate_com_errors():
            indices = self._range_indices(members)
            self._com_collection.Range(indices).Align(cmd, rel)

    def distribute(
        self, shapes: Sequence[Shape], how: str | int, *, relative_to: str | int | bool = "slide"
    ) -> None:
        """Space a set of shapes evenly on one axis.

        `how` is `"horizontal"` or `"vertical"` (`constants.DISTRIBUTE_CHOICES`);
        `relative_to` is `"slide"` (default) or `"selection"`.
        `Shapes.Range([...]).Distribute(cmd, RelativeTo)`. Distributing is only
        meaningful for three or more shapes (the two outermost pin the span and the
        rest are evenly spaced between them).

        Raises `ValueError` (before any COM) for an unknown `how`/`relative_to` or
        fewer than three shapes. A mutation: wrap in `deck.edit(...)`.
        """
        cmd = distribute_cmd_for(how)  # ValueError before any COM
        rel = relative_to_for(relative_to)
        members = list(shapes)
        if len(members) < 3:
            raise ValueError(f"distribute() needs at least 3 shapes, got {len(members)}")
        with _com.translate_com_errors():
            indices = self._range_indices(members)
            self._com_collection.Range(indices).Distribute(cmd, rel)

    def add_connector(
        self,
        connector_type: str | int = "straight",
        *,
        begin: Shape | None = None,
        end: Shape | None = None,
        begin_site: int = 1,
        end_site: int = 1,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> ShapeById:
        """Add a connector line and return its handle (`Shapes.AddConnector`).

        Two forms:

        - **Attached (primary):** pass `begin=` and `end=` shape handles to glue the
          two ends to those shapes (via `ConnectorFormat.BeginConnect` /
          `EndConnect` + `RerouteConnections`), so the line follows them when they
          move. `begin_site` / `end_site` request a 1-based connection site on each
          shape, but are **advisory** — `RerouteConnections()` re-chooses the
          shortest sites (spike finding), and the resulting glue reads back under
          the shape's `connector` field.
        - **Geometry:** omit both shapes and pass explicit `left`/`top`/`width`/
          `height` (points) to draw a free-floating connector across that box.

        `connector_type` is `"straight"`/`"elbow"`/`"curved"`
        (`constants.CONNECTOR_CHOICES`). Raises `ValueError` (before any COM) for a
        bad type, an out-of-range site, or an incomplete spec (need both `begin` and
        `end`, or full explicit geometry). A mutation: wrap in `deck.edit(...)`.
        """
        type_int = connector_type_for(connector_type)  # ValueError before any COM
        if (begin is None) != (end is None):
            raise ValueError("add_connector() needs BOTH begin= and end= (or neither)")
        attaching = begin is not None and end is not None
        if not attaching and None in (left, top, width, height):
            raise ValueError(
                "add_connector() needs begin= and end= shapes, or explicit "
                "left/top/width/height geometry"
            )
        with _com.translate_com_errors():
            if attaching:
                assert begin is not None and end is not None  # narrowed above
                begin_com = begin._com_shape()
                end_com = end._com_shape()
                _check_site(begin_com, begin_site, "begin_site")
                _check_site(end_com, end_site, "end_site")
                x1, y1 = _center_of(begin_com)
                x2, y2 = _center_of(end_com)
            else:
                x1, y1 = float(left), float(top)  # type: ignore[arg-type]
                x2 = float(left) + float(width)  # type: ignore[arg-type]
                y2 = float(top) + float(height)  # type: ignore[arg-type]
            conn = self._com_collection.AddConnector(type_int, x1, y1, x2, y2)
            conn_id = int(conn.Id)
            if attaching:
                cf = conn.ConnectorFormat
                cf.BeginConnect(begin_com, int(begin_site))
                cf.EndConnect(end_com, int(end_site))
                conn.RerouteConnections()
        return ShapeById(self._slide, conn_id)

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
        fill: str | int | tuple[int, int, int] | None = None,
        line: str | int | tuple[int, int, int] | None = None,
        line_width: float | None = None,
    ) -> Shape:
        """Add a horizontal text box and return it (`Shapes.AddTextbox`).

        Geometry is in points; omitted values default to a 4×1 in box near the
        top-left. `text`, if given, is written into the new frame. `fill`/`line`
        set a solid fill / border color (or `"none"` for transparent / no border)
        and `line_width` the border weight in points — see `Shape.set_fill`. A
        text box defaults to no fill and no line. Raises `ValueError` for a bad
        color before any COM.
        """
        _precheck_fill(fill, line)  # ValueError before any COM
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
            if fill is not None or line is not None or line_width is not None:
                apply_shape_fill(com_shape, fill=fill, line=line, line_width=line_width)
            return self._added()

    def add_shape(
        self,
        shape_type: str | int,
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        fill: str | int | tuple[int, int, int] | None = None,
        line: str | int | tuple[int, int, int] | None = None,
        line_width: float | None = None,
    ) -> Shape:
        """Add an autoshape and return it (`Shapes.AddShape`).

        `shape_type` is a friendly name (`"rectangle"`, `"oval"`, `"arrow"`, …;
        see `constants.AUTOSHAPE_CHOICES`) or a raw `MsoAutoShapeType` int.
        Geometry is in points; omitted values default to a 2×2 in box near the
        top-left. `fill`/`line` set a solid fill / border color (or `"none"` for
        transparent / no border) and `line_width` the border weight in points —
        see `Shape.set_fill`; omitted, the shape takes the theme's default accent
        fill. Raises `ValueError` for an unknown shape name or bad color (before
        any COM).
        """
        type_int = autoshape_type_for(shape_type)  # ValueError before COM
        _precheck_fill(fill, line)  # ValueError before COM
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = _DEFAULT_SHAPE_WIDTH if width is None else float(width)
        height = _DEFAULT_SHAPE_HEIGHT if height is None else float(height)
        with _com.translate_com_errors():
            com_shape = self._com_collection.AddShape(type_int, left, top, width, height)
            if fill is not None or line is not None or line_width is not None:
                apply_shape_fill(com_shape, fill=fill, line=line, line_width=line_width)
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

    def _add_media(
        self,
        path: str | os.PathLike[str],
        *,
        left: float | None,
        top: float | None,
        width: float | None,
        height: float | None,
        default_width: float,
        default_height: float,
        link: bool,
        autoplay: bool,
        hide_icon: bool,
        pace_slide: bool,
        alt_text: str | None,
    ) -> Shape:
        """Shared core for `add_audio` / `add_video` (`Shapes.AddMediaObject2`).

        Validates the file before any COM, embeds (or links) the clip, optionally
        wires auto-play / hide-while-not-playing, and — when `pace_slide` — sets the
        slide to auto-advance to the clip length by reusing `Slide.set_transition`.
        """
        fs_path = os.fspath(path)
        if not os.path.isfile(fs_path):
            raise FileNotFoundError(f"media not found: {fs_path}")
        abs_path = os.path.abspath(fs_path)
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = default_width if width is None else float(width)
        height = default_height if height is None else float(height)
        with _com.translate_com_errors():
            com_shape = self._com_collection.AddMediaObject2(
                abs_path,
                int(MsoTriState.TRUE) if link else int(MsoTriState.FALSE),  # LinkToFile
                int(MsoTriState.FALSE) if link else int(MsoTriState.TRUE),  # SaveWithDocument
                left,
                top,
                width,
                height,
            )
            if alt_text is not None:
                com_shape.AlternativeText = str(alt_text)
            if autoplay or hide_icon:
                # PlaySettings drives auto-play / hide-while-not-playing. A media
                # object inserted via AddMediaObject2 always exposes it, so a
                # failure here is genuine — surface it (translated) rather than
                # silently drop the caller's request. `autoplay` defaults True and
                # the narrate → auto-advance → export_video flow depends on it, so a
                # silent swallow would produce a wrong video with no error.
                ps = com_shape.AnimationSettings.PlaySettings
                if autoplay:
                    ps.PlayOnEntry = int(MsoTriState.TRUE)
                if hide_icon:
                    ps.HideWhileNotPlaying = int(MsoTriState.TRUE)
            length_ms = _safe(lambda: float(com_shape.MediaFormat.Length), 0.0)
            shape = self._added()
        if pace_slide and length_ms and length_ms > 0:
            # Reuse the shipped transition writer (sets AdvanceOnTime + AdvanceTime).
            self._slide.set_transition(advance_after=max(1.0, length_ms / 1000.0))
        return shape

    def add_audio(
        self,
        path: str | os.PathLike[str],
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        link: bool = False,
        autoplay: bool = True,
        hide_icon: bool = True,
        pace_slide: bool = True,
        alt_text: str | None = None,
    ) -> Shape:
        """Insert an audio clip and return its `Shape` (`Shapes.AddMediaObject2`).

        The narration path: the clip is **embedded** (set `link=True` to keep it on
        disk and shrink the deck). `autoplay` plays it on slide entry, `hide_icon`
        hides the speaker icon while it isn't playing, and `pace_slide` sets the
        slide to auto-advance to the clip's length (so an exported video paces
        itself to the narration). Geometry is in points; omitted values default to a
        small icon box near the top-left. The shape's `.has_media` is True. Raises
        `FileNotFoundError` if `path` doesn't exist (before any COM). Wrap in
        `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        return self._add_media(
            path,
            left=left,
            top=top,
            width=width,
            height=height,
            default_width=_DEFAULT_AUDIO_WIDTH,
            default_height=_DEFAULT_AUDIO_HEIGHT,
            link=link,
            autoplay=autoplay,
            hide_icon=hide_icon,
            pace_slide=pace_slide,
            alt_text=alt_text,
        )

    def add_video(
        self,
        path: str | os.PathLike[str],
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
        link: bool = False,
        autoplay: bool = True,
        pace_slide: bool = True,
        alt_text: str | None = None,
    ) -> Shape:
        """Insert a video clip and return its `Shape` (`Shapes.AddMediaObject2`).

        Like `add_audio`, but the clip stays visible (there is no `hide_icon` — a
        video frame is meant to be seen). Geometry is in points; omitted values
        default to a 16:9 frame near the top-left. `autoplay` plays it on slide
        entry; `pace_slide` auto-advances the slide to the clip length. The shape's
        `.has_media` is True. Raises `FileNotFoundError` if `path` doesn't exist
        (before any COM). Wrap in `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        return self._add_media(
            path,
            left=left,
            top=top,
            width=width,
            height=height,
            default_width=_DEFAULT_VIDEO_WIDTH,
            default_height=_DEFAULT_VIDEO_HEIGHT,
            link=link,
            autoplay=autoplay,
            hide_icon=False,
            pace_slide=pace_slide,
            alt_text=alt_text,
        )

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

    def _resolve_smartart_layout(self, kind: str) -> Any:
        """The live `SmartArtLayout` COM object for a friendly `kind`.

        Resolves the friendly name -> URN segment (`smartart_layout_for`, raising
        `ValueError` before COM), then matches it against the running app's
        `SmartArtLayouts` by stable URN `.Id` (the collection index drifts).
        Raises `AnchorNotFoundError` (kind "smartart_layout") if the layout isn't
        installed in this PowerPoint build.
        """
        seg = smartart_layout_for(kind)  # ValueError before any COM
        layouts = self._com_collection.Application.SmartArtLayouts
        suffix = "/" + seg
        for i in range(1, int(layouts.Count) + 1):
            layout = layouts.Item(i)
            if str(layout.Id).endswith(suffix):
                return layout
        raise AnchorNotFoundError("smartart_layout", kind)

    def add_smartart(
        self,
        kind: str,
        nodes: Sequence[NodeInput] | None = None,
        *,
        left: float | None = None,
        top: float | None = None,
        width: float | None = None,
        height: float | None = None,
    ) -> Shape:
        """Add a SmartArt diagram and return its `Shape` (`Shapes.AddSmartArt`).

        `kind` is a friendly layout name (`"process"`, `"cycle"`, `"orgchart"`,
        …; see `constants.SMARTART_CHOICES`). Geometry is in points; omitted values
        default to a ~6.7×4.2 in diagram near the top-left. Address the diagram's
        nodes through the returned shape's `.smartart` (or `anchor_id`); the
        shape's `.has_smartart` is True.

        If `nodes` is given, the diagram's nodes are replaced with it (via
        `SmartArt.set_nodes`) — a list of strings (flat) and/or `{text, children}`
        mappings (nested); otherwise the layout keeps its default placeholder
        nodes. Raises `ValueError` for an unknown `kind` (before any COM) and
        `AnchorNotFoundError` if the layout isn't installed. Wrap in
        `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        left = _DEFAULT_LEFT if left is None else float(left)
        top = _DEFAULT_TOP if top is None else float(top)
        width = _DEFAULT_SMARTART_WIDTH if width is None else float(width)
        height = _DEFAULT_SMARTART_HEIGHT if height is None else float(height)
        with _com.translate_com_errors():
            layout = self._resolve_smartart_layout(kind)
            self._com_collection.AddSmartArt(layout, left, top, width, height)
            shape = self._added()
        if nodes is not None:
            shape.smartart.set_nodes(nodes)
        return shape

    def list(self) -> list[dict[str, Any]]:
        """Every shape as a structured dict, in z-order."""
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, sh in enumerate(self._com_collection, start=1):
                out.append(shape_to_dict(sh, self._slide.index, idx))
        return out

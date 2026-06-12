"""Spike harness — pin the *deferred* fill types against real PowerPoint.

v1.2 shipped solid fill + border; gradient, picture, and pattern fills were
deferred (roadmap "Still open"). This probes the raw COM each upcoming wrapper
will lean on, before any wrapper exists (so it pokes `.com` / raw dispatch):

- **Two-colour gradient** — `Fill.TwoColorGradient(style, variant)` then setting
  `Fill.ForeColor.RGB` / `Fill.BackColor.RGB`; read back `Fill.Type` (==3
  msoFillGradient), `GradientStyle`, `GradientVariant`, `GradientColorType`.
- **One-colour gradient** — `Fill.OneColorGradient(style, variant, degree)`.
- **Preset gradient** — `Fill.PresetGradient(style, variant, presetType)` and
  whether `GradientColorType` reads back as preset (==2?).
- **GradientStops** — can we read each stop (`.Color.RGB`, `.Position`,
  `.Transparency`) and insert/clear via `Insert2(rgb, position, transparency,
  brightness)`? This is the multi-stop control the roadmap flagged as fiddly.
- **Picture fill** — `Fill.UserPicture(absPath)` (confirm absolute-path need),
  read back `Fill.Type` (==6 msoFillPicture?) and `Fill.TextureType`.
- **Pattern fill** — `Fill.Patterned(MsoPatternType)` + fore/back colour, read
  back `Fill.Type` (==2 msoFillPatterned) and `Fill.Pattern`.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/fill_advanced_spike.py

Net-zero and polite exactly like `scripts/style_spike.py`: everything is built on
a single temporary slide appended then deleted in a `finally`, the viewed slide
restored, `net_zero_ok` confirms the deck slide count is unchanged.
"""

from __future__ import annotations

import base64
import json
import os
import tempfile
from typing import Any

import pptlive as pl
from pptlive import _selection

_RED = 0x0000FF  # COM RGB is 0xBBGGRR
_BLUE = 0xFF0000
_GREEN = 0x00FF00

# MsoGradientStyle
_GRAD_HORIZONTAL = 1
_GRAD_DIAGONAL_UP = 3
# MsoPresetGradientType (msoGradientEarlySunset=1, msoGradientBrass=20, ...)
_PRESET_EARLY_SUNSET = 1
# MsoPatternType (msoPattern10Percent=2, msoPatternHorizontalBrick=29, ...)
_PATTERN_10_PERCENT = 2
_PATTERN_LARGE_GRID = 33  # msoPatternLargeGrid (a visible one)

_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _read_fill(fill: Any) -> dict[str, Any]:
    """Best-effort dump of the readable MsoFillFormat surface."""
    out: dict[str, Any] = {}
    for name in (
        "Type",
        "GradientStyle",
        "GradientVariant",
        "GradientColorType",
        "GradientDegree",
        "Pattern",
        "TextureType",
        "TextureName",
        "Visible",
    ):
        try:
            out[name] = (
                int(getattr(fill, name))
                if name not in ("GradientDegree", "TextureName")
                else getattr(fill, name)
            )
        except Exception as exc:
            out[name] = _err(exc)
    return out


def _read_stops(fill: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        stops = fill.GradientStops
        n = int(stops.Count)
        out["count"] = n
        items = []
        for i in range(1, n + 1):
            s = stops(i)
            entry: dict[str, Any] = {}
            for prop, conv in (
                ("Position", float),
                ("Transparency", float),
            ):
                try:
                    entry[prop] = conv(getattr(s, prop))
                except Exception as exc:
                    entry[prop] = _err(exc)
            try:
                entry["Color.RGB"] = int(s.Color.RGB)
            except Exception as exc:
                entry["Color.RGB"] = _err(exc)
            items.append(entry)
        out["stops"] = items
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_two_color_gradient(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=20.0, top=20.0, width=120.0, height=90.0)
        f = sh.com.Fill
        f.TwoColorGradient(_GRAD_HORIZONTAL, 1)
        # Order matters? set colours after declaring the gradient.
        f.ForeColor.RGB = _RED
        f.BackColor.RGB = _BLUE
        out["fill"] = _read_fill(f)
        out["fore_rgb"] = int(f.ForeColor.RGB)
        out["back_rgb"] = int(f.BackColor.RGB)
        out["fore_ok"] = int(f.ForeColor.RGB) == _RED
        out["back_ok"] = int(f.BackColor.RGB) == _BLUE
        out["type_is_gradient"] = out["fill"].get("Type") == 3
        out["stops_after_two_color"] = _read_stops(f)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_one_color_gradient(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=160.0, top=20.0, width=120.0, height=90.0)
        f = sh.com.Fill
        f.ForeColor.RGB = _GREEN
        # OneColorGradient(Style, Variant, Degree) — degree 0..1
        f.OneColorGradient(_GRAD_DIAGONAL_UP, 1, 0.5)
        out["fill"] = _read_fill(f)
        out["fore_rgb"] = int(f.ForeColor.RGB)
        out["type_is_gradient"] = out["fill"].get("Type") == 3
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_preset_gradient(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=300.0, top=20.0, width=120.0, height=90.0)
        f = sh.com.Fill
        f.PresetGradient(_GRAD_HORIZONTAL, 1, _PRESET_EARLY_SUNSET)
        out["fill"] = _read_fill(f)
        out["type_is_gradient"] = out["fill"].get("Type") == 3
        out["stops"] = _read_stops(f)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_gradient_stops(shapes: Any) -> dict[str, Any]:
    """Insert2 + clear: can we build a multi-stop gradient and read it back?"""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=20.0, top=130.0, width=120.0, height=90.0)
        f = sh.com.Fill
        f.TwoColorGradient(_GRAD_HORIZONTAL, 1)
        out["initial_stops"] = _read_stops(f)
        # Insert2(RGB, Position, Transparency, Brightness)
        try:
            f.GradientStops.Insert2(_GREEN, 0.5, 0.0, 0.0)
            out["insert2_ok"] = True
        except Exception as exc:
            out["insert2_error"] = _err(exc)
            # Fallback to legacy Insert(RGB, Position)
            try:
                f.GradientStops.Insert(_GREEN, 0.5)
                out["insert_legacy_ok"] = True
            except Exception as exc2:
                out["insert_legacy_error"] = _err(exc2)
        out["stops_after_insert"] = _read_stops(f)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_picture_fill(shapes: Any, png_path: str) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=160.0, top=130.0, width=120.0, height=90.0)
        f = sh.com.Fill
        # absolute path required?
        f.UserPicture(png_path)
        out["abs_path"] = _read_fill(f)
        out["type"] = out["abs_path"].get("Type")
        # Now try a relative path to confirm the footgun.
        try:
            rel = os.path.basename(png_path)
            sh2 = shapes.add_shape("rectangle", left=300.0, top=130.0, width=120.0, height=90.0)
            sh2.com.Fill.UserPicture(rel)
            out["rel_path_ok"] = True
            out["rel_path_type"] = int(sh2.com.Fill.Type)
        except Exception as exc:
            out["rel_path_error"] = _err(exc)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_pattern_fill(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=20.0, top=240.0, width=120.0, height=90.0)
        f = sh.com.Fill
        f.Patterned(_PATTERN_LARGE_GRID)
        f.ForeColor.RGB = _RED
        f.BackColor.RGB = _BLUE
        out["fill"] = _read_fill(f)
        out["fore_rgb"] = int(f.ForeColor.RGB)
        out["back_rgb"] = int(f.BackColor.RGB)
        out["type_is_patterned"] = out["fill"].get("Type") == 2
        out["pattern_reads_back"] = out["fill"].get("Pattern") == _PATTERN_LARGE_GRID
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    png_path: str | None = None
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        findings["slide_count_before"] = count_before

        snap = _selection.snapshot(ppt)
        findings["viewed_slide"] = snap.slide_index

        fd, png_path = tempfile.mkstemp(suffix=".png", prefix="pptlive_fill_")
        os.write(fd, _PNG_1X1)
        os.close(fd)

        temp_ids: list[int] = []
        try:
            with deck.edit("fill spike: gradient / picture / pattern"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["two_color_gradient"] = probe_two_color_gradient(shapes)
                findings["one_color_gradient"] = probe_one_color_gradient(shapes)
                findings["preset_gradient"] = probe_preset_gradient(shapes)
                findings["gradient_stops"] = probe_gradient_stops(shapes)
                findings["picture_fill"] = probe_picture_fill(shapes, png_path)
                findings["pattern_fill"] = probe_pattern_fill(shapes)
                findings["final_shape_count"] = len(shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("fill spike: cleanup"):
                    for _ in range(len(temp_ids) + 2):
                        victim = None
                        for idx in range(len(deck.slides), 0, -1):
                            try:
                                sid = deck.slides[idx].id
                            except Exception:
                                continue
                            if sid in temp_ids and sid not in deleted:
                                victim = (idx, sid)
                                break
                        if victim is None:
                            break
                        deck.slides[victim[0]].delete()
                        deleted.append(victim[1])
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        count_after = len(deck.slides)
        findings["slide_count_after"] = count_after
        findings["net_zero_ok"] = count_after == count_before

    if png_path and os.path.exists(png_path):
        os.remove(png_path)

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

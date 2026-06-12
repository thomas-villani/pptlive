"""Spike harness — pin fill/line transparency + line dash + arrowheads.

v0.5.0 shipped gradient/picture/pattern fills + shadow/glow/soft-edge/reflection
effects. Still deferred (roadmap "Still open"): partial-alpha transparency and
line dash/arrowheads. This probes the raw COM each upcoming wrapper will lean on:

- **Fill transparency** — `Fill.Transparency` (0..1) round-trips after `Solid()`.
- **Line transparency** — `Line.Transparency` (0..1).
- **Line dash** — `Line.DashStyle` (MsoLineDashStyle 1..); set each, read back so
  we pin the *actual* integer→name mapping this build uses.
- **Arrowheads** — on a connector line: `Begin/EndArrowheadStyle` (MsoArrowheadStyle
  1..6), `Begin/EndArrowheadLength` (1..3), `Begin/EndArrowheadWidth` (1..3). Pin
  whether a rectangle (closed path) also accepts them without error.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/line_alpha_spike.py

Net-zero and polite: one temp slide appended then deleted in a `finally`, the
viewed slide restored, `net_zero_ok` confirms the deck slide count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_RED = 0x0000FF  # COM RGB is 0xBBGGRR
_BLUE = 0xFF0000


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_fill_transparency(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=20.0, top=20.0, width=120.0, height=90.0)
        f = sh.com.Fill
        f.Solid()
        f.ForeColor.RGB = _RED
        out["before"] = float(f.Transparency)
        f.Transparency = 0.5
        out["after"] = float(f.Transparency)
        out["round_trips"] = abs(float(f.Transparency) - 0.5) < 1e-6
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_line_transparency(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=160.0, top=20.0, width=120.0, height=90.0)
        ln = sh.com.Line
        ln.Visible = -1
        ln.ForeColor.RGB = _BLUE
        out["before"] = float(ln.Transparency)
        ln.Transparency = 0.4
        out["after"] = float(ln.Transparency)
        out["round_trips"] = abs(float(ln.Transparency) - 0.4) < 1e-6
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_dash_styles(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=300.0, top=20.0, width=120.0, height=90.0)
        ln = sh.com.Line
        ln.Visible = -1
        ln.Weight = 2.0
        readback: dict[int, Any] = {}
        for val in range(1, 11):
            try:
                ln.DashStyle = val
                readback[val] = int(ln.DashStyle)
            except Exception as exc:
                readback[val] = _err(exc)
        out["set_then_read"] = readback
        out["all_round_trip"] = all(
            isinstance(v, int) and v == k for k, v in readback.items() if k <= 9
        )
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def _add_connector(shapes: Any) -> Any:
    """Add a straight connector via raw COM (AddConnector: type, x1, y1, x2, y2)."""
    # msoConnectorStraight = 1
    return shapes._com_collection.AddConnector(1, 20.0, 140.0, 220.0, 140.0)


def probe_arrowheads(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        conn = _add_connector(shapes)
        ln = conn.Line
        ln.Weight = 2.0
        ln.BeginArrowheadStyle = 2  # triangle
        ln.EndArrowheadStyle = 4  # stealth
        ln.BeginArrowheadLength = 3  # long
        ln.EndArrowheadWidth = 3  # wide
        out["begin_style"] = int(ln.BeginArrowheadStyle)
        out["end_style"] = int(ln.EndArrowheadStyle)
        out["begin_length"] = int(ln.BeginArrowheadLength)
        out["begin_width"] = int(ln.BeginArrowheadWidth)
        out["end_length"] = int(ln.EndArrowheadLength)
        out["end_width"] = int(ln.EndArrowheadWidth)
        out["round_trips"] = (
            int(ln.BeginArrowheadStyle) == 2
            and int(ln.EndArrowheadStyle) == 4
            and int(ln.BeginArrowheadLength) == 3
            and int(ln.EndArrowheadWidth) == 3
        )
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_arrowheads_on_rect(shapes: Any) -> dict[str, Any]:
    """Does a closed rectangle accept arrowhead props without error?"""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=300.0, top=140.0, width=120.0, height=90.0)
        ln = sh.com.Line
        ln.EndArrowheadStyle = 2
        out["accepted"] = int(ln.EndArrowheadStyle)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        findings["slide_count_before"] = count_before

        snap = _selection.snapshot(ppt)
        findings["viewed_slide"] = snap.slide_index

        temp_ids: list[int] = []
        try:
            with deck.edit("line/alpha spike"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["fill_transparency"] = probe_fill_transparency(shapes)
                findings["line_transparency"] = probe_line_transparency(shapes)
                findings["dash_styles"] = probe_dash_styles(shapes)
                findings["arrowheads"] = probe_arrowheads(shapes)
                findings["arrowheads_on_rect"] = probe_arrowheads_on_rect(shapes)
        finally:
            try:
                with deck.edit("line/alpha spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[idx].id
                        except Exception:
                            continue
                        if sid in temp_ids:
                            deck.slides[idx].delete()
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            _selection.restore(ppt, snap)

        count_after = len(deck.slides)
        findings["slide_count_after"] = count_after
        findings["net_zero_ok"] = count_after == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Spike — verify the v0.7b chart wrappers against real PowerPoint.

The *exploratory* pass (git history) drove raw COM to design the surface; its
findings are recorded in `_charts.py`'s module docstring and IMPLEMENTATION.md.
The two that shaped the code:

1. `Chart.SetSourceData` takes a **string** range (`"Sheet1!$A$1:$C$4"`), not a
   `Range` object (the Range form raised `E_FAIL`).
2. `SetSourceData` **dissolves the default Excel Table**, so the wrapper relies on
   `UsedRange.ClearContents` + `SetSourceData` (no ListObject) — which makes
   re-writes (shrink *and* grow) clean.

This is the regression check that the *shipped* wrappers behave on a live deck. It
exercises `ShapeCollection.add_chart` (with and without data), the `has_chart`
gate, and `Chart.read` / `set_type` / `set_data` (including a shrink then a grow).
Run against a *running* PowerPoint with a deck open:

    uv run python scripts/chart_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a single
temporary slide appended then deleted in a `finally`, the embedded workbook is
closed by `set_data`, and the viewed slide + Selection are restored.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection
from pptlive.exceptions import AnchorNotFoundError


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe(deck: pl.Presentation, slide: pl.Slide) -> dict[str, Any]:
    out: dict[str, Any] = {}

    # add_chart with data + the gate ------------------------------------------
    chart_shape = slide.shapes.add_chart(
        "column",
        ["Q1", "Q2", "Q3"],
        {"Revenue": [10, 20, 30], "Profit": [3, 6, 9]},
    )
    out["add_chart"] = {
        "anchor_id": chart_shape.anchor_id,
        "type": chart_shape.shape_type,
        "has_chart": chart_shape.has_chart,
    }
    out["read_after_add"] = chart_shape.chart.read()

    chart = chart_shape.chart

    # set_type round-trip -----------------------------------------------------
    out["set_type"] = {}
    for t in ("line", "pie", "bar"):
        chart.set_type(t)
        out["set_type"][t] = chart.chart_type

    # set_data: shrink then grow ----------------------------------------------
    chart.set_data(["A", "B"], {"Only": [7, 8]})
    out["after_shrink"] = chart.read()
    chart.set_data(
        ["Jan", "Feb", "Mar", "Apr"],
        {"North": [1, 2, 3, 4], "South": [5, 6, 7, 8], "East": [9, 10, 11, 12]},
    )
    out["after_grow"] = chart.read()

    # add_chart with no data -> PowerPoint's default placeholder data ----------
    default_chart = slide.shapes.add_chart("doughnut")
    out["default_data"] = default_chart.chart.read()

    # the gate on a non-chart shape -------------------------------------------
    box = slide.shapes.add_textbox("not a chart")
    try:
        _ = box.chart
        out["non_chart_gate"] = "NO ERROR (unexpected)"
    except AnchorNotFoundError as exc:
        out["non_chart_gate"] = _err(exc)

    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        temp_ids: list[int] = []
        try:
            with deck.edit("chart spike: build"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["probe"] = probe(deck, deck.slides[sidx])
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("chart spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[idx].id
                        except Exception:
                            continue
                        if sid in temp_ids and sid not in deleted:
                            deck.slides[idx].delete()
                            deleted.append(sid)
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            try:
                _selection.restore(ppt, snap)
            except Exception as exc:
                findings["restore_error"] = _err(exc)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Spike harness — PPTLIVE-009: can we recolor SmartArt / chart *text* over COM?

SmartArt diagrams and charts report `has_text_frame: false` and expose no text
anchor, so pptlive has no path to their internal text color. The finding asks for
"at minimum a way to set SmartArt node text color and chart text-element colors
(axis, legend, data labels, title)", or even a coarse "recolor all text in this
shape to X". This spike pokes the raw COM the eventual wrapper would lean on,
*before* any of it exists:

SmartArt (each node carries text on `TextFrame2`, like the node-text recipe):
- **Per-node font color** — `node.TextFrame2.TextRange.Font.Fill.ForeColor.RGB`
  (TextFrame2 colors live on `Font.Fill.ForeColor`, NOT `Font.Color`). Set + read.
- **Recolor-all** — iterate `SmartArt.AllNodes`, set each node's font fill; count
  how many took, and whether they read back the literal RGB.
- **Host-shape TextFrame2** — does `shape.TextFrame2.TextRange.Font...` exist / do
  anything on a SmartArt host (a coarse one-call recolor), or raise?

Chart (a dual object model — classic `.Font.Color.RGB` on elements, plus the
modern `.Format.TextFrame2`):
- **ChartArea global** — `chart.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB`
  as a single "default all chart text" call; does it read back, does it cascade?
- **Legend / axes / title / data labels** — the classic per-element
  `.Font.Color.RGB` (legend, category + value axis tick labels, chart title,
  series data labels). Set + read each.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/composite_text_spike.py

Net-zero and polite exactly like `scripts/style_spike.py`: everything is built on
a single temporary slide appended then deleted in a `finally`, the viewed slide
is restored, and `net_zero_ok` confirms the deck's slide count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_RED = 0x0000FF  # COM RGB is 0xBBGGRR, so this is pure red
_GREEN = 0x00FF00
_BLUE = 0xFF0000
_THEME_SENTINEL = 0x80000000  # what a non-literal (theme/auto) color reads back as

_XL_CATEGORY = 1  # XlAxisType.xlCategory
_XL_VALUE = 2  # XlAxisType.xlValue


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _try(fn: Any) -> Any:
    """Run a thunk, returning its value or an "error: ..." string."""
    try:
        return fn()
    except Exception as exc:
        return f"error: {_err(exc)}"


def probe_smartart(shapes: Any) -> dict[str, Any]:
    """Per-node + all-nodes + host-shape recolor of SmartArt text."""
    out: dict[str, Any] = {}
    try:
        shape = shapes.add_smartart(
            "process", ["Discover", "Design", "Build", "Ship"], left=72.0, top=72.0
        )
        sa = shape.com.SmartArt

        # 1. Per-node font color via TextFrame2.Font.Fill.ForeColor.RGB
        node1 = sa.Nodes.Item(1)
        tr1 = node1.TextFrame2.TextRange
        out["node1_default_rgb"] = _try(lambda: int(tr1.Font.Fill.ForeColor.RGB))

        def set_node1() -> int:
            tr1.Font.Fill.ForeColor.RGB = _RED
            return int(tr1.Font.Fill.ForeColor.RGB)

        out["node1_set_rgb"] = _try(set_node1)
        out["node1_set_ok"] = out["node1_set_rgb"] == _RED

        # 2. Recolor-all via AllNodes
        def recolor_all() -> dict[str, Any]:
            allnodes = sa.AllNodes
            total = int(allnodes.Count)
            applied = 0
            readback_ok = 0
            for i in range(1, total + 1):
                nd = allnodes.Item(i)
                font = nd.TextFrame2.TextRange.Font
                font.Fill.ForeColor.RGB = _GREEN
                applied += 1
                if int(font.Fill.ForeColor.RGB) == _GREEN:
                    readback_ok += 1
            return {"total": total, "applied": applied, "readback_ok": readback_ok}

        out["all_nodes"] = _try(recolor_all)

        # 3. Host-shape TextFrame2 — a single coarse recolor on the host?
        def host_tf2() -> Any:
            com = shape.com
            has = bool(com.HasTextFrame)
            com.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _BLUE
            return {
                "has_text_frame": has,
                "set_rgb": int(com.TextFrame2.TextRange.Font.Fill.ForeColor.RGB),
            }

        out["host_textframe2"] = _try(host_tf2)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_chart(shapes: Any) -> dict[str, Any]:
    """ChartArea global + classic per-element font recolor of chart text."""
    out: dict[str, Any] = {}
    try:
        shape = shapes.add_chart(
            "column",
            ["Q1", "Q2", "Q3"],
            {"Revenue": [10, 20, 30], "Profit": [3, 6, 9]},
            left=72.0,
            top=320.0,
        )
        chart = shape.com.Chart

        # 1. ChartArea global text default via modern Format.TextFrame2
        def chartarea_global() -> Any:
            f = chart.ChartArea.Format.TextFrame2.TextRange.Font
            f.Fill.ForeColor.RGB = _RED
            return int(f.Fill.ForeColor.RGB)

        out["chartarea_global_rgb"] = _try(chartarea_global)

        # The classic chart Font model: `Font.Color` is itself an RGB long
        # property (set/read directly), NOT a ColorFormat with `.RGB`.
        # 2. Legend
        def legend() -> Any:
            chart.HasLegend = True
            fobj = chart.Legend.Font
            fobj.Color = _GREEN
            return {"has_legend": True, "rgb": int(fobj.Color)}

        out["legend"] = _try(legend)

        # 3. Category + value axis tick labels
        def axis(axis_type: int) -> Any:
            fobj = chart.Axes(axis_type).TickLabels.Font
            fobj.Color = _BLUE
            return int(fobj.Color)

        out["category_axis_rgb"] = _try(lambda: axis(_XL_CATEGORY))
        out["value_axis_rgb"] = _try(lambda: axis(_XL_VALUE))

        # 4. Chart title
        def title() -> Any:
            chart.HasTitle = True
            chart.ChartTitle.Text = "Revenue"
            fobj = chart.ChartTitle.Font
            fobj.Color = _RED
            return {"rgb": int(fobj.Color)}

        out["title"] = _try(title)

        # 5. Data labels on every series (DataLabels is a METHOD -> call it)
        def data_labels() -> Any:
            sc = chart.SeriesCollection()
            n = int(sc.Count)
            ok = 0
            for i in range(1, n + 1):
                s = sc(i)
                s.HasDataLabels = True
                fobj = s.DataLabels().Font
                fobj.Color = _GREEN
                if int(fobj.Color) == _GREEN:
                    ok += 1
            return {"series": n, "readback_ok": ok}

        out["data_labels"] = _try(data_labels)

        # 6. Modern read on the legend after ChartArea global set (cascade check)
        def legend_modern_read() -> Any:
            return int(chart.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB)

        out["legend_modern_readback"] = _try(legend_modern_read)
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
            with deck.edit("composite-text spike: smartart / chart text color"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["smartart"] = probe_smartart(shapes)
                findings["chart"] = probe_chart(shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("composite-text spike: cleanup"):
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

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

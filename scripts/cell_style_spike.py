"""Spike — is table-cell styling COM-reachable? (the open v-next gap)

CLAUDE.md records "PowerPoint's COM exposes no cell-level fill/shading/border
automation" as the reason table cells can only have their *text* styled. This
spike tests that claim directly against a live deck, because the table model
(`_tables.py`) routes a cell through `Table.Cell(r, c)`, whose COM object exposes
`.Shape` (a real `Shape` with `.Fill`/`.Line`/`.TextFrame`) and `.Borders` — both
of which *look* writable. If they round-trip, cell fill/border becomes a thin
reuse of the existing `apply_shape_fill` helper rather than a delete-and-recreate.

Probes, all on `Table.Cell(r, c)`:
  1. fill   — `cell.Shape.Fill.ForeColor.RGB` (via `apply_shape_fill`), read back
  2. fill="none" — `Fill.Visible = msoFalse`, read `Fill.Visible`
  3. border — `cell.Borders(idx)` color + weight per edge, read back
  4. text   — confirm cell text styling still co-exists (font color on the same cell)

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/cell_style_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a
single temporary slide appended then deleted in a `finally`; the viewed slide +
Selection are restored. `net_zero_ok` confirms the deck's slide count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection
from pptlive._shapes import apply_shape_fill

# msoBorder edge indices (PpBorderType / MsoBorderType): top/left/bottom/right.
_BORDER_EDGES = {"top": 1, "left": 2, "bottom": 3, "right": 4}
_RED = 0x0000FF  # COM RGB long is BGR-ordered; 0x0000FF == pure red
_BLUE = 0xFF0000


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe(deck: pl.Presentation, slide: pl.Slide) -> dict[str, Any]:
    out: dict[str, Any] = {}

    try:
        sh = slide.shapes.add_table(3, 3)
        table = sh.table
        out["table"] = {"anchor_id": sh.anchor_id, "has_table": sh.has_table}
    except Exception as exc:
        out["table"] = {"error": _err(exc)}
        return out

    # 1. cell fill via the existing helper -----------------------------------
    try:
        com_cell = table.cell(1, 1)._cell_com()
        with deck.edit("cell spike: fill"):
            apply_shape_fill(com_cell.Shape, fill=_RED)
        out["fill_solid"] = {
            "wrote": _RED,
            "read_back": int(com_cell.Shape.Fill.ForeColor.RGB),
            "visible": int(com_cell.Shape.Fill.Visible),
            "ok": int(com_cell.Shape.Fill.ForeColor.RGB) == _RED,
        }
    except Exception as exc:
        out["fill_solid"] = {"error": _err(exc)}

    # 2. fill = none (transparent) -------------------------------------------
    try:
        com_cell = table.cell(1, 2)._cell_com()
        with deck.edit("cell spike: fill none"):
            apply_shape_fill(com_cell.Shape, fill="none")
        out["fill_none"] = {
            "visible_after": int(com_cell.Shape.Fill.Visible),  # expect msoFalse (0)
            "ok": int(com_cell.Shape.Fill.Visible) == 0,
        }
    except Exception as exc:
        out["fill_none"] = {"error": _err(exc)}

    # 3. cell borders --------------------------------------------------------
    try:
        com_cell = table.cell(2, 2)._cell_com()
        edge_results: dict[str, Any] = {}
        with deck.edit("cell spike: borders"):
            for _name, idx in _BORDER_EDGES.items():
                border = com_cell.Borders(idx)
                border.Visible = -1  # msoTrue
                border.ForeColor.RGB = _BLUE
                border.Weight = 2.5
        for name, idx in _BORDER_EDGES.items():
            border = com_cell.Borders(idx)
            edge_results[name] = {
                "color": int(border.ForeColor.RGB),
                "weight": float(border.Weight),
                "visible": int(border.Visible),
            }
        out["borders"] = {
            "edges": edge_results,
            "ok": all(e["color"] == _BLUE for e in edge_results.values()),
        }
    except Exception as exc:
        out["borders"] = {"error": _err(exc)}

    # 4. cell fill + text color co-exist -------------------------------------
    try:
        cell = table.cell(3, 3)
        com_cell = cell._cell_com()
        with deck.edit("cell spike: fill+text"):
            apply_shape_fill(com_cell.Shape, fill=_RED)
            cell.set_text("hi")
            cell.format_text(color="#FFFFFF", bold=True)
        out["fill_plus_text"] = {
            "fill": int(com_cell.Shape.Fill.ForeColor.RGB),
            "font_color": int(com_cell.Shape.TextFrame.TextRange.Font.Color.RGB),
            "text": cell.text,
        }
    except Exception as exc:
        out["fill_plus_text"] = {"error": _err(exc)}

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
            with deck.edit("cell spike: build"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["probe"] = probe(deck, deck.slides[sidx])
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("cell spike: cleanup"):
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
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

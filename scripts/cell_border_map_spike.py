"""Spike — map the table-cell `Borders(index)` edge ordering, visually.

`cell_style_spike.py` proved cell borders round-trip, but it set all four edges
the same color, so it could not tell *which* index is which physical edge. This
spike paints each index a distinct, thick color on one large cell, exports the
slide to a PNG, and leaves the file for a human/vision check. The conventional
Office order is top/left/bottom/right (1..4) + the two diagonals (5/6); this
confirms it on the live build before we bake friendly edge names.

    uv run python scripts/cell_border_map_spike.py

Writes the PNG path into the findings JSON. Net-zero on the *deck* (temp slide
appended then deleted, view/Selection restored); the PNG is left on disk under
the system temp dir for inspection and printed in the output.
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path
from typing import Any

import pptlive as pl
from pptlive import _selection

# Distinct, saturated COM RGB longs (BGR order) — easy to tell apart in the PNG.
_COLORS = {
    1: ("red", 0x0000FF),
    2: ("green", 0x00FF00),
    3: ("blue", 0xFF0000),
    4: ("yellow", 0x00FFFF),
    5: ("magenta", 0xFF00FF),
    6: ("cyan", 0xFFFF00),
}


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def main() -> int:
    findings: dict[str, Any] = {}
    out_png = Path(tempfile.gettempdir()) / "pptlive_border_map.png"
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)
        temp_ids: list[int] = []
        try:
            with deck.edit("border map: build"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                sidx = temp.index
            slide = deck.slides[sidx]
            sh = slide.shapes.add_table(2, 2, left=80, top=80, width=560, height=360)
            cell = sh.table.cell(1, 1)._cell_com()
            applied: dict[int, str] = {}
            with deck.edit("border map: paint"):
                for idx, (name, rgb) in _COLORS.items():
                    try:
                        border = cell.Borders(idx)
                        border.Visible = -1  # msoTrue
                        border.ForeColor.RGB = rgb
                        border.Weight = 6.0
                        applied[idx] = name
                    except Exception as exc:  # noqa: BLE001 — record which indices are valid
                        findings.setdefault("invalid_index", {})[idx] = _err(exc)
            findings["applied_colors"] = applied
            slide.export_image(str(out_png))
            findings["png"] = str(out_png)
        except Exception as exc:  # noqa: BLE001
            findings["error"] = _err(exc)
        finally:
            try:
                with deck.edit("border map: cleanup"):
                    for i in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[i].id
                        except Exception:
                            continue
                        if sid in temp_ids:
                            deck.slides[i].delete()
            except Exception as exc:  # noqa: BLE001
                findings["cleanup_error"] = _err(exc)
            _selection.restore(ppt, snap)
        findings["net_zero_ok"] = len(deck.slides) == count_before
    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

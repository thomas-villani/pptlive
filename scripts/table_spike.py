"""Spike — verify the v0.5 table wrappers against real PowerPoint.

The exploratory pass (git history) probed raw COM to design the surface — the
headline finding being that `Shapes.AddTable` can return a shape whose `Type`
reports *placeholder* (14), not table (19), so `HasTable` is the only reliable
gate. This is the regression check that the shipped wrappers behave on a live
deck. It exercises `ShapeCollection.add_table`, `Shape.has_table` / `Shape.table`,
`Table.read` / `cell` text round-trip, `add_row` / `delete_row`, and the
`cell:S:N:R:C` anchor through `Presentation.anchor_by_id`. Run against a
*running* PowerPoint with a deck open:

    uv run python scripts/table_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a
single temporary slide that is appended then deleted in a `finally`, and the
viewed slide + Selection are restored. `net_zero_ok` confirms the deck's slide
count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import Cell, _selection
from pptlive.exceptions import AnchorNotFoundError


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe(deck: pl.Presentation, slide: pl.Slide) -> dict[str, Any]:
    out: dict[str, Any] = {}

    # add_table + the gate ----------------------------------------------------
    try:
        sh = slide.shapes.add_table(3, 2)
        out["add_table"] = {
            "anchor_id": sh.anchor_id,
            "has_table": sh.has_table,
            "shape_type": sh.shape_type,  # may be "placeholder", not "table"
        }
    except Exception as exc:
        out["add_table"] = {"error": _err(exc)}
        return out

    # dimensions + cell text round-trip + read() ------------------------------
    try:
        table = sh.table
        with deck.edit("table spike: fill"):
            table.cell(1, 1).set_text("R1C1")
            table.cell(2, 2).set_text("two\rlines")
        grid = table.read()
        out["read"] = {
            "rows": grid["rows"],
            "columns": grid["columns"],
            "r1c1": table.cell(1, 1).text,
            "r2c2_anchor": grid["cells"][1][1]["anchor_id"],
            "r2c2_text": table.cell(2, 2).text,
        }
    except Exception as exc:
        out["read"] = {"error": _err(exc)}
        return out

    # add_row (with values) + delete_row -------------------------------------
    try:
        before = table.row_count
        with deck.edit("table spike: add row"):
            table.add_row(["added-a", "added-b"])
        after_add = table.row_count
        added = [table.cell(after_add, 1).text, table.cell(after_add, 2).text]
        with deck.edit("table spike: delete row"):
            table.delete_row(after_add)
        out["row_ops"] = {
            "before": before,
            "after_add": after_add,
            "added_row_text": added,
            "after_delete": table.row_count,
            "ok": after_add == before + 1 and added == ["added-a", "added-b"],
        }
    except Exception as exc:
        out["row_ops"] = {"error": _err(exc)}

    # cell:S:N:R:C anchor through anchor_by_id + a not-found gate ------------
    try:
        cell = deck.anchor_by_id(f"cell:{slide.index}:{sh.index}:1:1")
        out["anchor"] = {
            "is_cell": isinstance(cell, Cell),
            "anchor_id": cell.anchor_id,
            "text": cell.text,
        }
        try:
            deck.anchor_by_id(f"cell:{slide.index}:{sh.index}:9:9")
            out["anchor"]["out_of_range_raised"] = False
        except AnchorNotFoundError:
            out["anchor"]["out_of_range_raised"] = True
    except Exception as exc:
        out["anchor"] = {"error": _err(exc)}

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
            with deck.edit("table spike: build"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["probe"] = probe(deck, deck.slides[sidx])
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("table spike: cleanup"):
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

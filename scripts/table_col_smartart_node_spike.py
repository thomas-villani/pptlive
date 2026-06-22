"""Spike — verify table column add/delete + SmartArt per-node Font2 formatting.

Two open edit-surface items, probed on one temporary slide against a *running*
PowerPoint:

1. **Table columns.** `Table.Columns.Add([BeforeColumn])` appends a column when
   called with no arg, and inserts before a 1-based column when given one;
   `Columns(n).Delete()` removes one. Confirms the column count moves and that a
   freshly added column's cells are addressable + writable.
2. **SmartArt per-node format.** A node's text lives on `TextFrame2` (a `Font2`,
   not the classic `Font`): probe whether `Font.Bold` / `.Italic` / `.Size` /
   `.Name` / `.Fill.ForeColor.RGB` / `.UnderlineStyle` round-trip on a single
   node, and whether `AllNodes` enumerates in the same order as a depth-first
   walk of the nested `Nodes` tree (so a flat 1-based `AllNodes` index is a stable
   per-node address).

    uv run python scripts/table_col_smartart_node_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a
single temporary slide appended then deleted in a `finally`, with the viewed
slide + Selection restored. `net_zero_ok` confirms the slide count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_table_columns(deck: pl.Presentation, slide: pl.Slide) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = slide.shapes.add_table(2, 2)
        com = sh.com.Table
        before = int(com.Columns.Count)

        # append (no arg)
        com.Columns.Add()
        after_append = int(com.Columns.Count)
        # write into the new last column to prove it is addressable
        com.Cell(1, after_append).Shape.TextFrame.TextRange.Text = "appended"
        appended_text = str(com.Cell(1, after_append).Shape.TextFrame.TextRange.Text or "")

        # insert before column 1
        com.Columns.Add(1)
        after_insert = int(com.Columns.Count)
        com.Cell(1, 1).Shape.TextFrame.TextRange.Text = "inserted-first"
        inserted_text = str(com.Cell(1, 1).Shape.TextFrame.TextRange.Text or "")

        # delete a column
        com.Columns(after_insert).Delete()
        after_delete = int(com.Columns.Count)

        out = {
            "before": before,
            "after_append": after_append,
            "appended_text": appended_text,
            "after_insert": after_insert,
            "inserted_text": inserted_text,
            "after_delete": after_delete,
            "append_ok": after_append == before + 1 and appended_text == "appended",
            "insert_ok": after_insert == before + 2 and inserted_text == "inserted-first",
            "delete_ok": after_delete == after_insert - 1,
        }
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_smartart_node_format(deck: pl.Presentation, slide: pl.Slide) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = slide.shapes.add_smartart("hierarchy", left=40, top=120, width=400, height=200)
        sa = sh.smartart
        with deck.edit("spike: seed smartart"):
            sa.set_nodes(
                [
                    {"text": "A", "children": ["A1", "A2"]},
                ]
            )
    except Exception as exc:
        out["seed_error"] = _err(exc)
        return out

    # AllNodes ordering vs depth-first nested walk -----------------------------
    try:
        com = sa.com
        allnodes = com.AllNodes
        flat = [
            str(allnodes.Item(i).TextFrame2.TextRange.Text or "")
            for i in range(1, int(allnodes.Count) + 1)
        ]

        def _walk(nodes: Any) -> list[str]:
            acc: list[str] = []
            for i in range(1, int(nodes.Count) + 1):
                nd = nodes.Item(i)
                acc.append(str(nd.TextFrame2.TextRange.Text or ""))
                if int(nd.Nodes.Count):
                    acc.extend(_walk(nd.Nodes))
            return acc

        depth_first = _walk(com.Nodes)
        out["ordering"] = {
            "allnodes": flat,
            "depth_first": depth_first,
            "same": flat == depth_first,
            "count": int(allnodes.Count),
        }
    except Exception as exc:
        out["ordering"] = {"error": _err(exc)}

    # Font2 property round-trip on a single node -------------------------------
    try:
        com = sa.com
        target = com.AllNodes.Item(1)  # node "A"
        f = target.TextFrame2.TextRange.Font
        with deck.edit("spike: format node"):
            f.Bold = -1  # msoTrue
            f.Italic = -1
            f.Size = 28.0
            f.Name = "Georgia"
            f.Fill.ForeColor.RGB = 0x0000FF  # red in BGR-stored RGB long -> reads back
            try:
                f.UnderlineStyle = 2  # msoUnderlineSingleLine
                underline_set = True
            except Exception as exc:
                underline_set = f"err: {_err(exc)}"
        # read back
        f2 = com.AllNodes.Item(1).TextFrame2.TextRange.Font
        out["font_roundtrip"] = {
            "bold": int(f2.Bold),
            "italic": int(f2.Italic),
            "size": float(f2.Size),
            "name": str(f2.Name),
            "color_rgb": int(f2.Fill.ForeColor.RGB),
            "underline_style": _safe(lambda: int(f2.UnderlineStyle)),
            "underline_set": underline_set,
        }
    except Exception as exc:
        out["font_roundtrip"] = {"error": _err(exc)}

    return out


def _safe(fn: Any) -> Any:
    try:
        return fn()
    except Exception as exc:
        return _err(exc)


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        temp_ids: list[int] = []
        try:
            with deck.edit("spike: build"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["table_columns"] = probe_table_columns(deck, deck.slides[sidx])
            findings["smartart_node"] = probe_smartart_node_format(deck, deck.slides[sidx])
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("spike: cleanup"):
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

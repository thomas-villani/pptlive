"""Spike — design the SmartArt surface against real PowerPoint (exploratory pass).

SmartArt was never specced (it appears nowhere in `spec.md`/`IMPLEMENTATION.md`);
this is the exploratory raw-COM probe that nails down the surface *before* a
`_smartart.py` wrapper exists — the analog of the chart spike's first pass. It
drives `Shapes.AddSmartArt` / `SmartArt.Nodes` / `SmartArtNode.TextFrame2`
directly via `.com` and records the findings the wrapper design needs:

1. **Layout identity.** `Application.SmartArtLayouts` index drifts, but each
   layout's `.Id` is a stable URN (`…/layout/process1`). So a friendly-name ->
   URN table is the durable mapping (parallels `chart_type_for`). Probes which of
   the 7 core layouts resolve, and their default node count (varies per layout —
   the wrapper must size nodes to the caller's item list).
2. **The gate.** `Shape.HasSmartArt` is the reliable gate (Type reports
   msoSmartArt=24, but follow the table/chart lesson and gate on Has*).
3. **Populate (flat).** `Nodes.Add()` / `Item(i).Delete()` size the list;
   `Item(i).TextFrame2.TextRange.Text` sets text (note **TextFrame2**).
4. **Populate (tree) — two gotchas.** (a) Tree layouts ship a *pre-built
   skeleton* (`hierarchy1`/`orgChart1` default to `Nodes.Count==1` but
   `AllNodes==6/5` — the root already has empty placeholder children), so the
   wrapper must **clear to one empty root first**. (b) `node.Nodes.Add()` adds a
   *sibling*, not a child — true nesting needs
   `node.AddNode(msoSmartArtNodeBelow=5, type)`; `type=msoSmartArtNodeAssistant=4`
   gives an org-chart assistant.
4. **Read back.** Recurse `.Nodes` capturing text + level + children, or flat via
   `.AllNodes`; confirm round-trip.

Run against a *running* PowerPoint with any deck open (a blank deck is fine):

    uv run python scripts/smartart_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on
temporary slides appended then deleted in a `finally`, and the viewed slide +
Selection are restored.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

# The 7 core layouts, by the stable trailing URN segment of SmartArtLayout.Id.
CORE_URNS = {
    "list": "list1",
    "process": "process1",
    "cycle": "cycle1",
    "hierarchy": "hierarchy1",
    "orgchart": "orgChart1",
    "pyramid": "pyramid1",
    "venn": "venn1",
}

PP_LAYOUT_BLANK = 12
MSO_SMARTART = 24  # msoSmartArt (Shape.Type) — informational; gate on HasSmartArt

# MsoSmartArtNodePosition / MsoSmartArtNodeType (the ones the wrapper needs)
NODE_BELOW = 5  # msoSmartArtNodeBelow — add as a *child* (Nodes.Add adds a sibling)
NODE_DEFAULT = 1  # msoSmartArtNodeTypeDefault
NODE_ASSISTANT = 4  # msoSmartArtNodeTypeAssistant — org-chart assistant


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _resolve_layouts(app: Any) -> dict[str, Any]:
    """Map each friendly name -> the live SmartArtLayout COM object (by URN Id)."""
    layouts = app.SmartArtLayouts
    by_urn: dict[str, Any] = {}
    by_meta: dict[str, Any] = {}
    catalog: dict[str, Any] = {}
    for i in range(1, layouts.Count + 1):
        L = layouts.Item(i)
        catalog[L.Id] = (i, L.Name)
    for friendly, seg in CORE_URNS.items():
        match = next((urn for urn in catalog if urn.endswith("/" + seg)), None)
        if match is None:
            by_meta[friendly] = {"resolved": False}
            continue
        idx, name = catalog[match]
        by_urn[friendly] = layouts.Item(idx)
        by_meta[friendly] = {"resolved": True, "index": idx, "name": name, "id": match}
    return {"objects": by_urn, "meta": by_meta, "installed_count": layouts.Count}


def _dump_nodes(nodes: Any) -> list[dict[str, Any]]:
    """Recurse a SmartArtNodes collection -> nested {text, level, children}."""
    out: list[dict[str, Any]] = []
    for i in range(1, nodes.Count + 1):
        nd = nodes.Item(i)
        out.append(
            {
                "text": nd.TextFrame2.TextRange.Text,
                "level": nd.Level,
                "children": _dump_nodes(nd.Nodes) if nd.Nodes.Count else [],
            }
        )
    return out


def _set_flat(sa: Any, items: list[str]) -> None:
    """Size the top-level node list to len(items) and set each node's text."""
    nodes = sa.Nodes
    while nodes.Count < len(items):
        nodes.Add()
    while nodes.Count > len(items):
        nodes.Item(nodes.Count).Delete()
    for i, txt in enumerate(items, 1):
        nodes.Item(i).TextFrame2.TextRange.Text = txt


def probe(app: Any, slide: Any, layouts: dict[str, Any]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    out["layouts"] = layouts["meta"]
    out["installed_count"] = layouts["installed_count"]
    objs = layouts["objects"]

    # --- per-layout default node count (the wrapper must size to the item list) -
    defaults: dict[str, Any] = {}
    for friendly, layout in objs.items():
        try:
            sh = slide.Shapes.AddSmartArt(layout, 40, 40, 400, 240)
            defaults[friendly] = {
                "has_smartart": bool(sh.HasSmartArt),
                "type": sh.Type,
                "type_is_smartart": sh.Type == MSO_SMARTART,
                "default_nodes": sh.SmartArt.Nodes.Count,
                "default_allnodes": sh.SmartArt.AllNodes.Count,
            }
            sh.Delete()
        except Exception as exc:
            defaults[friendly] = {"error": _err(exc)}
    out["per_layout"] = defaults

    # --- flat populate + read-back (process) ---------------------------------
    if "process" in objs:
        sh = slide.Shapes.AddSmartArt(objs["process"], 40, 40, 600, 300)
        items = ["Discover", "Design", "Build", "Ship"]
        _set_flat(sh.SmartArt, items)
        read = _dump_nodes(sh.SmartArt.Nodes)
        out["flat"] = {
            "wrote": items,
            "read_nested": read,
            "read_allnodes": [
                sh.SmartArt.AllNodes.Item(i).TextFrame2.TextRange.Text
                for i in range(1, sh.SmartArt.AllNodes.Count + 1)
            ],
            "round_trip_ok": [n["text"] for n in read] == items,
        }
        sh.Delete()

    # --- nested tree populate + read-back (orgChart) -------------------------
    # The resolved recipe: clear the pre-built skeleton to one empty root, then
    # build with AddNode(BELOW, type) — Nodes.Add() would add siblings.
    tree_layout = objs.get("orgchart") or objs.get("hierarchy")
    if tree_layout is not None:
        sh = slide.Shapes.AddSmartArt(tree_layout, 40, 40, 600, 300)
        sa = sh.SmartArt
        while sa.Nodes.Count > 1:
            sa.Nodes.Item(sa.Nodes.Count).Delete()
        root = sa.Nodes.Item(1)
        while root.Nodes.Count:  # strip the skeleton's placeholder descendants
            root.Nodes.Item(root.Nodes.Count).Delete()
        root.TextFrame2.TextRange.Text = "CEO"
        ve = root.AddNode(NODE_BELOW, NODE_DEFAULT)
        ve.TextFrame2.TextRange.Text = "VP Eng"
        vs = root.AddNode(NODE_BELOW, NODE_DEFAULT)
        vs.TextFrame2.TextRange.Text = "VP Sales"
        lead = ve.AddNode(NODE_BELOW, NODE_DEFAULT)
        lead.TextFrame2.TextRange.Text = "Eng Lead"
        asst = root.AddNode(NODE_BELOW, NODE_ASSISTANT)
        asst.TextFrame2.TextRange.Text = "Chief of Staff"
        out["tree"] = {
            "read": _dump_nodes(sa.Nodes),
            "nests_correctly": _dump_nodes(sa.Nodes)[0]["children"][0]["children"] != [],
        }
        sh.Delete()

    # --- the gate on a non-smartart shape ------------------------------------
    box = slide.Shapes.AddTextbox(1, 10, 10, 100, 30)
    box.TextFrame.TextRange.Text = "not smartart"
    out["non_smartart_gate"] = {"has_smartart": bool(box.HasSmartArt), "type": box.Type}
    box.Delete()

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
            layouts = _resolve_layouts(ppt.com)
            with deck.edit("smartart spike: build"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["probe"] = probe(ppt.com, deck.slides[sidx].com, layouts)
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("smartart spike: cleanup"):
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

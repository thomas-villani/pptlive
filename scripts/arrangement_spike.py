"""Spike harness — pin the shape-arrangement COM paths against real PowerPoint.

Roadmap "Shape arrangement beyond z-order" is greenfield (no group/ungroup/align/
distribute/connector code exists). This pokes the raw COM the upcoming
`ShapeCollection.group/align/distribute/add_connector` + `Shape.ungroup` wrappers
will lean on, before they exist (so it drives `.com` / raw dispatch directly). The
real unknowns:

- **Group identity** — `Shapes.Range([names]).Group()` returns the group Shape.
  Does the group get a NEW `Shape.Id`? What `Type` (msoGroup=6)? Are the children
  reachable via `group.GroupItems` and do they keep their original Ids?
- **Ungroup identity churn** — `group.Ungroup()` returns a ShapeRange of the freed
  children. Do they get NEW Ids or keep their originals? (Decides whether
  `ungroup()` returns handles by id or must re-read the slide.)
- **Align / Distribute** — `Range.Align(cmd, RelativeTo)` /
  `Range.Distribute(cmd, RelativeTo)`; confirm the cmd ints and that
  RelativeTo=msoTrue(slide) vs msoFalse(selection) both apply.
- **Connectors** — `Shapes.AddConnector(type, x1,y1,x2,y2)` then
  `ConnectorFormat.BeginConnect(shape, site)` / `EndConnect(shape, site)` +
  `RerouteConnections()`. Confirm `Shape.ConnectionSiteCount`, that the glue
  round-trips (`BeginConnected` / `BeginConnectedShape.Id`), and what a connector
  reads back as (`Connector`, `ConnectorFormat.Type`).

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/arrangement_spike.py

Net-zero and polite (temp slide appended then deleted in a `finally`, viewed slide
restored, `net_zero_ok` confirms the slide count is unchanged), exactly like
`scripts/hyperlink_spike.py`.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_MSO_TRUE = -1
_MSO_FALSE = 0
_MSO_GROUP = 6  # msoGroup
_MSO_ALIGN_LEFTS = 1
_MSO_DISTRIBUTE_HORIZONTALLY = 0
_MSO_CONNECTOR_ELBOW = 2


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _add_rect(shapes: Any, left: float, top: float) -> Any:
    """Add a named rectangle and return the raw COM shape."""
    sh = shapes.add_shape("rectangle", left=left, top=top, width=120.0, height=60.0)
    return sh.com


def probe_group_ungroup(slide_com: Any, shapes: Any) -> dict[str, Any]:
    """Group three rectangles, inspect identity, then ungroup and inspect again."""
    out: dict[str, Any] = {}
    try:
        a = _add_rect(shapes, 60.0, 60.0)
        b = _add_rect(shapes, 220.0, 60.0)
        c = _add_rect(shapes, 380.0, 60.0)
        names = [str(a.Name), str(b.Name), str(c.Name)]
        child_ids_before = [int(a.Id), int(b.Id), int(c.Id)]
        out["child_names"] = names
        out["child_ids_before"] = child_ids_before

        rng = slide_com.Shapes.Range(names)
        out["range_count"] = int(rng.Count)
        group = rng.Group()
        out["group_id"] = int(group.Id)
        out["group_type"] = int(group.Type)
        out["group_is_msogroup"] = int(group.Type) == _MSO_GROUP
        out["group_id_is_new"] = int(group.Id) not in child_ids_before
        try:
            items = group.GroupItems
            out["group_items_count"] = int(items.Count)
            out["group_item_ids"] = [int(items(i).Id) for i in range(1, int(items.Count) + 1)]
            out["group_items_keep_ids"] = out["group_item_ids"] == child_ids_before
        except Exception as exc:
            out["group_items_error"] = _err(exc)

        # Ungroup and inspect the freed children's identity.
        freed = group.Ungroup()
        out["ungroup_count"] = int(freed.Count)
        freed_ids = [int(freed(i).Id) for i in range(1, int(freed.Count) + 1)]
        out["freed_ids"] = freed_ids
        out["freed_keep_original_ids"] = sorted(freed_ids) == sorted(child_ids_before)
        out["freed_names"] = [str(freed(i).Name) for i in range(1, int(freed.Count) + 1)]
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_align_distribute(slide_com: Any, shapes: Any) -> dict[str, Any]:
    """Align lefts (relative to slide) and distribute horizontally (relative to selection)."""
    out: dict[str, Any] = {}
    try:
        a = _add_rect(shapes, 80.0, 240.0)
        b = _add_rect(shapes, 300.0, 300.0)
        c = _add_rect(shapes, 500.0, 360.0)
        names = [str(a.Name), str(b.Name), str(c.Name)]

        rng = slide_com.Shapes.Range(names)
        rng.Align(_MSO_ALIGN_LEFTS, _MSO_TRUE)  # relative to slide
        lefts_after = [round(float(a.Left), 1), round(float(b.Left), 1), round(float(c.Left), 1)]
        out["lefts_after_align_to_slide"] = lefts_after
        out["align_to_slide_applied"] = len(set(lefts_after)) == 1

        rng2 = slide_com.Shapes.Range(names)
        rng2.Align(_MSO_ALIGN_LEFTS, _MSO_FALSE)  # relative to each other
        lefts_after2 = [round(float(a.Left), 1), round(float(b.Left), 1), round(float(c.Left), 1)]
        out["lefts_after_align_to_selection"] = lefts_after2
        out["align_to_selection_applied"] = len(set(lefts_after2)) == 1

        rng3 = slide_com.Shapes.Range(names)
        rng3.Distribute(_MSO_DISTRIBUTE_HORIZONTALLY, _MSO_FALSE)
        out["distribute_ok"] = True
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_connector(slide_com: Any, shapes: Any) -> dict[str, Any]:
    """Add an elbow connector, glue it to two shapes, reroute, and read it back."""
    out: dict[str, Any] = {}
    try:
        a = _add_rect(shapes, 80.0, 440.0)
        b = _add_rect(shapes, 400.0, 440.0)
        out["a_connection_sites"] = int(a.ConnectionSiteCount)
        out["b_connection_sites"] = int(b.ConnectionSiteCount)

        conn = slide_com.Shapes.AddConnector(_MSO_CONNECTOR_ELBOW, 0.0, 0.0, 100.0, 100.0)
        out["connector_id"] = int(conn.Id)
        out["connector_is_connector"] = bool(conn.Connector)
        cf = conn.ConnectorFormat
        out["connector_type"] = int(cf.Type)

        cf.BeginConnect(a, 1)
        cf.EndConnect(b, 3)
        conn.RerouteConnections()
        out["begin_connected"] = bool(cf.BeginConnected)
        out["end_connected"] = bool(cf.EndConnected)
        try:
            out["begin_shape_id"] = int(cf.BeginConnectedShape.Id)
            out["end_shape_id"] = int(cf.EndConnectedShape.Id)
            out["glue_round_trips"] = int(cf.BeginConnectedShape.Id) == int(a.Id) and int(
                cf.EndConnectedShape.Id
            ) == int(b.Id)
        except Exception as exc:
            out["connected_shape_error"] = _err(exc)
        out["begin_site"] = int(cf.BeginConnectionSite)
        out["end_site"] = int(cf.EndConnectionSite)
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
            with deck.edit("arrangement spike"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide = deck.slides[temp.index]
                slide_com = slide.com
                shapes = slide.shapes
                findings["group_ungroup"] = probe_group_ungroup(slide_com, shapes)
                findings["align_distribute"] = probe_align_distribute(slide_com, shapes)
                findings["connector"] = probe_connector(slide_com, shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("arrangement spike: cleanup"):
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

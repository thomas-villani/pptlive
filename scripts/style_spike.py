"""Spike harness — verify the pt3 styling COM paths against real PowerPoint.

Validates the raw COM the upcoming wrappers will lean on, before any of them
exist (so it pokes `.com` / raw dispatch directly):

- **Fill** — `Shape.Fill.Solid()` then `Shape.Fill.ForeColor.RGB = rgb` sets a
  solid fill; `Shape.Fill.Visible = msoFalse` makes it transparent;
  `Shape.Fill.ForeColor.RGB` reads back the literal RGB (and a theme default
  reads as the `0x80000000` sentinel, like font color).
- **Line** — `Shape.Line.ForeColor.RGB`, `Shape.Line.Weight` (points), and
  `Shape.Line.Visible` round-trip.
- **Z-order** — `Shape.ZOrder(cmd)` with the `MsoZOrderCmd` ints
  (BringToFront=0, SendToBack=1, BringForward=2, SendBackward=3) reorders the
  z-stack; `Shape.ZOrderPosition` reports the new 1-based slot.
- **Resolve by Id** — `Shape.Id` is stable across a delete that shifts the
  z-order index, so scanning `Shapes` for a known `.Id` re-finds the same shape
  (the `shapeid:S:ID` anchor's mechanism).

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/style_spike.py

Net-zero and polite, exactly like `scripts/shape_spike.py`: everything is built
on a single temporary slide that's appended then deleted in a `finally`, the
viewed slide is restored, and `net_zero_ok` confirms the deck's slide count is
unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_RED = 0x0000FF  # COM RGB is 0xBBGGRR, so this is pure red
_BLUE = 0xFF0000
_THEME_SENTINEL = 0x80000000  # what a non-literal (theme/auto) color reads back as


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_fill(shapes: Any) -> dict[str, Any]:
    """Solid fill set/read + transparent (Visible=msoFalse)."""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=72.0, top=72.0, width=120.0, height=120.0)
        com = sh.com
        default_rgb = int(com.Fill.ForeColor.RGB)
        com.Fill.Solid()
        com.Fill.ForeColor.RGB = _RED
        set_rgb = int(com.Fill.ForeColor.RGB)
        com.Fill.Visible = 0  # msoFalse -> transparent
        hidden = int(com.Fill.Visible)
        out = {
            "default_rgb": default_rgb,
            "default_is_theme_sentinel": default_rgb == _THEME_SENTINEL,
            "set_rgb": set_rgb,
            "set_rgb_ok": set_rgb == _RED,
            "visible_after_hide": hidden,
            "hide_ok": hidden == 0,
        }
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_line(shapes: Any) -> dict[str, Any]:
    """Line color / weight / visible round-trip."""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=220.0, top=72.0, width=120.0, height=120.0)
        com = sh.com
        com.Line.ForeColor.RGB = _BLUE
        com.Line.Weight = 4.5
        com.Line.Visible = -1  # msoTrue
        out = {
            "line_rgb": int(com.Line.ForeColor.RGB),
            "line_rgb_ok": int(com.Line.ForeColor.RGB) == _BLUE,
            "weight": float(com.Line.Weight),
            "weight_ok": abs(float(com.Line.Weight) - 4.5) < 0.01,
            "visible": int(com.Line.Visible),
        }
        com.Line.Visible = 0
        out["hide_ok"] = int(com.Line.Visible) == 0
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_zorder(shapes: Any) -> dict[str, Any]:
    """ZOrder(cmd) reorders; ZOrderPosition reports the slot.

    Holds **raw COM dispatch** refs (not the index-resolving wrappers), since a
    wrapper's `.com` re-resolves by z-order index, and that index is exactly what
    a reorder changes (PPTLIVE-010). A raw dispatch ref tracks the same object.
    """
    out: dict[str, Any] = {}
    try:
        col = shapes._com_collection
        # Three overlapping shapes; we'll shuffle them. Keep raw refs.
        a = col.AddShape(9, 360.0, 72.0, 100.0, 100.0)  # oval
        b = col.AddShape(9, 380.0, 92.0, 100.0, 100.0)
        c = col.AddShape(9, 400.0, 112.0, 100.0, 100.0)
        before = (int(a.ZOrderPosition), int(b.ZOrderPosition), int(c.ZOrderPosition))
        c.ZOrder(1)  # msoSendToBack
        c_back = int(c.ZOrderPosition)
        a.ZOrder(2)  # msoBringForward
        a_fwd = int(a.ZOrderPosition)
        out = {
            "positions_before": before,
            "c_after_send_to_back": c_back,
            "a_after_bring_forward": a_fwd,
            "send_to_back_ok": c_back < before[2],
            "bring_forward_ok": a_fwd > before[0],
        }
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_resolve_by_id(deck: Any, slide_index: int) -> dict[str, Any]:
    """Shape.Id survives a delete that renumbers the z-order index."""
    out: dict[str, Any] = {}
    try:
        com_shapes = deck.slides[slide_index].com.Shapes
        first = com_shapes.AddShape(1, 60.0, 240.0, 80.0, 80.0)  # rectangle
        target = com_shapes.AddShape(9, 160.0, 240.0, 80.0, 80.0)  # oval
        target_id = int(target.Id)
        idx_before = int(target.ZOrderPosition)
        first.Delete()  # renumbers every higher index down by one

        def find_by_id(wanted: int) -> Any:
            for i in range(1, int(com_shapes.Count) + 1):
                sh = com_shapes(i)
                if int(sh.Id) == wanted:
                    return sh
            return None

        refound = find_by_id(target_id)
        idx_after = int(refound.ZOrderPosition) if refound is not None else None
        out = {
            "target_id": target_id,
            "index_before_delete": idx_before,
            "index_after_delete": idx_after,
            "index_shifted": idx_after != idx_before,
            "refound_same_id": refound is not None and int(refound.Id) == target_id,
        }
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
            with deck.edit("style spike: fill / line / zorder / resolve-by-id"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["fill"] = probe_fill(shapes)
                findings["line"] = probe_line(shapes)
                findings["zorder"] = probe_zorder(shapes)
                findings["resolve_by_id"] = probe_resolve_by_id(deck, temp.index)
                findings["final_shape_count"] = len(shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("style spike: cleanup"):
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

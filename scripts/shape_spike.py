"""Spike harness — verify the v0.2 shape & geometry COM verbs against real PowerPoint.

v0.2 encodes assumptions the fake can't prove: that `Shapes.AddTextbox(Orientation,
Left, Top, Width, Height)`, `Shapes.AddShape(Type, …)`, and `Shapes.AddPicture(
FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)` each add a shape
and return it; that a newly added shape lands at the **top of the z-order** (the
last slot, so its 1-based index is the post-add `Shapes.Count` — what
`ShapeCollection._added()` relies on); that the friendly autoshape ints pptlive
ships map to the geometries it claims (`rectangle`==1, `oval`==9, `right_arrow`==33,
`five_point_star`==92 — the higher values are the ones most worth checking); and
that `Shape.move`/`resize`/`Delete` behave. Run against a *running* PowerPoint with
a deck open:

    uv run python scripts/shape_spike.py

Prints one JSON findings object to stdout. Like `scripts/layout_spike.py` it is
**net-zero and polite**: every shape is created on a single *temporary* slide that
is appended and then deleted in a `finally` (so no real slide is touched), the
viewed slide is restored at the end, and `net_zero_ok` confirms the deck's slide
count is unchanged. It exercises the shipped `pptlive` wrappers, so it doubles as a
live integration check of the v0.2 surface.
"""

from __future__ import annotations

import base64
import json
import os
import tempfile
from typing import Any

import pptlive as pl
from pptlive import _selection

# A valid 1x1 PNG, so AddPicture has a real image to embed.
_PNG_1X1 = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
)

# Friendly name -> the MsoAutoShapeType int pptlive claims it maps to.
_AUTOSHAPE_CHECKS = {
    "rectangle": 1,
    "oval": 9,
    "right_arrow": 33,
    "star": 92,
}


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_textbox(shapes: Any) -> dict[str, Any]:
    """AddTextbox returns a usable shape that lands at the last z-order slot."""
    out: dict[str, Any] = {}
    try:
        count_before = len(shapes)
        box = shapes.add_textbox("MARKER-TB", left=72.0, top=72.0, width=288.0, height=72.0)
        # If `_added()` indexed the wrong shape, this text read would not match.
        roundtrip = shapes[len(shapes)].text
        out = {
            "index": box.index,
            "id": box.shape_id,
            "type": box.shape_type,
            "text": box.text,
            "geometry": box.geometry(),
            "count_grew_by_one": len(shapes) == count_before + 1,
            "lands_at_last_index": box.index == len(shapes),
            "text_roundtrips_at_last_index": roundtrip == "MARKER-TB",
        }
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_autoshapes(shapes: Any) -> dict[str, Any]:
    """Each friendly name produces an auto_shape whose AutoShapeType is as claimed."""
    out: dict[str, Any] = {}
    for name, expected in _AUTOSHAPE_CHECKS.items():
        try:
            sh = shapes.add_shape(name, left=72.0, top=200.0, width=120.0, height=120.0)
            actual = int(sh.com.AutoShapeType)
            out[name] = {
                "type": sh.shape_type,
                "auto_shape_type": actual,
                "expected": expected,
                "ok": sh.shape_type == "auto_shape" and actual == expected,
            }
        except Exception as exc:
            out[name] = {"error": _err(exc), "expected": expected}
    return out


def probe_picture(shapes: Any, png_path: str) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        pic = shapes.add_picture(png_path, left=72.0, top=340.0, width=96.0, height=96.0)
        out = {
            "index": pic.index,
            "type": pic.shape_type,
            "has_text_frame": pic.has_text_frame,
            "geometry": pic.geometry(),
            "ok": pic.shape_type in ("picture", "linked_picture"),
        }
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_geometry(shapes: Any) -> dict[str, Any]:
    """move / resize / delete on a throwaway shape."""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=400.0, top=72.0, width=100.0, height=100.0)
        sh.move(left=420.0, top=90.0)
        sh.resize(width=160.0, height=120.0)
        geo = sh.geometry()
        out["move_resize"] = {
            "geometry": geo,
            "ok": geo["left"] == 420.0
            and geo["top"] == 90.0
            and geo["width"] == 160.0
            and geo["height"] == 120.0,
        }
        count_before = len(shapes)
        sh.delete()
        out["delete"] = {
            "count_before": count_before,
            "count_after": len(shapes),
            "ok": len(shapes) == count_before - 1,
        }
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    png_path: str | None = None
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        findings["slide_count_before"] = count_before

        snap = _selection.snapshot(ppt)
        findings["viewed_slide"] = snap.slide_index

        fd, png_path = tempfile.mkstemp(suffix=".png", prefix="pptlive_spike_")
        os.write(fd, _PNG_1X1)
        os.close(fd)

        temp_ids: list[int] = []
        try:
            with deck.edit("shape spike: add textbox / autoshapes / picture / geometry"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["textbox"] = probe_textbox(shapes)
                findings["autoshapes"] = probe_autoshapes(shapes)
                findings["picture"] = probe_picture(shapes, png_path)
                findings["geometry"] = probe_geometry(shapes)
                findings["final_shape_count"] = len(shapes)
        finally:
            # Backstop cleanup (same robust, index-based pass as layout_spike): the
            # temp slide carries all the shapes we made, so deleting it is net-zero.
            deleted: list[int] = []
            try:
                with deck.edit("shape spike: cleanup"):
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

    if png_path and os.path.exists(png_path):
        os.remove(png_path)

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Spike harness — verify the v0.1 slide-lifecycle COM verbs against real PowerPoint.

v0.1 encodes a handful of assumptions the fake can't prove: that
`SlideMaster.CustomLayouts` enumerates with `.Name`, that
`Slides.AddSlide(Index, CustomLayout)` adds and returns a slide, that
`Slide.Duplicate()` yields a 1-based `SlideRange`, that `Slide.MoveTo(toPos)`
reorders, that assigning `Slide.CustomLayout` re-applies a layout, and that
`Slide.Delete()` removes. Run against a *running* PowerPoint with a deck open:

    uv run python scripts/layout_spike.py

Prints one JSON findings object to stdout. It is deliberately **net-zero and
polite**: every slide it creates is tracked by `SlideID` and deleted in a
`finally`, the move/duplicate probes operate only on those appended temp slides
(never reordering the user's real slides), and the viewed slide is restored at
the end. A `net_zero_ok` flag confirms the deck's slide count is unchanged.

Like `scripts/spike.py`, this exercises the shipped `pptlive` wrappers (not just
raw COM) so it doubles as a live integration check of the v0.1 surface.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_layouts(deck: pl.Presentation) -> dict[str, Any]:
    """List the deck's CustomLayouts and confirm the default alias resolves."""
    out: dict[str, Any] = {}
    try:
        rows = deck.layouts()
        out["layouts"] = rows
        out["count"] = len(rows)
        # Resolve the add() default the way the library will, without mutating.
        custom = deck._resolve_layout(None)
        out["default_layout_name"] = None if custom is None else str(custom.Name)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_lifecycle(deck: pl.Presentation, temp_ids: list[int]) -> dict[str, Any]:
    """Append a temp slide, then set-layout / duplicate / move it (all net-zero).

    Appends to `temp_ids` as slides are created so the caller's `finally` can
    always clean them up, even if a later step raises.
    """
    out: dict[str, Any] = {}

    # add ------------------------------------------------------------------
    try:
        temp = deck.slides.add(layout="title_and_content")
        temp_ids.append(temp.id)
        out["add"] = {"index": temp.index, "id": temp.id, "layout": temp.layout_name}
    except Exception as exc:
        out["add"] = {"error": _err(exc)}
        return out  # nothing else can run without the temp slide

    # set_layout -----------------------------------------------------------
    try:
        temp.set_layout("two_content")
        out["set_layout"] = {
            "layout_after": temp.layout_name,
            "ok": temp.layout_name == "Two Content",
        }
    except Exception as exc:
        out["set_layout"] = {"error": _err(exc)}

    # duplicate ------------------------------------------------------------
    dup = None
    try:
        temp_index_before = temp.index
        dup = temp.duplicate()
        temp_ids.append(dup.id)
        out["duplicate"] = {
            "dup_index": dup.index,
            "dup_id": dup.id,
            "dup_layout": dup.layout_name,
            "inserted_immediately_after": dup.index == temp_index_before + 1,
            "fresh_id": dup.id != temp.id,
        }
    except Exception as exc:
        out["duplicate"] = {"error": _err(exc)}

    # move_to (only the appended temp slides move — real slides untouched) --
    if dup is not None:
        try:
            target = temp.index  # slot the duplicate just behind the temp slide
            dup.move_to(target)
            out["move_to"] = {
                "requested": target,
                "index_after": dup.index,
                "ok": dup.index == target,
            }
        except Exception as exc:
            out["move_to"] = {"error": _err(exc)}

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

        findings["layouts"] = probe_layouts(deck)

        temp_ids: list[int] = []
        try:
            with deck.edit("layout spike: add / set-layout / duplicate / move"):
                findings["lifecycle"] = probe_lifecycle(deck, temp_ids)
        finally:
            # Backstop cleanup: delete every slide we created, by SlideID, no
            # matter what the probes did. Re-resolve each slide *by index* fresh
            # on every pass (a bulk enumerator can briefly yield a stale handle
            # right after a move/duplicate), guard each .id read, and delete one
            # victim per pass so indices never go stale mid-loop.
            deleted: list[int] = []
            try:
                with deck.edit("layout spike: cleanup"):
                    for _ in range(len(temp_ids) + 2):  # bounded; one delete per pass
                        victim = None
                        for idx in range(len(deck.slides), 0, -1):
                            try:
                                sid = deck.slides[idx].id
                            except Exception:
                                continue  # skip a transiently-stale handle
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

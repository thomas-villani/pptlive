"""Spike harness — verify the slide-transition COM paths against real PowerPoint.

Pins the raw COM the upcoming `Slide.set_transition` wrapper (v1.5 cut) will lean
on, before it exists (so it pokes `.com` directly). The unknowns:

- **EntryEffect** — `Slide.SlideShowTransition.EntryEffect = <PpEntryEffect int>`
  writes and reads back cleanly (probe a few known values: none=0, cut=257,
  fade=1793, push-up?, so the wrapper's curated friendly map is grounded in what
  this build actually accepts/normalizes).
- **Duration** — `.Duration` (the transition's animation length, seconds) round-trips.
- **Auto-advance model** — does `.AdvanceTime` (seconds) take effect on its own, or
  must `.AdvanceOnTime = msoTrue` be set too? And `.AdvanceOnClick` round-trips.

`msoTrue` = -1, `msoFalse` = 0. Run against a *running* PowerPoint with a deck open:

    uv run python scripts/transition_spike.py

Net-zero and polite (temp slide appended then deleted in a `finally`, viewed
slide restored), exactly like `scripts/style_spike.py`. The transition lives on
the temp slide itself, so no extra shapes are needed.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

# The curated PpEntryEffect set the wrapper will expose, with the AUTHORITATIVE
# Microsoft enum values — confirm this build accepts/round-trips every one before
# baking them into `constants.PpEntryEffect` (round-trip = accepted & stored; it
# does NOT verify the visual, same bar as chart_type_for / autoshape_type_for).
_EFFECTS = {
    "none": 0,
    "cut": 257,
    "cut_through_black": 258,
    "random": 513,
    "blinds_horizontal": 769,
    "blinds_vertical": 770,
    "checkerboard_across": 1025,
    "checkerboard_down": 1026,
    "cover_left": 1281,
    "cover_up": 1282,
    "cover_right": 1283,
    "cover_down": 1284,
    "dissolve": 1537,
    "fade": 1793,
    "uncover_left": 2049,
    "uncover_up": 2050,
    "uncover_right": 2051,
    "uncover_down": 2052,
    "wipe_left": 3329,
    "wipe_right": 3330,
    "wipe_up": 3331,
    "wipe_down": 3332,
    "box_out": 3585,
    "box_in": 3586,
}


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_entry_effects(slide_com: Any) -> dict[str, Any]:
    """Set each candidate EntryEffect and read it back."""
    out: dict[str, Any] = {}
    trans = slide_com.SlideShowTransition
    out["default_entry_effect"] = int(trans.EntryEffect)
    for name, value in _EFFECTS.items():
        try:
            trans.EntryEffect = value
            readback = int(trans.EntryEffect)
            out[name] = {
                "set": value,
                "readback": readback,
                "round_trips": readback == value,
            }
        except Exception as exc:
            out[name] = {"set": value, "error": _err(exc)}
    return out


def probe_advance_and_duration(slide_com: Any) -> dict[str, Any]:
    """Duration + the auto-advance model (AdvanceOnTime / AdvanceTime / AdvanceOnClick)."""
    out: dict[str, Any] = {}
    trans = slide_com.SlideShowTransition
    try:
        trans.Duration = 0.75
        out["duration_readback"] = float(trans.Duration)
        out["duration_ok"] = abs(float(trans.Duration) - 0.75) < 0.01
    except Exception as exc:
        out["duration_error"] = _err(exc)
    try:
        out["advance_on_click_default"] = int(trans.AdvanceOnClick)
        out["advance_on_time_default"] = int(trans.AdvanceOnTime)
        out["advance_time_default"] = float(trans.AdvanceTime)
        # Set auto-advance: both the flag and the seconds.
        trans.AdvanceOnTime = -1  # msoTrue
        trans.AdvanceTime = 3.0
        trans.AdvanceOnClick = 0  # msoFalse
        out["advance_on_time_set"] = int(trans.AdvanceOnTime)
        out["advance_time_set"] = float(trans.AdvanceTime)
        out["advance_on_click_set"] = int(trans.AdvanceOnClick)
        out["advance_round_trips"] = (
            int(trans.AdvanceOnTime) == -1
            and abs(float(trans.AdvanceTime) - 3.0) < 0.01
            and int(trans.AdvanceOnClick) == 0
        )
    except Exception as exc:
        out["advance_error"] = _err(exc)
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
            with deck.edit("transition spike: entry effect / advance / duration"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide_com = deck.slides[temp.index].com
                findings["entry_effects"] = probe_entry_effects(slide_com)
                findings["advance_duration"] = probe_advance_and_duration(slide_com)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("transition spike: cleanup"):
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

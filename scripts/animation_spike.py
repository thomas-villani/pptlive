"""Spike harness — pin slide *animations* against real PowerPoint.

The roadmap's v1.5 "Animations (fiddly, second cut)" is deferred with an explicit
open question: does `AddEffect` round-trip its `EffectType`/`Timing`, or is some of
it write-only (like SmartArt assistant nodes)? This probes the raw `TimeLine` /
`MainSequence` COM directly (no wrapper exists) to answer it and pin the 80% asks
"show this" (appear) and "fade this in" (fade):

- **Add** — `Slide.TimeLine.MainSequence.AddEffect(Shape, EffectId, Level,
  Trigger)`. `MsoAnimEffect`: appear=1, fade=10, flyIn=2. `Trigger`
  (`MsoAnimTriggerType`): onPageClick=1, withPrevious=2, afterPrevious=3.
- **Read back** — the returned `Effect`: `.EffectType`, `.Shape.Id` (map effect →
  shape), `.Exit`, `.Timing.Duration`/`.TriggerType`/`.TriggerDelayTime`, and the
  `MainSequence` count / iteration. Does `EffectType` survive the round-trip?
- **Tune** — set `Effect.Timing.Duration` (seconds) and trigger; make it an
  **exit** effect (`Effect.Exit = msoTrue`) and confirm read-back.
- **Remove** — `Effect.Delete()` drops the sequence count.

Run:  uv run python scripts/animation_spike.py

Net-zero / polite exactly like `scripts/style_spike.py`.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

# MsoAnimEffect
_APPEAR = 1
_FADE = 10
_FLY_IN = 2
# MsoAnimTriggerType
_TRIGGER_ON_CLICK = 1
_TRIGGER_WITH_PREVIOUS = 2
_TRIGGER_AFTER_PREVIOUS = 3


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:200]


def _read_effect(eff: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for name in ("EffectType", "Exit", "DisplayName"):
        try:
            out[name] = getattr(eff, name)
        except Exception as exc:
            out[name] = _err(exc)
    try:
        out["shape_id"] = int(eff.Shape.Id)
    except Exception as exc:
        out["shape_id"] = _err(exc)
    try:
        out["shape_name"] = str(eff.Shape.Name)
    except Exception as exc:
        out["shape_name"] = _err(exc)
    t = None
    try:
        t = eff.Timing
    except Exception as exc:
        out["timing_error"] = _err(exc)
    if t is not None:
        for name in ("Duration", "TriggerType", "TriggerDelayTime", "Speed"):
            try:
                out[f"Timing.{name}"] = getattr(t, name)
            except Exception as exc:
                out[f"Timing.{name}"] = _err(exc)
    return out


def probe_add_fade(slide_com: Any, shape_com: Any) -> dict[str, Any]:
    """The headline ask: 'fade this in'."""
    out: dict[str, Any] = {}
    try:
        seq = slide_com.TimeLine.MainSequence
        out["count_before"] = int(seq.Count)
        eff = seq.AddEffect(shape_com, _FADE, 0, _TRIGGER_ON_CLICK)
        out["count_after_add"] = int(seq.Count)
        out["effect_read"] = _read_effect(eff)
        out["effecttype_round_trips"] = out["effect_read"].get("EffectType") == _FADE
        # Tune duration + trigger.
        try:
            eff.Timing.Duration = 2.0
            eff.Timing.TriggerType = _TRIGGER_AFTER_PREVIOUS
            out["tuned"] = True
            out["effect_after_tune"] = _read_effect(eff)
        except Exception as exc:
            out["tune_error"] = _err(exc)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_add_appear(slide_com: Any, shape_com: Any) -> dict[str, Any]:
    """'show this' — appear, with-previous."""
    out: dict[str, Any] = {}
    try:
        seq = slide_com.TimeLine.MainSequence
        eff = seq.AddEffect(shape_com, _APPEAR, 0, _TRIGGER_WITH_PREVIOUS)
        out["effect_read"] = _read_effect(eff)
        out["effecttype_round_trips"] = out["effect_read"].get("EffectType") == _APPEAR
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_exit_effect(slide_com: Any, shape_com: Any) -> dict[str, Any]:
    """Make a fade an *exit* effect and confirm .Exit reads back."""
    out: dict[str, Any] = {}
    try:
        seq = slide_com.TimeLine.MainSequence
        eff = seq.AddEffect(shape_com, _FADE, 0, _TRIGGER_ON_CLICK)
        try:
            eff.Exit = -1  # msoTrue
            out["set_exit_ok"] = True
        except Exception as exc:
            out["set_exit_error"] = _err(exc)
        out["effect_read"] = _read_effect(eff)
        out["exit_round_trips"] = out["effect_read"].get("Exit") in (-1, True)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_iterate_and_delete(slide_com: Any) -> dict[str, Any]:
    """Iterate the whole sequence (read), then delete one effect."""
    out: dict[str, Any] = {}
    try:
        seq = slide_com.TimeLine.MainSequence
        n = int(seq.Count)
        out["count"] = n
        listing = []
        for i in range(1, n + 1):
            listing.append(_read_effect(seq(i)))
        out["all_effects"] = listing
        if n > 0:
            seq(1).Delete()
            out["count_after_delete_one"] = int(seq.Count)
            out["delete_ok"] = int(seq.Count) == n - 1
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
            with deck.edit("animation spike: add/read/tune/exit/delete effects"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide = deck.slides[temp.index]
                slide_com = slide.com
                shapes = slide.shapes
                s1 = shapes.add_shape("rectangle", left=40.0, top=40.0, width=120.0, height=90.0)
                s2 = shapes.add_shape("oval", left=200.0, top=40.0, width=120.0, height=90.0)
                s3 = shapes.add_shape("star", left=360.0, top=40.0, width=120.0, height=90.0)
                findings["add_fade"] = probe_add_fade(slide_com, s1.com)
                findings["add_appear"] = probe_add_appear(slide_com, s2.com)
                findings["exit_effect"] = probe_exit_effect(slide_com, s3.com)
                findings["iterate_and_delete"] = probe_iterate_and_delete(slide_com)
                findings["final_shape_count"] = len(shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("animation spike: cleanup"):
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

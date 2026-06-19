"""Confirmation spike — pin the *curated effect set* + the `delay` knob.

The original `animation_spike.py` proved the AddEffect round-trip mechanism with
appear/fade and Duration/TriggerType. This narrows the last open risk before
hard-coding constants: does every curated `MsoAnimEffect` int we plan to name
accept via `AddEffect` and read its `EffectType` back, and does
`Timing.TriggerDelayTime` (the `delay=` knob) round-trip? Net-zero / polite.

Run:  uv run python scripts/animation_curated_spike.py
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

# The curated friendly -> MsoAnimEffect int set we intend to name.
CURATED = {
    "appear": 1,
    "fly_in": 2,
    "grow_turn": 14,
    "float_in": 21,
    "split": 23,
    "swivel": 26,
    "wheel": 28,
    "wipe": 29,
    "zoom": 31,
    "fade": 10,
}
_TRIGGER_ON_CLICK = 1


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:200]


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)
        temp_ids: list[int] = []
        try:
            with deck.edit("curated-anim spike"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide = deck.slides[temp.index]
                slide_com = slide.com
                seq = slide_com.TimeLine.MainSequence
                per_effect: dict[str, Any] = {}
                for name, eid in CURATED.items():
                    try:
                        sh = slide.shapes.add_shape(
                            "rectangle", left=20.0, top=20.0, width=80.0, height=60.0
                        )
                        eff = seq.AddEffect(sh.com, eid, 0, _TRIGGER_ON_CLICK)
                        got = int(eff.EffectType)
                        per_effect[name] = {"requested": eid, "read_back": got, "ok": got == eid}
                    except Exception as exc:
                        per_effect[name] = {"requested": eid, "error": _err(exc)}
                findings["curated"] = per_effect
                # delay round-trip
                try:
                    sh = slide.shapes.add_shape(
                        "oval", left=20.0, top=120.0, width=80.0, height=60.0
                    )
                    eff = seq.AddEffect(sh.com, 10, 0, _TRIGGER_ON_CLICK)
                    eff.Timing.Duration = 1.5
                    eff.Timing.TriggerDelayTime = 0.75
                    findings["delay_roundtrip"] = {
                        "duration": float(eff.Timing.Duration),
                        "delay": float(eff.Timing.TriggerDelayTime),
                        "ok": abs(float(eff.Timing.TriggerDelayTime) - 0.75) < 1e-6,
                    }
                except Exception as exc:
                    findings["delay_roundtrip"] = {"error": _err(exc)}
        finally:
            try:
                with deck.edit("curated-anim cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            if deck.slides[idx].id in temp_ids:
                                deck.slides[idx].delete()
                        except Exception:
                            continue
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            _selection.restore(ppt, snap)
        findings["net_zero_ok"] = len(deck.slides) == count_before
    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

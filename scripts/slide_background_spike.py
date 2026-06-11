"""Spike harness — verify the per-slide background COM paths against real PowerPoint.

Pins the raw COM the upcoming `Slide.set_background` wrapper (v1.2 cut) will lean
on, before it exists (so it pokes `.com` directly). The per-slide override of
v0.9's deck-wide master background. The unknowns:

- **Override flag** — `Slide.FollowMasterBackground = msoFalse` lets the slide
  carry its own background; default reads as msoTrue.
- **Solid fill** — `Slide.Background.Fill.Solid()` then `.ForeColor.RGB = rgb`
  sets the slide's own background colour, mirroring `Master.set_background`
  (`_theme.py`). Confirm `.Fill.ForeColor.RGB` and `.Fill.Type` read back.
- **Revert** — re-setting `FollowMasterBackground = msoTrue` cleanly drops the
  override and falls back to the master background.

`msoTrue` = -1, `msoFalse` = 0; COM RGB is 0xBBGGRR. Run against a *running*
PowerPoint with a deck open:

    uv run python scripts/slide_background_spike.py

Net-zero and polite (temp slide appended then deleted in a `finally`, viewed
slide restored), exactly like `scripts/style_spike.py`.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_TEAL = 0x807040  # 0xBBGGRR -> R=0x40, G=0x70, B=0x80
_MSO_FILL_SOLID = 1  # msoFillSolid


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_background(slide_com: Any) -> dict[str, Any]:
    """FollowMasterBackground toggle + solid fill set/read + revert."""
    out: dict[str, Any] = {}
    try:
        out["follow_master_default"] = int(slide_com.FollowMasterBackground)
        # Take over the background.
        slide_com.FollowMasterBackground = 0  # msoFalse
        out["follow_master_after_override"] = int(slide_com.FollowMasterBackground)
        fill = slide_com.Background.Fill
        fill.Solid()
        fill.ForeColor.RGB = _TEAL
        out["bg_rgb_readback"] = int(fill.ForeColor.RGB)
        out["bg_rgb_ok"] = int(fill.ForeColor.RGB) == _TEAL
        try:
            out["bg_fill_type"] = int(fill.Type)
            out["bg_fill_type_is_solid"] = int(fill.Type) == _MSO_FILL_SOLID
        except Exception as exc:
            out["bg_fill_type_error"] = _err(exc)
        # Revert to the master background.
        slide_com.FollowMasterBackground = -1  # msoTrue
        out["follow_master_after_revert"] = int(slide_com.FollowMasterBackground)
        out["revert_ok"] = int(slide_com.FollowMasterBackground) == -1
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
            with deck.edit("slide background spike: override / solid / revert"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                slide_com = deck.slides[temp.index].com
                findings["background"] = probe_background(slide_com)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("slide background spike: cleanup"):
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

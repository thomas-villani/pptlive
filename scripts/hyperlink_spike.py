"""Spike harness — verify the shape-hyperlink COM paths against real PowerPoint.

Pins the raw COM the upcoming `Shape.set_hyperlink` wrapper (v1.4 cut) will lean
on, before it exists (so it pokes `.com` / raw dispatch directly). The two real
unknowns:

- **URL link** — `Shape.ActionSettings(ppMouseClick).Hyperlink.Address = url`
  round-trips, and whether setting `.Address` implicitly flips `.Action` to
  `ppActionHyperlink` (7). Also: what does `.Address` read back as when *no* link
  is set (empty string vs None vs raise)?
- **Slide-jump link** — the `Hyperlink.SubAddress` string format for "jump to
  slide N". The PowerPoint-UI convention is `"<SlideID>,<index>,<title>"`; this
  probe sets that form against a real target slide and reads it back, so the
  wrapper knows the exact encoding COM accepts (and what it normalizes to).
- **Removal** — `Hyperlink.Delete()` (and/or `Action = ppActionNone` = 0) clears
  the link.

`ppMouseClick` = 1. Run against a *running* PowerPoint with a deck open:

    uv run python scripts/hyperlink_spike.py

Net-zero and polite (temp slide appended then deleted in a `finally`, viewed
slide restored, `net_zero_ok` confirms the slide count is unchanged), exactly
like `scripts/style_spike.py`.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_PP_MOUSE_CLICK = 1
_PP_ACTION_NONE = 0
_PP_ACTION_HYPERLINK = 7
_URL = "https://www.anthropic.com/"


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe_url_link(shapes: Any) -> dict[str, Any]:
    """Set Hyperlink.Address to a URL; read back Address + Action."""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=72.0, top=72.0, width=160.0, height=80.0)
        com = sh.com
        acts = com.ActionSettings(_PP_MOUSE_CLICK)
        # What does an un-linked shape read back as?
        try:
            out["address_before"] = repr(acts.Hyperlink.Address)
        except Exception as exc:
            out["address_before_error"] = _err(exc)
        out["action_before"] = int(acts.Action)
        acts.Hyperlink.Address = _URL
        out["address_after"] = str(acts.Hyperlink.Address)
        out["address_ok"] = str(acts.Hyperlink.Address) == _URL
        out["action_after"] = int(acts.Action)
        out["action_became_hyperlink"] = int(acts.Action) == _PP_ACTION_HYPERLINK
        # Remove via Delete()
        acts.Hyperlink.Delete()
        out["action_after_delete"] = int(acts.Action)
        try:
            out["address_after_delete"] = repr(acts.Hyperlink.Address)
        except Exception as exc:
            out["address_after_delete_error"] = _err(exc)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_slide_jump(deck: Any, shapes: Any) -> dict[str, Any]:
    """Set Hyperlink.SubAddress to jump to slide 1; read back the normalized form."""
    out: dict[str, Any] = {}
    try:
        target = deck.slides[1].com
        target_id = int(target.SlideID)
        target_index = int(target.SlideIndex)
        # Title text, if the slide has a title placeholder.
        try:
            target_title = str(target.Shapes.Title.TextFrame.TextRange.Text)
        except Exception:
            target_title = ""
        out["target_slide_id"] = target_id
        out["target_index"] = target_index
        out["target_title"] = target_title

        sh = shapes.add_shape("rectangle", left=72.0, top=200.0, width=160.0, height=80.0)
        acts = sh.com.ActionSettings(_PP_MOUSE_CLICK)
        candidate = f"{target_id},{target_index},{target_title}"
        out["subaddress_set"] = candidate
        acts.Hyperlink.SubAddress = candidate
        out["subaddress_readback"] = str(acts.Hyperlink.SubAddress)
        out["address_readback"] = repr(acts.Hyperlink.Address)
        out["action_after"] = int(acts.Action)
        # Does a bare "index" form also work? (probe a second shape)
        sh2 = shapes.add_shape("rectangle", left=260.0, top=200.0, width=160.0, height=80.0)
        acts2 = sh2.com.ActionSettings(_PP_MOUSE_CLICK)
        try:
            acts2.Hyperlink.SubAddress = str(target_index)
            out["bare_index_readback"] = str(acts2.Hyperlink.SubAddress)
            out["bare_index_ok"] = True
        except Exception as exc:
            out["bare_index_error"] = _err(exc)
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
            with deck.edit("hyperlink spike: url / slide-jump / remove"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["url_link"] = probe_url_link(shapes)
                findings["slide_jump"] = probe_slide_jump(deck, shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("hyperlink spike: cleanup"):
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

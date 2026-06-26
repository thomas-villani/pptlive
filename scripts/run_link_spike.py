"""Spike harness — pin text-run-level hyperlink COM paths against real PowerPoint.

Shape-level hyperlinks shipped (`Shape.set_hyperlink`); the remaining v1.4 piece is
links on a SUB-SPAN of a text frame (a linked word inside a textbox / cell / notes).
`scripts/hyperlink_spike.py` only probed `Shape.ActionSettings`; this confirms the
SAME COM works on a character range — the unknowns the wrapper rests on:

- **Span link** — `tr.Characters(start+1, length).ActionSettings(ppMouseClick=1)
  .Hyperlink.Address = url` round-trips and auto-flips that range's `.Action` to
  `ppActionHyperlink=7` (mirrors the shape finding, but scoped to the run).
- **Slide jump on a span** — `.SubAddress = "<SlideID>,<index>,<title>"` round-trips
  on a character range.
- **Read-back** — iterate `tr.Runs()`, read each run's
  `ActionSettings(1).Action` / `.Hyperlink.Address` / `.SubAddress` to reconstruct
  which runs carry a link and at what offsets. Confirm a linked span shows up as its
  own run (PowerPoint splits a run at a formatting/action boundary).
- **Removal** — `Characters(...).ActionSettings(1).Hyperlink.Delete()` clears just
  that span, leaving the rest of the frame's links intact.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/run_link_spike.py

Net-zero / polite exactly like `scripts/hyperlink_spike.py`.
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
_TEXT = "Visit Anthropic now and later"
#         0123456789...        ^ "Anthropic" at offset 6, length 9


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _walk_runs(tr: Any) -> list[dict[str, Any]]:
    """Reconstruct per-run text + offset + any link, the way `links()` will read."""
    rows: list[dict[str, Any]] = []
    runs = tr.Runs()
    offset = 0
    for i in range(1, int(runs.Count) + 1):
        run = tr.Runs(i, 1)
        text = str(run.Text or "")
        row: dict[str, Any] = {"index": i, "start": offset, "length": len(text), "text": text}
        try:
            acts = run.ActionSettings(_PP_MOUSE_CLICK)
            if int(acts.Action) == _PP_ACTION_HYPERLINK:
                row["address"] = str(acts.Hyperlink.Address or "") or None
                row["sub_address"] = str(acts.Hyperlink.SubAddress or "") or None
        except Exception as exc:
            row["run_action_error"] = _err(exc)
        rows.append(row)
        offset += len(text)
    return rows


def probe_span_link(shapes: Any) -> dict[str, Any]:
    """Link just the word 'Anthropic' (offset 6, length 9) to a URL; read it back."""
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_textbox(_TEXT, left=72.0, top=72.0, width=400.0, height=60.0)
        tr = sh.com.TextFrame.TextRange
        out["frame_text"] = str(tr.Text)
        span = tr.Characters(7, 9)  # 1-based: chars 7..15 == "Anthropic"
        out["span_text"] = str(span.Text)
        acts = span.ActionSettings(_PP_MOUSE_CLICK)
        out["action_before"] = int(acts.Action)
        acts.Hyperlink.Address = _URL
        out["address_after"] = str(acts.Hyperlink.Address)
        out["address_ok"] = str(acts.Hyperlink.Address) == _URL
        out["action_after"] = int(acts.Action)
        out["action_became_hyperlink"] = int(acts.Action) == _PP_ACTION_HYPERLINK
        out["runs_after_link"] = _walk_runs(tr)
        # Remove just this span's link.
        span2 = tr.Characters(7, 9)
        span2.ActionSettings(_PP_MOUSE_CLICK).Hyperlink.Delete()
        out["action_after_delete"] = int(tr.Characters(7, 9).ActionSettings(_PP_MOUSE_CLICK).Action)
        out["runs_after_delete"] = _walk_runs(tr)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_span_slide_jump(deck: Any, shapes: Any) -> dict[str, Any]:
    """Link the word 'later' (offset 24, length 5) to an in-deck slide jump."""
    out: dict[str, Any] = {}
    try:
        target = deck.slides[1].com
        target_id = int(target.SlideID)
        target_index = int(target.SlideIndex)
        try:
            target_title = str(target.Shapes.Title.TextFrame.TextRange.Text)
        except Exception:
            target_title = ""
        sh = shapes.add_textbox(_TEXT, left=72.0, top=200.0, width=400.0, height=60.0)
        tr = sh.com.TextFrame.TextRange
        span = tr.Characters(25, 5)  # "later"
        out["span_text"] = str(span.Text)
        candidate = f"{target_id},{target_index},{target_title}"
        out["subaddress_set"] = candidate
        acts = span.ActionSettings(_PP_MOUSE_CLICK)
        acts.Hyperlink.SubAddress = candidate
        out["subaddress_readback"] = str(acts.Hyperlink.SubAddress)
        out["action_after"] = int(acts.Action)
        out["runs"] = _walk_runs(tr)
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
            with deck.edit("run-link spike: span url / slide-jump / remove"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["span_link"] = probe_span_link(shapes)
                findings["span_slide_jump"] = probe_span_slide_jump(deck, shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("run-link spike: cleanup"):
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

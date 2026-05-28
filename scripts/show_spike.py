"""Spike — verify the v0.6 live slide-show wrappers against real PowerPoint.

Exercises `deck.show` end to end on a *running* PowerPoint with a deck open:

    uv run python scripts/show_spike.py

It records `state()` before/while/after a show, drives `start`/`goto`/`next`/
`previous`/`black`/`white`/`resume`/`end`, and — the one genuine unknown — probes
**what happens to an edit while a show is running** (the cross-cutting "widen
`_BUSY_HRESULTS` with the slide-show-running rejection" item): it tries a
`set_text` mid-show and reports the resulting exception type + HRESULT so the busy
classification can be hardened from real data.

Prints one JSON findings object. Net-zero and polite: it adds/deletes nothing, the
slide-show window is always `Exit()`ed in a `finally`, any text it touches for the
busy probe is captured first and restored, and the viewed slide + Selection are
restored at the end. A slide show *does* take over the screen while it runs (that
is the feature) — it is ended within the same run.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import Shape, _selection
from pptlive.exceptions import PptliveError


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _first_text_shape(deck: pl.Presentation, slide_index: int) -> Shape | None:
    for sh in deck.slides[slide_index].shapes:
        if isinstance(sh, Shape) and sh.has_text_frame:
            return sh
    return None


def probe_show(deck: pl.Presentation) -> dict[str, Any]:
    out: dict[str, Any] = {}
    count = len(deck.slides)
    out["before"] = deck.show.state()

    out["started"] = deck.show.start()
    if count >= 2:
        out["next"] = deck.show.next()
        out["previous"] = deck.show.previous()
        out["goto_last"] = deck.show.goto(count)
        out["goto_first"] = deck.show.goto(1)

    try:
        out["black"] = deck.show.black()
        out["white"] = deck.show.white()
        out["resume"] = deck.show.resume()
    except Exception as exc:
        out["blank_error"] = _err(exc)

    return out


def probe_busy_edit(deck: pl.Presentation) -> dict[str, Any]:
    """While a show runs, try to edit — capture how PowerPoint rejects it."""
    box = _first_text_shape(deck, 1)
    if box is None:
        return {"skipped": "no text-bearing shape on slide 1"}
    original = box.text
    result: dict[str, Any] = {"anchor": box.anchor_id, "original_len": len(original)}
    try:
        with deck.edit("show spike: busy edit probe"):
            box.set_text(original + " (probe)")
        result["edit_succeeded_during_show"] = True
        # undo our change politely (we're not testing undo here, just net-zero)
        with deck.edit("show spike: busy edit restore"):
            box.set_text(original)
    except PptliveError as exc:
        result["edit_rejected"] = {
            "type": type(exc).__name__,
            "hresult": getattr(exc, "hresult", None),
            "message": str(exc)[:200],
        }
    except Exception as exc:  # noqa: BLE001 — we want to see *whatever* it is
        result["edit_raised_other"] = _err(exc)
    return result


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        try:
            findings["show"] = probe_show(deck)
            findings["busy_edit"] = probe_busy_edit(deck)
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            try:
                findings["ended"] = deck.show.end()
            except Exception as exc:
                findings["end_error"] = _err(exc)
            try:
                _selection.restore(ppt, snap)
            except Exception as exc:
                findings["restore_error"] = _err(exc)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

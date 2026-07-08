"""Spike — which COM ops reject with a busy/RPC HRESULT (widen ``_BUSY_HRESULTS``)?

The cross-cutting item (IMPLEMENTATION.md §"Cross-cutting"): *"widen ``_BUSY_HRESULTS``
as real ``com_error``s show up in smoke runs (add the slide-show-running rejection)."*

The 2026-05-28 show spike found a **text edit succeeds mid-show** (no rejection),
so ``PowerPointBusyError`` no longer claims "a show blocks edits." But that spike
only probed a ``TextRange.Text`` set. This spike probes the ops it *didn't*:
**structural** mutations (add / delete / duplicate / move slide, add shape, set
layout) while a slide show is running — the plausible remaining source of a
slide-show-running rejection HRESULT. For anything that raises, it decodes the
HRESULT + SCODE + description via the shipped ``exceptions._decode_com_error`` and
reports whether it's already in ``_BUSY_HRESULTS``.

**This spike takes over the screen** (it starts a real slide show), exactly like
``scripts/show_spike.py``. It ends the show and deletes any temp slides in a
``finally`` (net-zero).

    uv run python scripts/busy_hresult_spike.py

Prints one JSON findings object.
"""

from __future__ import annotations

import json
from collections.abc import Callable
from typing import Any

import pptlive as pl
from pptlive import _selection
from pptlive.exceptions import _BUSY_HRESULTS, _decode_com_error


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _classify_com(exc: BaseException) -> dict[str, Any]:
    """Decode a (possibly COM) exception into the busy-taxonomy view."""
    rec: dict[str, Any] = {"type": type(exc).__name__, "str": str(exc)[:300]}
    try:
        hresult, description, message = _decode_com_error(exc)
    except Exception:  # noqa: BLE001
        hresult, description, message = None, None, str(exc)
    rec["hresult"] = None if hresult is None else f"0x{hresult & 0xFFFFFFFF:08X}"
    rec["hresult_int"] = hresult
    rec["description"] = description
    rec["message"] = message[:300]
    rec["already_in_BUSY_HRESULTS"] = hresult is not None and hresult in _BUSY_HRESULTS
    return rec


def attempt(label: str, fn: Callable[[], Any]) -> dict[str, Any]:
    """Run ``fn``; report success or a decoded failure. Never raises."""
    rec: dict[str, Any] = {"op": label}
    try:
        result = fn()
        rec["outcome"] = "succeeded"
        rec["result"] = str(result)[:120]
    except BaseException as exc:  # noqa: BLE001 — we want to classify everything
        rec["outcome"] = "raised"
        rec["error"] = _classify_com(exc)
    return rec


def main() -> int:
    findings: dict[str, Any] = {}
    findings["BUSY_HRESULTS_before"] = sorted(f"0x{h & 0xFFFFFFFF:08X}" for h in _BUSY_HRESULTS)

    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)
        temp_ids: list[int] = []
        results: list[dict[str, Any]] = []

        show_started = False
        try:
            # Seed one temp slide *before* the show so delete/duplicate/move have a
            # safe target that isn't the user's content.
            with deck.edit("busy spike: seed"):
                seed = deck.slides.add(layout="title_and_content")
                temp_ids.append(seed.id)
                seed_idx = seed.index

            # --- Baseline: the same ops with NO show running (should all succeed). ---
            findings["baseline_no_show"] = attempt(
                "add_slide (no show)",
                lambda: deck.slides.add(layout="blank"),
            )
            # Track whatever that added for cleanup.
            try:
                temp_ids.append(deck.slides[len(deck.slides)].id)
            except Exception:
                pass

            # --- Start the show; then probe structural ops mid-show. ---
            state = deck.show.start()
            show_started = bool(state.get("running"))
            findings["show_state"] = state

            raw_deck = deck.com

            # Each op is attempted on the RAW COM object (bypassing pptlive's own
            # edit()/politeness) so we see PowerPoint's unfiltered rejection, if any.
            results.append(
                attempt(
                    "AddSlide mid-show",
                    lambda: raw_deck.Slides.AddSlide(
                        len(deck.slides) + 1,
                        raw_deck.SlideMaster.CustomLayouts(1),
                    ),
                )
            )
            results.append(
                attempt(
                    "Slide.Duplicate mid-show",
                    lambda: raw_deck.Slides(seed_idx).Duplicate(),
                )
            )
            results.append(
                attempt(
                    "Slide.Delete mid-show (last slide)",
                    lambda: raw_deck.Slides(len(deck.slides)).Delete(),
                )
            )
            results.append(
                attempt(
                    "Slide.MoveTo mid-show",
                    lambda: raw_deck.Slides(len(deck.slides)).MoveTo(1),
                )
            )
            results.append(
                attempt(
                    "AddShape mid-show",
                    lambda: raw_deck.Slides(seed_idx).Shapes.AddShape(1, 40, 40, 80, 80),
                )
            )
            results.append(
                attempt(
                    "set CustomLayout mid-show",
                    lambda: setattr(
                        raw_deck.Slides(seed_idx),
                        "CustomLayout",
                        raw_deck.SlideMaster.CustomLayouts(2),
                    ),
                )
            )
            results.append(
                attempt(
                    "Presentation.Save mid-show",
                    lambda: raw_deck.Save(),
                )
            )
            # A control-flow op that legitimately competes with the running show:
            results.append(
                attempt(
                    "second SlideShowSettings.Run mid-show",
                    lambda: raw_deck.SlideShowSettings.Run(),
                )
            )
            findings["mid_show_ops"] = results
        except Exception as exc:  # noqa: BLE001
            findings["fatal"] = _err(exc)
        finally:
            # 1. End the show FIRST, so cleanup edits aren't racing the presentation.
            try:
                if show_started:
                    deck.show.end()
            except Exception as exc:  # noqa: BLE001
                findings["end_show_error"] = _err(exc)

            # 2. Delete every temp slide by stable id (tail-first).
            deleted: list[int] = []
            try:
                with deck.edit("busy spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[idx].id
                        except Exception:
                            continue
                        if sid in temp_ids and sid not in deleted:
                            deck.slides[idx].delete()
                            deleted.append(sid)
                    # Anything the mid-show AddSlide/Duplicate added won't be in
                    # temp_ids; delete down to the pre-spike count as a backstop.
                    while len(deck.slides) > count_before:
                        deck.slides[len(deck.slides)].delete()
            except Exception as exc:  # noqa: BLE001
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    # Verdict: did any mid-show op raise, and was it a *new* HRESULT?
    raised = [
        r
        for r in findings.get("mid_show_ops", [])
        if isinstance(r, dict) and r.get("outcome") == "raised"
    ]
    new_hresults = sorted(
        {
            r["error"]["hresult"]
            for r in raised
            if r.get("error", {}).get("hresult") and not r["error"].get("already_in_BUSY_HRESULTS")
        }
    )
    if not raised:
        findings["VERDICT"] = (
            "No structural op rejected mid-show — nothing to add to _BUSY_HRESULTS "
            "(confirms the show doesn't block edits, extended to structural ops)."
        )
    elif new_hresults:
        findings["VERDICT"] = f"NEW busy HRESULTs seen mid-show: {new_hresults}"
    else:
        findings["VERDICT"] = (
            "Ops raised mid-show but every HRESULT is already classified "
            "(or had no HRESULT) — review the raised list."
        )

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Interactive probe for the one open spike question: does PowerPoint's
`Application.StartNewUndoEntry` coalesce a block of edits into a SINGLE undo
entry (one Ctrl-Z), or does each COM mutation stay its own entry?

PowerPoint exposes no programmatic `Undo()`, so this can only be settled by a
human pressing Ctrl-Z and observing. Usage:

    uv run python scripts/undo_test.py group      # 1x StartNewUndoEntry, then 2 edits
    uv run python scripts/undo_test.py baseline    # 2 edits, NO StartNewUndoEntry (control)
    uv run python scripts/undo_test.py split       # 2 edits, StartNewUndoEntry, then a 3rd
    uv run python scripts/undo_test.py xproc1      # one fenced edit() (run as process 1)
    uv run python scripts/undo_test.py xproc2      # one fenced edit() (run as process 2)
    uv run python scripts/undo_test.py restore     # put the original text back

The edit phases stash the slide's original title+body text in a state file so
`restore` can undo the experiment no matter what the Ctrl-Z did. The active slide
is brought into view so the edits are visible.

Findings (2026-05-26): `group` and `baseline` both revert both edits with one
Ctrl-Z (PowerPoint groups in-session edits by default); `split`'s 1st Ctrl-Z
reverts only edit 3 and the 2nd reverts edits 1+2 (so StartNewUndoEntry is a
boundary); `xproc1`+`xproc2` run as separate processes → 1 Ctrl-Z reverts only
process 2's edit (separate invocations stay distinct undo entries). Conclusion:
`edit()` fences a block with StartNewUndoEntry → one Ctrl-Z per block, and
separate CLI invocations are isolated.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import pptlive as pl

STATE = Path(__file__).with_name(".undo_test_state.json")
TARGET_SLIDE = 3
TITLE_MARK = "UNDO TEST -- title edit (1 of 2)"
BODY_MARK = "UNDO TEST -- body edit (2 of 2)"

# `split` phase: 2 edits, then StartNewUndoEntry, then a 3rd edit.
SPLIT_E1 = "EDIT 1 of 3 (title) -- before StartNewUndoEntry"
SPLIT_E2 = "EDIT 2 of 3 (body) -- before StartNewUndoEntry"
SPLIT_E3 = "EDIT 3 of 3 (title) -- AFTER StartNewUndoEntry"

# `xproc1`/`xproc2` phases: one fenced edit() each, run as SEPARATE processes,
# to test whether two CLI-style invocations stay distinct undo entries.
XPROC_TITLE = "XPROC edit A (title) -- process 1"
XPROC_BODY = "XPROC edit B (body) -- process 2"

_MARKERS = ("EDIT ", "UNDO TEST", "XPROC ")


def _goto(ppt: pl.PowerPoint, index: int) -> None:
    try:
        ppt.com.ActiveWindow.View.GotoSlide(index)
    except Exception:
        pass


def do_edits(use_undo_entry: bool) -> dict[str, Any]:
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        slide = deck.slides[TARGET_SLIDE]
        title = slide.placeholder("title")
        body = slide.placeholder("body")

        # Capture originals so `restore` can always recover them.
        originals = {
            "deck": deck.name,
            "slide": TARGET_SLIDE,
            "title": title.text,
            "body": body.text,
        }
        STATE.write_text(json.dumps(originals), encoding="utf-8")

        _goto(ppt, TARGET_SLIDE)

        if use_undo_entry:
            ppt.com.StartNewUndoEntry()  # the primitive under test

        # Two distinct edits, bare (no edit() scope), back to back.
        title.set_text(TITLE_MARK)
        body.set_text(BODY_MARK)

        return {
            "phase": "group" if use_undo_entry else "baseline",
            "called_StartNewUndoEntry": use_undo_entry,
            "edited_slide": TARGET_SLIDE,
            "title_now": title.text,
            "body_now": body.text,
            "originals_saved_to": str(STATE),
        }


def do_split_edits() -> dict[str, Any]:
    """2 edits, then StartNewUndoEntry, then a 3rd edit — tests the boundary."""
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        slide = deck.slides[TARGET_SLIDE]
        title = slide.placeholder("title")
        body = slide.placeholder("body")

        orig_title, orig_body = title.text, body.text
        # Never capture our own markers as "originals" (would corrupt restore).
        if orig_title.startswith(_MARKERS) or orig_body.startswith(_MARKERS):
            return {
                "error": "slide still shows test markers; run `restore` or Ctrl-Z first",
                "title_now": orig_title,
                "body_now": orig_body,
            }
        STATE.write_text(
            json.dumps(
                {"deck": deck.name, "slide": TARGET_SLIDE, "title": orig_title, "body": orig_body}
            ),
            encoding="utf-8",
        )

        _goto(ppt, TARGET_SLIDE)

        title.set_text(SPLIT_E1)  # edit 1  ┐ expected to share one undo entry
        body.set_text(SPLIT_E2)  # edit 2   ┘ (default in-session grouping)
        ppt.com.StartNewUndoEntry()  # forced boundary
        title.set_text(SPLIT_E3)  # edit 3  → expected to be its own undo entry

        return {
            "phase": "split",
            "edited_slide": TARGET_SLIDE,
            "title_now": title.text,
            "body_now": body.text,
            "expectation": (
                "1st Ctrl-Z reverts title to EDIT 1 (only edit 3); "
                "2nd Ctrl-Z reverts edits 1+2 to originals"
            ),
        }


def do_xproc(process: int) -> dict[str, Any]:
    """One fenced `edit()` edit, mimicking a single `pptlive write` invocation.

    `xproc1` (process 1) captures originals + edits the title; `xproc2`
    (process 2) edits the body and leaves the state file alone. Run as two
    separate processes, then press Ctrl-Z once: if only the body reverts, the
    two CLI-style invocations are distinct undo entries (the `edit()` fence
    isolates them across processes).
    """
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        slide = deck.slides[TARGET_SLIDE]
        title = slide.placeholder("title")
        body = slide.placeholder("body")

        if process == 1:
            orig_title, orig_body = title.text, body.text
            if orig_title.startswith(_MARKERS) or orig_body.startswith(_MARKERS):
                return {
                    "error": "slide still shows test markers; run `restore` or Ctrl-Z first",
                    "title_now": orig_title,
                    "body_now": orig_body,
                }
            STATE.write_text(
                json.dumps(
                    {
                        "deck": deck.name,
                        "slide": TARGET_SLIDE,
                        "title": orig_title,
                        "body": orig_body,
                    }
                ),
                encoding="utf-8",
            )
            _goto(ppt, TARGET_SLIDE)
            with deck.edit("xproc process 1"):  # fences StartNewUndoEntry
                title.set_text(XPROC_TITLE)
        else:
            with deck.edit("xproc process 2"):  # separate process, separate fence
                body.set_text(XPROC_BODY)

        return {
            "phase": f"xproc{process}",
            "edited_slide": TARGET_SLIDE,
            "title_now": title.text,
            "body_now": body.text,
        }


def restore() -> dict[str, Any]:
    if not STATE.exists():
        return {"restored": False, "reason": "no state file; nothing to restore"}
    saved = json.loads(STATE.read_text(encoding="utf-8"))
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        slide = deck.slides[int(saved["slide"])]
        with deck.edit("restore undo-test slide"):
            slide.placeholder("title").set_text(saved["title"])
            slide.placeholder("body").set_text(saved["body"])
    STATE.unlink(missing_ok=True)
    return {"restored": True, "slide": saved["slide"]}


def main(argv: list[str]) -> int:
    phase = argv[1] if len(argv) > 1 else "group"
    if phase == "group":
        out = do_edits(use_undo_entry=True)
    elif phase == "baseline":
        out = do_edits(use_undo_entry=False)
    elif phase == "split":
        out = do_split_edits()
    elif phase == "xproc1":
        out = do_xproc(1)
    elif phase == "xproc2":
        out = do_xproc(2)
    elif phase == "restore":
        out = restore()
    else:
        out = {"error": f"unknown phase {phase!r}; use group|baseline|split|xproc1|xproc2|restore"}
    print(json.dumps(out, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

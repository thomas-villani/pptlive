"""Spike — characterize PowerPoint ``TextRange.Find`` / ``.Replace`` for v1.0.

The roadmap's v1.0 find/replace build rests on two COM behaviours the docs are
vague about. This exploratory spike nails them down on a live deck:

  1. **Empty-match sentinel** — what does ``TextRange.Find`` return on no match?
     ``None`` (VBA ``Nothing``), an empty ``TextRange``, or does it raise? The
     traversal loop's stop condition depends on the answer.
  2. **``.Replace`` semantics** — does one call replace the *first* occurrence or
     *all*? What does it return? And the offset-drift hazard: if the replacement
     re-contains the search text, does a naive replace-until-None loop spin
     forever (so we must advance ``After``)?
  3. **Reach** — does ``.Find`` see text in table cells and notes (separate
     ``TextRange``s our traversal would visit), confirming the slide×shape×frame
     walk is the right shape.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/findreplace_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a
single temporary slide appended then deleted in a ``finally``, with the viewed
slide + Selection restored. ``net_zero_ok`` confirms the slide count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _describe(r: Any) -> dict[str, Any] | None:
    """Describe a returned TextRange (or None for VBA Nothing)."""
    if r is None:
        return None
    try:
        return {"start": r.Start, "length": r.Length, "text": r.Text, "bool": bool(r)}
    except Exception as exc:
        return {"error": _err(exc)}


def probe(deck: pl.Presentation, slide: pl.Slide) -> dict[str, Any]:
    out: dict[str, Any] = {}

    sh = slide.shapes.add_textbox("alpha beta alpha gamma alpha")
    tf = sh.com.TextFrame

    # 1. empty-match sentinel ------------------------------------------------
    try:
        miss = tf.TextRange.Find("zzqq_absent")
        out["find_miss"] = {"is_none": miss is None, "desc": _describe(miss)}
    except Exception as exc:
        out["find_miss"] = {"raised": _err(exc)}

    # 2. a successful single find --------------------------------------------
    try:
        out["find_hit_beta"] = _describe(tf.TextRange.Find("beta"))
    except Exception as exc:
        out["find_hit_beta"] = {"raised": _err(exc)}

    # 3. iterate every 'alpha' via the After parameter -----------------------
    #    Question: is After 0- or 1-based, and does Start+Length advance cleanly?
    try:
        positions: list[int] = []
        cur = tf.TextRange.Find("alpha")
        guard = 0
        while cur is not None and guard < 12:
            positions.append(cur.Start)
            cur = tf.TextRange.Find("alpha", cur.Start + cur.Length)
            guard += 1
        out["find_iter_alpha"] = {"positions": positions, "n": len(positions)}
    except Exception as exc:
        out["find_iter_alpha"] = {"raised": _err(exc)}

    # 4. .Replace return value + first-vs-all on ONE call --------------------
    try:
        sh.com.TextFrame.TextRange.Text = "alpha beta alpha gamma alpha"
        with deck.edit("fr spike: one replace"):
            r1 = sh.com.TextFrame.TextRange.Replace("alpha", "X")
        out["replace_one_call"] = {
            "returned": _describe(r1),
            "text_after": sh.com.TextFrame.TextRange.Text,
        }
    except Exception as exc:
        out["replace_one_call"] = {"raised": _err(exc)}

    # 5. replace-all loop (re-fetch TextRange each iter; stop on None) -------
    try:
        sh.com.TextFrame.TextRange.Text = "alpha beta alpha gamma alpha"
        iters = 0
        with deck.edit("fr spike: replace all"):
            r = sh.com.TextFrame.TextRange.Replace("alpha", "X")
            while r is not None and iters < 20:
                iters += 1
                r = sh.com.TextFrame.TextRange.Replace("alpha", "X")
        out["replace_loop"] = {
            "iterations": iters,
            "text_after": sh.com.TextFrame.TextRange.Text,
        }
    except Exception as exc:
        out["replace_loop"] = {"raised": _err(exc)}

    # 6. offset-drift hazard: replacement RE-CONTAINS the search text -------
    #    Does a second from-scratch Replace re-match inside the replacement?
    try:
        sh.com.TextFrame.TextRange.Text = "alpha"
        with deck.edit("fr spike: drift"):
            first = sh.com.TextFrame.TextRange.Replace("alpha", "alpha_X")
            second = sh.com.TextFrame.TextRange.Replace("alpha", "alpha_X")
        out["replace_drift"] = {
            "first_returned": _describe(first),
            "second_returned": _describe(second),  # not None => naive loop would spin
            "text_after": sh.com.TextFrame.TextRange.Text,
            "naive_loop_would_spin": second is not None,
        }
    except Exception as exc:
        out["replace_drift"] = {"raised": _err(exc)}

    # 7. reach: does Find see table-cell text and notes text? ---------------
    try:
        with deck.edit("fr spike: table"):
            tbl_sh = slide.shapes.add_table(2, 2)
            tbl_sh.table.cell(1, 1).set_text("needle_in_cell")
        cell_tr = tbl_sh.com.Table.Cell(1, 1).Shape.TextFrame.TextRange
        out["reach_table_cell"] = _describe(cell_tr.Find("needle"))
    except Exception as exc:
        out["reach_table_cell"] = {"raised": _err(exc)}

    try:
        with deck.edit("fr spike: notes"):
            slide.notes.set_text("needle_in_notes")
        notes_tr = slide.notes.com  # the notes-body TextRange
        out["reach_notes"] = _describe(notes_tr.Find("needle"))
    except Exception as exc:
        out["reach_notes"] = {"raised": _err(exc)}

    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        temp_ids: list[int] = []
        try:
            with deck.edit("fr spike: build"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["probe"] = probe(deck, deck.slides[sidx])
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("fr spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[idx].id
                        except Exception:
                            continue
                        if sid in temp_ids and sid not in deleted:
                            deck.slides[idx].delete()
                            deleted.append(sid)
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Spike — does a live ``TextRange.Text`` read ever contain a bare ``\\n``?

The carried-over cross-cutting item (IMPLEMENTATION.md §"Cross-cutting"): confirm
whether a raw ``TextRange.Text`` **read** ever yields a bare ``\\n`` (0x0A) and, if
so, whether COM ``.Paragraphs()`` treats it as a paragraph break. This decides
whether ``_selection.read_selection``'s ``\\r``-only paragraph count can *undercount*.

``_selection.read_selection`` computes the caret's paragraph as::

    paragraph = full[: start - 1].count("\\r") + 1

If the frame text can contain bare ``\\n`` breaks that COM counts as separate
paragraphs, that count is wrong (it would land the ``para:S:N:P`` anchor on the
wrong paragraph). pptlive's *write* path already normalises ``\\n``/``\\r\\n`` →
``\\r`` (``_anchors.normalize_paragraph_breaks``) and folds soft breaks to ``\\v``,
so writes are safe — this spike is entirely about **reads** of text that may have
been set by the user, a paste, or a raw ``.com`` write that bypassed us.

It sets ``TextRange.Text`` **raw** (bypassing pptlive normalisation) with each
newline spelling — ``\\r`` (CR), ``\\n`` (LF), ``\\r\\n`` (CRLF), ``\\v`` (VT / soft
break) — then reads back and reports, for each:

  * the exact control chars present in the read-back (which of CR/LF/VT appear),
  * ``TextRange.Paragraphs().Count`` and each paragraph's ``.Text`` repr,
  * ``\\r``-count + 1 (what ``read_selection`` would compute) vs the real
    paragraph count — a mismatch is the undercount bug.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/newline_reads_spike.py

Prints one JSON findings object; net-zero (probes live on a temp slide deleted in
a ``finally``).
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

CR = "\r"
LF = "\n"
CRLF = "\r\n"
VT = "\v"  # vertical tab — PowerPoint's soft (Shift+Enter) line break


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _char_names(text: str) -> dict[str, int]:
    """Count the control chars that matter for paragraph splitting."""
    return {
        "CR (\\r 0x0D)": text.count(CR),
        "LF (\\n 0x0A)": text.count(LF),
        "VT (\\v 0x0B)": text.count(VT),
        "codepoints": [hex(ord(c)) for c in text if ord(c) < 0x20],
    }


def _read_selection_paragraph_count(full: str) -> int:
    """Reproduce what ``_selection.read_selection`` counts for a caret at end."""
    # read_selection uses full[: start-1].count("\r") + 1; at end of frame the
    # whole text is scanned, so the count equals total paragraphs *as it sees them*.
    return full.count(CR) + 1


def probe_spelling(tf: Any, label: str, raw: str) -> dict[str, Any]:
    """Set ``TextRange.Text`` raw to ``raw`` and report the read-back structure."""
    rec: dict[str, Any] = {"set_repr": repr(raw)}
    tr = tf.TextRange
    rec["set_error"] = None
    try:
        tr.Text = raw
    except Exception as exc:  # noqa: BLE001
        rec["set_error"] = _err(exc)
        return rec

    # Re-read fresh off the live frame.
    try:
        back = str(tf.TextRange.Text)
    except Exception as exc:  # noqa: BLE001
        rec["read_error"] = _err(exc)
        return rec

    rec["read_repr"] = repr(back)
    rec["read_chars"] = _char_names(back)

    # Real paragraph structure per COM.
    try:
        paras = tf.TextRange.Paragraphs()
        pcount = int(paras.Count)
        rec["paragraphs_count"] = pcount
        rec["paragraphs_text"] = [repr(str(paras.Item(i + 1).Text)) for i in range(pcount)]
    except Exception as exc:  # noqa: BLE001
        rec["paragraphs_error"] = _err(exc)
        pcount = None

    # The read_selection heuristic vs. the truth.
    heuristic = _read_selection_paragraph_count(back)
    rec["read_selection_would_count"] = heuristic
    if pcount is not None:
        rec["MATCHES_COM"] = heuristic == pcount
        rec["undercount_by"] = pcount - heuristic

    return rec


def probe_insert_soft_break(tf: Any) -> dict[str, Any]:
    """What char does an *inserted* soft line break read back as (\\v vs \\n)?"""
    rec: dict[str, Any] = {}
    tr = tf.TextRange
    try:
        tr.Text = "one"
        # InsertAfter a VT — PowerPoint's soft break — then more text.
        tr.InsertAfter(VT + "two")
        back = str(tf.TextRange.Text)
        rec["read_repr"] = repr(back)
        rec["read_chars"] = _char_names(back)
        rec["paragraphs_count"] = int(tf.TextRange.Paragraphs().Count)
    except Exception as exc:  # noqa: BLE001
        rec["error"] = _err(exc)
    return rec


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        temp_ids: list[int] = []
        try:
            with deck.edit("newline-reads spike: build"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
                # A plain text box — no placeholder autofit/inheritance in the way.
                tb = deck.slides[sidx].shapes.add_textbox(
                    "seed", left=72, top=72, width=360, height=180
                )
                shp = tb.com

            tf = shp.TextFrame
            findings["probes"] = {
                "CR (paragraph break)": probe_spelling(tf, "CR", "A" + CR + "B" + CR + "C"),
                "LF (raw \\n)": probe_spelling(tf, "LF", "A" + LF + "B" + LF + "C"),
                "CRLF": probe_spelling(tf, "CRLF", "A" + CRLF + "B" + CRLF + "C"),
                "VT (soft break)": probe_spelling(tf, "VT", "A" + VT + "B"),
                "mixed CR+LF+VT": probe_spelling(tf, "mixed", "A" + CR + "B" + LF + "C" + VT + "D"),
            }
            findings["inserted_soft_break"] = probe_insert_soft_break(tf)
        except Exception as exc:  # noqa: BLE001
            findings["fatal"] = _err(exc)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("newline-reads spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[idx].id
                        except Exception:
                            continue
                        if sid in temp_ids and sid not in deleted:
                            deck.slides[idx].delete()
                            deleted.append(sid)
            except Exception as exc:  # noqa: BLE001
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    # A one-line verdict for the reader.
    verdict = "inconclusive"
    probes = findings.get("probes", {})
    if isinstance(probes, dict) and probes:
        any_lf = any(
            isinstance(p, dict) and p.get("read_chars", {}).get("LF (\\n 0x0A)", 0)
            for p in probes.values()
        )
        any_undercount = any(
            isinstance(p, dict) and (p.get("undercount_by") or 0) != 0 for p in probes.values()
        )
        if not any_lf and not any_undercount:
            verdict = "SAFE: no bare \\n survives a read; \\r-only paragraph count matches COM"
        elif any_lf and not any_undercount:
            verdict = (
                "\\n survives reads but does NOT add a paragraph (soft break) — count still OK"
            )
        else:
            verdict = "UNDERCOUNT RISK: a read separator is a paragraph break \\r-count misses"
    findings["VERDICT"] = verdict

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

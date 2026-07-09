"""Spike — does PowerPoint's COM expose spelling errors? (linter §9 / P6)

The linter spec defers `proofing` (typos) because pptlive has no `proofing()`
surface and it's unconfirmed whether PowerPoint's object model exposes spelling
the way Word's `Range.SpellingErrors` does. This exploratory spike settles it on a
live deck, sorting the outcome into three tiers:

  * **A — Word-style span API.** `TextRange.SpellingErrors` (or a TextFrame2
    equivalent) yields the misspelled sub-ranges directly. Ideal: exact offsets,
    so a rule can anchor + fix per span.
  * **B — token-level checker.** `Application.CheckSpelling(word)` returns a bool
    and `GetSpellingSuggestions(word)` returns candidates. Viable fallback: we
    tokenize each text frame ourselves and check word-by-word.
  * **C — neither.** Proofing needs an external spell engine; stays out of scope.

**FINDING (2026-07-08, live).** PowerPoint's own COM is verdict **C** — no
`Application.CheckSpelling`/`GetSpellingSuggestions`, no `TextRange.SpellingErrors`
(only `LanguageSettings` on the app and `TextRange.LanguageID` exist). But a
**hidden `Word.Application` borrowed over COM** is a fully working token checker:
`CheckSpelling(word)` returns the right bool and `GetSpellingSuggestions(word)`
returns candidates, both on **bare strings with no document open**, and Word (unlike
PowerPoint) runs `Visible=False`. Since we tokenize each frame ourselves (regex →
`(start, length)`), we get **exact spans + suggestions** — tier-A anchoring on a
tier-B primitive. The real path forward, gated only on Word being installed.

**Setup the user provides:** a running PowerPoint whose *first slide* has a
deliberately misspelled word in a textbox and in the heading. The spike reads
those live shapes, prints their text, and demonstrates whatever tier is available
against the real misspelling.

    uv run python scripts/proofing_spike.py

Prints one JSON findings object. Pure read — mutates nothing, adds/deletes no
slides (inherently net-zero); it snapshots + restores the viewed slide/Selection
only as a courtesy since a couple of probes may touch focus.
"""

from __future__ import annotations

import json
import re
from typing import Any

import pptlive as pl
from pptlive import _selection

# Candidate members to probe by name (getattr — late-bound COM has no useful dir()).
_APP_MEMBERS = (
    "CheckSpelling",
    "GetSpellingSuggestions",
    "CheckGrammar",
    "SpellingChecked",
    "Language",
    "LanguageSettings",
)
_TEXTRANGE_MEMBERS = (
    "SpellingErrors",
    "GrammarErrors",
    "LanguageID",
    "SpellCheck",
    "NoProofing",
)


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _probe_members(obj: Any, names: tuple[str, ...]) -> dict[str, Any]:
    """For each candidate name: present? and a short repr (without calling)."""
    out: dict[str, Any] = {}
    for name in names:
        try:
            val = getattr(obj, name)
            out[name] = {"present": True, "repr": repr(val)[:120]}
        except Exception as exc:
            out[name] = {"present": False, "err": _err(exc)}
    return out


def _slide1_text_shapes(slide: pl.Slide) -> list[dict[str, Any]]:
    """Every text-bearing shape on the slide + its raw text and TextRange handle."""
    shapes: list[dict[str, Any]] = []
    for i, shape in enumerate(slide.shapes, start=1):
        raw = shape.com
        try:
            if not raw.HasTextFrame:
                continue
            tf = raw.TextFrame
            if not tf.HasText:
                continue
            tr = tf.TextRange
            shapes.append(
                {
                    "zorder": i,
                    "name": raw.Name,
                    "text": tr.Text,
                    "_tr": tr,  # stripped before JSON
                }
            )
        except Exception as exc:
            shapes.append({"zorder": i, "err": _err(exc)})
    return shapes


def probe_tier_a(shape_entries: list[dict[str, Any]]) -> dict[str, Any]:
    """Tier A: does a TextRange expose the misspelled sub-ranges directly?"""
    out: dict[str, Any] = {"members_on_first_textrange": None, "spellingerrors_call": None}
    tr = next((s["_tr"] for s in shape_entries if "_tr" in s), None)
    if tr is None:
        out["note"] = "no text-bearing shape found on slide 1"
        return out
    out["members_on_first_textrange"] = _probe_members(tr, _TEXTRANGE_MEMBERS)

    # If SpellingErrors exists, try to enumerate it (Word returns a ProofreadingErrors
    # collection with .Count and 1-based items, each a sub-Range with .Text/.Start).
    try:
        errs = tr.SpellingErrors
        try:
            count = int(errs.Count)
        except Exception as exc:
            out["spellingerrors_call"] = {"present": True, "count_err": _err(exc)}
            return out
        items = []
        for i in range(1, count + 1):
            try:
                e = errs.Item(i)
                items.append({"text": e.Text, "start": getattr(e, "Start", None)})
            except Exception as exc:
                items.append({"err": _err(exc)})
        out["spellingerrors_call"] = {"present": True, "count": count, "items": items}
    except Exception as exc:
        out["spellingerrors_call"] = {"present": False, "err": _err(exc)}
    return out


def probe_tier_b(app: Any, shape_entries: list[dict[str, Any]]) -> dict[str, Any]:
    """Tier B: token-level `CheckSpelling` / `GetSpellingSuggestions`."""
    out: dict[str, Any] = {}

    # Known-good / known-bad sanity, so we can tell a working checker from one that
    # rubber-stamps everything true.
    for probe_word in ("hello", "recieve", "teh", "PowerPoint"):
        try:
            out.setdefault("check_spelling", {})[probe_word] = bool(app.CheckSpelling(probe_word))
        except Exception as exc:
            out.setdefault("check_spelling", {})[probe_word] = {"err": _err(exc)}

    try:
        sugg = app.GetSpellingSuggestions("recieve")
        try:
            out["suggestions_recieve"] = [sugg.Item(i).Name for i in range(1, int(sugg.Count) + 1)][
                :8
            ]
        except Exception as exc:
            out["suggestions_recieve"] = {"present": True, "iter_err": _err(exc)}
    except Exception as exc:
        out["suggestions_recieve"] = {"present": False, "err": _err(exc)}

    # Demonstrate the token-level rule end-to-end against slide 1's real words: for
    # each text frame, split into word tokens and flag those CheckSpelling rejects.
    if "check_spelling" in out and all(isinstance(v, bool) for v in out["check_spelling"].values()):
        flagged: list[dict[str, Any]] = []
        for s in shape_entries:
            text = s.get("text")
            if not text:
                continue
            for m in re.finditer(r"[A-Za-z][A-Za-z'’]*", text):
                token = m.group(0)
                try:
                    if not app.CheckSpelling(token):
                        flagged.append(
                            {
                                "shape": s.get("name"),
                                "word": token,
                                "start": m.start(),
                                "length": len(token),
                            }
                        )
                except Exception:
                    pass
        out["live_tokenized_flags"] = flagged
    return out


def probe_tier_b_word_borrow(shape_entries: list[dict[str, Any]]) -> dict[str, Any]:
    """Tier B via a *borrowed* checker: a hidden Word.Application. PowerPoint's own
    COM has no spelling; Word's does, runs invisibly, and checks bare strings with no
    document open. We tokenize each frame ourselves, so we recover exact spans."""
    import win32com.client  # spike-only raw COM (pptlive keeps pywin32 in _com.py)

    out: dict[str, Any] = {}
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        try:
            word.Visible = False
        except Exception as exc:
            out["visible_err"] = _err(exc)
        # A scratch document gives GetSpellingSuggestions its required context (it
        # raises "no document is open" otherwise; CheckSpelling works without one).
        doc = word.Documents.Add()

        # Sanity: a working checker rejects known-bad and passes known-good.
        for probe_word in ("hello", "recieve", "teh", "PowerPoint"):
            try:
                out.setdefault("check_spelling", {})[probe_word] = bool(
                    word.CheckSpelling(probe_word)
                )
            except Exception as exc:
                out.setdefault("check_spelling", {})[probe_word] = {"err": _err(exc)}

        # End-to-end against slide 1's real words, recovering per-typo spans.
        flagged: list[dict[str, Any]] = []
        for s in shape_entries:
            text = s.get("text")
            if not text:
                continue
            for m in re.finditer(r"[A-Za-z][A-Za-z'’]*", text):
                token = m.group(0)
                try:
                    if not word.CheckSpelling(token):
                        sugg = word.GetSpellingSuggestions(token)
                        cands = [sugg.Item(i).Name for i in range(1, int(sugg.Count) + 1)][:5]
                        flagged.append(
                            {
                                "shape": s.get("name"),
                                "word": token,
                                "start": m.start(),
                                "length": len(token),
                                "suggestions": cands,
                            }
                        )
                except Exception as exc:
                    flagged.append({"word": token, "err": _err(exc)})
        out["live_flags_with_spans"] = flagged
    except Exception as exc:
        out["dispatch_err"] = _err(exc)
    finally:
        if doc is not None:
            try:
                doc.Close(0)  # wdDoNotSaveChanges
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
    b = out.get("check_spelling", {})
    out["works"] = b.get("hello") is True and b.get("recieve") is False
    return out


def main() -> None:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        snap = _selection.snapshot(ppt)
        try:
            slide = deck.slides[1]
            entries = _slide1_text_shapes(slide)
            findings["slide1_shapes"] = [
                {k: v for k, v in e.items() if k != "_tr"} for e in entries
            ]

            app = ppt.com  # the raw Application
            findings["app_members"] = _probe_members(app, _APP_MEMBERS)
            findings["tier_a_spans"] = probe_tier_a(entries)
            findings["tier_b_native"] = probe_tier_b(app, entries)
            findings["tier_b_word_borrow"] = probe_tier_b_word_borrow(entries)

            # Verdict.
            a = findings["tier_a_spans"].get("spellingerrors_call") or {}
            b_native = findings["tier_b_native"].get("check_spelling") or {}
            b_native_works = (
                isinstance(b_native, dict)
                and b_native.get("recieve") is False
                and b_native.get("hello") is True
            )
            if a.get("present") and "count" in a:
                findings["verdict"] = "A — TextRange.SpellingErrors works (exact spans)"
            elif b_native_works:
                findings["verdict"] = "B(native) — PowerPoint CheckSpelling works"
            elif findings["tier_b_word_borrow"].get("works"):
                findings["verdict"] = (
                    "B(borrow) — PowerPoint COM has no spelling, but a hidden "
                    "Word.Application does; we tokenize → exact spans + suggestions"
                )
            else:
                findings["verdict"] = "C — no usable spelling COM; needs external engine"
        finally:
            try:
                _selection.restore(ppt, snap)
            except Exception as exc:
                findings["restore_err"] = _err(exc)

    print(json.dumps(findings, indent=2, default=str))


if __name__ == "__main__":
    main()

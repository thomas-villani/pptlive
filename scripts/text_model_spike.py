"""Spike — pin PowerPoint's text-model sharp edges before hardening v1.6.

The gpt-5.4 review (``docs/reviews/gpt-5.4-review.md``) flagged three places where
PowerPoint's text model leaks through pptlive's abstraction. This spike pins the
exact COM behaviour for each, on a *temporary* slide (net-zero), so the v1.6 build
can harden against verified facts rather than guesses:

  1. **LineRule semantics — the ``line_spacing: 24`` footgun.** ``ParagraphFormat``
     stores spacing as a bare number whose *unit* is selected by a companion bool:
     ``LineRuleWithin`` pairs with ``SpaceWithin``, ``LineRuleBefore`` with
     ``SpaceBefore``, ``LineRuleAfter`` with ``SpaceAfter``. Confirm
     ``msoTrue`` ⇒ the value is a **multiple/lines**, ``msoFalse`` ⇒ **points**,
     and that all three ``LineRule*`` flags *read back* (so we can report the mode).

  2. **The reset primitive.** Once a placeholder is in a bad state (5 pt font, giant
     spacing), can we clear *direct* formatting so text re-inherits the layout/master
     defaults? Probe, in order: (a) does re-setting ``TextRange.Text`` drop run
     overrides? (b) what does nulling each override do? (c) can we read the matching
     ``CustomLayout`` placeholder's geometry + font (the ``reset_to_layout`` source)?

  3. **Autofit / text-frame reads.** Confirm ``TextFrame.AutoSize`` /
     ``.WordWrap`` / the four ``Margin*`` and a shrink-to-fit signal
     (``Font.AutoScale`` / ``TextFrame2 ...``) are *readable* for the diagnostics
     read op, and decide the overflow-risk heuristic.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/text_model_spike.py

Prints one JSON findings object; net-zero (all probes live on a temp slide deleted
in a ``finally``).
"""

from __future__ import annotations

import json
from collections.abc import Callable
from typing import Any

import pptlive as pl
from pptlive import _selection

MSO_TRUE = -1
MSO_FALSE = 0


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _get(obj: Any, name: str) -> Any:
    try:
        return getattr(obj, name)
    except Exception as exc:
        return f"<{_err(exc)}>"


def _try(fn: Callable[[], Any]) -> Any:
    try:
        return fn()
    except Exception as exc:
        return f"<{_err(exc)}>"


def probe_linerule(tf: Any) -> dict[str, Any]:
    """Set each Space* with its LineRule both ways; report the read-back."""
    out: dict[str, Any] = {}
    tr = tf.TextRange
    para = tr.Paragraphs(1, 1)
    pf = para.ParagraphFormat

    # Within (line spacing): multiple vs points.
    for mode, rule in (("multiple", MSO_TRUE), ("points", MSO_FALSE)):
        rec: dict[str, Any] = {}
        rec["set_error"] = _try(
            lambda rule=rule: (
                setattr(pf, "LineRuleWithin", rule),
                setattr(pf, "SpaceWithin", 24.0),
                None,
            )[-1]
        )
        rec["read_LineRuleWithin"] = _get(pf, "LineRuleWithin")
        rec["read_SpaceWithin"] = _get(pf, "SpaceWithin")
        out[f"within_{mode}"] = rec

    # Before / After: confirm the same pairing + read-back.
    for prop, rule_attr in (
        ("SpaceBefore", "LineRuleBefore"),
        ("SpaceAfter", "LineRuleAfter"),
    ):
        rec = {}
        rec["set_points_error"] = _try(
            lambda prop=prop, rule_attr=rule_attr: (
                setattr(pf, rule_attr, MSO_FALSE),
                setattr(pf, prop, 18.0),
                None,
            )[-1]
        )
        rec[f"read_{rule_attr}"] = _get(pf, rule_attr)
        rec[f"read_{prop}"] = _get(pf, prop)
        out[prop] = rec

    return out


def probe_reset(shp: Any) -> dict[str, Any]:
    """Can we clear direct run/paragraph formatting back to inherited defaults?"""
    out: dict[str, Any] = {}
    tf = shp.TextFrame
    tr = tf.TextRange
    font = tr.Font

    # Capture the inherited-ish baseline first (before we vandalise it).
    out["baseline_size"] = _get(font, "Size")

    # Vandalise: tiny bold font + huge multiple spacing (the reviewer's bad state).
    out["vandalise_error"] = _try(
        lambda: (
            setattr(font, "Size", 5.0),
            setattr(font, "Bold", MSO_TRUE),
            setattr(tr.ParagraphFormat, "LineRuleWithin", MSO_TRUE),
            setattr(tr.ParagraphFormat, "SpaceWithin", 4.0),
            None,
        )[-1]
    )
    out["after_vandalise"] = {
        "size": _get(font, "Size"),
        "bold": _get(font, "Bold"),
        "space_within": _get(tr.ParagraphFormat, "SpaceWithin"),
    }

    # (a) Does re-setting .Text drop the run overrides? (expected: no)
    txt = str(_get(tr, "Text") or "Reset probe text")
    out["resettext_error"] = _try(lambda: setattr(tr, "Text", txt))
    out["after_resettext"] = {
        "size": _get(shp.TextFrame.TextRange.Font, "Size"),
        "bold": _get(shp.TextFrame.TextRange.Font, "Bold"),
    }

    # (c) reset_to_layout source: the matching CustomLayout placeholder.
    out["custom_layout"] = probe_custom_layout(shp)
    return out


def probe_custom_layout(shp: Any) -> dict[str, Any]:
    """Read the CustomLayout placeholder matching our shape (geometry + font)."""
    out: dict[str, Any] = {}
    layout = _try(lambda: shp.Parent.CustomLayout)
    out["layout_name"] = _get(layout, "Name")
    want_type = _try(lambda: int(shp.PlaceholderFormat.Type)) if shp is not None else None
    out["shape_placeholder_type"] = want_type
    phs = _try(lambda: layout.Shapes.Placeholders)
    rows: list[dict[str, Any]] = []
    cnt = _try(lambda: int(phs.Count))
    if isinstance(cnt, int):
        for i in range(1, cnt + 1):
            ph = phs.Item(i)
            rows.append(
                {
                    "name": _get(ph, "Name"),
                    "type": _try(lambda ph=ph: int(ph.PlaceholderFormat.Type)),
                    "left": _get(ph, "Left"),
                    "top": _get(ph, "Top"),
                    "width": _get(ph, "Width"),
                    "height": _get(ph, "Height"),
                    "font_size": _try(lambda ph=ph: ph.TextFrame.TextRange.Font.Size),
                }
            )
    out["layout_placeholders"] = rows
    return out


def probe_autofit(shp: Any) -> dict[str, Any]:
    """Are the text-frame container props readable for the diagnostics read op?"""
    tf = shp.TextFrame
    out: dict[str, Any] = {
        "AutoSize": _get(tf, "AutoSize"),
        "WordWrap": _get(tf, "WordWrap"),
        "MarginLeft": _get(tf, "MarginLeft"),
        "MarginRight": _get(tf, "MarginRight"),
        "MarginTop": _get(tf, "MarginTop"),
        "MarginBottom": _get(tf, "MarginBottom"),
        "HasText": _get(tf, "HasText"),
    }
    font = _try(lambda: tf.TextRange.Font)
    out["Font.AutoScale"] = _get(font, "AutoScale")
    # TextFrame2 (newer) carries the shrink-to-fit knobs when present.
    tf2 = _try(lambda: shp.TextFrame2)
    out["has_TextFrame2"] = not (isinstance(tf2, str) and tf2.startswith("<"))
    if out["has_TextFrame2"]:
        out["TextFrame2.AutoSize"] = _get(tf2, "AutoSize")
        out["TextFrame2.WordWrap"] = _get(tf2, "WordWrap")
        out["TextFrame2.HasText"] = _get(tf2, "HasText")
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
            with deck.edit("text-model spike: build"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
                body = deck.anchor_by_id(f"ph:{sidx}:body")
                body.set_text("Alpha line\nBeta line\nGamma line")
                shp = body.com

            findings["linerule"] = probe_linerule(shp.TextFrame)
            findings["autofit"] = probe_autofit(shp)
            findings["reset"] = probe_reset(shp)
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("text-model spike: cleanup"):
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

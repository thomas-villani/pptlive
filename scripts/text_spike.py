"""Spike harness — verify the v0.3 text-structure wrappers against real PowerPoint.

The exploratory pass (git history) probed raw COM to design the wrappers; this is
the regression check that the shipped surface behaves on a live deck. It exercises
`Shape.paragraphs` / `paragraph(p)`, `Paragraph.set_text` (break-preserving),
`insert_paragraph_before/after` (incl. the last-paragraph end-detection),
`Paragraph.delete`, `format_text`, `format_paragraph` (alignment/spacing/indent),
and `apply_list`/`remove_list`. Run against a *running* PowerPoint with a deck
open:

    uv run python scripts/text_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a single
temporary slide that is appended and then deleted in a `finally`, and the viewed
slide is restored. `net_zero_ok` confirms the deck's slide count is unchanged.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection
from pptlive.constants import color_hex

_SEED = "Intro\rDemo\rQ&A"


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def probe(box: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}

    # addressing + reads ---------------------------------------------------
    box.set_text(_SEED)
    out["read"] = {
        "count": len(box.paragraphs),
        "para2_text": box.paragraph(2).text,
        "list_texts": [r["text"] for r in box.paragraphs.list()],
        "ok": len(box.paragraphs) == 3 and box.paragraph(2).text == "Demo",
    }

    # set_text preserves siblings -----------------------------------------
    box.set_text(_SEED)
    box.paragraph(2).set_text("DEMO")
    out["set_text"] = {"text": box.text, "ok": box.text == "Intro\rDEMO\rQ&A"}

    # insert after a middle paragraph -------------------------------------
    box.set_text(_SEED)
    box.paragraph(2).insert_paragraph_after("New")
    out["insert_after_middle"] = {"text": box.text, "ok": box.text == "Intro\rDemo\rNew\rQ&A"}

    # insert after the LAST paragraph (end-detection) ----------------------
    box.set_text(_SEED)
    last_raw = repr(box.com.TextFrame.TextRange.Paragraphs(3, 1).Text)
    box.paragraph(3).insert_paragraph_after("End")
    out["insert_after_last"] = {
        "last_para_raw": last_raw,
        "text": box.text,
        "ok": box.text == "Intro\rDemo\rQ&A\rEnd",
    }

    # insert before the first paragraph -----------------------------------
    box.set_text(_SEED)
    box.paragraph(1).insert_paragraph_before("Top")
    out["insert_before"] = {"text": box.text, "ok": box.text == "Top\rIntro\rDemo\rQ&A"}

    # append via the whole-shape anchor -----------------------------------
    box.set_text(_SEED)
    box.insert_paragraph_after("Tail")
    out["append_via_shape"] = {"text": box.text, "ok": box.text == "Intro\rDemo\rQ&A\rTail"}

    # delete a paragraph ---------------------------------------------------
    box.set_text(_SEED)
    box.paragraph(2).delete()
    out["delete"] = {
        "text": box.text,
        "count": len(box.paragraphs),
        "ok": box.text == "Intro\rQ&A" and len(box.paragraphs) == 2,
    }

    # font formatting ------------------------------------------------------
    box.set_text(_SEED)
    box.paragraph(1).format_text(bold=True, size=24.0, color="#FF0000")
    f = box.com.TextFrame.TextRange.Paragraphs(1, 1).Font
    out["format_text"] = {
        "bold": int(f.Bold),
        "size": float(f.Size),
        "color": color_hex(f.Color.RGB),
        "ok": int(f.Bold) == -1 and float(f.Size) == 24.0 and color_hex(f.Color.RGB) == "#FF0000",
    }

    # paragraph formatting -------------------------------------------------
    box.set_text(_SEED)
    box.paragraph(1).format_paragraph(
        alignment="center", space_before=12.0, line_spacing=1.5, indent_level=2
    )
    p1 = box.com.TextFrame.TextRange.Paragraphs(1, 1)
    out["format_paragraph"] = {
        "alignment": int(p1.ParagraphFormat.Alignment),
        "space_before": float(p1.ParagraphFormat.SpaceBefore),
        "line_spacing": float(p1.ParagraphFormat.SpaceWithin),
        "indent_level": int(p1.IndentLevel),
        "ok": int(p1.ParagraphFormat.Alignment) == 2 and int(p1.IndentLevel) == 2,
    }

    # bullets / lists ------------------------------------------------------
    box.set_text(_SEED)
    box.apply_list("numbered")
    b = box.com.TextFrame.TextRange.ParagraphFormat.Bullet
    numbered = {"visible": int(b.Visible), "type": int(b.Type)}
    box.apply_list("bulleted", character="•")
    bulleted = {"type": int(b.Type), "character": int(b.Character)}
    box.remove_list()
    out["list"] = {
        "numbered": numbered,
        "bulleted": bulleted,
        "removed_visible": int(b.Visible),
        "ok": numbered["type"] == 2 and bulleted["type"] == 1 and int(b.Visible) == 0,
    }
    return out


def probe_body_indent(deck: pl.Presentation, slide_index: int) -> dict[str, Any]:
    """Indent levels on a real body placeholder (its native outline behavior)."""
    out: dict[str, Any] = {}
    try:
        body = deck.slides[slide_index].placeholder("body")
        body.set_text("A\rB\rC")
        body.paragraph(2).format_paragraph(indent_level=3)
        com = body.com.TextFrame.TextRange.Paragraphs(2, 1)
        out = {"indent_level": int(com.IndentLevel), "ok": int(com.IndentLevel) == 3}
    except Exception as exc:
        out["error"] = _err(exc)
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
            with deck.edit("text spike: verify v0.3 wrappers"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
                box = deck.slides[sidx].shapes.add_textbox(
                    _SEED, left=72, top=72, width=480, height=300
                )
                findings["text"] = probe(box)
                findings["body_indent"] = probe_body_indent(deck, sidx)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("text spike: cleanup"):
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

        findings["net_zero_ok"] = len(deck.slides) == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

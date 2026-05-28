"""Spike — verify the v0.4 render + selection wrappers against real PowerPoint.

The exploratory pass (git history) probed raw COM to design the surface; this is
the regression check that the shipped wrappers behave on a live deck. It exercises
`Slide.export_image` (default/native + requested + aspect-fill dims, that it
captures an unsaved edit, and that it's polite), `Presentation.selection()` for a
shape and a text selection, and `anchor_by_id("here:")`. Run against a *running*
PowerPoint with a deck open:

    uv run python scripts/render_select_spike.py

Prints one JSON findings object. Net-zero and polite: all work happens on a single
temporary slide that is appended then deleted in a `finally`, exported PNGs go to a
TemporaryDirectory, and the viewed slide + Selection are restored. `net_zero_ok`
confirms the deck's slide count is unchanged.
"""

from __future__ import annotations

import json
import os
import struct
import tempfile
from typing import Any

import pptlive as pl
from pptlive import Paragraph, Shape, _selection

_SEED = "Intro\rDemo\rQ&A"
_PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _png_dims(path: str) -> tuple[int, int] | None:
    try:
        with open(path, "rb") as fh:
            head = fh.read(24)
        if len(head) < 24 or head[:8] != _PNG_SIG:
            return None
        w, h = struct.unpack(">II", head[16:24])
        return int(w), int(h)
    except Exception:
        return None


def probe_render(deck: pl.Presentation, slide: pl.Slide, box: Any, tmpdir: str) -> dict[str, Any]:
    out: dict[str, Any] = {}
    win = deck._ppt.com.ActiveWindow  # noqa: SLF001  (spike-only introspection)

    # native dims --------------------------------------------------------------
    try:
        p = slide.export_image(os.path.join(tmpdir, "native.png"))
        out["native"] = {
            "path_abs": p.is_absolute(),
            "dims": _png_dims(str(p)),
            "exists": p.is_file(),
        }
    except Exception as exc:
        out["native"] = {"error": _err(exc)}

    # requested + aspect-fill --------------------------------------------------
    try:
        p = slide.export_image(os.path.join(tmpdir, "sized.png"), width=640, height=480)
        sized = _png_dims(str(p))
        p2 = slide.export_image(os.path.join(tmpdir, "fill.png"), width=1920)
        out["sized"] = {"sized": sized, "aspect_fill_width_only": _png_dims(str(p2))}
    except Exception as exc:
        out["sized"] = {"error": _err(exc)}

    # captures an unsaved live edit -------------------------------------------
    try:
        box.set_text("BEFORE")
        b1 = str(slide.export_image(os.path.join(tmpdir, "b1.png"), width=480, height=270))
        box.set_text("AFTER is different")
        b2 = str(slide.export_image(os.path.join(tmpdir, "b2.png"), width=480, height=270))
        out["captures_live_edit"] = {"differs": os.path.getsize(b1) != os.path.getsize(b2)}
        box.set_text(_SEED)
    except Exception as exc:
        out["captures_live_edit"] = {"error": _err(exc)}

    # temp default + polite ----------------------------------------------------
    try:
        before = (int(win.View.Slide.SlideIndex), int(win.Selection.Type))
        tmp = slide.export_image()  # no path -> temp file
        after = (int(win.View.Slide.SlideIndex), int(win.Selection.Type))
        out["polite_temp"] = {
            "temp_exists": tmp.is_file(),
            "temp_suffix": tmp.suffix,
            "view_selection_unchanged": before == after,
        }
        os.remove(tmp)
    except Exception as exc:
        out["polite_temp"] = {"error": _err(exc)}

    return out


def probe_selection(deck: pl.Presentation, slide: pl.Slide, box: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    win = deck._ppt.com.ActiveWindow  # noqa: SLF001
    box.set_text(_SEED)
    try:
        win.View.GotoSlide(slide.index)
    except Exception as exc:
        out["goto_error"] = _err(exc)

    # shape selection ----------------------------------------------------------
    try:
        slide.com.Shapes.Range([box.name]).Select()
        info = deck.selection()
        here = deck.anchor_by_id("here:")
        out["shapes"] = {
            "type": info.type,
            "anchor_id": info.anchor_id,
            "first_shape": info.shapes[0] if info.shapes else None,
            "here_is_shape": isinstance(here, Shape),
            "here_anchor": here.anchor_id,
        }
    except Exception as exc:
        out["shapes"] = {"error": _err(exc)}

    # text selection (paragraph 2) --------------------------------------------
    try:
        box.com.TextFrame.TextRange.Paragraphs(2, 1).Select()
        info = deck.selection()
        here = deck.anchor_by_id("here:")
        out["text"] = {
            "type": info.type,
            "paragraph": info.paragraph,
            "text": info.text,
            "anchor_id": info.anchor_id,
            "here_is_paragraph": isinstance(here, Paragraph),
            "here_text": here.text,
            "ok": info.type == "text" and info.paragraph == 2 and here.text == "Demo",
        }
    except Exception as exc:
        out["text"] = {"error": _err(exc)}

    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        temp_ids: list[int] = []
        with tempfile.TemporaryDirectory(prefix="pptlive_v04_") as tmpdir:
            try:
                with deck.edit("render/selection spike: build"):
                    temp = deck.slides.add(layout="title_and_content")
                    temp_ids.append(temp.id)
                    sidx = temp.index
                    box = deck.slides[sidx].shapes.add_textbox(
                        _SEED, left=72, top=72, width=480, height=300
                    )
                findings["render"] = probe_render(deck, deck.slides[sidx], box, tmpdir)
                findings["selection"] = probe_selection(deck, deck.slides[sidx], box)
            finally:
                deleted: list[int] = []
                try:
                    with deck.edit("render/selection spike: cleanup"):
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

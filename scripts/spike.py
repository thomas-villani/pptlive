"""Spike harness — verify the four real-COM assumptions pptlive's v0 encodes.

Run against a *running* PowerPoint with at least one presentation open
(ideally a couple of slides with title/body placeholders and some speaker
notes on at least one slide):

    uv run python scripts/spike.py

Prints one JSON findings object to stdout. It is deliberately polite:

  * snapshots the viewed slide + Selection up front and restores them at the
    end (the probes move both);
  * reads the original `Application.Visible` before touching it and restores
    that exact value;
  * runs the most disruptive probe (Visible) *last*.

This is a one-off verification tool, not part of the shipped package — it pokes
the raw COM object directly (via the `.com` escape hatch) on purpose, to test
things pptlive deliberately does *not* wrap.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection
from pptlive.constants import MsoTriState, PpPlaceholderType, is_true


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _ph_type_name(value: int) -> str:
    try:
        return PpPlaceholderType(value).name
    except ValueError:
        return f"<{value}>"


# ---------------------------------------------------------------------------
# Spike #1 — UndoRecord absence
# ---------------------------------------------------------------------------


def probe_undo_record(app: Any) -> dict[str, Any]:
    """Does any undo-grouping primitive exist on this build's Application?"""
    out: dict[str, Any] = {}
    for name in ("UndoRecord", "StartNewUndoEntry"):
        try:
            val = getattr(app, name)
            out[name] = {"present": True, "repr": repr(val)[:120]}
        except Exception as exc:
            out[name] = {"present": False, "error": _err(exc)}
    return out


# ---------------------------------------------------------------------------
# Spike #4 — notes placeholder resolves by Type == BODY (not a fixed index)
# ---------------------------------------------------------------------------


def probe_notes_placeholder(deck: pl.Presentation) -> dict[str, Any]:
    """Enumerate every notes-page placeholder; locate the body by Type == BODY."""
    rows: list[dict[str, Any]] = []
    for slide in deck.slides:
        row: dict[str, Any] = {"slide": slide.index}
        try:
            placeholders = slide.com.NotesPage.Shapes.Placeholders
            phs: list[dict[str, Any]] = []
            body_index: int | None = None
            for i in range(1, int(placeholders.Count) + 1):
                ph = placeholders.Item(i)
                t = int(ph.PlaceholderFormat.Type)
                has_tf = is_true(ph.HasTextFrame)
                phs.append(
                    {"index": i, "type": t, "type_name": _ph_type_name(t), "has_text_frame": has_tf}
                )
                if t == int(PpPlaceholderType.BODY) and has_tf and body_index is None:
                    body_index = i
            row["placeholders"] = phs
            row["body_index"] = body_index
            row["body_found_by_type"] = body_index is not None
        except Exception as exc:
            row["error"] = _err(exc)
        rows.append(row)
    return {"slides": rows}


# ---------------------------------------------------------------------------
# Spike #3 — shape-range Selection round-trips by name
# ---------------------------------------------------------------------------


def probe_selection_roundtrip(ppt: pl.PowerPoint) -> dict[str, Any]:
    """Select the active slide's first shape by name; read the Selection back."""
    out: dict[str, Any] = {}
    try:
        win = ppt.com.ActiveWindow
        shapes = win.View.Slide.Shapes
        if int(shapes.Count) == 0:
            return {"skipped": "active slide has no shapes"}
        name = str(shapes.Item(1).Name)
        out["target_name"] = name
        shapes.Range([name]).Select()
        sel = win.Selection
        out["selection_type_after"] = int(sel.Type)
        got = [str(sh.Name) for sh in sel.ShapeRange]
        out["selected_names_after"] = got
        out["roundtrip_ok"] = got == [name]
    except Exception as exc:
        out["error"] = _err(exc)
    return out


# ---------------------------------------------------------------------------
# Spike #2 — Application.Visible = False raises (run LAST; restore immediately)
# ---------------------------------------------------------------------------


def probe_visible(app: Any) -> dict[str, Any]:
    """Try to hide the app; record whether it raised; restore the original value."""
    out: dict[str, Any] = {}
    original: int | None = None
    try:
        original = int(app.Visible)
        out["original"] = original
    except Exception as exc:
        out["read_error"] = _err(exc)

    try:
        app.Visible = int(MsoTriState.FALSE)
        out["set_false_raised"] = False
        try:
            out["value_after_set_false"] = int(app.Visible)
        except Exception as exc:
            out["readback_error"] = _err(exc)
    except Exception as exc:
        out["set_false_raised"] = True
        out["error"] = _err(exc)

    # Always try to put it back exactly as we found it (default to visible).
    restore_to = (
        original if original not in (None, int(MsoTriState.FALSE)) else int(MsoTriState.TRUE)
    )
    try:
        app.Visible = restore_to
        out["restored_to"] = restore_to
    except Exception as exc:
        out["restore_error"] = _err(exc)
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        findings["attach_ok"] = True
        app = ppt.com

        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        findings["slide_count"] = len(deck.slides)

        # Snapshot the user's view + selection before we disturb anything.
        snap = _selection.snapshot(ppt)
        findings["snapshot"] = {
            "slide_index": snap.slide_index,
            "selection_type": snap.selection_type,
            "shape_names": list(snap.shape_names),
        }

        # Confirm the real read paths work end-to-end (v0 integration smoke).
        findings["library_reads"] = {
            "slides": deck.slides.list(),
            "outline": deck.outline(),
            "page_setup": deck.page_setup(),
        }

        # The four spike probes — read-only ones first, Visible last.
        findings["spike"] = {
            "undo_record": probe_undo_record(app),
            "notes_placeholder": probe_notes_placeholder(deck),
            "selection_roundtrip": probe_selection_roundtrip(ppt),
        }

        # Restore selection/view before the disruptive Visible probe.
        _selection.restore(ppt, snap)

        findings["spike"]["visible"] = probe_visible(app)

        # Final restore in case the Visible flip disturbed the window.
        _selection.restore(ppt, snap)

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

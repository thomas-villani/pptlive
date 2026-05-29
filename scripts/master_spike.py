"""Spike — drive the shipped v0.9 master/theme wrappers against real PowerPoint.

The regression pass for `deck.theme` / `deck.master` (the analog of the chart /
SmartArt regression spikes): it exercises the *shipped* high-level wrappers
end-to-end on a live deck and confirms each write round-trips. Unlike the
SmartArt spike, master/theme edits touch the deck's one shared master — there's
no temp slide to throw away — so net-zero is achieved by **reading every value
first and restoring it** in a `finally`. The viewed slide + Selection are
restored too.

It also resolves the three open questions the v0.9 design left for its spike:

  (a) **multi-master reality** — reports `Presentation.Designs.Count` and that a
      primary `SlideMaster` exists (v0.9 targets the primary master; this says
      how often a deck carries more than one).
  (b) **undo granularity** — master edits are wrapped in `deck.edit()`, so they
      carry the same `StartNewUndoEntry` fence as content edits; whether one
      Ctrl-Z reverts the block is an interactive check (see `undo_test.py`) and
      is *not* asserted here — recorded as a note.
  (c) **background fill round-trip** — only when the master's background is
      already a plain solid (so the write is fully reversible); otherwise the
      destructive solid write is **skipped** to stay net-zero, and that's
      reported.

Run against a *running* PowerPoint with any deck open:

    uv run python scripts/master_spike.py

Prints one JSON findings object. Net-zero + polite.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        snap = _selection.snapshot(ppt)

        # multi-master reality (open question a)
        try:
            findings["designs_count"] = int(deck.com.Designs.Count)
        except Exception as exc:
            findings["designs_count_error"] = _err(exc)
        findings["has_slide_master"] = bool(deck.com.SlideMaster is not None)
        findings["undo_note"] = (
            "master edits are wrapped in deck.edit() (StartNewUndoEntry fence) like "
            "content edits; one-Ctrl-Z reversion is an interactive check, not asserted here"
        )

        # snapshot originals via the wrappers themselves (so we restore exactly)
        orig_theme = deck.theme.read()
        orig_master = deck.master.read()
        orig_accent1 = orig_theme["colors"]["accent1"]
        orig_major = orig_theme["fonts"]["major"]
        orig_body_l1 = orig_master["text_styles"]["body"]["levels"][0]
        orig_bg = orig_master["background"]
        bg_is_solid = orig_bg.get("type") == "solid" and orig_bg.get("color")

        try:
            with deck.edit("master spike: write"):
                deck.theme.set_color("accent1", "#C00000")
                deck.theme.set_font("major", "Georgia")
                deck.master.format_text_style("body", 1, font="Georgia", size=40.0)
                if bg_is_solid:
                    deck.master.set_background("#102030")

            after_theme = deck.theme.read()
            after_master = deck.master.read()
            after_body_l1 = after_master["text_styles"]["body"]["levels"][0]

            findings["theme_color_round_trip"] = {
                "wrote": "#C00000",
                "read": after_theme["colors"]["accent1"],
                "ok": after_theme["colors"]["accent1"] == "#C00000",
            }
            findings["theme_font_round_trip"] = {
                "wrote": "Georgia",
                "read": after_theme["fonts"]["major"],
                "ok": after_theme["fonts"]["major"] == "Georgia",
            }
            findings["text_style_round_trip"] = {
                "wrote": {"font": "Georgia", "size": 40.0},
                "read": {"font": after_body_l1.get("font"), "size": after_body_l1.get("size")},
                "ok": after_body_l1.get("font") == "Georgia" and after_body_l1.get("size") == 40.0,
            }
            if bg_is_solid:
                findings["background_round_trip"] = {
                    "wrote": "#102030",
                    "read": after_master["background"].get("color"),
                    "type": after_master["background"].get("type"),
                    "ok": after_master["background"].get("color") == "#102030",
                }
            else:
                findings["background_skipped"] = (
                    f"master background is not a plain solid (type={orig_bg.get('type')!r}); "
                    "skipped the destructive solid write to stay net-zero"
                )
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            try:
                with deck.edit("master spike: restore"):
                    deck.theme.set_color("accent1", orig_accent1)
                    if orig_major:
                        deck.theme.set_font("major", orig_major)
                    restore_kwargs: dict[str, Any] = {}
                    if orig_body_l1.get("font"):
                        restore_kwargs["font"] = orig_body_l1["font"]
                    if orig_body_l1.get("size") is not None:
                        restore_kwargs["size"] = orig_body_l1["size"]
                    if restore_kwargs:
                        deck.master.format_text_style("body", 1, **restore_kwargs)
                    if bg_is_solid:
                        deck.master.set_background(orig_bg["color"])
            except Exception as exc:
                findings["restore_error"] = _err(exc)
            try:
                _selection.restore(ppt, snap)
            except Exception as exc:
                findings["restore_selection_error"] = _err(exc)

        # net-zero check: everything back to the captured originals
        final_theme = deck.theme.read()
        final_master = deck.master.read()
        final_body_l1 = final_master["text_styles"]["body"]["levels"][0]
        findings["net_zero_ok"] = bool(
            final_theme["colors"]["accent1"] == orig_accent1
            and final_theme["fonts"]["major"] == orig_major
            and final_body_l1.get("font") == orig_body_l1.get("font")
            and final_body_l1.get("size") == orig_body_l1.get("size")
            and (not bg_is_solid or final_master["background"].get("color") == orig_bg["color"])
        )

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Spike harness for the v0.6.0 batch-2 candidates — pin three open questions
against real PowerPoint before hardening any of them. Net-zero / polite: every
mutation rides temp slides/shapes that are torn down, and the view + Selection are
restored.

1. **Sections** (`Presentation.SectionProperties`) — which add primitive exists
   (`AddSection` vs `AddBeforeSlide`), the index base, what `FirstSlide` /
   `SlidesCount` / `Name` read back, and the `Delete` signature (does it take a
   delete-slides bool?). Probed over two appended temp slides so the user's real
   sections are never touched.
2. **Headers / footers** (`Slide.HeadersFooters`, `SlideMaster.HeadersFooters`) —
   what reads back, whether a slide-level footer set round-trips, and what the
   master exposes (the inheritance question).
3. **Direct-vs-inherited font color** — the open Claude Desktop friction item: is
   there ANY COM signal that distinguishes a *directly set* run color from a
   *theme-cascaded* one? Probes `ColorFormat.Type` (msoColorType) +
   `ObjectThemeColor` on both classic `Font.Color` and modern
   `TextFrame2 ... Font.Fill.ForeColor`, before and after setting an explicit RGB.

Run:  uv run python scripts/batch2_spike.py
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:200]


def _read(obj: Any, names: tuple[str, ...]) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for name in names:
        try:
            val = getattr(obj, name)
            out[name] = val if isinstance(val, (int, float, str, bool)) else type(val).__name__
        except Exception as exc:
            out[name] = _err(exc)
    return out


def _call(obj: Any, name: str, *args: Any) -> Any:
    try:
        return getattr(obj, name)(*args)
    except Exception as exc:
        return _err(exc)


# ---------------------------------------------------------------------------
# 1. Sections
# ---------------------------------------------------------------------------


def probe_sections(deck: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sp = deck.com.SectionProperties
    except Exception as exc:
        return {"error": f"no SectionProperties: {_err(exc)}"}
    out["count_before"] = _call(sp, "Count") if not callable(getattr(sp, "Count", None)) else None
    try:
        out["count_before"] = int(sp.Count)
    except Exception as exc:
        out["count_before"] = _err(exc)
    # Which add primitive marshals? Try AddBeforeSlide(slideIndex, name) first
    # (it ties a section to a concrete slide), then AddSection(index, name).
    last_slide = len(deck.slides)
    out["addbeforeslide"] = _call(sp, "AddBeforeSlide", last_slide, "SpikeSec A")
    out["addsection"] = _call(sp, "AddSection", int(sp.Count) + 1, "SpikeSec B")
    try:
        n = int(sp.Count)
        out["count_after_add"] = n
        out["per_section"] = [
            {
                "i": i,
                "Name": _call(sp, "Name", i),
                "SlidesCount": _call(sp, "SlidesCount", i),
                "FirstSlide": _call(sp, "FirstSlide", i),
            }
            for i in range(1, n + 1)
        ]
        # Rename + Move probes on the last section we added.
        out["rename"] = _call(sp, "Rename", n, "SpikeSec Renamed")
        out["name_after_rename"] = _call(sp, "Name", n)
    except Exception as exc:
        out["read_error"] = _err(exc)
    return out


def cleanup_sections(deck: Any, count_before: int) -> dict[str, Any]:
    """Delete sections back down to the original count, keeping slides."""
    out: dict[str, Any] = {}
    try:
        sp = deck.com.SectionProperties
        attempts = []
        guard = 0
        while int(sp.Count) > count_before and guard < 10:
            guard += 1
            idx = int(sp.Count)
            # Try the delete-slides=False signature first; fall back to 1-arg.
            res = _call(sp, "Delete", idx, False)
            if isinstance(res, str) and res.startswith(("TypeError", "ComError", "PptliveError")):
                res = _call(sp, "Delete", idx)
            attempts.append({"idx": idx, "res": res, "count_now": int(sp.Count)})
        out["delete_attempts"] = attempts
        out["count_final"] = int(sp.Count)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


# ---------------------------------------------------------------------------
# 2. Headers / footers
# ---------------------------------------------------------------------------


def probe_headers_footers(slide: Any, deck: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        hf = slide.com.HeadersFooters
    except Exception as exc:
        return {"error": f"slide has no HeadersFooters: {_err(exc)}"}
    out["slide_default"] = {
        "Footer": _read(hf.Footer, ("Text", "Visible")),
        "SlideNumber": _read(hf.SlideNumber, ("Visible",)),
        "DateAndTime": _read(hf.DateAndTime, ("Visible", "Text", "Format", "UseFormat")),
        "DisplayOnTitleSlide": _read(hf, ("DisplayOnTitleSlide",)),
    }
    # Set a slide-level footer + slide number, read back.
    try:
        hf.Footer.Visible = -1
        hf.Footer.Text = "Spike footer"
        hf.SlideNumber.Visible = -1
        out["slide_after_set"] = {
            "Footer": _read(hf.Footer, ("Text", "Visible")),
            "SlideNumber": _read(hf.SlideNumber, ("Visible",)),
        }
    except Exception as exc:
        out["set_error"] = _err(exc)
    # Master-level read (don't mutate the master — just see what's exposed).
    try:
        master = deck.com.SlideMaster.HeadersFooters
        out["master"] = {
            "Footer": _read(master.Footer, ("Text", "Visible")),
            "SlideNumber": _read(master.SlideNumber, ("Visible",)),
            "DateAndTime": _read(master.DateAndTime, ("Visible", "Format", "UseFormat")),
        }
    except Exception as exc:
        out["master_error"] = _err(exc)
    return out


# ---------------------------------------------------------------------------
# 3. Direct-vs-inherited font color
# ---------------------------------------------------------------------------

# msoColorType: 1=RGB, 2=Scheme, 3=CMYK, 4=CMS, 5=Ink, -2=Mixed
_COLOR_TYPE = {1: "rgb", 2: "scheme", 3: "cmyk", 4: "cms", 5: "ink", -2: "mixed"}


def _color_signal(font_color: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        out["RGB"] = int(font_color.RGB)
    except Exception as exc:
        out["RGB"] = _err(exc)
    try:
        t = int(font_color.Type)
        out["Type"] = {"raw": t, "name": _COLOR_TYPE.get(t, f"type:{t}")}
    except Exception as exc:
        out["Type"] = _err(exc)
    try:
        out["ObjectThemeColor"] = int(font_color.ObjectThemeColor)
    except Exception as exc:
        out["ObjectThemeColor"] = _err(exc)
    try:
        out["SchemeColor"] = int(font_color.SchemeColor)
    except Exception as exc:
        out["SchemeColor"] = _err(exc)
    return out


def probe_direct_vs_inherited(slide: Any) -> dict[str, Any]:
    """Add a textbox + read a body placeholder; compare the color signal on a
    freshly-inherited run vs. one with an explicit RGB set."""
    out: dict[str, Any] = {}
    shapes = slide.shapes
    tb = shapes.add_textbox("Inherited then explicit", left=40.0, top=200.0, width=300.0)
    com = tb.com
    # Classic ColorFormat (Font.Color) and modern TextFrame2 ForeColor.
    classic = com.TextFrame.TextRange.Font.Color
    try:
        modern = com.TextFrame2.TextRange.Font.Fill.ForeColor
    except Exception as exc:
        modern = None
        out["modern_error"] = _err(exc)
    out["textbox_inherited"] = {
        "classic": _color_signal(classic),
        "modern": _color_signal(modern) if modern is not None else None,
    }
    # Now set an explicit RGB and re-read — does Type flip rgb/scheme?
    try:
        classic.RGB = 0x0000FF  # red in BGR -> 0x0000FF == RGB red? RGB long is 0xBBGGRR
    except Exception as exc:
        out["set_error"] = _err(exc)
    out["textbox_explicit"] = {
        "classic": _color_signal(classic),
        "modern": _color_signal(modern) if modern is not None else None,
    }
    # A body placeholder (inherits the master/theme) for contrast, if present.
    try:
        body = None
        for sh in slide.com.Shapes:
            try:
                if int(sh.PlaceholderFormat.Type) in (2, 7) and bool(sh.HasTextFrame):
                    body = sh
                    break
            except Exception:
                continue
        if body is not None:
            out["placeholder_inherited"] = _color_signal(body.TextFrame.TextRange.Font.Color)
        else:
            out["placeholder_inherited"] = "no body placeholder on this layout"
    except Exception as exc:
        out["placeholder_error"] = _err(exc)
    return out


# ---------------------------------------------------------------------------


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        findings["slide_count_before"] = count_before
        snap = _selection.snapshot(ppt)

        sec_count_before = 0
        try:
            sec_count_before = int(deck.com.SectionProperties.Count)
        except Exception:
            pass
        findings["section_count_before"] = sec_count_before

        temp_ids: list[int] = []
        try:
            with deck.edit("batch2 spike: sections / headersfooters / color"):
                t1 = deck.slides.add(layout="title_and_content")
                t2 = deck.slides.add(layout="blank")
                temp_ids += [t1.id, t2.id]
                findings["sections"] = probe_sections(deck)
                findings["headers_footers"] = probe_headers_footers(
                    deck.slides[t1.index], deck
                )
                findings["direct_vs_inherited"] = probe_direct_vs_inherited(deck.slides[t1.index])
        finally:
            findings["sections_cleanup"] = cleanup_sections(deck, sec_count_before)
            try:
                with deck.edit("batch2 spike: cleanup slides"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            if deck.slides[idx].id in temp_ids:
                                deck.slides[idx].delete()
                        except Exception:
                            continue
            except Exception as exc:
                findings["slide_cleanup_error"] = _err(exc)
            _selection.restore(ppt, snap)

        findings["slide_count_after"] = len(deck.slides)
        findings["section_count_after"] = (
            int(deck.com.SectionProperties.Count)
            if hasattr(deck.com, "SectionProperties")
            else None
        )
        findings["net_zero_slides"] = len(deck.slides) == count_before
        findings["net_zero_sections"] = findings["section_count_after"] == sec_count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

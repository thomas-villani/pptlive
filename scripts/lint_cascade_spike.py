"""Spike — does the placeholder → CustomLayout → master `TextStyles` cascade resolve a
concrete *baseline* font for each placeholder kind, so `format_info()` can report
`override` for placeholder text?  (`spec-linter.md` §7b, the one spike the linter
design asked for before hardening.)

Questions, each answered with a value a wrong assumption would change:

1. Does a **layout** placeholder's `TextFrame.TextRange.Font` read a concrete
   `Name`/`Size`/`Bold` (i.e. does the layout fall through to the master text style
   rather than reading 0 / "" for an un-overridden placeholder)?
2. Does an **untouched** slide placeholder read the SAME values as its layout
   placeholder (→ `override=False` baseline works with no config)?
3. After a **direct** `Font.Size` write on the slide placeholder, does the layout
   placeholder stay put (→ the diff is a true override, not a shared object)?
4. Does the layout placeholder for `body` expose per-**level** values (via
   `Paragraphs(n)` on its prompt text, whose indent levels are 1..5 on a stock
   layout), matching `SlideMaster.TextStyles(ppBodyStyle).Levels(n)`?
5. What does a **free textbox** (no cascade) report — sanity for `baseline=None`.

Net-zero: works on a scratch presentation it creates and closes without saving.

    uv run python scripts/lint_cascade_spike.py
"""

from __future__ import annotations

import json

import pptlive as pl
from pptlive._anchors import font_to_dict

PP_TITLE_STYLE, PP_BODY_STYLE = 2, 3  # PpTextStyleType


def _font(tr):  # type: ignore[no-untyped-def]
    d = font_to_dict(tr)
    return {k: d[k] for k in ("font", "size", "bold", "color", "color_source")}


def main() -> int:
    out: dict[str, object] = {}
    cm = pl.connect()
    app = cm.__enter__()
    app.com.Presentations.Add()
    deck = app.presentations.active
    try:
        with deck.edit("cascade spike"):
            s1 = deck.slides.add(
                layout="title_slide", content={"ctrtitle": "Hello", "subtitle": "Sub"}
            )
            s2 = deck.slides.add(
                layout="title_and_content",
                title="Body slide",
                body=["L1", {"text": "L2", "indent_level": 2}, {"text": "L3", "indent_level": 3}],
            )
            tb = deck.slides[s2.index].shapes.add_textbox(
                "free", left=10, top=10, width=200, height=40
            )

        # Q1: layout placeholders read concrete fonts?
        layout_phs = {}
        lay = deck.slides[s2.index].com.CustomLayout
        for i in range(1, int(lay.Shapes.Placeholders.Count) + 1):
            ph = lay.Shapes.Placeholders(i)
            tr = ph.TextFrame.TextRange
            entry = {
                "type": int(ph.PlaceholderFormat.Type),
                "font": _font(tr),
                "text": str(tr.Text)[:40],
            }
            # Q4: per-level on the body layout placeholder
            if int(ph.PlaceholderFormat.Type) in (2, 7):  # body / object
                levels = []
                for p in range(1, int(tr.Paragraphs().Count) + 1):
                    pr = tr.Paragraphs(p)
                    levels.append(
                        {
                            "indent": int(pr.IndentLevel),
                            "size": float(pr.Font.Size),
                            "font": str(pr.Font.Name),
                        }
                    )
                entry["levels"] = levels
            layout_phs[str(ph.Name)] = entry
        out["q1_layout_placeholders"] = layout_phs

        # master text styles for comparison
        master = deck.com.SlideMaster
        out["master_title_l1"] = {
            "font": str(master.TextStyles(PP_TITLE_STYLE).Levels(1).Font.Name),
            "size": float(master.TextStyles(PP_TITLE_STYLE).Levels(1).Font.Size),
        }
        out["master_body_levels"] = [
            {
                "font": str(master.TextStyles(PP_BODY_STYLE).Levels(n).Font.Name),
                "size": float(master.TextStyles(PP_BODY_STYLE).Levels(n).Font.Size),
            }
            for n in range(1, 6)
        ]

        # Q2: untouched slide placeholder == layout placeholder?
        title = deck.slides[s2.index].placeholder("title")
        body = deck.slides[s2.index].placeholder("body")
        out["q2_slide_title_untouched"] = _font(title.com.TextFrame.TextRange)
        out["q2_slide_body_paras"] = [
            {"indent": int(pr.IndentLevel), "size": float(pr.Font.Size), "font": str(pr.Font.Name)}
            for pr in (body.com.TextFrame.TextRange.Paragraphs(p) for p in range(1, 4))
        ]
        # Q2b: title slide's ctrtitle/subtitle
        s1s = deck.slides[s1.index]
        out["q2_title_slide"] = {
            "ctrtitle": _font(s1s.placeholder("ctrtitle").com.TextFrame.TextRange),
            "subtitle": _font(s1s.placeholder("subtitle").com.TextFrame.TextRange),
        }
        lay1 = s1s.com.CustomLayout
        out["q2_title_slide_layout"] = {
            str(lay1.Shapes.Placeholders(i).Name): _font(
                lay1.Shapes.Placeholders(i).TextFrame.TextRange
            )
            for i in range(1, int(lay1.Shapes.Placeholders.Count) + 1)
        }

        # Q3: direct override on the slide leaves the layout alone
        with deck.edit("override"):
            title.format_text(size=30)
        out["q3_after_override"] = {
            "slide_title_size": float(title.com.TextFrame.TextRange.Font.Size),
            "layout_title_size": float(
                next(
                    lay.Shapes.Placeholders(i)
                    for i in range(1, int(lay.Shapes.Placeholders.Count) + 1)
                    if int(lay.Shapes.Placeholders(i).PlaceholderFormat.Type) == 1
                ).TextFrame.TextRange.Font.Size
            ),
        }

        # Q5: free textbox
        out["q5_free_textbox"] = _font(tb.com.TextFrame.TextRange)
        out["q5_textbox_is_placeholder"] = int(tb.com.Type) == 14
    finally:
        deck.com.Saved = True
        deck.com.Close()
    print(json.dumps(out, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

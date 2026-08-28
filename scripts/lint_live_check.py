"""Live end-to-end check of the linter against real PowerPoint (net-zero).

Builds a scratch deck with four Title-and-Content slides — one of them drifted
(title shrunk to 36 pt + nudged 15 pt down, a level-1 bullet at 20 pt, a text box
hanging off the right edge, two cards 2 pt out of alignment) — then:

1. `deck.lint()`            → expect the five default-rule findings on the drift
2. `deck.regularize()`      → applies the three fixable ones in ONE undo entry
3. `deck.lint()` again      → only the report-only `shape-off-slide` remains
4. `regularize()` again     → applies nothing (the idempotency contract)
5. `--rule alignment`       → edge-alignment fires, regularize snaps the cards,
                              a third lint of that rule is clean

Prints a JSON summary; exits 1 if any expectation fails. Closes the scratch deck
without saving.

    uv run python scripts/lint_live_check.py
"""

from __future__ import annotations

import json
import sys

import pptlive as pl


def main() -> int:
    out: dict[str, object] = {}
    ok = True

    def expect(name: str, cond: bool, detail: object = None) -> None:
        nonlocal ok
        out[name] = {"ok": cond, "detail": detail}
        ok = ok and cond

    cm = pl.connect()
    app = cm.__enter__()
    app.com.Presentations.Add()
    deck = app.presentations.active
    try:
        with deck.edit("build"):
            for i in range(4):
                deck.slides.add(
                    layout="title_and_content",
                    title=f"Slide {i + 1}",
                    body=["Point one", "Point two"],
                )
            drift = deck.slides[3]
            title = drift.placeholder("title")
            title.format_text(size=36)
            title.move(top=title.geometry()["top"] + 15)
            drift.placeholder("body").paragraph(2).format_text(size=20)
            drift.shapes.add_textbox("bleeds", left=900, top=300, width=200, height=40)
            cards = deck.slides[4]
            cards.shapes.add_textbox("A", left=100, top=300, width=150, height=60)
            cards.shapes.add_textbox("B", left=102, top=400, width=150, height=60)

        first = deck.lint()
        rules = sorted({f["rule"] for f in first})
        out["first_findings"] = [
            {k: f[k] for k in ("rule", "slide", "anchor_id", "observed", "expected")} for f in first
        ]
        expect(
            "default rules fire on the drift",
            rules
            == [
                "body-font-consistent",
                "shape-off-slide",
                "title-font-consistent",
                "title-position-consistent",
            ],
            rules,
        )
        expect("findings all on slide 3", all(f["slide"] == 3 for f in first), None)

        report = deck.regularize()
        expect(
            "regularize applied three fixes",
            sorted(f["rule"] for f in report["applied"])
            == ["body-font-consistent", "title-font-consistent", "title-position-consistent"],
            [f["rule"] for f in report["applied"]],
        )
        info = deck.slides[3].placeholder("title").format_info()
        expect("title size restored", info["font"]["size"]["value"] == 44.0, info["font"]["size"])
        expect("title override cleared", info["font"]["size"]["override"] is False, None)
        geo = deck.slides[3].placeholder("title").geometry()
        peer = deck.slides[2].placeholder("title").geometry()
        expect(
            "title box back on its peers",
            abs(geo["top"] - peer["top"]) < 0.5,
            (geo["top"], peer["top"]),
        )

        second = deck.lint()
        expect(
            "only the report-only finding remains",
            [f["rule"] for f in second] == ["shape-off-slide"],
            [f["rule"] for f in second],
        )
        again = deck.regularize()
        expect("second regularize is a no-op", again["applied"] == [], None)

        align = deck.lint(rules=["edge-alignment"])
        expect(
            "edge-alignment finds the 2 pt drift",
            any(f["slide"] == 4 and f["fix"]["how"] == "left" for f in align),
            [f["observed"] for f in align],
        )
        deck.regularize(rules=["edge-alignment"], within="slide:4")
        lefts = sorted(
            s["geometry"]["left"] for s in deck.slides[4].shapes.list() if s["text"] in ("A", "B")
        )
        expect("cards snapped to one left edge", lefts[0] == lefts[1], lefts)
        expect(
            "alignment lint clean after fix",
            not [
                f
                for f in deck.lint(rules=["edge-alignment"], within="slide:4")
                if f["fix"]["how"] == "left"
            ],
            None,
        )
    finally:
        deck.com.Saved = True
        deck.com.Close()
    out["ok"] = ok
    print(json.dumps(out, indent=2, default=str))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

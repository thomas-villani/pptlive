"""Net-zero live probe: does `set_text_frame` actually stick in PowerPoint?

The typelib pinned `MsoVerticalAnchor`/`MsoAutoSize` (see tests/test_typelib_parity.py),
but a constant being *right* doesn't prove the *write* lands — `TextFrame2.AutoSize`
in particular is the property whose classic twin returns the mixed sentinel, and
`ExportAsFixedFormat` has already taught us that a nominally-correct COM call can
simply refuse to marshal under our late-bound dispatch.

So this checks the two things a unit test against the fake cannot:

1. Every knob round-trips through a **fresh read** of live PowerPoint.
2. `autosize="none"` makes a passed `height` binding — the reviewer's actual
   complaint was that a text box silently grew and their layout drifted. We set a
   height, write text long enough to overflow it, and assert the height held.

Net-zero: adds one temp slide, deletes it, restores the viewed slide. Run with
    uv run python scripts/text_frame_setter_spike.py

Findings (2026-08-03, first run against live PowerPoint) — it caught two real bugs
and pinned one behavior worth documenting:

1. **BUG (fixed): every canonical name was rejected.** `autosize="shape_to_fit_text"`
   raised "unknown autosize; expected one of: none, shape_to_fit_text, ..." — the
   lookup keys were written with underscores, but `_normalize_name` strips
   non-alphanumerics, so only the single-word aliases ("off"/"grow"/"shrink") ever
   matched. The unit tests missed it because they only exercised aliases. Guarded
   now by `test_every_advertised_choice_coerces`.
2. **BUG (fixed): `add_textbox` applied the text before the frame.** A new box
   autofits, so writing text first resized it and the later `autosize="none"`
   merely pinned the already-wrong height — a 40 pt box arrived at 29.1 pt.
   Creation now configures the frame first.
3. **Behavior (documented, not a bug): autofit does not re-fit retroactively.**
   Turning `autosize` back on for a frame that was laid out while autofit was off
   leaves it at its pinned size, even after rewriting the text. Set the mode before
   the text lands. The measured drift this prevents: a 40 pt box holding a long
   paragraph grows to **312.6 pt** under the default autofit, and stays at 40 pt
   with `autosize="none"`.
"""

from __future__ import annotations

import sys

import pptlive as pl

EXPECTED_MARGINS = {"left": 0.0, "right": 0.0, "top": 0.0, "bottom": 0.0}
BOX_HEIGHT = 40.0
OVERFLOWING_TEXT = "Overflow " * 40


def main() -> int:
    with pl.connect() as ppt:  # attaches, or launches PowerPoint if it isn't running
        if not len(ppt.presentations):
            ppt.com.Presentations.Add()  # a launched instance starts with no deck
        return run(ppt)


def run(ppt: pl.PowerPoint) -> int:
    deck = ppt.presentations.active
    if not len(deck.slides):
        with deck.edit("spike: seed a slide"):
            deck.slides.add(layout="blank")
    before_slides = len(deck.slides)
    viewed = ppt.viewed_slide_index()
    failures: list[str] = []

    def check(label: str, got: object, want: object) -> None:
        ok = got == want
        print(f"  [{'ok' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f" != {want!r}"))
        if not ok:
            failures.append(label)

    slide = deck.slides.add(layout="blank")
    try:
        with deck.edit("spike: text frame setters"):
            box = slide.shapes.add_textbox(
                "Pinned",
                left=60.0,
                top=60.0,
                width=300.0,
                height=BOX_HEIGHT,
                autosize="none",
                margins=0.0,
                vertical_anchor="middle",
                word_wrap=True,
            )

        print("1. creation kwargs round-trip through a fresh read")
        st = box.text_frame_status()
        check("autosize", st.autosize, "none")
        check("vertical_anchor", st.vertical_anchor, "middle")
        check("word_wrap", st.word_wrap, True)
        check("margins", {k: round(v, 1) for k, v in st.margins.items()}, EXPECTED_MARGINS)
        check("overflow_risk", st.overflow_risk, "possible")

        print("2. autosize='none' makes the passed height binding")
        with deck.edit("spike: overflow the box"):
            box.set_text(OVERFLOWING_TEXT)
        # The whole point: with autofit off the frame must NOT have grown.
        check("height after overflow", round(box.geometry()["height"], 1), BOX_HEIGHT)

        print("3. set_text_frame flips each knob on an existing shape")
        with deck.edit("spike: flip the knobs"):
            st = box.set_text_frame(
                autosize="shape_to_fit_text",
                word_wrap=False,
                vertical_anchor="bottom",
                margins=9.0,
                margin_left=18.0,
            )
        check("autosize", st.autosize, "shape_to_fit_text")
        check("word_wrap", st.word_wrap, False)
        check("vertical_anchor", st.vertical_anchor, "bottom")
        check(
            "margins",
            {k: round(v, 1) for k, v in st.margins.items()},
            {"left": 18.0, "right": 9.0, "top": 9.0, "bottom": 9.0},
        )
        # Re-read from COM rather than trusting the returned object.
        check("autosize (re-read)", box.text_frame_status().autosize, "shape_to_fit_text")

        print("4. control: autofit set at CREATION does grow the frame")
        # Not a claim the library makes, but it pins *when* PowerPoint's autofit
        # actually runs — which is what makes autosize="none" worth reaching for.
        with deck.edit("spike: a box that is allowed to grow"):
            grower = slide.shapes.add_textbox(
                OVERFLOWING_TEXT,
                left=60.0,
                top=200.0,
                width=300.0,
                height=BOX_HEIGHT,
                autosize="shape_to_fit_text",
                word_wrap=True,
            )
        grown = grower.geometry()["height"]
        check("height grew past the passed 40pt", grown > BOX_HEIGHT, True)
        print(f"     (height is now {grown:.1f}pt — the drift the reviewer hit)")

        print("5. observation: flipping AutoSize later does NOT reflow an existing frame")
        # Word wrap goes back ON first: step 3 turned it off, and an unwrapped
        # frame grows *sideways* on one long line, so a still height would prove
        # nothing about autofit.
        with deck.edit("spike: re-wrap, then rewrite the text"):
            box.set_text_frame(word_wrap=True)
            box.set_text(OVERFLOWING_TEXT + "!")
        print(
            f"     pinned box is {box.geometry()['height']:.1f}pt with "
            f"autosize={box.text_frame_status().autosize!r} — PowerPoint does not "
            "retroactively re-fit a frame that was laid out while autofit was off."
        )
    finally:
        with deck.edit("spike: clean up"):
            slide.delete()
        if viewed is not None and viewed <= len(deck.slides):
            deck.go_to(deck.slides[viewed])

    net_zero = len(deck.slides) == before_slides
    print(f"\nnet-zero: {net_zero} ({before_slides} slides before and after)")
    if failures:
        print(f"FAILED: {', '.join(failures)}")
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

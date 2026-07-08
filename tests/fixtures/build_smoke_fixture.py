"""Regenerate ``tests/fixtures/smoke_deck.pptx`` — the stable smoke-test target.

The smoke suite drives a *live* PowerPoint. Most cases build a throwaway deck and
read back what they just wrote (a round-trip), which never exercises the **read
path against content pptlive did not author**. This fixture is that missing
target: a small deck with *known* slides / a placeholder body / speaker notes / a
table, checked in so a smoke test can open it and assert exact values.

pptlive has no ``python-pptx`` dependency (its whole premise is the live app), so
the fixture is authored by **pptlive itself** over COM and saved with
``deck.save_as``. This script builds it in a **fresh throwaway presentation** (via
``Presentations.Add()``) so the user's open decks are never touched, then closes
that deck without leaving a window behind.

Run on a Windows box with PowerPoint open, from the repo root:

    uv run python tests/fixtures/build_smoke_fixture.py

Commit the resulting ``smoke_deck.pptx``. Keep the known values below in sync with
the assertions in ``tests/test_smoke.py`` (``test_fixture_reads_back``).
"""

from __future__ import annotations

from pathlib import Path

import pptlive as pl

OUT = Path(__file__).with_name("smoke_deck.pptx")

# The known content the smoke test asserts. Edit here + in test_smoke.py together.
TITLE_1 = "pptlive smoke fixture"
TITLE_2 = "Agenda"
BODY_2 = "Intro\nDemo\nQ&A"
NOTES_2 = "Lead with the roadmap."
TITLE_3 = "Metrics"
TABLE_3 = [["Team", "People"], ["Eng", "24"]]


def main() -> int:
    with pl.attach() as ppt:
        # A brand-new deck — never the user's active presentation.
        ppt.com.Presentations.Add()
        deck = ppt.presentations.active

        mine: list[int] = []
        with deck.edit("build smoke fixture"):
            s1 = deck.slides.add("title_and_content")
            mine.append(s1.id)
            deck.anchor_by_id(f"ph:{s1.index}:title").set_text(TITLE_1)

            s2 = deck.slides.add("title_and_content")
            mine.append(s2.id)
            deck.anchor_by_id(f"ph:{s2.index}:title").set_text(TITLE_2)
            deck.anchor_by_id(f"ph:{s2.index}:body").set_text(BODY_2)
            deck.slides[s2.index].notes.set_text(NOTES_2)

            s3 = deck.slides.add("title_and_content")
            mine.append(s3.id)
            deck.anchor_by_id(f"ph:{s3.index}:title").set_text(TITLE_3)
            table_shape = deck.slides[s3.index].shapes.add_table(rows=2, columns=2)
            table = table_shape.table
            for r, row in enumerate(TABLE_3, start=1):
                for c, value in enumerate(row, start=1):
                    table.cell(r, c).set_text(value)

            # Drop any default slide(s) the new presentation shipped with, so the
            # fixture is exactly our three regardless of the PowerPoint build.
            for idx in range(len(deck.slides), 0, -1):
                if deck.slides[idx].id not in mine:
                    deck.slides[idx].delete()

        deck.save_as(str(OUT), overwrite=True)
        print(f"wrote {OUT} ({len(deck.slides)} slides)")

        # Close the throwaway deck without leaving a window behind.
        com = deck.com
        try:
            com.Saved = True
            com.Close()
        except Exception as exc:  # noqa: BLE001
            print(f"(note: could not close fixture deck cleanly: {exc})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

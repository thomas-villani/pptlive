"""Restyle a whole deck: theme palette + fonts, master text styles, background.

    uv run python examples/python/03_restyle_deck.py

Where `format_text` styles one anchor, `deck.theme` and `deck.master` restyle the
*entire* deck by editing what every slide inherits. These ops are deliberately
global and anti-polite — one call recolors or re-fonts every inheriting slide —
so they still go through `deck.edit(...)` for the one-Ctrl-Z fence (your view
won't move, but the whole deck changes).
"""

from __future__ import annotations

import pptlive as pl
from pptlive.exceptions import PptliveError


def fresh_presentation(ppt: pl.PowerPoint) -> pl.Presentation:
    ppt.com.Presentations.Add()
    return ppt.presentations.active


def seed_content(deck: pl.Presentation) -> None:
    """A couple of slides so the restyle has something visible to act on."""
    with deck.edit("Seed content"):
        title = deck.slides.add("title")
        deck.anchor_by_id(f"ph:{title.index}:ctrtitle").set_text("Acme Q3 Review")
        deck.anchor_by_id(f"ph:{title.index}:subtitle").set_text("Company-wide update")

        content = deck.slides.add("title_and_content")
        deck.anchor_by_id(f"ph:{content.index}:title").set_text("Highlights")
        deck.anchor_by_id(f"ph:{content.index}:body").set_text(
            "Revenue up 18%\nTwo new markets\nNPS at an all-time high"
        )


def restyle(deck: pl.Presentation) -> None:
    with deck.edit("Restyle the deck"):
        # Theme: the 12-slot palette + heading/body typefaces.
        deck.theme.set_color("accent1", "#2E5BFF")
        deck.theme.set_color("accent2", "#00C2A8")
        deck.theme.set_font("major", "Georgia")  # headings
        deck.theme.set_font("minor", "Calibri")  # body

        # Master: text styles (title/body/default x 5 levels) + background.
        deck.master.format_text_style("title", 1, bold=True, size=40)
        deck.master.format_text_style("body", 1, size=20)
        deck.master.set_background("#0B1021")  # dark slide background


def main() -> None:
    with pl.connect() as ppt:
        deck = fresh_presentation(ppt)
        seed_content(deck)
        restyle(deck)

        # Read the theme + master back (reads never move the view).
        theme = deck.theme.read()
        print("Palette:")
        for slot, color in theme["colors"].items():
            print(f"  {slot}: {color}")
        print(f"Fonts: heading={theme['fonts']['major']}, body={theme['fonts']['minor']}")
        print(f"Background: {deck.master.read()['background']}")


if __name__ == "__main__":
    try:
        main()
    except PptliveError as exc:
        raise SystemExit(f"pptlive error: {exc}") from exc

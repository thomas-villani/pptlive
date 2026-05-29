"""pptlive quickstart — attach, add two slides, set and format text.

Run it against a live PowerPoint:

    uv run python examples/python/01_quickstart.py

It creates a fresh presentation (it won't touch a deck you already have open),
adds a title slide and a content slide, and formats a word on the title. Watch
PowerPoint while it runs — every step is polite (your view is preserved) and each
`deck.edit(...)` block is a single Ctrl-Z.
"""

from __future__ import annotations

import pptlive as pl
from pptlive.exceptions import PptliveError


def fresh_presentation(ppt: pl.PowerPoint) -> pl.Presentation:
    """A brand-new, empty presentation to draw into.

    Uses the `.com` escape hatch for `Presentations.Add()` — pptlive wraps editing
    existing decks, not deck creation, and the raw COM call is right there when we
    need it.
    """
    ppt.com.Presentations.Add()  # opens + activates a blank deck
    return ppt.presentations.active


def main() -> None:
    # connect() attaches to a running PowerPoint, launching one if none is open.
    with pl.connect() as ppt:
        deck = fresh_presentation(ppt)

        # A title slide. `ph:S:KIND` is the drift-proof way to reach a placeholder.
        with deck.edit("Add the title slide"):
            title_slide = deck.slides.add("title")
            s = title_slide.index
            deck.anchor_by_id(f"ph:{s}:ctrtitle").set_text("pptlive")
            deck.anchor_by_id(f"ph:{s}:subtitle").set_text("Drive a running PowerPoint from Python")

        # Emphasize the product name: make the whole title bold + brand-colored.
        with deck.edit("Emphasize the title"):
            deck.anchor_by_id(f"ph:{title_slide.index}:ctrtitle").format_text(
                bold=True, color="#2E5BFF", size=54
            )

        # A content slide with a bulleted body. `\n` makes paragraphs.
        with deck.edit("Add a content slide"):
            content = deck.slides.add("title_and_content")
            s = content.index
            deck.anchor_by_id(f"ph:{s}:title").set_text("Why pptlive")
            deck.anchor_by_id(f"ph:{s}:body").set_text(
                "Talks to the app you already have open\n"
                "Edits stay polite — your view never moves\n"
                "Every change is one clean undo"
            )

        # Read it back — reads never move the user's view.
        print(f"Created a {len(deck.slides)}-slide deck: {deck.name!r}")
        for slide in deck.slides:
            data = slide.read()
            title = data.get("title") or "(no title)"
            print(f"  slide {slide.index}: {title}")


if __name__ == "__main__":
    try:
        main()
    except PptliveError as exc:
        # e.g. PowerPoint not running and couldn't be launched.
        raise SystemExit(f"pptlive error: {exc}") from exc

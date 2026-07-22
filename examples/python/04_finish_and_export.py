"""Finish a deck: fix a typo, snapshot it, then save + export a PDF.

    uv run python examples/python/04_finish_and_export.py

Where 01-03 build content, this is the *finishing pass* an agent runs right
before handing a deck back: `find_replace` catches a leftover typo,
`snapshot()` renders cheap low-res PNGs for a vision-model sanity check, then
`save_as` + `export_pdf` actually hand something back on disk. To keep the
script self-contained (nothing to attach to beforehand) it plants its own
typo via a normal edit first.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pptlive as pl
from pptlive.exceptions import PptliveError


def fresh_presentation(ppt: pl.PowerPoint) -> pl.Presentation:
    ppt.com.Presentations.Add()
    return ppt.presentations.active


def seed_content(deck: pl.Presentation) -> None:
    """A couple of slides, with a deliberate typo to fix later."""
    with deck.edit("Seed content"):
        title = deck.slides.add("title")
        deck.anchor_by_id(f"ph:{title.index}:ctrtitle").set_text("Q3 Reveiw")
        deck.anchor_by_id(f"ph:{title.index}:subtitle").set_text("Draft for review")

        content = deck.slides.add("title_and_content")
        deck.anchor_by_id(f"ph:{content.index}:title").set_text("Highlights")
        deck.anchor_by_id(f"ph:{content.index}:body").set_text(
            "Revenue up 18%\nTwo new markets\nNPS at an all-time high"
        )


def fix_typo(deck: pl.Presentation) -> None:
    """`find_replace` catches the "Reveiw" typo `seed_content` planted."""
    with deck.edit("Fix a typo"):
        applied = deck.find_replace("Reveiw", "Review")
        for hit in applied:
            print(f"  fixed {hit['anchor_id']!r}: {hit['text']!r} -> 'Review'")


def main() -> None:
    # connect() attaches to a running PowerPoint, launching one if none is open.
    with pl.connect() as ppt:
        deck = fresh_presentation(ppt)
        seed_content(deck)
        fix_typo(deck)

        out_dir = Path(tempfile.mkdtemp(prefix="pptlive_"))

        # Snapshot: cheap low-res PNGs, one per slide, long edge capped to
        # 1000px -- plenty for a vision model to eyeball the layout. A read,
        # so it needs no deck.edit() fence. Multiple slides land next to `out`
        # as `<stem>-s<N><suffix>`, e.g. deck-s1.png, deck-s2.png.
        snaps = deck.snapshot(out=out_dir / "deck.png", max_dim=1000)
        print("Snapshot:")
        for snap in snaps:
            print(f"  slide {snap.slide}: {snap.path}")

        # Persist: save_as rebinds the open deck to the new .pptx path.
        pptx_path = deck.save_as(out_dir / "deck_demo.pptx", overwrite=True)
        print(f"Saved: {pptx_path}")

        # Export a PDF too -- also a read: doesn't rebind or touch the dirty flag.
        pdf_path = deck.export_pdf(out_dir / "deck_demo.pdf")
        print(f"Exported PDF: {pdf_path}")


if __name__ == "__main__":
    try:
        main()
    except PptliveError as exc:
        raise SystemExit(f"pptlive error: {exc}") from exc

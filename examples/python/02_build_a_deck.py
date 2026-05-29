"""Build a small deck: a table, a chart, and a SmartArt diagram.

    uv run python examples/python/02_build_a_deck.py

Each of the three object types is a *shape on a slide*. You add it through the
slide's `shapes` collection and then reach its structured content through the
returned shape (`.table`, `.chart`, `.smartart`). Geometry is in points; use
`pl.units.inches(...)` so you don't hardcode the math.
"""

from __future__ import annotations

import pptlive as pl
from pptlive.exceptions import PptliveError
from pptlive.units import inches


def fresh_presentation(ppt: pl.PowerPoint) -> pl.Presentation:
    ppt.com.Presentations.Add()
    return ppt.presentations.active


def add_table_slide(deck: pl.Presentation) -> None:
    with deck.edit("Add a table slide"):
        slide = deck.slides.add("title_and_content")
        deck.anchor_by_id(f"ph:{slide.index}:title").set_text("Headcount by team")

        shape = slide.shapes.add_table(
            rows=3, columns=2, left=inches(1), top=inches(2), width=inches(5)
        )
        table = shape.table
        # Cells are anchors too: address them cell:S:N:R:C or via the table.
        rows = [("Team", "People"), ("Engineering", "24"), ("Sales", "11")]
        for r, (left, right) in enumerate(rows, start=1):
            table.cell(r, 1).set_text(left)
            table.cell(r, 2).set_text(right)
        # Bold the header row.
        table.cell(1, 1).format_text(bold=True)
        table.cell(1, 2).format_text(bold=True)


def add_chart_slide(deck: pl.Presentation) -> None:
    with deck.edit("Add a chart slide"):
        slide = deck.slides.add("title_and_content")
        deck.anchor_by_id(f"ph:{slide.index}:title").set_text("Quarterly revenue")

        # add_chart can take its data up front: categories + named series.
        slide.shapes.add_chart(
            "column",
            ["Q1", "Q2", "Q3", "Q4"],
            {"Revenue": [10, 14, 19, 23], "Profit": [3, 5, 8, 11]},
            left=inches(1),
            top=inches(1.8),
        )


def add_smartart_slide(deck: pl.Presentation) -> None:
    with deck.edit("Add a SmartArt slide"):
        slide = deck.slides.add("title_and_content")
        deck.anchor_by_id(f"ph:{slide.index}:title").set_text("How it ships")

        # A flat layout takes many top-level nodes (strings). A tree layout
        # (orgchart/hierarchy) would take a single {text, children} root instead.
        slide.shapes.add_smartart(
            "process",
            ["Discover", "Design", "Build", "Ship"],
            left=inches(0.7),
            top=inches(2),
        )


def main() -> None:
    with pl.connect() as ppt:
        deck = fresh_presentation(ppt)
        add_table_slide(deck)
        add_chart_slide(deck)
        add_smartart_slide(deck)

        # Read the chart back from its embedded data to confirm the round-trip.
        chart_shape = deck.slides[2].shapes[1]
        if chart_shape.has_chart:
            data = chart_shape.chart.read()
            print(f"Chart type: {data['chart_type']}")
            print(f"Categories: {data['categories']}")
            for series in data["series"]:
                print(f"  {series['name']}: {series['values']}")

        print(f"Built a {len(deck.slides)}-slide deck.")


if __name__ == "__main__":
    try:
        main()
    except PptliveError as exc:
        raise SystemExit(f"pptlive error: {exc}") from exc

"""Live-PowerPoint smoke tests (skipped by default; `uv run pytest -m smoke`).

These need a reachable PowerPoint install — the `real_powerpoint` fixture skips
the suite otherwise. Every test is **net-zero**: it opens its own fresh
presentation, does its work, and closes that deck without saving, so it never
touches a deck the user already has open and leaves nothing behind.

The headline case is the embedded-Excel `wb.Close()` after a chart data write.
The 2026-05-28 spike saw a clean close on English Office; the open question for
0.1.0 was whether a non-default Excel config (a non-English first-sheet name like
"Feuil1"/"Hoja1", or a build that prompts on close) breaks it. Run this on such a
box to confirm: `_sheet_ref(ws.Name)` quotes the localized sheet name, and the
no-arg `Workbook.Close()` commits the embedded data back without a Save prompt.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path
from typing import Any

import pytest

import pptlive as pl

pytestmark = pytest.mark.smoke

#: The checked-in fixture deck (regenerate via tests/fixtures/build_smoke_fixture.py).
#: Keep these known values in sync with that script's constants.
FIXTURE = Path(__file__).parent / "fixtures" / "smoke_deck.pptx"


@pytest.fixture
def fresh_deck(real_powerpoint: pl.PowerPoint) -> Iterator[pl.Presentation]:
    """A brand-new presentation, closed (unsaved) on teardown — never the user's."""
    real_powerpoint.com.Presentations.Add()
    deck = real_powerpoint.presentations.active
    try:
        yield deck
    finally:
        # Close without saving; suppress the dialog by clearing the dirty flag.
        # Both calls are best-effort: a chart-automation stress failure can leave
        # the connection's COM proxies dead (RPC_S_SERVER_UNAVAILABLE), in which
        # case there's nothing left to close — don't turn that into a teardown error
        # on top of the test failure.
        com = deck.com
        try:
            com.Saved = True
        except Exception:
            pass
        try:
            com.Close()
        except Exception:
            pass


def test_chart_set_data_round_trips_and_closes_cleanly(fresh_deck: pl.Presentation) -> None:
    """End-to-end: add a chart, rewrite its data, read it back.

    Exercises the `ChartData.Activate()` → write cells → `SetSourceData` →
    `wb.Close()` sequence against real Excel. A non-English first-sheet name or a
    Close prompt would surface here as a COM error rather than a silent pass.

    Reliability note (2026-06-10): PowerPoint's embedded-Excel chart automation is
    fragile under *repeated* use — a chart's data write spins up and tears down an
    Excel server, and a tight create/rewrite loop eventually trips transient RPC
    failures (RPC_S_CALL_FAILED) or, under enough stress, takes the whole COM
    connection down (RPC_S_SERVER_UNAVAILABLE — unrecoverable in-process, so not
    retried). `Chart.set_data` now reads the data back and retries the *recoverable*
    silent-commit race, which makes a single run reliable; but stress-looping this
    test many times in quick succession can still destabilize the live instance.
    Run it once (as the smoke suite does), not in a hammer loop.
    """
    with fresh_deck.edit("smoke: add chart"):
        slide = fresh_deck.slides.add("title_and_content")
        shape = slide.shapes.add_chart(
            "column",
            ["Q1", "Q2", "Q3", "Q4"],
            {"Revenue": [10, 14, 19, 23], "Profit": [3, 5, 8, 11]},
        )

    chart = shape.chart
    with fresh_deck.edit("smoke: rewrite chart data"):
        chart.set_data(["A", "B", "C"], {"Only": [1.0, 2.0, 3.0]})

    data: dict[str, Any] = chart.read()
    assert data["categories"] == ["A", "B", "C"]
    assert len(data["series"]) == 1
    assert data["series"][0]["name"] == "Only"
    assert data["series"][0]["values"] == [1.0, 2.0, 3.0]


def test_table_round_trips(fresh_deck: pl.Presentation) -> None:
    """A table write/read round-trip against real PowerPoint (cheap sanity case)."""
    with fresh_deck.edit("smoke: add table"):
        slide = fresh_deck.slides.add("title_and_content")
        shape = slide.shapes.add_table(rows=2, columns=2)
        table = shape.table
        table.cell(1, 1).set_text("Team")
        table.cell(1, 2).set_text("People")
        table.cell(2, 1).set_text("Eng")
        table.cell(2, 2).set_text("24")

    grid = shape.table.grid()
    assert grid[0] == ["Team", "People"]
    assert grid[1] == ["Eng", "24"]


def test_align_uses_the_real_office_enum(fresh_deck: pl.Presentation) -> None:
    """`shapes.align` maps each `how` to the right live `ShapeRange.Align` cmd.

    Regression for the 0-vs-1-based `MsoAlignCmd` bug: the enum shipped 1-based, so
    `left` aligned *centers*, `right` aligned *tops*, and `bottom` sent an
    out-of-range 6 that errored outright. The unit suite cannot catch this class of
    bug — the fake COM mirrors whatever the enum says — so it has to be pinned live.

    The three rectangles have **distinct widths** on purpose: on equal-width shapes
    an align-lefts and an align-centers produce identical `Left` values, so uniform
    shapes cannot tell the two apart (exactly why the original spike passed against
    the wrong constant).
    """
    with fresh_deck.edit("smoke: align"):
        slide = fresh_deck.slides.add("blank")
        a = slide.shapes.add_shape("rectangle", left=80.0, top=240.0, width=120.0, height=60.0)
        b = slide.shapes.add_shape("rectangle", left=300.0, top=300.0, width=180.0, height=60.0)
        c = slide.shapes.add_shape("rectangle", left=500.0, top=360.0, width=240.0, height=90.0)
        members = [a, b, c]
        slide.shapes.align(members, "left", relative_to="selection")

    lefts = [round(m.geometry()["left"], 1) for m in members]
    assert len(set(lefts)) == 1, f"align 'left' did not align left edges: {lefts}"
    # Centers must now *differ* — if they matched, we aligned centers, not lefts.
    centers = [round(m.geometry()["left"] + m.geometry()["width"] / 2, 1) for m in members]
    assert len(set(centers)) == 3, f"align 'left' aligned centers instead: {centers}"

    # `bottom` is the case that errored outright under the 1-based enum (cmd 6).
    with fresh_deck.edit("smoke: align bottom"):
        slide.shapes.align(members, "bottom", relative_to="selection")

    bottoms = [round(m.geometry()["top"] + m.geometry()["height"], 1) for m in members]
    assert len(set(bottoms)) == 1, f"align 'bottom' did not align bottom edges: {bottoms}"


@pytest.fixture
def fixture_deck(real_powerpoint: pl.PowerPoint) -> Iterator[pl.Presentation]:
    """Open the checked-in smoke fixture read-only; close it on teardown.

    Read-only so we never touch the committed file or trip a dirty-flag prompt.
    Skips (not fails) if the fixture is missing — the file is regenerated by
    tests/fixtures/build_smoke_fixture.py on a box with PowerPoint.
    """
    if not FIXTURE.exists():
        pytest.skip(f"smoke fixture not found: {FIXTURE}")
    real_powerpoint.com.Presentations.Open(str(FIXTURE), ReadOnly=True, WithWindow=True)
    deck = real_powerpoint.presentations.active
    try:
        yield deck
    finally:
        com = deck.com
        try:
            com.Saved = True
        except Exception:
            pass
        try:
            com.Close()
        except Exception:
            pass


def test_fixture_reads_back(fixture_deck: pl.Presentation) -> None:
    """Read the checked-in fixture and assert its known content.

    Unlike the create-then-read cases above, this exercises the read path against
    content pptlive is *opening*, not one it just wrote in this session — the
    thing a plain round-trip can't catch (a save/reopen that mangles text, notes,
    or a table).
    """
    assert len(fixture_deck.slides) == 3

    assert fixture_deck.anchor_by_id("ph:1:title").text == "pptlive smoke fixture"

    assert fixture_deck.anchor_by_id("ph:2:title").text == "Agenda"
    body = fixture_deck.anchor_by_id("ph:2:body").text
    bullets = [b for b in body.replace("\v", "\r").split("\r") if b]
    assert bullets == ["Intro", "Demo", "Q&A"]
    assert fixture_deck.slides[2].notes.text == "Lead with the roadmap."

    assert fixture_deck.anchor_by_id("ph:3:title").text == "Metrics"
    table_shape = next(s for s in fixture_deck.slides[3].shapes if s.has_table)
    grid = table_shape.table.grid()
    assert grid == [["Team", "People"], ["Eng", "24"]]

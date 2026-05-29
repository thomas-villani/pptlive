"""Charts (v0.7): add_chart, the has_chart gate, and the Chart wrapper.

Against the fake, a chart shape carries a `_FakeChart` whose embedded workbook is
a {(row,col): value} dict; `SetSourceData(string)` records the plotted range and
`SeriesCollection` parses it back — so the real write sequence (ClearContents →
Cells writes → SetSourceData) round-trips, and shrink/grow re-writes leave no
stale data, exactly as the live spike found.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.exceptions import AnchorNotFoundError


def _json(result):  # type: ignore[no-untyped-def]
    return json.loads(result.output)


# -- add_chart + gate (wrapper) ---------------------------------------------


def test_add_chart_appends_and_has_chart(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[3].shapes
    before = len(shapes)
    ch = shapes.add_chart("column")
    assert len(shapes) == before + 1
    assert ch.index == before + 1  # top of z-order
    assert ch.shape_type == "chart"
    assert ch.has_chart is True


def test_add_chart_default_data(deck) -> None:  # type: ignore[no-untyped-def]
    # No categories/series -> PowerPoint's default placeholder data (3 series x 4).
    info = deck.slides[3].shapes.add_chart("bar").chart.read()
    assert info["chart_type"] == "bar_clustered"
    assert len(info["series"]) == 3
    assert info["categories"] == ["Category 1", "Category 2", "Category 3", "Category 4"]


def test_add_chart_with_data(deck) -> None:  # type: ignore[no-untyped-def]
    ch = deck.slides[3].shapes.add_chart(
        "line", ["Q1", "Q2", "Q3"], {"Revenue": [10, 20, 30], "Profit": [3, 6, 9]}
    )
    info = ch.chart.read()
    assert info["chart_type"] == "line"
    assert info["categories"] == ["Q1", "Q2", "Q3"]
    assert [s["name"] for s in info["series"]] == ["Revenue", "Profit"]
    assert info["series"][0]["values"] == [10.0, 20.0, 30.0]


def test_add_chart_unknown_type_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown chart type"):
        deck.slides[3].shapes.add_chart("piechart")


def test_add_chart_categories_without_series_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="both categories and series"):
        deck.slides[3].shapes.add_chart("pie", ["A", "B"])


def test_has_chart_false_for_plain_shape(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[2].shapes[3].has_chart is False  # the picture


def test_chart_on_non_chart_shape_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[2].shapes[3].chart


def test_shape_listing_emits_has_chart(deck) -> None:  # type: ignore[no-untyped-def]
    deck.slides[3].shapes.add_chart("column")
    rows = deck.slides[3].shapes.list()
    assert rows[-1]["has_chart"] is True
    assert rows[0]["has_chart"] is False


# -- Chart wrapper: set_type / set_data -------------------------------------


def test_set_type_round_trips(deck) -> None:  # type: ignore[no-untyped-def]
    ch = deck.slides[3].shapes.add_chart("column").chart
    ch.set_type("pie")
    assert ch.chart_type == "pie"


def test_set_data_shrink_then_grow(deck) -> None:  # type: ignore[no-untyped-def]
    ch = deck.slides[3].shapes.add_chart("column").chart
    # shrink to 1 series x 2 categories
    ch.set_data(["A", "B"], {"Only": [7, 8]})
    info = ch.read()
    assert info["categories"] == ["A", "B"]
    assert [s["name"] for s in info["series"]] == ["Only"]
    assert info["series"][0]["values"] == [7.0, 8.0]
    # grow to 3 series x 4 categories -> no stale data from the shrink
    ch.set_data(
        ["Jan", "Feb", "Mar", "Apr"],
        {"N": [1, 2, 3, 4], "S": [5, 6, 7, 8], "E": [9, 10, 11, 12]},
    )
    info = ch.read()
    assert info["categories"] == ["Jan", "Feb", "Mar", "Apr"]
    assert [s["name"] for s in info["series"]] == ["N", "S", "E"]
    assert info["series"][2]["values"] == [9.0, 10.0, 11.0, 12.0]


def test_set_data_accepts_pair_sequence(deck) -> None:  # type: ignore[no-untyped-def]
    ch = deck.slides[3].shapes.add_chart("column").chart
    ch.set_data(["A", "B"], [("X", [1, 2]), ("Y", [3, 4])])
    assert [s["name"] for s in ch.read()["series"]] == ["X", "Y"]


def test_set_data_length_mismatch_raises(deck) -> None:  # type: ignore[no-untyped-def]
    ch = deck.slides[3].shapes.add_chart("column").chart
    with pytest.raises(ValueError, match="values but there are"):
        ch.set_data(["A", "B", "C"], {"Bad": [1, 2]})


def test_set_data_empty_raises(deck) -> None:  # type: ignore[no-untyped-def]
    ch = deck.slides[3].shapes.add_chart("column").chart
    with pytest.raises(ValueError, match="at least one category"):
        ch.set_data([], {"X": []})


def test_set_data_closes_workbook(deck) -> None:  # type: ignore[no-untyped-def]
    # Politeness/cleanup: the embedded workbook is Close()d after a write.
    shape = deck.slides[3].shapes.add_chart("column")
    shape.chart.set_data(["A"], {"X": [1]})
    assert shape.com._chart._wb.closed is True


# -- CLI --------------------------------------------------------------------


def test_cli_shape_add_chart(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "chart",
            "--chart-type",
            "line",
            "--categories",
            "Q1,Q2,Q3",
            "--series",
            '{"Revenue":[10,20,30]}',
        ],
    )
    assert result.exit_code == 0, result.output
    payload = _json(result)
    assert payload["type"] == "chart"
    assert payload["has_chart"] is True


def test_cli_chart_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "chart",
            "--chart-type",
            "pie",
            "--categories",
            "A,B",
            "--series",
            '{"S":[1,2]}',
        ],
    )
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(main, ["chart", "read", "--slide", "3", "--shape", str(n)])
    assert result.exit_code == 0
    info = _json(result)
    assert info["chart_type"] == "pie"
    assert info["categories"] == ["A", "B"]
    assert info["series"][0]["values"] == [1.0, 2.0]


def test_cli_chart_set_type(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(main, ["shape", "add", "--slide", "3", "--kind", "chart"])
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(
        main, ["chart", "set-type", "--slide", "3", "--shape", str(n), "--chart-type", "bar"]
    )
    assert result.exit_code == 0
    assert _json(result)["chart_type"] == "bar_clustered"


def test_cli_chart_set_data(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(main, ["shape", "add", "--slide", "3", "--kind", "chart"])
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(
        main,
        [
            "chart",
            "set-data",
            "--slide",
            "3",
            "--shape",
            str(n),
            "--categories",
            '["X","Y"]',
            "--series",
            '[["A",[1,2]],["B",[3,4]]]',
        ],
    )
    assert result.exit_code == 0
    info = _json(result)
    assert info["categories"] == ["X", "Y"]
    assert [s["name"] for s in info["series"]] == ["A", "B"]


def test_cli_chart_read_non_chart_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # slide 2 shape 3 is a picture, not a chart.
    result = CliRunner().invoke(main, ["chart", "read", "--slide", "2", "--shape", "3"])
    assert result.exit_code == 2


def test_cli_shape_add_chart_one_of_data_opts_errors(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "add", "--slide", "3", "--kind", "chart", "--categories", "A,B"]
    )
    assert result.exit_code != 0
    assert "both --categories and --series" in result.output

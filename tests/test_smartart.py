"""SmartArt (v0.8): add_smartart, the has_smartart gate, and the SmartArt wrapper.

Against the fake, a SmartArt shape carries a `_FakeSmartArt` whose content is a
node tree (text on `TextFrame2`). The fake mirrors the live findings the wrapper
relies on: `Nodes.Add()` adds a top-level sibling, `SmartArtNode.AddNode(BELOW)`
adds a child, tree layouts seed a clearable skeleton, and a node's `Type` never
reflects assistant — so the populate/read recipe round-trips exactly here.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.constants import parse_color, smartart_layout_for, smartart_layout_name
from pptlive.exceptions import AnchorNotFoundError


def _json(result):  # type: ignore[no-untyped-def]
    return json.loads(result.output)


# -- constants --------------------------------------------------------------


def test_smartart_layout_for_is_separator_insensitive() -> None:
    assert smartart_layout_for("Org Chart") == "orgChart1"
    assert smartart_layout_for("orgchart") == "orgChart1"
    assert smartart_layout_for("PROCESS") == "process1"


def test_smartart_layout_for_unknown_raises() -> None:
    with pytest.raises(ValueError, match="unknown SmartArt layout"):
        smartart_layout_for("flowchart")


def test_smartart_layout_name_round_trip_and_fallback() -> None:
    assert smartart_layout_name("urn:.../layout/process1") == "process"
    assert smartart_layout_name("urn:.../layout/orgChart1") == "orgchart"
    # unknown URN falls back to its trailing segment, never raises
    assert smartart_layout_name("urn:.../layout/gear1") == "gear1"


# -- add_smartart + gate (wrapper) ------------------------------------------


def test_add_smartart_appends_and_has_smartart(deck) -> None:  # type: ignore[no-untyped-def]
    shapes = deck.slides[3].shapes
    before = len(shapes)
    sa = shapes.add_smartart("process")
    assert len(shapes) == before + 1
    assert sa.index == before + 1  # top of z-order
    assert sa.shape_type == "smart_art"
    assert sa.has_smartart is True


def test_add_smartart_default_nodes(deck) -> None:  # type: ignore[no-untyped-def]
    # No nodes -> the layout's default placeholder nodes (flat process seeds 3).
    info = deck.slides[3].shapes.add_smartart("process").smartart.read()
    assert info["layout"] == "process"
    assert info["node_count"] == 3
    assert all(n["level"] == 1 for n in info["nodes"])


def test_read_warns_when_walk_count_diverges_from_allnodes(deck, monkeypatch) -> None:  # type: ignore[no-untyped-def]
    from types import SimpleNamespace

    sh = deck.slides[3].shapes.add_smartart("process")  # 3 flat nodes
    fake = sh.com.SmartArt  # the _FakeSmartArt instance
    # Inflate AllNodes.Count beyond the recursive walk (3) so the depth-first
    # assumption fails — read() must warn that node_index may not address nodes.
    monkeypatch.setattr(
        type(fake),
        "AllNodes",
        property(lambda self: SimpleNamespace(Count=99, Item=lambda i: self._nodes.Item(i))),
    )
    with pytest.warns(UserWarning, match="node_index"):
        sh.smartart.read()


def test_add_smartart_with_flat_nodes(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["Discover", "Design", "Build", "Ship"])
    info = sa.smartart.read()
    assert [n["text"] for n in info["nodes"]] == ["Discover", "Design", "Build", "Ship"]
    assert info["node_count"] == 4


def test_add_smartart_with_tree_nodes(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart(
        "orgchart",
        [{"text": "CEO", "children": ["VP Eng", {"text": "VP Sales", "children": ["AE"]}]}],
    )
    info = sa.smartart.read()
    assert info["layout"] == "orgchart"
    root = info["nodes"][0]
    assert root["text"] == "CEO" and root["level"] == 1
    assert [c["text"] for c in root["children"]] == ["VP Eng", "VP Sales"]
    assert [c["level"] for c in root["children"]] == [2, 2]
    # grandchild nests under VP Sales at level 3
    vp_sales = root["children"][1]
    assert [g["text"] for g in vp_sales["children"]] == ["AE"]
    assert vp_sales["children"][0]["level"] == 3


def test_add_smartart_unknown_kind_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown SmartArt layout"):
        deck.slides[3].shapes.add_smartart("flowchart")


def test_has_smartart_false_for_plain_shape(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[2].shapes[3].has_smartart is False  # the picture


def test_smartart_on_non_smartart_shape_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        deck.slides[2].shapes[3].smartart


def test_shape_listing_emits_has_smartart(deck) -> None:  # type: ignore[no-untyped-def]
    deck.slides[3].shapes.add_smartart("cycle")
    rows = deck.slides[3].shapes.list()
    assert rows[-1]["has_smartart"] is True
    assert rows[0]["has_smartart"] is False


# -- SmartArt wrapper: set_nodes --------------------------------------------


def test_set_nodes_flat_round_trip(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("list").smartart
    sa.set_nodes(["A", "B"])
    assert [n["text"] for n in sa.read()["nodes"]] == ["A", "B"]


def test_set_nodes_clears_tree_skeleton(deck) -> None:  # type: ignore[no-untyped-def]
    # orgChart seeds a 1-root + 2-empty-children skeleton; set_nodes must clear it.
    sa = deck.slides[3].shapes.add_smartart("orgchart").smartart
    sa.set_nodes([{"text": "Root", "children": ["Only"]}])
    info = sa.read()
    assert info["node_count"] == 2  # root + one child, no stale placeholders
    assert [c["text"] for c in info["nodes"][0]["children"]] == ["Only"]


def test_set_nodes_seeds_root_when_diagram_is_empty(deck) -> None:  # type: ignore[no-untyped-def]
    # Regression: a blank layout (or one left empty by a prior edit) has zero
    # top-level nodes. set_nodes must seed a root before sizing the list rather
    # than blowing up on Nodes.Item(1).
    sa = deck.slides[3].shapes.add_smartart("process").smartart
    sa.com.Nodes._nodes.clear()
    assert sa.com.Nodes.Count == 0
    sa.set_nodes(["Alpha", "Beta"])
    assert [n["text"] for n in sa.read()["nodes"]] == ["Alpha", "Beta"]


def test_set_nodes_empty_raises(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process").smartart
    with pytest.raises(ValueError, match="at least one node"):
        sa.set_nodes([])


def test_set_nodes_tree_layout_rejects_multiple_roots(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("hierarchy").smartart
    with pytest.raises(ValueError, match="single root node"):
        sa.set_nodes(["A", "B"])


def test_set_nodes_rejects_bare_string(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process").smartart
    with pytest.raises(TypeError, match="list of strings"):
        sa.set_nodes("just a string")  # type: ignore[arg-type]


def test_read_is_side_effect_free(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["X", "Y"]).smartart
    first = sa.read()
    second = sa.read()
    assert first == second


# -- CLI --------------------------------------------------------------------


def test_cli_shape_add_smartart(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "smartart",
            "--smartart-kind",
            "process",
            "--nodes",
            '["Discover","Design","Build"]',
        ],
    )
    assert result.exit_code == 0, result.output
    payload = _json(result)
    assert payload["type"] == "smart_art"
    assert payload["has_smartart"] is True


def test_cli_smartart_read(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(
        main,
        ["shape", "add", "--slide", "3", "--kind", "smartart", "--smartart-kind", "cycle"],
    )
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(main, ["smartart", "read", "--slide", "3", "--shape", str(n)])
    assert result.exit_code == 0, result.output
    info = _json(result)
    assert info["layout"] == "cycle"
    assert "nodes" in info


def test_cli_smartart_set_nodes(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(
        main,
        ["shape", "add", "--slide", "3", "--kind", "smartart", "--smartart-kind", "orgchart"],
    )
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(
        main,
        [
            "smartart",
            "set-nodes",
            "--slide",
            "3",
            "--shape",
            str(n),
            "--nodes",
            '[{"text":"CEO","children":["VP Eng","VP Sales"]}]',
        ],
    )
    assert result.exit_code == 0, result.output
    info = _json(result)
    assert info["nodes"][0]["text"] == "CEO"
    assert [c["text"] for c in info["nodes"][0]["children"]] == ["VP Eng", "VP Sales"]


def test_cli_smartart_read_non_smartart_is_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["smartart", "read", "--slide", "2", "--shape", "3"])
    assert result.exit_code == 2, result.output


# -- recolor_text (PPTLIVE-009) ---------------------------------------------


def test_smartart_recolor_text_recolors_every_node(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart(
        "process", ["A", "B", {"text": "C", "children": ["c1", "c2"]}]
    )
    info = sa.smartart.recolor_text("#112233")
    assert info["ok"] is True
    assert info["color"] == "#112233"
    assert info["nodes_recolored"] == 5  # A, B, C, c1, c2 (AllNodes, depth-first)
    allnodes = sa.com.SmartArt.AllNodes
    expected = parse_color("#112233")
    assert allnodes.Count == 5
    for i in range(1, allnodes.Count + 1):
        node = allnodes.Item(i)
        assert int(node.TextFrame2.TextRange.Font.Fill.ForeColor.RGB) == expected


def test_smartart_recolor_text_bad_color_raises(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["A"])
    with pytest.raises(ValueError):
        sa.smartart.recolor_text("nope")


# -- read echoes node_index + format_node (per-node text) --------------------


def test_read_echoes_depth_first_node_index(deck) -> None:  # type: ignore[no-untyped-def]
    # node_index is a 1-based depth-first walk == AllNodes order (spike-verified).
    sa = (
        deck.slides[3]
        .shapes.add_smartart(
            "orgchart", [{"text": "Root", "children": ["A", {"text": "B", "children": ["b1"]}]}]
        )
        .smartart
    )
    info = sa.read()
    root = info["nodes"][0]
    assert root["node_index"] == 1
    assert [c["node_index"] for c in root["children"]] == [2, 3]
    assert root["children"][1]["children"][0]["node_index"] == 4  # b1
    # and the index addresses the matching AllNodes item
    allnodes = sa.com.AllNodes
    assert str(allnodes.Item(4).TextFrame2.TextRange.Text) == "b1"


def test_format_node_sets_font_on_one_node(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["A", "B", "C"]).smartart
    info = sa.format_node(2, bold=True, italic=True, size=28, font="Georgia", color="#0000FF")
    assert info["ok"] is True
    assert info["node_index"] == 2
    assert info["text"] == "B"
    target = sa.com.AllNodes.Item(2).TextFrame2.TextRange.Font
    assert int(target.Bold) == -1
    assert int(target.Italic) == -1
    assert float(target.Size) == 28.0
    assert str(target.Name) == "Georgia"
    assert int(target.Fill.ForeColor.RGB) == parse_color("#0000FF")
    # neighbors untouched (still the automatic-color sentinel)
    assert int(sa.com.AllNodes.Item(1).TextFrame2.TextRange.Font.Fill.ForeColor.RGB) == 0x80000000


def test_format_node_underline_maps_to_style(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["A"]).smartart
    sa.format_node(1, underline=True)
    assert int(sa.com.AllNodes.Item(1).TextFrame2.TextRange.Font.UnderlineStyle) == 2
    sa.format_node(1, underline=False)
    assert int(sa.com.AllNodes.Item(1).TextFrame2.TextRange.Font.UnderlineStyle) == 0


def test_format_node_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["A", "B"]).smartart
    with pytest.raises(AnchorNotFoundError):
        sa.format_node(9, bold=True)


def test_format_node_bad_color_raises(deck) -> None:  # type: ignore[no-untyped-def]
    sa = deck.slides[3].shapes.add_smartart("process", ["A"]).smartart
    with pytest.raises(ValueError):
        sa.format_node(1, color="nope")


def test_cli_smartart_format_node(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "smartart",
            "--smartart-kind",
            "process",
            "--nodes",
            '["A","B","C"]',
        ],
    )
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(
        main,
        [
            "smartart",
            "format-node",
            "--slide",
            "3",
            "--shape",
            str(n),
            "--node",
            "1",
            "--bold",
            "--color",
            "#FF0000",
        ],
    )
    assert result.exit_code == 0, result.output
    info = _json(result)
    assert info["node_index"] == 1
    assert info["text"] == "A"


def test_cli_smartart_recolor_text(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    add = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "smartart",
            "--smartart-kind",
            "process",
            "--nodes",
            '["A","B","C"]',
        ],
    )
    n = int(_json(add)["anchor_id"].split(":")[2])
    result = CliRunner().invoke(
        main,
        ["smartart", "recolor-text", "--slide", "3", "--shape", str(n), "--color", "#FFFFFF"],
    )
    assert result.exit_code == 0, result.output
    info = _json(result)
    assert info["color"] == "#FFFFFF"
    assert info["nodes_recolored"] == 3

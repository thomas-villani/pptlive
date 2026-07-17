"""Shape arrangement (v-next) — group/ungroup, align/distribute, connectors.

Against the fake deck: slide 1 has Title 1 (shape:1:1, id 2) + Subtitle 2
(shape:1:2, id 3); slide 2 has Title 1 (shape:2:1, id 2), Content Placeholder 2
(shape:2:2, id 3), Picture 3 (shape:2:3, id 4). Live COM behaviour was pinned in
`scripts/arrangement_spike.py` (group gets a new id; ungroup keeps the members'
ids; reroute reassigns connection sites).
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, _edit_core
from pptlive.cli.main import main
from pptlive.constants import MsoAlignCmd, MsoDistributeCmd

# -- constants: pinned against the live Office typelib ----------------------


def test_align_cmd_matches_the_office_typelib() -> None:
    """`MsoAlignCmd` is 0-based — the one Mso* enum here that doesn't start at 1.

    Shipped 1-based in v0.7.0-dev, which silently shifted every align by one slot
    (`left` aligned centers; `bottom` sent an out-of-range 6). The fake COM can't
    catch this — it mirrors whatever the enum says — so these values are pinned
    against the real Office typelib
    (`gencache.EnsureModule('{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}', 0, 2, 8)`):

        msoAlignLefts 0  msoAlignCenters 1  msoAlignRights  2
        msoAlignTops  3  msoAlignMiddles 4  msoAlignBottoms 5

    Distribute is pinned alongside it because the two adjacent enums disagreeing
    (Distribute 0-based, Align 1-based) was the tell that went unnoticed.
    """
    assert (MsoAlignCmd.LEFTS, MsoAlignCmd.CENTERS, MsoAlignCmd.RIGHTS) == (0, 1, 2)
    assert (MsoAlignCmd.TOPS, MsoAlignCmd.MIDDLES, MsoAlignCmd.BOTTOMS) == (3, 4, 5)
    assert (MsoDistributeCmd.HORIZONTALLY, MsoDistributeCmd.VERTICALLY) == (0, 1)


# -- library: group / ungroup -----------------------------------------------


def test_group_returns_new_id_members_keep_theirs(deck) -> None:  # type: ignore[no-untyped-def]
    s1 = deck.slides[1]
    a, b = s1.shapes[1], s1.shapes[2]
    with deck.edit("t"):
        group = s1.shapes.group([a, b])
    assert group.anchor_id.startswith("shapeid:1:")
    d = group.to_dict()
    assert d["type"] == "group"
    assert sorted(d["group_item_ids"]) == [2, 3]  # members kept their ids
    # The group replaced its members at the top level — one shape now.
    assert len(s1.shapes) == 1


def test_group_then_ungroup_round_trips(deck) -> None:  # type: ignore[no-untyped-def]
    s1 = deck.slides[1]
    with deck.edit("t"):
        group = s1.shapes.group([s1.shapes[1], s1.shapes[2]])
    with deck.edit("t"):
        freed = group.ungroup()
    assert sorted(c.shape_id for c in freed) == [2, 3]  # original ids preserved
    assert len(s1.shapes) == 2


def test_group_needs_two_shapes(deck) -> None:  # type: ignore[no-untyped-def]
    s1 = deck.slides[1]
    with pytest.raises(ValueError, match="at least 2"):
        s1.shapes.group([s1.shapes[1]])


def test_ungroup_non_group_is_value_error(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"), pytest.raises(ValueError, match="group shape"):
        deck.slides[1].shapes[1].ungroup()


# -- library: align / distribute --------------------------------------------


def test_align_snaps_lefts_and_records_cmd(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    members = [s2.shapes[1], s2.shapes[2], s2.shapes[3]]
    with deck.edit("t"):
        s2.shapes.align(members, "left")
    lefts = {round(m.geometry()["left"], 3) for m in members}
    assert len(lefts) == 1  # all aligned to one left edge
    assert s2.com.Shapes._align_calls[-1] == (0, -1)  # msoAlignLefts, relative-to-slide


def test_align_middle_uses_vertical_cmd(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with deck.edit("t"):
        s2.shapes.align([s2.shapes[1], s2.shapes[2]], "middle", relative_to="selection")
    assert s2.com.Shapes._align_calls[-1] == (4, 0)  # msoAlignMiddles, relative-to-selection


def test_align_selection_needs_two(deck) -> None:  # type: ignore[no-untyped-def]
    s3 = deck.slides[3]
    with pytest.raises(ValueError, match="selection needs"):
        s3.shapes.align([s3.shapes[1]], "left", relative_to="selection")


def test_distribute_records_cmd(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with deck.edit("t"):
        s2.shapes.distribute([s2.shapes[1], s2.shapes[2], s2.shapes[3]], "horizontal")
    assert s2.com.Shapes._distribute_calls[-1] == (0, -1)  # msoDistributeHorizontally


def test_distribute_needs_three(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with pytest.raises(ValueError, match="at least 3"):
        s2.shapes.distribute([s2.shapes[1], s2.shapes[2]], "horizontal")


# -- library: connectors ----------------------------------------------------


def test_add_connector_attaches_to_two_shapes(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with deck.edit("t"):
        conn = s2.shapes.add_connector("elbow", begin=s2.shapes[1], end=s2.shapes[2])
    d = conn.to_dict()
    assert d["connector"] == {"type": "elbow", "begin_shape_id": 2, "end_shape_id": 3}


def test_add_connector_geometry_form(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with deck.edit("t"):
        conn = s2.shapes.add_connector("straight", left=10, top=10, width=200, height=0)
    d = conn.to_dict()
    assert d["connector"]["type"] == "straight"
    assert d["connector"]["begin_shape_id"] is None  # free-floating


def test_add_connector_needs_both_ends_or_geometry(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with pytest.raises(ValueError, match="BOTH begin"):
        s2.shapes.add_connector("straight", begin=s2.shapes[1])
    with pytest.raises(ValueError, match="explicit"):
        s2.shapes.add_connector("straight")


def test_add_connector_bad_site_rejected(deck) -> None:  # type: ignore[no-untyped-def]
    s2 = deck.slides[2]
    with pytest.raises(ValueError, match="out of range"):
        s2.shapes.add_connector("straight", begin=s2.shapes[1], end=s2.shapes[2], begin_site=99)


def test_non_connector_reads_connector_none(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[1].shapes[1].to_dict()["connector"] is None


# -- batch ------------------------------------------------------------------


def test_batch_group_align_connect(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"):
        out = _edit_core(deck, EditOp.SHAPE_GROUP, {"anchors": ["shape:1:1", "shape:1:2"]})
    assert sorted(out["group_item_ids"]) == [2, 3]
    with deck.edit("t"):
        out = _edit_core(deck, EditOp.SHAPE_ALIGN, {"anchors": "shape:2:1,shape:2:2", "how": "top"})
    assert out["how"] == "top"
    with deck.edit("t"):
        out = _edit_core(
            deck,
            EditOp.SHAPE_ADD_CONNECTOR,
            {"type": "curved", "begin": "shape:2:1", "end": "shape:2:2"},
        )
    assert out["connector"]["type"] == "curved"


def test_batch_align_requires_how(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"), pytest.raises(BatchOpError, match="requires `how`"):
        _edit_core(deck, EditOp.SHAPE_ALIGN, {"anchors": ["shape:2:1", "shape:2:2"]})


# -- CLI --------------------------------------------------------------------


def test_cli_shape_group(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "group", "--anchors", "shape:1:1,shape:1:2"])
    assert result.exit_code == 0
    assert sorted(json.loads(result.output)["group_item_ids"]) == [2, 3]


def test_cli_shape_align(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "align", "--anchors", "shape:2:1,shape:2:2", "--how", "middle"]
    )
    assert result.exit_code == 0
    assert json.loads(result.output)["how"] == "middle"


def test_cli_shape_connect(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["shape", "connect", "--type", "elbow", "--begin", "shape:2:1", "--end", "shape:2:2"],
    )
    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["connector"] == {"type": "elbow", "begin_shape_id": 2, "end_shape_id": 3}


def test_cli_shape_connect_needs_ends_or_slide(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "connect", "--type", "straight"])
    assert result.exit_code == 2  # click UsageError

"""Text-run-level hyperlinks (v-next) — link a span of a text anchor.

Against the fake deck: slide 3 TextBox 1 (shape:3:1) holds "Free text"; slide 2
Content Placeholder 2 (shape:2:2) holds "Intro\\rDemo\\rQ&A" (para:2:2:2 == "Demo").
The COM path was pinned in `scripts/run_link_spike.py` (Characters(...).
ActionSettings(ppMouseClick).Hyperlink round-trips; linking splits the runs).
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import BatchOpError, EditOp, ReadOp, _edit_core, _read_core
from pptlive.cli.main import main
from pptlive.exceptions import AmbiguousMatchError, AnchorNotFoundError, SlideNotFoundError

# -- library: set by substring / offset -------------------------------------


def test_set_link_by_substring(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]  # "Free text"
    with deck.edit("t"):
        link = sh.set_link(text="Free", url="https://x.test")
    assert link == {
        "text": "Free",
        "start": 0,
        "length": 4,
        "address": "https://x.test",
        "sub_address": None,
    }
    assert sh.links() == [
        {
            "text": "Free",
            "start": 0,
            "length": 4,
            "address": "https://x.test",
            "sub_address": None,
        }
    ]


def test_set_link_by_offset(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]  # "Free text" — "text" at offset 5
    with deck.edit("t"):
        link = sh.set_link(start=5, length=4, url="https://y.test")
    assert link["text"] == "text"
    assert sh.to_dict()["links"][0]["address"] == "https://y.test"


def test_set_link_slide_jump(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with deck.edit("t"):
        link = sh.set_link(text="Free", slide=1)
    assert link["sub_address"] == "256,1,Welcome"
    assert link["address"] is None


def test_set_link_on_paragraph_anchor(deck) -> None:  # type: ignore[no-untyped-def]
    para = deck.anchor_by_id("para:2:2:2")  # "Demo"
    with deck.edit("t"):
        para.set_link(text="Demo", url="https://demo.test")
    rows = para.links()
    assert rows == [
        {
            "text": "Demo",
            "start": 0,
            "length": 4,
            "address": "https://demo.test",
            "sub_address": None,
        }
    ]


# -- library: validation ----------------------------------------------------


def test_set_link_requires_one_destination(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with pytest.raises(ValueError, match="exactly one"):
        sh.set_link(text="Free")
    with pytest.raises(ValueError, match="exactly one"):
        sh.set_link(text="Free", url="https://x", slide=1)


def test_set_link_zero_match_is_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with deck.edit("t"), pytest.raises(AnchorNotFoundError):
        sh.set_link(text="zzz", url="https://x")


def test_set_link_ambiguous_match(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]  # "Free text" — "e" appears 3 times
    with deck.edit("t"), pytest.raises(AmbiguousMatchError):
        sh.set_link(text="e", url="https://x")


def test_set_link_offset_out_of_range(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with deck.edit("t"), pytest.raises(ValueError, match="out of range"):
        sh.set_link(start=100, length=5, url="https://x")


def test_set_link_text_and_offset_rejected(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with deck.edit("t"), pytest.raises(ValueError, match="not both"):
        sh.set_link(text="Free", start=0, length=4, url="https://x")


def test_set_link_slide_out_of_range(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with deck.edit("t"), pytest.raises(SlideNotFoundError):
        sh.set_link(text="Free", slide=99)


# -- library: remove --------------------------------------------------------


def test_remove_link_span(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]
    with deck.edit("t"):
        sh.set_link(text="Free", url="https://x")
    with deck.edit("t"):
        n = sh.remove_link(text="Free")
    assert n == 1
    assert sh.links() == []


def test_remove_all_links(deck) -> None:  # type: ignore[no-untyped-def]
    sh = deck.slides[3].shapes[1]  # "Free text"
    with deck.edit("t"):
        sh.set_link(start=0, length=4, url="https://a")  # "Free"
        sh.set_link(start=5, length=4, url="https://b")  # "text"
    assert len(sh.links()) == 2
    with deck.edit("t"):
        n = sh.remove_link()
    assert n == 2
    assert sh.links() == []


# -- batch ------------------------------------------------------------------


def test_batch_link_set_remove_and_read(deck, ppt) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"):
        out = _edit_core(
            deck, EditOp.LINK_SET, {"anchor_id": "shape:3:1", "text": "Free", "url": "https://x"}
        )
    assert out["link"]["text"] == "Free"
    read = _read_core(ppt, ReadOp.LINKS, {"anchor_id": "shape:3:1"})
    assert read["links"][0]["address"] == "https://x"
    with deck.edit("t"):
        out = _edit_core(deck, EditOp.LINK_REMOVE, {"anchor_id": "shape:3:1", "text": "Free"})
    assert out["removed"] == 1


def test_batch_link_set_needs_one_destination(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"), pytest.raises(BatchOpError, match="exactly one"):
        _edit_core(deck, EditOp.LINK_SET, {"anchor_id": "shape:3:1", "text": "Free"})


# -- CLI --------------------------------------------------------------------


def test_cli_link_set_and_list(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    result = runner.invoke(
        main, ["link", "set", "--anchor-id", "shape:3:1", "--text", "Free", "--url", "https://x"]
    )
    assert result.exit_code == 0
    assert json.loads(result.output)["link"]["text"] == "Free"
    listed = runner.invoke(main, ["link", "list", "--anchor-id", "shape:3:1"])
    assert json.loads(listed.output)[0]["address"] == "https://x"


def test_cli_link_set_both_is_usage_error(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["link", "set", "--anchor-id", "shape:3:1", "--text", "Free", "--url", "x", "--slide", "1"],
    )
    assert result.exit_code == 2


def test_cli_link_remove(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    runner.invoke(
        main, ["link", "set", "--anchor-id", "shape:3:1", "--text", "Free", "--url", "https://x"]
    )
    result = runner.invoke(main, ["link", "remove", "--anchor-id", "shape:3:1"])
    assert result.exit_code == 0
    assert json.loads(result.output)["removed"] == 1

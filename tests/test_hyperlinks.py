"""Shape hyperlinks (v0.4.0, the v1.4 cut) — set/read/remove a click action.

Against the fake deck: slide 1 (id 256) title "Welcome"; slide 2 has a frameless
Picture 3 (shape:2:3) proving a link needs no text frame.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive._batch import EditOp, _edit_core
from pptlive.cli.main import main
from pptlive.exceptions import SlideNotFoundError

# -- library: url link ------------------------------------------------------


def test_set_hyperlink_url(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with deck.edit("t"):
        link = shape.set_hyperlink(url="https://anthropic.com/")
    assert link == {"address": "https://anthropic.com/", "sub_address": None}
    assert shape.to_dict()["hyperlink"] == {
        "address": "https://anthropic.com/",
        "sub_address": None,
    }


def test_set_hyperlink_on_frameless_shape(deck) -> None:  # type: ignore[no-untyped-def]
    # Picture 3 (shape:2:3) has no text frame — a link is a shape-level action.
    picture = deck.slides[2].shapes[3]
    assert picture.has_text_frame is False
    with deck.edit("t"):
        picture.set_hyperlink(url="mailto:team@example.com")
    assert picture.to_dict()["hyperlink"]["address"] == "mailto:team@example.com"


def test_set_hyperlink_screen_tip(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with deck.edit("t"):
        shape.set_hyperlink(url="https://x.test", screen_tip="Go to X")
    acts = shape.com.ActionSettings(1)
    assert acts.Hyperlink.ScreenTip == "Go to X"


# -- library: slide jump ----------------------------------------------------


def test_set_hyperlink_slide_jump_subaddress(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with deck.edit("t"):
        link = shape.set_hyperlink(slide=1)
    # "<SlideID>,<index>,<title>" — slide 1 is id 256, position 1, title "Welcome".
    assert link == {"address": None, "sub_address": "256,1,Welcome"}


def test_set_hyperlink_slide_out_of_range_is_not_found(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with deck.edit("t"), pytest.raises(SlideNotFoundError):
        shape.set_hyperlink(slide=99)


# -- library: remove + read default -----------------------------------------


def test_remove_hyperlink(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with deck.edit("t"):
        shape.set_hyperlink(url="https://x.test")
    with deck.edit("t"):
        shape.remove_hyperlink()
    assert shape.to_dict()["hyperlink"] is None


def test_unlinked_shape_reads_hyperlink_none(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[2].shapes[1].to_dict()["hyperlink"] is None


# -- library: validation (before any COM) -----------------------------------


def test_set_hyperlink_requires_exactly_one_destination(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with pytest.raises(ValueError, match="exactly one"):
        shape.set_hyperlink()
    with pytest.raises(ValueError, match="exactly one"):
        shape.set_hyperlink(url="https://x.test", slide=1)


def test_set_hyperlink_blank_url_rejected(deck) -> None:  # type: ignore[no-untyped-def]
    shape = deck.slides[2].shapes[1]
    with pytest.raises(ValueError, match="non-empty"):
        shape.set_hyperlink(url="   ")


# -- batch op ---------------------------------------------------------------


def test_batch_shape_set_and_remove_hyperlink(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("t"):
        out = _edit_core(
            deck,
            EditOp.SHAPE_SET_HYPERLINK,
            {"anchor_id": "shape:2:1", "url": "https://anthropic.com/"},
        )
    assert out["hyperlink"]["address"] == "https://anthropic.com/"
    with deck.edit("t"):
        out = _edit_core(deck, EditOp.SHAPE_REMOVE_HYPERLINK, {"anchor_id": "shape:2:1"})
    assert out["hyperlink"] is None


def test_batch_shape_set_hyperlink_needs_one_destination(deck) -> None:  # type: ignore[no-untyped-def]
    from pptlive._batch import BatchOpError

    with deck.edit("t"), pytest.raises(BatchOpError, match="exactly one"):
        _edit_core(deck, EditOp.SHAPE_SET_HYPERLINK, {"anchor_id": "shape:2:1"})


# -- CLI --------------------------------------------------------------------


def test_cli_shape_set_link_url(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "set-link", "--anchor-id", "shape:2:1", "--url", "https://x.test"]
    )
    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["hyperlink"]["address"] == "https://x.test"


def test_cli_shape_set_link_slide_jump(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "set-link", "--anchor-id", "shape:2:1", "--slide", "1"]
    )
    assert result.exit_code == 0
    payload = json.loads(result.output)
    assert payload["hyperlink"]["sub_address"] == "256,1,Welcome"


def test_cli_shape_set_link_both_is_usage_error(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main,
        ["shape", "set-link", "--anchor-id", "shape:2:1", "--url", "https://x", "--slide", "1"],
    )
    assert result.exit_code == 2  # click UsageError


def test_cli_shape_remove_link(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    runner = CliRunner()
    runner.invoke(main, ["shape", "set-link", "--anchor-id", "shape:2:1", "--url", "https://x"])
    result = runner.invoke(main, ["shape", "remove-link", "--anchor-id", "shape:2:1"])
    assert result.exit_code == 0
    assert json.loads(result.output)["hyperlink"] is None

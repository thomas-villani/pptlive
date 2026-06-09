"""Comments (v1.3): the slide-attached, threaded, identity-bound review channel.

Against the fake, each slide carries a `_FakeCommentCollection`; the default deck
seeds slide 1 with one parent comment + one threaded reply, both bearing a real
`ProviderID`/`UserID`. So reads/threads, identity-sourced `Add2`, the legacy
identity-free `Add` fallback (on a comment-less deck), reply, and delete all
round-trip here without a live PowerPoint.
"""

from __future__ import annotations

import json

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.exceptions import AnchorNotFoundError


def _json(result):  # type: ignore[no-untyped-def]
    return json.loads(result.output)


# -- read (wrapper) ---------------------------------------------------------


def test_slide_comments_read_fields(deck) -> None:  # type: ignore[no-untyped-def]
    comments = deck.slides[1].comments
    assert len(comments) == 1
    c = comments[1]
    assert c.index == 1
    assert c.author == "Thomas Villani"
    assert c.author_initials == "TV"
    assert c.text == "Tighten this headline."
    assert c.datetime == "2026-06-07T10:30:00+00:00"
    assert (c.left, c.top) == (12.0, 12.0)


def test_slide_comments_read_thread(deck) -> None:  # type: ignore[no-untyped-def]
    replies = deck.slides[1].comments[1].replies
    assert [r.text for r in replies] == ["Agreed — will do."]
    assert replies[0].index == 1


def test_comment_to_dict_nests_replies(deck) -> None:  # type: ignore[no-untyped-def]
    d = deck.slides[1].comments[1].to_dict()
    assert d["text"] == "Tighten this headline."
    assert d["author"] == "Thomas Villani"
    assert d["initials"] == "TV"
    assert d["datetime"] == "2026-06-07T10:30:00+00:00"
    assert len(d["replies"]) == 1
    assert d["replies"][0]["text"] == "Agreed — will do."


def test_reply_reports_no_subthread_no_infinite_recursion(deck) -> None:  # type: ignore[no-untyped-def]
    # A live reply's `.Replies` is self-referential (returns the sibling list,
    # which contains the reply), so a naive thread walk recurses forever. A reply
    # must report no replies of its own; the nested dict stops at one level.
    reply = deck.slides[1].comments[1].replies[0]
    assert reply.replies == []
    assert reply.to_dict()["replies"] == []


def test_comment_list_empty_slide(deck) -> None:  # type: ignore[no-untyped-def]
    assert deck.slides[2].comments.list() == []


def test_comment_index_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(AnchorNotFoundError):
        _ = deck.slides[1].comments[5]
    with pytest.raises(AnchorNotFoundError):
        _ = deck.slides[2].comments[1]  # comment-less slide


# -- deck-wide rollup -------------------------------------------------------


def test_deck_comments_rollup(deck) -> None:  # type: ignore[no-untyped-def]
    roll = deck.comments()
    assert roll["total"] == 1
    assert [s["slide"] for s in roll["slides"]] == [1]
    assert roll["slides"][0]["comments"][0]["text"] == "Tighten this headline."


# -- add (modern Add2 vs legacy Add fallback) -------------------------------


def test_add_uses_modern_add2_when_identity_discoverable(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # slide 1 already has a comment, so identity is discoverable deck-wide.
    with deck.edit("add"):
        c = deck.slides[2].comments.add("Please add a source.")
    assert c.text == "Please add a source."
    coll = fake_powerpoint.ActivePresentation.Slides(2).Comments
    assert coll.last_add_method == "Add2"
    # the identity was sourced off slide 1's comment
    assert c.com.ProviderID == "AD"
    assert c.com.UserID == "S::tom@example.com::abc-123"


def test_add_falls_back_to_legacy_on_commentless_deck(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # Remove the only comment so the deck has no identity to source.
    with deck.edit("clear"):
        deck.slides[1].comments[1].delete()
    assert deck.comments()["total"] == 0
    with deck.edit("add"):
        c = deck.slides[3].comments.add("First comment.", author="Bot", initials="B")
    coll = fake_powerpoint.ActivePresentation.Slides(3).Comments
    assert coll.last_add_method == "Add"
    assert c.text == "First comment."
    assert c.author == "Bot"


def test_add_honors_left_top(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("add"):
        c = deck.slides[2].comments.add("Note", left=100.0, top=50.0)
    assert (c.left, c.top) == (100.0, 50.0)


# -- reply ------------------------------------------------------------------


def test_reply_appends_and_lifts_parent_identity(deck) -> None:  # type: ignore[no-untyped-def]
    parent = deck.slides[1].comments[1]
    with deck.edit("reply"):
        rep = parent.reply("On it.")
    assert rep.text == "On it."
    # the new reply inherited the parent's identity + anchor
    assert rep.com.ProviderID == "AD"
    assert (rep.left, rep.top) == (parent.left, parent.top)
    assert [r.text for r in deck.slides[1].comments[1].replies] == [
        "Agreed — will do.",
        "On it.",
    ]


# -- delete -----------------------------------------------------------------


def test_delete_removes_comment_and_thread(deck) -> None:  # type: ignore[no-untyped-def]
    with deck.edit("delete"):
        deck.slides[1].comments[1].delete()
    assert len(deck.slides[1].comments) == 0
    assert deck.comments()["total"] == 0


# -- CLI --------------------------------------------------------------------


def test_cli_comment_list_slide(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "comment", "list", "--slide", "1"])
    assert res.exit_code == 0
    payload = _json(res)
    assert payload[0]["text"] == "Tighten this headline."
    assert payload[0]["replies"][0]["text"] == "Agreed — will do."


def test_cli_comment_list_deck(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "comment", "list"])
    assert res.exit_code == 0
    payload = _json(res)
    assert payload["total"] == 1
    assert payload["slides"][0]["slide"] == 1


def test_cli_comment_add(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "comment", "add", "--slide", "2", "--text", "A note"])
    assert res.exit_code == 0
    payload = _json(res)
    assert payload["ok"] is True
    assert payload["comment"]["text"] == "A note"
    assert fake_powerpoint.ActivePresentation.Slides(2).Comments.Count == 1


def test_cli_comment_reply(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(
        main, ["--json", "comment", "reply", "--slide", "1", "--index", "1", "--text", "Done"]
    )
    assert res.exit_code == 0
    payload = _json(res)
    assert payload["ok"] is True
    assert payload["reply"]["text"] == "Done"


def test_cli_comment_delete(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "comment", "delete", "--slide", "1", "--index", "1"])
    assert res.exit_code == 0
    assert _json(res)["ok"] is True
    assert fake_powerpoint.ActivePresentation.Slides(1).Comments.Count == 0


def test_cli_comment_list_empty_slide(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--text", "comment", "list", "--slide", "2"])
    assert res.exit_code == 0
    assert "(no comments)" in res.output

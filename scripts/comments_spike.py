"""Spike — does the COM ``Slide.Comments`` collection see modern comments?

The roadmap's v1.3 (review loop) carries the one genuine COM risk: PowerPoint's
**modern threaded comments** (post-2019, with replies/mentions, stored in a
separate part) have historically been *invisible* to the classic COM
``Slide.Comments`` collection, which only surfaced the legacy "sticky-note"
comments. If that's still true, v1.3 ships as "read of whatever COM sees + an
honest caveat" rather than full coverage. This spike characterizes the surface:

  1. **Existing reach (read-only)** — walk every slide of the *user's open deck*
     and report ``Comments.Count`` + each comment's fields. If the user has
     modern threaded comments and this reports 0, that's the answer.
  2. **Per-comment shape** — ``Author`` / ``AuthorInitials`` / ``Text`` /
     ``DateTime`` / position, and whether a ``.Replies`` collection exists
     (the modern-threading tell).
  3. **Add surface** — ``Comments.Add2`` (modern) vs ``Comments.Add`` (legacy):
     which exists, what it returns, and whether the added comment exposes
     ``.Replies``. Done on a *temporary* slide, so it's net-zero.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/comments_spike.py

For the most informative result, open a deck that already has a comment or two
(ideally a modern threaded one) before running. Prints one JSON findings object;
net-zero (the add-probe lives on a temp slide deleted in a ``finally``).
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


# Candidate identity/state props on a modern Comment — dumped to learn which
# exist and what ProviderID/UserID the modern threaded API wants.
_ID_PROPS = (
    "ProviderID",
    "ProviderId",
    "UserID",
    "UserId",
    "AuthorIndex",
    "Status",
    "Resolved",
)


def _read_comment(c: Any, *, recurse: bool = True, props: bool = False) -> dict[str, Any]:
    d: dict[str, Any] = {}
    for attr in ("Author", "AuthorInitials", "Text", "DateTime", "Left", "Top"):
        try:
            d[attr] = getattr(c, attr)
        except Exception as exc:
            d[attr] = f"<{_err(exc)}>"
    if props:
        d["id_props"] = {n: _safe(c, n) for n in _ID_PROPS}
    # The modern-threading tell: a .Replies collection — recurse into the thread.
    try:
        replies = c.Replies
        d["has_replies_attr"] = True
        rc = int(replies.Count)
        d["replies_count"] = rc
        if recurse and rc:
            d["replies"] = [
                _read_comment(replies.Item(i), recurse=False, props=props) for i in range(1, rc + 1)
            ]
    except Exception:
        d["has_replies_attr"] = False
    return d


def _safe(obj: Any, name: str) -> Any:
    try:
        return getattr(obj, name)
    except Exception as exc:
        return f"<{_err(exc)}>"


def probe_existing(deck: pl.Presentation) -> list[dict[str, Any]]:
    """Read-only walk of the real deck — what does COM actually see?"""
    rows: list[dict[str, Any]] = []
    for s in range(1, len(deck.slides) + 1):
        try:
            comments = deck.slides[s].com.Comments
            cnt = int(comments.Count)
            items = [_read_comment(comments.Item(i), props=True) for i in range(1, cnt + 1)]
            rows.append({"slide": s, "count": cnt, "comments": items})
        except Exception as exc:
            rows.append({"slide": s, "error": _err(exc)})
    return rows


def _discover_identity(deck: pl.Presentation) -> dict[str, Any]:
    """Pull ProviderID/UserID off the first real modern comment, if any."""
    for s in range(1, len(deck.slides) + 1):
        try:
            comments = deck.slides[s].com.Comments
            if int(comments.Count):
                c = comments.Item(1)
                return {n: _safe(c, n) for n in _ID_PROPS}
        except Exception:
            continue
    return {}


def probe_add(slide: pl.Slide, identity: dict[str, Any]) -> dict[str, Any]:
    """Add a modern parent + reply on a temp slide using discovered IDs (net-zero)."""
    out: dict[str, Any] = {"identity_used": identity}
    sc = slide.com.Comments

    # ProviderID/UserID discovered from the user's real comment (string forms).
    provider = identity.get("ProviderID") or identity.get("ProviderId") or ""
    user = identity.get("UserID") or identity.get("UserId") or ""
    provider = "" if str(provider).startswith("<") else str(provider)
    user = "" if str(user).startswith("<") else str(user)

    added = None
    # Modern Add2(Left, Top, Author, Initials, Text, ProviderID, UserID).
    try:
        added = sc.Add2(12.0, 12.0, "Spike Author", "SA", "Spike parent", provider, user)
        out["add_method"] = "Add2(7-arg)"
    except Exception as exc:
        out["add2_error"] = _err(exc)
        try:  # legacy fallback
            added = sc.Add(12.0, 12.0, "Spike Author", "SA", "Spike parent")
            out["add_method"] = "Add(legacy)"
        except Exception as exc2:
            out["add_error"] = _err(exc2)

    if added is not None:
        out["added_read"] = _read_comment(added, props=True)
        out["count_after_add"] = int(slide.com.Comments.Count)

        # Reply: try a few plausible signatures, report which lands.
        replies = _safe(added, "Replies")
        reply_attempts = [
            (
                "Replies.Add2(Author,Init,Text,Prov,User)",
                lambda: replies.Add2("Spike Author", "SA", "Spike reply", provider, user),
            ),
            (
                "Replies.Add2(L,T,Author,Init,Text,Prov,User)",
                lambda: replies.Add2(0.0, 0.0, "Spike Author", "SA", "Spike reply", provider, user),
            ),
            (
                "Replies.Add(Author,Init,Text)",
                lambda: replies.Add("Spike Author", "SA", "Spike reply"),
            ),
        ]
        out["reply"] = {"ok": False}
        for label, fn in reply_attempts:
            try:
                rep = fn()
                out["reply"] = {
                    "ok": True,
                    "signature": label,
                    "read": _read_comment(rep, recurse=False, props=True),
                    "parent_replies_count": int(added.Replies.Count),
                }
                break
            except Exception as exc:
                out["reply"].setdefault("tried", []).append(
                    {"signature": label, "error": _err(exc)}
                )

        # Net-zero: delete the probe comment (takes its replies with it).
        try:
            added.Delete()
            out["deleted"] = True
            out["count_after_delete"] = int(slide.com.Comments.Count)
        except Exception as exc:
            out["delete_error"] = _err(exc)

    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        # 1. Read-only reach over the real deck.
        try:
            findings["existing"] = probe_existing(deck)
            findings["existing_total"] = sum(
                r.get("count", 0) for r in findings["existing"] if isinstance(r.get("count"), int)
            )
        except Exception as exc:
            findings["existing_error"] = _err(exc)

        # 2. Discover provider/user identity from a real modern comment.
        findings["discovered_identity"] = _discover_identity(deck)

        # 3. Add parent + reply / delete on a temp slide (net-zero).
        temp_ids: list[int] = []
        try:
            with deck.edit("comments spike: build"):
                temp = deck.slides.add(layout="title_and_content")
                temp_ids.append(temp.id)
                sidx = temp.index
            findings["add_probe"] = probe_add(deck.slides[sidx], findings["discovered_identity"])
        except Exception as exc:
            findings["add_probe"] = {"fatal": _err(exc)}
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("comments spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            sid = deck.slides[idx].id
                        except Exception:
                            continue
                        if sid in temp_ids and sid not in deleted:
                            deck.slides[idx].delete()
                            deleted.append(sid)
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Comments — PowerPoint's review-side annotation channel.

The PowerPoint analog of wordlive's `_comments.py`, with the three diffs the
2-D / shared-deck object model forces:

1. **Comments attach to a *slide***, at an `(x, y)` point — not to a text range.
   So the collection hangs off `Slide.comments` (`slide.comments[2]`,
   `slide.comments.add(...)`), 1-based, matching PowerPoint's own
   `Slide.Comments(n)` ordering; a deck-wide roll-up lives on
   `Presentation.comments()`.
2. **Comments are threaded.** A modern comment carries a `.Replies` collection;
   `Comment.replies` reads the thread and `Comment.reply(text)` appends to it.
3. **Adding needs an identity.** The modern `Comments.Add2` requires a
   `ProviderID`/`UserID` pair (the signed-in Office account) — it can't be
   invented. We lift it off any existing comment in the deck
   (`_discover_identity`); on a deck with **no** comments to lift from, `add()`
   falls back to the legacy `Comments.Add` (no IDs needed). A `reply` lifts the
   identity straight off its parent comment.

Honest caveats, both verified on a live deck (2026-06-07,
`scripts/comments_spike.py`) and documented for callers:

- **Authorship can't be spoofed.** `Add2` *ignores* the passed author/initials
  and binds the comment to the account behind `UserID` (only `Text` is honoured)
  — so an agent-authored comment is correctly attributed to the human's account.
  The `author`/`initials` args are **best-effort** (they target the legacy `Add`
  fallback) but on a modern Office build that path *also* binds to the signed-in
  account, so they may be ignored entirely (verified live 2026-06-09).
- **No resolve/reopen.** `Comment.Status` / `.Resolved` are "no longer supported
  by this version" over COM on current builds, so there is no `resolve()` verb
  (unlike wordlive's). Comment *resolution state* is simply not COM-readable.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import AnchorNotFoundError

if TYPE_CHECKING:
    from ._presentation import Presentation
    from ._slides import Slide

#: Default anchor point (points) for a new comment when the caller gives none —
#: the top-left corner area, where PowerPoint drops a fresh review comment.
_DEFAULT_LEFT = 12.0
_DEFAULT_TOP = 12.0


def _dt(value: Any) -> str | None:
    """Render a COM `DateTime` as an ISO-8601 string (or None)."""
    if value is None:
        return None
    try:
        return value.isoformat()
    except Exception:
        return str(value)


def _identity_from_comment(com: Any) -> tuple[str, str] | None:
    """Lift the `(ProviderID, UserID)` pair off a COM comment, or None.

    Both are mandatory for the modern `Add2`; if either is missing/blank (a
    legacy comment, or a COM build that doesn't expose them) we report None and
    the caller falls back to the legacy add path.
    """
    try:
        provider = str(com.ProviderID or "")
        user = str(com.UserID or "")
    except Exception:
        return None
    if provider and user:
        return provider, user
    return None


def _discover_identity(deck: Presentation) -> tuple[str, str] | None:
    """Walk the deck for the first comment and lift its identity, or None.

    The modern `Comments.Add2` needs a `ProviderID`/`UserID` we can't synthesize,
    so we borrow it from any comment already in the deck. None means "no comment
    to borrow from" — `add()` then uses the legacy, identity-free path.
    """
    for slide in deck.slides:
        with _com.translate_com_errors():
            comments = slide.com.Comments
            count = int(comments.Count)
            if count:
                identity = _identity_from_comment(comments.Item(1))
                if identity is not None:
                    return identity
    return None


class Comment:
    """A single review comment on a slide, located by its 1-based slide index.

    A comment carries its author/text/timestamp/anchor and, when modern, a thread
    of `replies`. `reply(text)` appends to the thread; `delete()` removes the
    comment (and its replies). All reads are side-effect-free; `reply`/`delete`
    mutate, so wrap them in `deck.edit(...)` (as the CLI/MCP do).
    """

    def __init__(self, slide: Slide, com: Any, index: int, *, is_reply: bool = False) -> None:
        self._slide = slide
        self._com = com
        self._index = index
        # PowerPoint threads are one level deep, but on a live build a *reply's*
        # `.Replies` returns the sibling-reply list (which contains the reply
        # itself), so recursing into it never terminates. A reply therefore
        # reports no sub-thread — only a top-level comment expands its replies.
        self._is_reply = is_reply

    @property
    def com(self) -> Any:
        """Raw COM `Comment` — escape hatch (replies, mentions, …)."""
        return self._com

    @property
    def index(self) -> int:
        """1-based position in the slide's `Comments` collection (or thread)."""
        return self._index

    @property
    def slide(self) -> Slide:
        return self._slide

    @property
    def author(self) -> str:
        with _com.translate_com_errors():
            return str(self._com.Author or "")

    @property
    def author_initials(self) -> str:
        with _com.translate_com_errors():
            return str(self._com.AuthorInitials or "")

    @property
    def text(self) -> str:
        """The comment body."""
        with _com.translate_com_errors():
            return str(self._com.Text or "")

    @property
    def datetime(self) -> str | None:
        """The comment's timestamp as an ISO-8601 string (tz-aware), or None."""
        with _com.translate_com_errors():
            return _dt(self._com.DateTime)

    @property
    def left(self) -> float:
        """The comment anchor's x position, in points."""
        with _com.translate_com_errors():
            return float(self._com.Left)

    @property
    def top(self) -> float:
        """The comment anchor's y position, in points."""
        with _com.translate_com_errors():
            return float(self._com.Top)

    @property
    def replies(self) -> list[Comment]:
        """The thread's replies, in order (empty for a legacy/unthreaded comment).

        Replies are themselves `Comment`s (same fields), indexed 1-based within
        the thread. PowerPoint threads are one level deep, so a reply reports no
        replies of its own — and crucially a live reply's `.Replies` is
        self-referential (it returns the sibling list, which includes the reply),
        so we must *not* read it, or a thread walk would never terminate.
        """
        if self._is_reply:
            return []
        out: list[Comment] = []
        try:
            with _com.translate_com_errors():
                replies = self._com.Replies
                count = int(replies.Count)
        except Exception:
            return out
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                rep_com = replies.Item(i)
            out.append(Comment(self._slide, rep_com, i, is_reply=True))
        return out

    def reply(
        self, text: str, *, author: str | None = None, initials: str | None = None
    ) -> Comment:
        """Append a threaded reply to this comment; return the new reply `Comment`.

        Uses the modern `Replies.Add2(Left, Top, Author, Initials, Text,
        ProviderID, UserID)`, lifting the `ProviderID`/`UserID` straight off this
        (the parent) comment — so the reply binds to the deck's comment identity.
        The reply inherits the parent's anchor position. `author`/`initials` are
        accepted for symmetry but, like `Add2` on a parent, are overridden by the
        identity account. Wrap in `deck.edit(...)` for the one-Ctrl-Z fence.
        """
        identity = _identity_from_comment(self._com)
        with _com.translate_com_errors():
            left, top = float(self._com.Left), float(self._com.Top)
            replies = self._com.Replies
            provider, user = identity if identity is not None else ("", "")
            rep = replies.Add2(left, top, author or "", initials or "", text, provider, user)
            new_index = int(replies.Count)
        return Comment(self._slide, rep, new_index, is_reply=True)

    def delete(self) -> None:
        """Delete this comment (and its replies). The wrapper is spent."""
        with _com.translate_com_errors():
            self._com.Delete()

    def to_dict(self) -> dict[str, Any]:
        """`{index, author, initials, text, datetime, left, top, replies:[...]}`."""
        with _com.translate_com_errors():
            d: dict[str, Any] = {
                "index": self._index,
                "author": str(self._com.Author or ""),
                "initials": str(self._com.AuthorInitials or ""),
                "text": str(self._com.Text or ""),
                "datetime": _dt(self._com.DateTime),
                "left": float(self._com.Left),
                "top": float(self._com.Top),
            }
        d["replies"] = [r.to_dict() for r in self.replies]
        return d

    def __repr__(self) -> str:
        return f"<Comment {self._index} on slide {self._slide.index} by {self.author!r}>"


class CommentCollection:
    """Indexable, iterable view over a slide's review comments (1-based)."""

    def __init__(self, slide: Slide) -> None:
        self._slide = slide

    @property
    def _com_collection(self) -> Any:
        return self._slide.com.Comments

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def __getitem__(self, index: int) -> Comment:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"comment index must be int, got {type(index).__name__}")
        n = len(self)
        if not (1 <= index <= n):
            raise AnchorNotFoundError("comment", f"{self._slide.index}:{index}")
        with _com.translate_com_errors():
            return Comment(self._slide, self._com_collection(index), index)

    def __iter__(self) -> Iterator[Comment]:
        with _com.translate_com_errors():
            count = int(self._com_collection.Count)
        for i in range(1, count + 1):
            with _com.translate_com_errors():
                com = self._com_collection(i)
            yield Comment(self._slide, com, i)

    def add(
        self,
        text: str,
        *,
        left: float = _DEFAULT_LEFT,
        top: float = _DEFAULT_TOP,
        author: str | None = None,
        initials: str | None = None,
    ) -> Comment:
        """Add a new top-level comment to the slide; return the new `Comment`.

        Prefers the modern threaded `Comments.Add2(Left, Top, Author, Initials,
        Text, ProviderID, UserID)`, sourcing the `ProviderID`/`UserID` from any
        comment already in the deck. On a deck with **no** comment to borrow an
        identity from, falls back to the legacy `Comments.Add(Left, Top, Author,
        Initials, Text)` (identity-free; may produce a non-threaded comment).

        `author`/`initials` are best-effort (they target the legacy path) — `Add2`
        binds to the identity account, and on a modern Office build even the legacy
        `Add` does, so they may be ignored (see the module docstring). `left`/`top`
        are the anchor point in points. Wrap in `deck.edit(...)` for the one-Ctrl-Z
        fence.
        """
        identity = _discover_identity(self._slide._deck)
        with _com.translate_com_errors():
            comments = self._com_collection
            if identity is not None:
                provider, user = identity
                com = comments.Add2(
                    float(left), float(top), author or "", initials or "", text, provider, user
                )
            else:
                com = comments.Add(float(left), float(top), author or "", initials or "", text)
            index = int(comments.Count)
        return Comment(self._slide, com, index)

    def list(self) -> list[dict[str, Any]]:
        """All top-level comments (with their reply threads) as dicts."""
        return [c.to_dict() for c in self]

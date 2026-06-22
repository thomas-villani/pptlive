"""LLM-friendly fuzzy plain-text find/replace — the matching core.

Ported almost verbatim from wordlive's `_findreplace.py`: the matching is
forgiving of cosmetic differences that show up when an LLM re-emits text it read
off a slide — smart quotes, dashes, NBSP, NFKC variants, whitespace runs — but
produces *original* character offsets so the actual PowerPoint `TextRange` can be
edited without disturbing the surrounding formatting.

This module is the pure, OS-independent half: `find_matches` locates a needle in
a haystack string and returns `Match`es in the haystack's original-offset
coordinates. The PowerPoint-side traversal that feeds it each text frame's text —
and writes replacements back through `TextRange.Characters` — lives on
`Presentation.find` / `find_replace` in `_presentation.py`.

It is *not* style-aware: replacing a span inherits the formatting of the first
character of the match, which is what PowerPoint's own Find/Replace does too. For
structured edits ("rewrite this paragraph but keep these runs bold"), reach for
the raw `.com` `TextRange` API instead. PowerPoint separates paragraphs with
`\\r` (and soft line breaks with `\\v`); both fold to whitespace here so a needle
can match across a bullet break.

The fold-table keys are written as `\\u`/`\\x` escapes (not literal glyphs) so the
exotic code points survive any source round-trip intact.
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass

_QUOTE_FOLDS = {
    "‘": "'",  # left single quote
    "’": "'",  # right single quote / apostrophe
    "‚": "'",  # single low-9 quote
    "‛": "'",  # single high-reversed-9 quote
    "“": '"',  # left double quote
    "”": '"',  # right double quote
    "„": '"',  # double low-9 quote
    "‟": '"',  # double high-reversed-9 quote
    "′": "'",  # prime
    "″": '"',  # double prime
    "«": '"',  # left guillemet
    "»": '"',  # right guillemet
}

_DASH_FOLDS = {
    "‐": "-",  # hyphen
    "‑": "-",  # non-breaking hyphen
    "‒": "-",  # figure dash
    "–": "-",  # en dash
    "—": "-",  # em dash
    "―": "-",  # horizontal bar
    "−": "-",  # minus sign
}

# NFKC (applied first in `_normalize`) already folds most width/space variants to
# an ASCII space; these escaped entries cover the no-break and exotic spaces
# explicitly, plus PowerPoint's own break characters.
_SPACE_FOLDS = {
    " ": " ",  # no-break space
    " ": " ",  # figure space
    " ": " ",  # narrow no-break space
    " ": " ",  # thin space
    " ": " ",  # hair space
    "\t": " ",
    "\v": " ",  # PowerPoint soft line break
    "\f": " ",
    "\r": "\n",  # PowerPoint paragraph break
}


def _fold_char(ch: str) -> str:
    """Map one character to its fuzzy-match equivalent."""
    if ch in _QUOTE_FOLDS:
        return _QUOTE_FOLDS[ch]
    if ch in _DASH_FOLDS:
        return _DASH_FOLDS[ch]
    if ch in _SPACE_FOLDS:
        return _SPACE_FOLDS[ch]
    return ch


@dataclass(frozen=True)
class _Normalized:
    """A normalized string and a mapping back to the original offsets.

    `text[i]` is the normalized character; `offsets[i]` is the *first* original
    offset that contributed to it. `offsets[len(text)]` is the original offset
    immediately after the last contributing character — i.e. you can use
    `offsets[match_start]` and `offsets[match_end]` directly as a span into the
    original text.
    """

    text: str
    offsets: list[int]


def _normalize(s: str, *, collapse_whitespace: bool = True) -> _Normalized:
    """NFKC + character folds + (optional) whitespace collapse.

    Tracks original offsets so a match in the normalized string maps cleanly
    back to a span over the original text.
    """
    out_chars: list[str] = []
    out_offsets: list[int] = []
    prev_space = True  # collapse leading whitespace, like text.strip() lite

    for i, raw_ch in enumerate(s):
        # NFKC may expand one char into several (e.g. ligatures). Each output
        # char shares the same source offset.
        decomposed = unicodedata.normalize("NFKC", raw_ch)
        for ch in decomposed:
            folded = _fold_char(ch)
            if folded == "":
                continue
            for fch in folded:
                is_space = fch in (" ", "\n")
                if collapse_whitespace and is_space:
                    if prev_space:
                        continue
                    out_chars.append(" ")
                    out_offsets.append(i)
                    prev_space = True
                else:
                    out_chars.append(fch)
                    out_offsets.append(i)
                    prev_space = False

    # Trailing space is harmless but ugly; strip it.
    while collapse_whitespace and out_chars and out_chars[-1] == " ":
        out_chars.pop()
        out_offsets.pop()

    # Sentinel: one-past-the-last retained source char so callers can use
    # offsets[end] as a half-open right boundary even when the match runs to the
    # end of the normalized string. It must NOT be len(s): the folds map the
    # paragraph mark \r -> \n and the soft break \v -> space, and trailing
    # whitespace is then stripped, so len(s) can point past those dropped chars.
    # A match ending at the last retained char would then span across the
    # stripped \r and fuse the paragraph into the next one on replace. Anchor the
    # sentinel to the last retained char's source offset + 1 instead.
    out_offsets.append(out_offsets[-1] + 1 if out_offsets else len(s))
    return _Normalized(text="".join(out_chars), offsets=out_offsets)


@dataclass(frozen=True)
class Match:
    """A located occurrence of a `find` string inside a text frame's text.

    `start` / `end` are 0-based offsets into the frame's text (the haystack),
    measured against the original (un-normalized) string. `text` is the actual
    original substring at those offsets — useful both for round-tripping
    formatting and for showing the user what was matched.
    """

    start: int
    end: int
    text: str


def find_matches(haystack: str, needle: str) -> list[Match]:
    """Locate every fuzzy occurrence of `needle` inside `haystack`.

    Both sides are normalized identically (NFKC, smart-quote / dash / NBSP
    folds, whitespace collapse). Returns non-overlapping matches in
    original-offset coordinates of `haystack`. Empty `needle` returns no matches.
    """
    if not needle:
        return []
    norm_h = _normalize(haystack)
    norm_n = _normalize(needle)
    if not norm_n.text:
        return []

    matches: list[Match] = []
    i = 0
    nlen = len(norm_n.text)
    while True:
        j = norm_h.text.find(norm_n.text, i)
        if j == -1:
            break
        start = norm_h.offsets[j]
        end = norm_h.offsets[j + nlen]
        matches.append(Match(start=start, end=end, text=haystack[start:end]))
        i = j + nlen
    return matches

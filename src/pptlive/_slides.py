"""Slides — the ordered containers, and the second level of the anchor hierarchy.

A `Slide` is a *container*, not a text anchor (like a whole table in wordlive):
slide-level verbs live here, while text lives on its shapes/placeholders/notes.
`deck.slides[S]` is 1-based. A slide exposes both its 1-based `index` (what users
say) and its stable `id` (`SlideID`, survives reordering) so listings can be
re-identified after a move.

Slide lifecycle — `SlideCollection.add()` and `Slide.delete/duplicate/move_to/
set_layout` — is the v0.1 track (the first with no Word analog). These verbs only
mutate; wrap a call in `deck.edit(label)` (as the CLI does) for view preservation
and a one-Ctrl-Z fence. Layout names resolve to a `CustomLayout` via
`Presentation._resolve_layout` (see `constants.match_layout_name`).
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterator
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Notes
from ._comments import CommentCollection
from ._headersfooters import HeadersFooters
from ._shapes import (
    PlaceholderShape,
    Shape,
    ShapeCollection,
    background_to_dict,
    effect_to_dict,
    is_placeholder,
)
from .constants import (
    DEFAULT_LEGACY_LAYOUT,
    MsoTriState,
    entry_effect_for,
    entry_effect_name,
    image_filter_for,
    is_true,
    parse_color,
    placeholder_types_for,
)
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    LayoutNotFoundError,
    PowerPointBusyError,
    SlideNotFoundError,
)

if TYPE_CHECKING:
    from ._presentation import Presentation


_PLACEHOLDER_GEOMETRY_KEYS = frozenset({"left", "top", "width", "height"})


def _validate_placeholders_arg(placeholders: dict[str, dict[str, float]] | None) -> None:
    """Guard the `add(placeholders=...)` map before any COM work (clean ValueError).

    Each value must be a dict with at least one of left/top/width/height (numbers);
    unknown keys are rejected so a typo (`with`) fails loudly rather than silently.
    """
    if placeholders is None:
        return
    if not isinstance(placeholders, dict):
        raise ValueError("placeholders must be a dict of {kind: {left/top/width/height}}")
    for kind, geo in placeholders.items():
        if not isinstance(geo, dict) or not geo:
            raise ValueError(f"placeholders[{kind!r}] must be a non-empty geometry dict")
        unknown = set(geo) - _PLACEHOLDER_GEOMETRY_KEYS
        if unknown:
            raise ValueError(
                f"placeholders[{kind!r}] has unknown key(s) {sorted(unknown)}; "
                "allowed: left, top, width, height"
            )
        for key, val in geo.items():
            if isinstance(val, bool) or not isinstance(val, (int, float)):
                raise ValueError(f"placeholders[{kind!r}][{key!r}] must be a number (points)")


def _paragraphs(text: str) -> list[str]:
    """Split a PowerPoint `TextRange.Text` into non-empty paragraph strings.

    PowerPoint separates paragraphs with `\\r` and soft line breaks with `\\v`;
    `\\n` shows up via COM too. We split on all three and drop blank lines so an
    outline read stays tidy.
    """
    out: list[str] = []
    for chunk in text.replace("\v", "\r").replace("\n", "\r").split("\r"):
        stripped = chunk.strip()
        if stripped:
            out.append(stripped)
    return out


class Slide:
    """Wraps a PowerPoint `Slide` COM object."""

    def __init__(self, deck: Presentation, slide_com: Any) -> None:
        self._deck = deck
        self._slide = slide_com

    @property
    def com(self) -> Any:
        return self._slide

    @property
    def index(self) -> int:
        """1-based position in the deck (`SlideIndex`). Shifts when slides move."""
        with _com.translate_com_errors():
            return int(self._slide.SlideIndex)

    @property
    def id(self) -> int:
        """Stable `SlideID` — survives reordering, unlike `index`."""
        with _com.translate_com_errors():
            return int(self._slide.SlideID)

    @property
    def deck(self) -> Presentation:
        """The owning `Presentation` (e.g. for resolving sibling slides)."""
        return self._deck

    @property
    def shapes(self) -> ShapeCollection:
        return ShapeCollection(self)

    @property
    def notes(self) -> Notes:
        """The speaker-notes anchor (`notes:S`)."""
        return Notes(self)

    @property
    def comments(self) -> CommentCollection:
        """The slide's review comments (`comments:S`) — read + add/reply/delete."""
        return CommentCollection(self)

    @property
    def headers_footers(self) -> HeadersFooters:
        """This slide's footer / slide-number / date placeholders (a per-slide
        override of the master default). See `_headersfooters.HeadersFooters`.
        """
        with _com.translate_com_errors():
            return HeadersFooters(self._slide.HeadersFooters)

    @property
    def layout_name(self) -> str | None:
        """The slide's custom-layout name (e.g. "Title and Content"), or None."""
        try:
            with _com.translate_com_errors():
                return str(self._slide.CustomLayout.Name)
        except PowerPointBusyError:
            # A transient busy must surface (exit 3, retryable), not masquerade
            # as "this slide has no layout".
            raise
        except Exception:
            return None

    def _find_placeholder(self, kind: str) -> tuple[Any, int]:
        """Resolve a placeholder KIND to (COM shape, 1-based z-order index).

        Picks the accepted `PpPlaceholderType` of highest preference (see
        `constants._PLACEHOLDER_KINDS`); a more-preferred type wins over a
        less-preferred one (so `body` prefers a real BODY over a generic OBJECT).
        But when **two or more** placeholders share the *same* best-preference
        type — e.g. the two OBJECT bodies of a Two Content / Comparison layout —
        the kind is genuinely ambiguous: rather than silently pick the first,
        raise `AmbiguousMatchError` (exit 5) listing the candidate `shape:S:N`
        anchors so the caller targets one explicitly. Raises `AnchorNotFoundError`
        if no matching placeholder exists, `ValueError` for a bad KIND.
        """
        accepted = placeholder_types_for(kind)  # ValueError on unknown kind
        accepted_ints = [int(t) for t in accepted]
        # Collect every accepted placeholder as (rank, idx, com_shape). The outer
        # translate covers the Shapes iteration + is_placeholder reads; the inner
        # one lets a transient PowerPointBusyError on a single Type read propagate
        # (it must not be swallowed as "skip this placeholder") while a genuinely
        # unreadable placeholder type is still skipped.
        matches: list[tuple[int, int, Any]] = []
        with _com.translate_com_errors():
            for idx, sh in enumerate(self._slide.Shapes, start=1):
                if not is_placeholder(sh):
                    continue
                try:
                    with _com.translate_com_errors():
                        ph_type = int(sh.PlaceholderFormat.Type)
                except PowerPointBusyError:
                    raise
                except Exception:
                    continue
                if ph_type in accepted_ints:
                    matches.append((accepted_ints.index(ph_type), idx, sh))
        if not matches:
            raise AnchorNotFoundError("placeholder", f"ph:{self.index}:{kind.lower()}")
        best_rank = min(rank for rank, _idx, _sh in matches)
        tied = [(idx, sh) for rank, idx, sh in matches if rank == best_rank]
        if len(tied) > 1:
            candidates = [
                {
                    "anchor_id": f"shape:{self.index}:{idx}",
                    "name": str(sh.Name),
                    "id": int(sh.Id),
                    "index": idx,
                }
                for idx, sh in tied
            ]
            raise AmbiguousMatchError.for_placeholder(f"ph:{self.index}:{kind.lower()}", candidates)
        idx, sh = tied[0]
        return sh, idx

    def placeholder(self, kind: str) -> PlaceholderShape:
        """Return the `ph:S:KIND` placeholder anchor (resolved live by kind).

        KIND ∈ title, ctrtitle, subtitle, body, footer, date, slidenum. Raises
        `AnchorNotFoundError` if the slide has no such placeholder.
        """
        # Resolve once now so a missing placeholder fails fast with a clean error;
        # the returned anchor still re-resolves live on each use.
        with _com.translate_com_errors():
            self._find_placeholder(kind)
        return PlaceholderShape(self, kind)

    @property
    def title(self) -> str | None:
        """Text of the slide's title placeholder, or None if it has no title."""
        try:
            with _com.translate_com_errors():
                shapes = self._slide.Shapes
                if not is_true(shapes.HasTitle):
                    return None
                title_shape = shapes.Title
                if not is_true(title_shape.HasTextFrame):
                    return None
                return str(title_shape.TextFrame.TextRange.Text or "")
        except PowerPointBusyError:
            # Don't let a momentary busy read back as "no title".
            raise
        except Exception:
            return None

    def has_notes(self) -> bool:
        """Whether the slide has non-empty speaker notes."""
        try:
            return bool(self.notes.text.strip())
        except PowerPointBusyError:
            # A transient busy is retryable (exit 3); don't bury it as "no notes".
            raise
        except Exception:
            # A missing notes body (AnchorNotFoundError) or any COM hiccup just
            # means "no notes" for the purpose of a listing.
            return False

    def read(self) -> dict[str, Any]:
        """Every shape on the slide plus its metadata — the `slide read S` payload."""
        return {
            "index": self.index,
            "id": self.id,
            "layout": self.layout_name,
            "title": self.title,
            "transition": self.transition(),
            "background": self.background(),
            "animations": self.animations(),
            "shapes": self.shapes.list(),
        }

    def transition(self) -> dict[str, Any]:
        """The slide's entrance transition — `{effect, duration, advance_on_click,
        advance_on_time, advance_time}`.

        `effect` is the friendly `PpEntryEffect` name (`"fade"`, `"none"`, …);
        `duration` is the transition animation length in seconds; the `advance_*`
        fields describe auto-advance (`advance_on_time` + `advance_time` seconds)
        vs. click-to-advance (`advance_on_click`). A read — no view move.
        """
        with _com.translate_com_errors():
            t = self._slide.SlideShowTransition
            return {
                "effect": entry_effect_name(int(t.EntryEffect)),
                "duration": float(t.Duration),
                "advance_on_click": is_true(t.AdvanceOnClick),
                "advance_on_time": is_true(t.AdvanceOnTime),
                "advance_time": float(t.AdvanceTime),
            }

    def animations(self) -> list[dict[str, Any]]:
        """The slide's shape animations, in play order — one row per effect.

        Reads `Slide.TimeLine.MainSequence`: each row is `{seq_index, shapeid,
        shape, effect, exit, trigger, duration, delay}` (see `effect_to_dict`),
        ordered by `seq_index` (the 1-based position the effect plays in). The
        `shapeid` maps each effect back to its target shape (drift-proof), so an
        agent can tell *what* animates *how* without a render. Empty when the slide
        has no animations. A read — no view move.
        """
        idx = self.index
        with _com.translate_com_errors():
            seq = self._slide.TimeLine.MainSequence
            count = int(seq.Count)
            return [{"seq_index": i, **effect_to_dict(seq(i), idx)} for i in range(1, count + 1)]

    def clear_animations(self, anchor: Shape | None = None) -> int:
        """Remove animation effects from the slide; return how many were deleted.

        With `anchor=None` (the default) wipes the **whole** slide's animation
        sequence; pass a `Shape` to remove only the effects targeting that shape
        (matched by stable `Shape.Id`, so a restack is irrelevant). Deletes from the
        end of the sequence so the live indices don't shift mid-loop. A no-op
        (returns 0) when there's nothing to remove. A mutation: wrap in
        `deck.edit(...)`.
        """
        target_id = None if anchor is None else anchor.shape_id  # COM read before the loop
        with _com.translate_com_errors():
            seq = self._slide.TimeLine.MainSequence
            removed = 0
            for i in range(int(seq.Count), 0, -1):
                eff = seq(i)
                if target_id is None or int(eff.Shape.Id) == target_id:
                    eff.Delete()
                    removed += 1
            return removed

    def background(self) -> dict[str, Any]:
        """The slide's background — `{follows_master, type, color}`.

        `follows_master` is True when the slide inherits the master/layout
        background (the default); when False it carries its own `{type, color}`
        fill (set via `set_background`). A read — no view move.
        """
        with _com.translate_com_errors():
            follows = is_true(self._slide.FollowMasterBackground)
            bg = background_to_dict(self._slide)
        return {"follows_master": follows, **bg}

    def geometry_report(self) -> dict[str, Any]:
        """A geometry-only spatial map of the slide — catch overlaps and off-slide
        shapes *before* rendering.

        Returns the slide size (points) and, per shape, its bounding `box`
        (`left`/`top`/`right`/`bottom`/`width`/`height`) plus an `off_slide` flag,
        then the list of `overlaps` (shape pairs whose boxes intersect, largest
        area first) and the `off_slide` anchor ids. The point of it is the feedback
        loop the snapshot can't give cheaply: an agent that just placed an arrow or
        a card can see "shape:5:3 overlaps shape:5:4" or "the arrow runs off the
        right edge" without a render round-trip or float math.

        Pure axis-aligned geometry on the boxes PowerPoint reports — shape
        **rotation is not accounted for** (the box is the unrotated extent), so a
        rotated shape's overlap / bounds are approximate (each shape carries its
        `rotation` so the caller can judge). A read — no view move.
        """
        with _com.translate_com_errors():
            ps = self._deck.com.PageSetup
            width, height = float(ps.SlideWidth), float(ps.SlideHeight)
        boxes: list[tuple[dict[str, Any], float, float, float, float]] = []
        shapes_out: list[dict[str, Any]] = []
        for s in self.shapes.list():
            geo = s.get("geometry")
            if not geo:
                continue
            left, top = float(geo["left"]), float(geo["top"])
            w, h = float(geo["width"]), float(geo["height"])
            right, bottom = left + w, top + h
            off = left < 0 or top < 0 or right > width or bottom > height
            entry = {
                "anchor_id": s["anchor_id"],
                "shapeid": s.get("shapeid"),
                "name": s["name"],
                "id": s["id"],
                "box": {
                    "left": left,
                    "top": top,
                    "right": right,
                    "bottom": bottom,
                    "width": w,
                    "height": h,
                },
                "rotation": float(geo.get("rotation", 0.0)),
                "off_slide": off,
            }
            shapes_out.append(entry)
            boxes.append((entry, left, top, right, bottom))
        overlaps: list[dict[str, Any]] = []
        for i in range(len(boxes)):
            ei, l1, t1, r1, b1 = boxes[i]
            for j in range(i + 1, len(boxes)):
                ej, l2, t2, r2, b2 = boxes[j]
                ix = min(r1, r2) - max(l1, l2)
                iy = min(b1, b2) - max(t1, t2)
                if ix > 0 and iy > 0:
                    overlaps.append(
                        {
                            "a": ei["anchor_id"],
                            "b": ej["anchor_id"],
                            "a_name": ei["name"],
                            "b_name": ej["name"],
                            "area": ix * iy,
                        }
                    )
        overlaps.sort(key=lambda o: o["area"], reverse=True)
        return {
            "slide": self.index,
            "slide_size": {"width": width, "height": height},
            "shapes": shapes_out,
            "overlaps": overlaps,
            "off_slide": [s["anchor_id"] for s in shapes_out if s["off_slide"]],
        }

    # -- render (v0.4; a read — no mutation, polite by nature) -----------------

    def export_image(
        self,
        path: str | os.PathLike[str] | None = None,
        *,
        width: int | None = None,
        height: int | None = None,
        fmt: str = "png",
    ) -> Path:
        """Render the slide to an image file and return its absolute path.

        Wraps `Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)`. The
        export renders the slide's **current in-memory state** — unsaved edits
        included — so an agent can edit over COM and immediately *see* the result;
        and it's polite (it doesn't move the viewed slide or change the Selection).

        `fmt` is a friendly token (`png`/`jpg`/`gif`/`bmp`/`tiff`; see
        `constants.IMAGE_FORMAT_CHOICES`). When `path` is None a temp file is
        created (so export-then-read is one step). `width`/`height` are output
        **pixels**; pass one and the other is filled from the slide's aspect
        ratio, pass neither for the slide's native pixel size. A relative `path`
        is resolved to absolute first — PowerPoint otherwise drops the file in
        its own working directory, not the caller's.
        """
        filter_name, ext = image_filter_for(fmt)  # ValueError before any COM
        if path is None:
            fd, tmp = tempfile.mkstemp(prefix="pptlive_slide_", suffix=f".{ext}")
            os.close(fd)
            os.remove(tmp)  # hand PowerPoint a clean path to write
            abs_path = tmp
        else:
            abs_path = os.path.abspath(os.fspath(path))
        with _com.translate_com_errors():
            w, h = self._export_dims(width, height)
            if w is not None and h is not None:
                self._slide.Export(abs_path, filter_name, int(round(w)), int(round(h)))
            else:
                self._slide.Export(abs_path, filter_name)
        return Path(abs_path)

    def _export_dims(
        self, width: int | None, height: int | None
    ) -> tuple[float | None, float | None]:
        """Resolve requested export pixels, filling a missing dimension from the
        slide's point aspect ratio so a single `width`/`height` keeps proportions."""
        if width is None and height is None:
            return None, None
        if width is not None and height is not None:
            return float(width), float(height)
        ps = self._deck.com.PageSetup
        sw, sh = float(ps.SlideWidth), float(ps.SlideHeight)
        if width is None:
            return float(height) * (sw / sh), float(height)  # type: ignore[arg-type]
        return float(width), float(width) * (sh / sw)

    # -- lifecycle (v0.1; wrap in deck.edit(...) for view + one-Ctrl-Z) --------

    def delete(self) -> None:
        """Delete this slide from the deck (`Slide.Delete`). The wrapper is spent."""
        with _com.translate_com_errors():
            self._slide.Delete()

    def duplicate(self) -> Slide:
        """Duplicate this slide; return the copy (inserted immediately after).

        Wraps `Slide.Duplicate`, which yields a one-item `SlideRange`. The copy
        gets a fresh `SlideID`; everything after the original shifts down by one.
        """
        with _com.translate_com_errors():
            new_range = self._slide.Duplicate()
            new_com = new_range(1)
        return Slide(self._deck, new_com)

    def move_to(self, index: int) -> Slide:
        """Move this slide to 1-based position `index` (`Slide.MoveTo`); return self.

        The wrapper keeps pointing at the same slide, which now reports the new
        `index`. Raises `SlideNotFoundError` if `index` is out of range (1..count).
        """
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"index must be int, got {type(index).__name__}")
        count = len(self._deck.slides)
        if index < 1 or index > count:
            raise SlideNotFoundError(index)
        with _com.translate_com_errors():
            self._slide.MoveTo(index)
        return self

    def set_layout(self, layout: str | int) -> Slide:
        """Re-apply a slide layout by friendly name or 1-based index; return self.

        Resolves `layout` to a `CustomLayout` (see `Presentation._resolve_layout`)
        and assigns `Slide.CustomLayout`. Raises `LayoutNotFoundError` (listing
        the deck's layout names) for an unknown layout.
        """
        custom = self._deck._resolve_layout(layout)
        if custom is None:
            raise LayoutNotFoundError(str(layout), [])
        with _com.translate_com_errors():
            self._slide.CustomLayout = custom
        return self

    def set_transition(
        self,
        effect: str | int | None = None,
        *,
        duration: float | None = None,
        advance_after: float | None = None,
        advance_on_click: bool | None = None,
    ) -> dict[str, Any]:
        """Set the slide's entrance transition; return the resulting transition dict.

        `effect` is a friendly `PpEntryEffect` name (`"fade"`, `"cut"`,
        `"cover_left"`, … — see `constants.ENTRY_EFFECT_CHOICES`) or a raw int.
        `duration` is the transition animation length in seconds. `advance_after`
        is the auto-advance delay in seconds — passing it sets **both**
        `AdvanceOnTime=msoTrue` and `AdvanceTime` (the spike confirmed both are
        needed); pass `0` to keep the timing but require a click. `advance_on_click`
        toggles click-to-advance independently. Only the kwargs passed are written.

        Raises `ValueError` (before any COM) for an unknown effect name or if
        nothing is passed. A mutation: wrap in `deck.edit(...)`.
        """
        if (
            effect is None
            and duration is None
            and advance_after is None
            and advance_on_click is None
        ):
            raise ValueError(
                "set_transition() requires at least one of effect, duration, "
                "advance_after, or advance_on_click"
            )
        effect_int = entry_effect_for(effect) if effect is not None else None  # ValueError first
        with _com.translate_com_errors():
            t = self._slide.SlideShowTransition
            if effect_int is not None:
                t.EntryEffect = effect_int
            if duration is not None:
                t.Duration = float(duration)
            if advance_after is not None:
                # Auto-advance needs the flag AND the seconds (spike finding).
                t.AdvanceOnTime = int(MsoTriState.TRUE)
                t.AdvanceTime = float(advance_after)
            if advance_on_click is not None:
                t.AdvanceOnClick = (
                    int(MsoTriState.TRUE) if advance_on_click else int(MsoTriState.FALSE)
                )
        return self.transition()

    def set_background(self, color: str | int | tuple[int, int, int]) -> dict[str, Any]:
        """Give the slide its own solid background color; return the background dict.

        The per-slide override of the deck-wide master background (`deck.master.
        set_background`). `color` is `"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB
        int. Sets `FollowMasterBackground=msoFalse` then a solid fill of that color.
        Raises `ValueError` for a bad color (before any COM). Revert with
        `follow_master_background()`. A mutation: wrap in `deck.edit(...)`.
        """
        rgb = parse_color(color)  # ValueError before any COM
        with _com.translate_com_errors():
            self._slide.FollowMasterBackground = int(MsoTriState.FALSE)
            fill = self._slide.Background.Fill
            fill.Solid()
            fill.ForeColor.RGB = rgb
        return self.background()

    def follow_master_background(self) -> dict[str, Any]:
        """Drop any per-slide background override and inherit the master's again.

        Sets `FollowMasterBackground=msoTrue` (the spike-verified revert). A
        mutation: wrap in `deck.edit(...)`.
        """
        with _com.translate_com_errors():
            self._slide.FollowMasterBackground = int(MsoTriState.TRUE)
        return self.background()

    def __repr__(self) -> str:
        return f"<Slide index={self.index}>"


class SlideCollection:
    """Indexable, iterable view over a presentation's slides (1-based)."""

    def __init__(self, deck: Presentation) -> None:
        self._deck = deck

    @property
    def _com_collection(self) -> Any:
        return self._deck.com.Slides

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def __getitem__(self, index: int) -> Slide:
        if isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"slide index must be int, got {type(index).__name__}")
        count = len(self)
        if index < 1 or index > count:
            raise SlideNotFoundError(index)
        with _com.translate_com_errors():
            slide_com = self._com_collection(index)
        return Slide(self._deck, slide_com)

    def __iter__(self) -> Iterator[Slide]:
        with _com.translate_com_errors():
            slides = list(self._com_collection)
        for slide_com in slides:
            yield Slide(self._deck, slide_com)

    def add(
        self,
        layout: str | int | None = None,
        index: int | None = None,
        *,
        placeholders: dict[str, dict[str, float]] | None = None,
    ) -> Slide:
        """Insert a new slide and return it (v0.1; wrap in `deck.edit(...)`).

        `layout` is a friendly name or 1-based layout index (default
        `title_and_content`); `index` is the 1-based insertion position
        (default: appended to the end). Prefers the modern
        `Slides.AddSlide(Index, CustomLayout)`, falling back to legacy
        `Slides.Add` only on a deck that exposes no custom layouts. Raises
        `LayoutNotFoundError` for an unknown layout and `SlideNotFoundError`
        for an out-of-range insertion position (1..count+1).

        `placeholders` repositions the layout's placeholders right after creation
        — `{KIND: {left, top, width, height}}` in points, any subset of the four
        keys per KIND — so "a content slide with the body on the left half" is one
        op instead of an add-then-resize fix-up. KIND is the same semantic name
        `ph:S:KIND` uses (`title`/`body`/…); an unknown-on-this-layout or ambiguous
        KIND raises (`AnchorNotFoundError` / `AmbiguousMatchError`) the same way
        addressing it would. Pair with `Slide.geometry_report()` to size the boxes.
        """
        _validate_placeholders_arg(placeholders)
        count = len(self)
        if index is None:
            target = count + 1
        elif isinstance(index, bool) or not isinstance(index, int):
            raise TypeError(f"index must be int, got {type(index).__name__}")
        elif index < 1 or index > count + 1:
            raise SlideNotFoundError(index)
        else:
            target = index
        with _com.translate_com_errors():
            custom = self._deck._resolve_layout(layout)
            if custom is not None:
                new_com = self._com_collection.AddSlide(target, custom)
            else:
                new_com = self._com_collection.Add(target, int(DEFAULT_LEGACY_LAYOUT))
        new_slide = Slide(self._deck, new_com)
        if placeholders:
            self._apply_placeholder_geometry(new_slide.index, placeholders)
        return new_slide

    def _apply_placeholder_geometry(
        self, slide_index: int, placeholders: dict[str, dict[str, float]]
    ) -> None:
        """Move/resize the new slide's placeholders by semantic KIND (points).

        Resolves every requested KIND *first*, so an unknown-on-this-layout or
        ambiguous KIND raises (`AnchorNotFoundError` / `AmbiguousMatchError`)
        before any placeholder is moved — otherwise a typo'd KIND late in the
        dict would leave the freshly-added slide half-positioned.
        """
        resolved = [
            (self._deck.anchor_by_id(f"ph:{slide_index}:{kind}"), geo)
            for kind, geo in placeholders.items()
        ]
        for ph, geo in resolved:
            left, top = geo.get("left"), geo.get("top")
            width, height = geo.get("width"), geo.get("height")
            if left is not None or top is not None:
                ph.move(left=left, top=top)  # type: ignore[attr-defined]
            if width is not None or height is not None:
                ph.resize(width=width, height=height)  # type: ignore[attr-defined]

    def list(self) -> list[dict[str, Any]]:
        """`[{index, id, layout, title, shape_count, has_notes}, ...]`."""
        out: list[dict[str, Any]] = []
        for slide in self:
            out.append(
                {
                    "index": slide.index,
                    "id": slide.id,
                    "layout": slide.layout_name,
                    "title": slide.title,
                    "shape_count": len(slide.shapes),
                    "has_notes": slide.has_notes(),
                }
            )
        return out

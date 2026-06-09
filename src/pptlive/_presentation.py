"""Presentation wrapper + PresentationCollection — the wordlive `Document` analog.

A `Presentation` is the deck: it owns the slides and is where `anchor_by_id`,
`outline`, `page_setup`, `edit`, and `go_to` live. Unlike Word there is no
document-wide character stream, so anchor ids are hierarchical and slide-first
(see spec.md §"The anchor model").
"""

from __future__ import annotations

import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import _com, _findreplace, _snapshot
from ._anchors import Anchor
from ._edit import EditScope
from ._selection import SelectionInfo, read_selection
from ._shapes import Shape, has_table, has_text_frame
from ._show import SlideShow
from ._slides import Slide, SlideCollection, _paragraphs
from ._snapshot import Snapshot
from ._theme import Master, Theme
from .constants import (
    DEFAULT_LAYOUT_ALIAS,
    PpSaveAsFileType,
    image_filter_for,
    match_layout_name,
    save_format_for,
)
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    LayoutNotFoundError,
    NoTextFrameError,
    PowerPointBusyError,
    PresentationNotFoundError,
    UnsavedPresentationError,
)

#: Chars of context shown on each side of a `find` match in its `context` snippet.
_CONTEXT_PAD = 30


def _match_context(text: str, start: int, end: int) -> str:
    """A short, single-line snippet of `text` around the half-open span [start, end).

    Paragraph/line separators are shown as visible glyphs rather than flattened to
    spaces, so the surrounding structure is legible in the preview (`⏎` for a `\\r`
    paragraph break, `↵` for a `\\n`/`\\v` soft line break). The `start` offsets are
    computed elsewhere and count each separator as one char, so they're unaffected.
    """
    lo = max(0, start - _CONTEXT_PAD)
    hi = min(len(text), end + _CONTEXT_PAD)
    snippet = text[lo:hi].replace("\r", "⏎").replace("\v", "↵").replace("\n", "↵")
    return ("…" if lo > 0 else "") + snippet + ("…" if hi < len(text) else "")


if TYPE_CHECKING:
    from ._app import PowerPoint


class Presentation:
    """Wraps a PowerPoint `Presentation` COM object."""

    def __init__(self, ppt: PowerPoint, pres_com: Any) -> None:
        self._ppt = ppt
        self._pres = pres_com

    @property
    def com(self) -> Any:
        return self._pres

    @property
    def name(self) -> str:
        with _com.translate_com_errors():
            return str(self._pres.Name)

    @property
    def path(self) -> str:
        """Full path, or just the name for a never-saved deck."""
        with _com.translate_com_errors():
            return str(self._pres.FullName)

    @property
    def saved(self) -> bool:
        """Whether the deck has no unsaved changes (PowerPoint's `Presentation.Saved`).

        `True` right after a save; `False` once an edit dirties it, and `False`
        for a brand-new never-saved deck. The same flag `pptlive status` reports
        per open deck — how an agent sees there's unsaved work before deciding to
        `save()`. (PowerPoint's COM value is an `MsoTriState`, `-1`/`0`; both
        coerce to the right bool.)
        """
        with _com.translate_com_errors():
            return bool(self._pres.Saved)

    def save(self) -> str:
        """Save the deck to its existing file; return the absolute path written.

        The **explicit, never-implicit** persist verb (pptlive never auto-saves).
        Raises [`UnsavedPresentationError`][pptlive.UnsavedPresentationError] if
        the deck has never been saved — it has no path yet, so call
        [`save_as`][pptlive.Presentation.save_as] with a destination first.

        The never-saved guard is in Python on purpose: the 2026-06-09 spike found
        PowerPoint's `Save()` does *not* raise on a path-less deck — on a
        OneDrive/SharePoint build it silently uploads to the user's default cloud
        folder — so relying on COM to refuse would let the deck escape somewhere
        the caller didn't choose.
        """
        with _com.translate_com_errors():
            folder = str(self._pres.Path)
            if not folder:
                raise UnsavedPresentationError(self.name)
            self._pres.Save()
            return str(self._pres.FullName)

    def save_as(
        self, path: str | os.PathLike[str], *, fmt: str = "pptx", overwrite: bool = False
    ) -> str:
        """Save the deck to `path`, returning the absolute path written.

        `fmt` is `"pptx"` (the modern Open XML format). For PDF use
        [`export_pdf`][pptlive.Presentation.export_pdf] — same COM call, but a read
        (it doesn't rebind the working file). **Rebinds** the working file: after
        this, the open deck *is* the new file (its `name`/`path` follow), matching
        PowerPoint's own Save-As. By default refuses to clobber an existing file —
        pass `overwrite=True` to allow it. Explicit-only, like
        [`save`][pptlive.Presentation.save].
        """
        file_format, _ext = save_format_for(fmt)  # validates; rejects "pdf"
        target = Path(os.fspath(path)).expanduser()
        if not overwrite and target.exists():
            raise FileExistsError(
                f"refusing to overwrite existing file {str(target)!r}; pass overwrite=True"
            )
        abspath = str(target.resolve())
        with _com.translate_com_errors():
            self._pres.SaveAs(abspath, file_format)
        return abspath

    def export_pdf(self, path: str | os.PathLike[str]) -> str:
        """Export the deck to a PDF at `path`; return the absolute path written.

        The recommended "hand back a deliverable" path — a pixel-faithful render
        of the deck's current (unsaved) state via PowerPoint's PDF engine. A
        **read**: unlike [`save_as`][pptlive.Presentation.save_as] it neither
        rebinds the working file nor clears its dirty flag (verified 2026-06-09),
        so the user's `.pptx` is untouched and no `edit()` fence is needed.
        Overwrites an existing PDF.

        Goes through `Presentation.SaveAs(path, ppSaveAsPDF=32)`:
        `ExportAsFixedFormat` is the nominal PDF API but won't marshal under
        pptlive's late-bound COM dispatch, and `SaveAs`-to-PDF produces the same
        faithful PDF while behaving as a pure export.
        """
        abspath = str(Path(os.fspath(path)).expanduser().resolve())
        with _com.translate_com_errors():
            self._pres.SaveAs(abspath, int(PpSaveAsFileType.PDF))
        return abspath

    @property
    def slides(self) -> SlideCollection:
        return SlideCollection(self)

    @property
    def show(self) -> SlideShow:
        """Live slide-show control (`start`/`next`/`goto`/`black`/`state`/…).

        Unlike the editing verbs, these deliberately drive what the user sees, so
        they are *not* wrapped in `edit()`. See `_show.SlideShow`.
        """
        return SlideShow(self)

    @property
    def theme(self) -> Theme:
        """Deck-wide theme styling — palette + typefaces (`set_color`/`set_font`).

        A **global, anti-polite** surface: one call recolors/re-fonts every slide
        that inherits the theme. Reaches `SlideMaster.Theme`. See `_theme.Theme`.
        """
        return Theme(self)

    @property
    def master(self) -> Master:
        """Deck-wide master styling — text styles + background.

        The counterpart to per-anchor `format_text`, applied to the primary
        `SlideMaster` so every inheriting slide re-renders. See `_theme.Master`.
        """
        return Master(self)

    def page_setup(self) -> dict[str, float]:
        """Slide canvas dimensions in points: `{width, height}`.

        From `Presentation.PageSetup.SlideWidth`/`SlideHeight`, so an agent can
        place shapes relative to the canvas. Points, never EMUs.
        """
        with _com.translate_com_errors():
            ps = self._pres.PageSetup
            return {"width": float(ps.SlideWidth), "height": float(ps.SlideHeight)}

    def export_images(
        self,
        directory: str | os.PathLike[str],
        *,
        fmt: str = "png",
        width: int | None = None,
        height: int | None = None,
    ) -> list[Path]:
        """Render every slide into `directory`; return the image paths, in order.

        Files are named `slide-001.<ext>`, `slide-002.<ext>`, …. A per-slide wrap
        of `Slide.export_image` (same `fmt`/`width`/`height` semantics) — the
        whole-deck "show me what I built" read.
        """
        _filter_name, ext = image_filter_for(fmt)  # validate before any work
        out_dir = os.path.abspath(os.fspath(directory))
        os.makedirs(out_dir, exist_ok=True)
        paths: list[Path] = []
        for slide in self.slides:
            target = os.path.join(out_dir, f"slide-{slide.index:03d}.{ext}")
            paths.append(slide.export_image(target, width=width, height=height, fmt=fmt))
        return paths

    def snapshot(
        self,
        out: str | os.PathLike[str] | None = None,
        *,
        slides: int | tuple[int, int] | None = None,
        fmt: str = "png",
        max_dim: int | None = None,
    ) -> list[Snapshot]:
        """Render slides to PNG so a vision model can *see* the whole deck cheaply.

        The token-cost-aware read: `max_dim` caps each slide's **long edge** in
        pixels (only ever lowering resolution), so the per-slide cost is a
        predictable budget — and because every slide shares one geometry, that
        budget is uniform across the deck. The lever for "render the whole deck
        and check my styling landed" without full-resolution token bloat
        (~1000 stays legible). `max_dim=None` renders at native size.

        `slides` selects what to render: `None` (default) every slide, an `int`
        a single 1-based slide, a `(start, end)` tuple an inclusive span. Returns
        one [`Snapshot`][pptlive.Snapshot] per slide (so a single slide is a
        one-element list); read `.png` for the bytes.

        If `out` is given the image is also written there: a single slide to `out`
        itself, multiple slides alongside it as `<stem>-s<N><suffix>`. `fmt` is a
        friendly image token (`png`/`jpg`/…). A read — the export renders the
        current unsaved state and leaves the viewed slide and Selection untouched
        (no `edit()` fence needed).
        """
        return _snapshot.snapshot(self, out, slides=slides, fmt=fmt, max_dim=max_dim)

    def selection(self) -> SelectionInfo:
        """The user's current selection, resolved to anchors — a polite read.

        Snapshots `ActiveWindow.Selection` without changing it (the complement to
        `status`, which reports the viewed slide): the selected shapes as
        `shape:S:N`, or a text caret as `para:S:N:P`. To *act on* the selection,
        resolve `anchor_by_id("here:")` — the explicit opt-in (pptlive never
        targets the live Selection unless asked).
        """
        return read_selection(self._ppt)

    def layouts(self) -> list[dict[str, Any]]:
        """The deck's slide layouts: `[{index, name}, ...]` (1-based index).

        Sourced from `SlideMaster.CustomLayouts`. Lists the exact names that
        `slides.add(layout=…)` / `Slide.set_layout(…)` accept on this template —
        useful when a theme has renamed the standard Office layouts.
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for idx, layout in enumerate(self._custom_layouts(), start=1):
                out.append({"index": idx, "name": str(layout.Name)})
        return out

    def _custom_layouts(self) -> list[Any]:
        """The deck's `CustomLayout` COM objects (empty if it exposes none)."""
        try:
            with _com.translate_com_errors():
                return list(self._pres.SlideMaster.CustomLayouts)
        except PowerPointBusyError:
            # Surface a transient busy (exit 3) rather than reporting "no layouts"
            # and silently falling back to the legacy Slides.Add path.
            raise
        except Exception:
            return []

    def _resolve_layout(self, requested: str | int | None) -> Any | None:
        """Map a friendly layout name/index to a `CustomLayout` COM object.

        Returns the matching `CustomLayout`, or None *only* when the deck exposes
        no custom layouts at all (so `slides.add` can fall back to the legacy
        `Slides.Add`). `requested=None` picks the default (`title_and_content`,
        else the first layout). A name matches case/separator-insensitively
        against the deck's real layout names, then a small friendly-alias table;
        an int is a 1-based index into the layouts. Raises `LayoutNotFoundError`
        — listing the available names — for an unknown name or bad index.
        """
        with _com.translate_com_errors():
            layouts = self._custom_layouts()
            if not layouts:
                return None
            names = [str(layout.Name) for layout in layouts]
            if requested is None:
                idx = match_layout_name(names, DEFAULT_LAYOUT_ALIAS)
                return layouts[idx - 1] if idx is not None else layouts[0]
            if isinstance(requested, bool):
                raise LayoutNotFoundError(str(requested), names)
            if isinstance(requested, int):
                if requested < 1 or requested > len(layouts):
                    raise LayoutNotFoundError(str(requested), names)
                return layouts[requested - 1]
            idx = match_layout_name(names, requested)
            if idx is None:
                raise LayoutNotFoundError(str(requested), names)
            return layouts[idx - 1]

    def outline(self) -> list[dict[str, Any]]:
        """The Outline-view analog: `[{slide, title, bullets:[...]}, ...]`.

        `bullets` are the non-empty paragraphs of the slide's body placeholder
        (when it has one); slides without a body — or whose body placeholder
        holds a chart/table/picture rather than text — just carry their title.
        """
        out: list[dict[str, Any]] = []
        for slide in self.slides:
            bullets: list[str] = []
            try:
                body = slide.placeholder("body")
                bullets = _paragraphs(body.text)
            except (AnchorNotFoundError, NoTextFrameError):
                # No body placeholder, or it's been filled with a chart/table/
                # picture (no text frame): the slide simply contributes no bullets.
                bullets = []
            out.append({"slide": slide.index, "title": slide.title, "bullets": bullets})
        return out

    def comments(self) -> dict[str, Any]:
        """Every review comment across the deck — the deck-wide roll-up.

        `{total, slides: [{slide, comments:[{index, author, initials, text,
        datetime, left, top, replies:[...]}, ...]}, ...]}`. Only slides that carry
        at least one comment appear; `total` counts top-level comments (not
        replies). A read — side-effect-free and polite (no view move). For one
        slide, use `deck.slides[S].comments.list()` (the `comments:S` read).
        """
        slides_out: list[dict[str, Any]] = []
        total = 0
        for slide in self.slides:
            items = slide.comments.list()
            if items:
                slides_out.append({"slide": slide.index, "comments": items})
                total += len(items)
        return {"total": total, "slides": slides_out}

    def anchor_by_id(self, anchor_id: str) -> Anchor:
        """Resolve an `anchor_id` string into an `Anchor`.

        Recognised:
          - `shape:S:N`   — Nth shape (1-based z-order) on slide S
          - `shapeid:S:ID`— shape with stable `Shape.Id` ID on slide S — the
                            delete-proof handle (the `id` in every shape listing)
          - `ph:S:KIND`   — placeholder of semantic KIND on slide S
                            (title/ctrtitle/subtitle/body/footer/date/slidenum)
          - `para:S:N:P`  — paragraph P of shape N on slide S (v0.3)
          - `cell:S:N:R:C`— cell (row R, col C) of the table in shape N on slide S (v0.5)
          - `notes:S`     — speaker-notes body of slide S
          - `here:`       — whatever the user has selected right now (v0.4): the
                            selected shape, or the paragraph holding the text caret

        `slide:S` is a *container*, not a text anchor — use `deck.slides[S]`.

        Raises `AnchorNotFoundError` for unknown schemes or missing anchors
        (`SlideNotFoundError`, a subclass, for an out-of-range slide).
        """
        if not isinstance(anchor_id, str) or ":" not in anchor_id:
            raise AnchorNotFoundError("anchor", str(anchor_id))
        kind, _, rest = anchor_id.partition(":")

        if kind == "shape":
            parts = rest.split(":")
            if len(parts) != 2:
                raise AnchorNotFoundError("shape", anchor_id)
            try:
                s, n = int(parts[0]), int(parts[1])
            except ValueError as e:
                raise AnchorNotFoundError("shape", anchor_id) from e
            return self.slides[s].shapes[n]

        if kind == "shapeid":
            parts = rest.split(":")
            if len(parts) != 2:
                raise AnchorNotFoundError("shape", anchor_id)
            try:
                s, sid = int(parts[0]), int(parts[1])
            except ValueError as e:
                raise AnchorNotFoundError("shape", anchor_id) from e
            return self.slides[s].shapes.by_id(sid)

        if kind == "para":
            parts = rest.split(":")
            if len(parts) != 3:
                raise AnchorNotFoundError("paragraph", anchor_id)
            try:
                s, n, p = int(parts[0]), int(parts[1]), int(parts[2])
            except ValueError as e:
                raise AnchorNotFoundError("paragraph", anchor_id) from e
            return self.slides[s].shapes[n].paragraph(p)

        if kind == "cell":
            parts = rest.split(":")
            if len(parts) != 4:
                raise AnchorNotFoundError("table cell", anchor_id)
            try:
                s, n, r, c = (int(x) for x in parts)
            except ValueError as e:
                raise AnchorNotFoundError("table cell", anchor_id) from e
            return self.slides[s].shapes[n].table.cell(r, c)

        if kind == "ph":
            s_str, sep, ph_kind = rest.partition(":")
            if not sep or not ph_kind:
                raise AnchorNotFoundError("placeholder", anchor_id)
            try:
                s = int(s_str)
            except ValueError as e:
                raise AnchorNotFoundError("placeholder", anchor_id) from e
            try:
                return self.slides[s].placeholder(ph_kind)
            except ValueError as e:
                # Unknown KIND — surface as a missing anchor (exit 2), not a crash.
                raise AnchorNotFoundError("placeholder", anchor_id) from e

        if kind == "notes":
            try:
                s = int(rest)
            except ValueError as e:
                raise AnchorNotFoundError("notes", anchor_id) from e
            return self.slides[s].notes

        if kind == "here":
            # The explicit opt-in to target the live Selection (politeness:
            # never otherwise). Resolves live to a Shape or Paragraph.
            info = read_selection(self._ppt)
            if (
                info.type == "text"
                and info.slide is not None
                and info.shape_index is not None
                and info.paragraph is not None
            ):
                return self.slides[info.slide].shapes[info.shape_index].paragraph(info.paragraph)
            if info.type == "shapes" and info.slide is not None and info.shape_index is not None:
                return self.slides[info.slide].shapes[info.shape_index]
            raise AnchorNotFoundError("selection", anchor_id)

        raise AnchorNotFoundError("anchor", anchor_id)

    # -- find / replace (v1.0) -------------------------------------------------
    #
    # PowerPoint has no deck-wide character stream, so search is a *traversal*:
    # slides × shapes → each text frame, table cells, and speaker notes. A search
    # *unit* is one text frame addressed by a resolvable anchor (`shape:S:N`,
    # `cell:S:N:R:C`, `notes:S`); offsets are 0-based char positions within that
    # frame's text. Matching reuses wordlive's fuzzy core (`_findreplace`);
    # replacement writes back through `TextRange.Characters` so only the matched
    # span changes and the rest of the frame keeps its run formatting.

    def _units_for_shape(self, shape: Shape) -> list[tuple[str, Any]]:
        """The (anchor_id, COM `TextRange`) search units a shape contributes.

        A text-framed shape contributes its frame (`shape:S:N`); a table shape
        contributes one unit per cell (`cell:S:N:R:C`). A shape may be both
        (rare) or neither (a picture / line — no text, no units).
        """
        units: list[tuple[str, Any]] = []
        sh = shape.com  # raw COM Shape, resolved live
        if has_text_frame(sh):
            units.append((shape.anchor_id, sh.TextFrame.TextRange))
        if has_table(sh):
            table = sh.Table
            nrows, ncols = int(table.Rows.Count), int(table.Columns.Count)
            s, n = shape.slide.index, shape.index
            for r in range(1, nrows + 1):
                for c in range(1, ncols + 1):
                    cell_tr = table.Cell(r, c).Shape.TextFrame.TextRange
                    units.append((f"cell:{s}:{n}:{r}:{c}", cell_tr))
        return units

    def _units_for_anchor(self, anchor: Anchor) -> list[tuple[str, Any]]:
        """Search units for a single anchor scope (a shape expands to its frames)."""
        if isinstance(anchor, Shape):
            return self._units_for_shape(anchor)
        return [(anchor.anchor_id, anchor._text_range())]

    def _units_for_slides(self, slides: list[Slide]) -> list[tuple[str, Any]]:
        units: list[tuple[str, Any]] = []
        for slide in slides:
            for shape in slide.shapes:
                units.extend(self._units_for_shape(shape))
            # Speaker notes — a text frame the traversal must visit. Slides with
            # no notes-body placeholder simply contribute no notes unit.
            try:
                notes_tr = slide.notes._text_range()
            except AnchorNotFoundError:
                notes_tr = None
            if notes_tr is not None:
                units.append((slide.notes.anchor_id, notes_tr))
        return units

    def _search_units(self, scope: str | Slide | Anchor | None) -> list[tuple[str, Any]]:
        """Resolve a find/replace `scope` to its list of (anchor_id, `TextRange`)."""
        with _com.translate_com_errors():
            if scope is None:
                return self._units_for_slides(list(self.slides))
            if isinstance(scope, Slide):
                return self._units_for_slides([scope])
            if isinstance(scope, Anchor):
                return self._units_for_anchor(scope)
            if isinstance(scope, str):
                if scope.split(":", 1)[0] == "slide":
                    parts = scope.split(":")
                    try:
                        s = int(parts[1])
                    except (IndexError, ValueError) as e:
                        raise AnchorNotFoundError("slide", scope) from e
                    return self._units_for_slides([self.slides[s]])
                return self._units_for_anchor(self.anchor_by_id(scope))
        raise TypeError(
            f"scope must be an anchor id, Slide, Anchor, or None, got {type(scope).__name__}"
        )

    def find(self, text: str, *, scope: str | Slide | Anchor | None = None) -> list[dict[str, Any]]:
        """Locate every fuzzy occurrence of `text` across the deck (or `scope`).

        Search is a traversal of slides × shapes → text frames, table cells, and
        speaker notes (there is no deck-wide character stream). Matching is
        whitespace- and Unicode-normalized (NFKC, smart quotes, dashes, NBSP), so
        text an LLM re-typed off a slide still matches the original glyphs.

        Returns `{anchor_id, start, length, text, context}` per hit, in document
        order: `anchor_id` is a resolvable text anchor (`shape:S:N`,
        `cell:S:N:R:C`, or `notes:S`), `start` is the 0-based char offset of the
        match within that anchor's text, `text` is the actual original substring,
        and `context` is a short surrounding snippet. The offsets are live — use
        them before further edits shift the text.

        A read — polite (no view move). `scope` restricts the search: a `slide:S`
        string (or a `Slide`) limits it to one slide; any text-anchor id (or an
        `Anchor`) limits it to that shape / cell / notes frame; `None` (default)
        searches the whole deck.
        """
        units = self._search_units(scope)
        results: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for anchor_id, tr in units:
                haystack = str(tr.Text or "")
                for m in _findreplace.find_matches(haystack, text):
                    results.append(
                        {
                            "anchor_id": anchor_id,
                            "start": m.start,
                            "length": m.end - m.start,
                            "text": m.text,
                            "context": _match_context(haystack, m.start, m.end),
                        }
                    )
        return results

    def find_replace(
        self,
        find: str,
        replace: str,
        *,
        scope: str | Slide | Anchor | None = None,
        all: bool = False,
        occurrence: int | None = None,
    ) -> list[dict[str, Any]]:
        """Fuzzy plain-text replace across the deck (or `scope`). See `find`.

        Args:
            find: the text to look for (fuzzy-matched, same semantics as `find`).
            replace: the replacement text.
            scope: restrict the search — a `slide:S` / anchor-id string, a `Slide`
                or `Anchor`, or `None` for the whole deck.
            all: replace every match.
            occurrence: 1-based index — replace only the Nth match (document order).

        Raises:
            AnchorNotFoundError: zero matches (`kind='find'`, exit 2), or an
                out-of-range `occurrence`.
            AmbiguousMatchError: more than one match and neither `all` nor
                `occurrence` was given (exit 5).

        Returns the replacements applied, each `{anchor_id, start, length, text}`
        in their pre-replacement coordinates. Only the matched span is rewritten
        (via `TextRange.Characters`), so the rest of each frame keeps its run
        formatting. Wrap the call in `deck.edit(...)` for view preservation and a
        one-Ctrl-Z fence.
        """
        units = self._search_units(scope)
        # (anchor_id, COM TextRange, start, end, original_text) in document order.
        matches: list[tuple[str, Any, int, int, str]] = []
        with _com.translate_com_errors():
            for anchor_id, tr in units:
                haystack = str(tr.Text or "")
                for m in _findreplace.find_matches(haystack, find):
                    matches.append((anchor_id, tr, m.start, m.end, m.text))

        if not matches:
            raise AnchorNotFoundError("find", find)

        if occurrence is not None:
            if occurrence < 1 or occurrence > len(matches):
                raise AnchorNotFoundError("find", f"{find} (occurrence {occurrence})")
            to_apply = [matches[occurrence - 1]]
        elif all:
            to_apply = list(matches)
        elif len(matches) == 1:
            to_apply = matches
        else:
            raise AmbiguousMatchError(
                find,
                [
                    {"anchor_id": a, "start": s, "length": e - s, "text": t}
                    for (a, _tr, s, e, t) in matches
                ],
            )

        # Apply in reverse document order so an earlier match's offsets stay valid
        # after a later match in the *same* frame is rewritten to a different
        # length (matches in different frames are independent).
        applied: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            for anchor_id, tr, start, end, text in reversed(to_apply):
                tr.Characters(start + 1, end - start).Text = replace
                applied.append(
                    {"anchor_id": anchor_id, "start": start, "length": end - start, "text": text}
                )
        applied.reverse()  # report in document order
        return applied

    @contextmanager
    def edit(self, label: str) -> Iterator[EditScope]:
        """Open an atomic-undo + view/Selection-preserving edit scope.

        Mutations inside the block collapse into a **single Ctrl-Z**: the scope
        fences a fresh undo entry with `Application.StartNewUndoEntry()` on entry
        and PowerPoint groups the rest. On clean exit the user is returned to the
        slide and selection they had. See `EditScope` for the mechanism and its
        caveats (no explicit "end" fence; always wrap mutations in `edit`).

        ```
        with deck.edit("Revise agenda slide"):
            deck.anchor_by_id("ph:2:title").set_text("Agenda")
            deck.anchor_by_id("ph:2:body").set_text("Intro\\nDemo\\nQ&A")
        ```
        """
        scope = EditScope(self._ppt, label)
        with scope:
            yield scope

    def go_to(self, target: Anchor | Slide, *, select: bool = True) -> None:
        """Move the user's view to a slide or shape (deliberate, opt-in jump).

        Rare — most operations preserve the view. `target` is an `Anchor`
        (resolved via `anchor_by_id`) or a `Slide`. Jumps the active window to
        that slide and, when `select` is True and the target is a shape, selects
        it. This intentionally moves the user, so inside a `deck.edit(...)` block
        call `scope.allow_view_move()` first or it'll be snapped back on exit.
        """
        slide: Slide | None
        shape: Shape | None
        if isinstance(target, Slide):
            slide, shape = target, None
        else:
            slide = getattr(target, "slide", None)
            shape = target if isinstance(target, Shape) else None
            if slide is None:
                raise TypeError(f"cannot go_to {type(target).__name__}: no slide context")

        assert slide is not None  # narrowed by the branches above
        slide_index = slide.index
        with _com.translate_com_errors():
            win = self._ppt.com.ActiveWindow
            win.View.GotoSlide(int(slide_index))
            if select and shape is not None:
                try:
                    with _com.translate_com_errors():
                        shape.com.Select()
                except PowerPointBusyError:
                    # The goto landed; only the shape-select failed. A busy here
                    # is still retryable — surface it rather than silently
                    # leaving nothing selected.
                    raise
                except Exception:
                    pass

    def __repr__(self) -> str:
        return f"<Presentation {self.name!r}>"


class PresentationCollection:
    """Indexable view over open presentations."""

    def __init__(self, ppt: PowerPoint) -> None:
        self._ppt = ppt

    @property
    def _com_collection(self) -> Any:
        return self._ppt.com.Presentations

    @property
    def active(self) -> Presentation:
        with _com.translate_com_errors():
            try:
                pres = self._ppt.com.ActivePresentation
            except Exception as e:
                raise PresentationNotFoundError("<active>") from e
        return Presentation(self._ppt, pres)

    def __getitem__(self, name: str) -> Presentation:
        with _com.translate_com_errors():
            for pres in self._com_collection:
                if str(pres.Name) == name:
                    return Presentation(self._ppt, pres)
        raise PresentationNotFoundError(name)

    def __iter__(self) -> Iterator[Presentation]:
        with _com.translate_com_errors():
            decks = list(self._com_collection)
        for pres in decks:
            yield Presentation(self._ppt, pres)

    def __len__(self) -> int:
        with _com.translate_com_errors():
            return int(self._com_collection.Count)

    def list(self) -> list[dict[str, Any]]:
        """`[{name, path, saved, is_active}, ...]` — used by `pptlive status`.

        `saved` is `Presentation.Saved` (False = unsaved changes / never saved),
        so an agent sees dirty state before deciding to `save()`; `path` is the
        full path (just the name for a never-saved deck).
        """
        out: list[dict[str, Any]] = []
        with _com.translate_com_errors():
            active_name: str | None
            try:
                active_name = str(self._ppt.com.ActivePresentation.Name)
            except Exception:
                active_name = None
            for pres in self._com_collection:
                name = str(pres.Name)
                out.append(
                    {
                        "name": name,
                        "path": str(pres.FullName),
                        "saved": bool(pres.Saved),
                        "is_active": name == active_name,
                    }
                )
        return out

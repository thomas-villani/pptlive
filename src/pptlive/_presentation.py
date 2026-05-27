"""Presentation wrapper + PresentationCollection — the wordlive `Document` analog.

A `Presentation` is the deck: it owns the slides and is where `anchor_by_id`,
`outline`, `page_setup`, `edit`, and `go_to` live. Unlike Word there is no
document-wide character stream, so anchor ids are hierarchical and slide-first
(see spec.md §"The anchor model").
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import TYPE_CHECKING, Any

from . import _com
from ._anchors import Anchor
from ._edit import EditScope
from ._shapes import Shape
from ._slides import Slide, SlideCollection, _paragraphs
from .constants import DEFAULT_LAYOUT_ALIAS, match_layout_name
from .exceptions import AnchorNotFoundError, LayoutNotFoundError, PresentationNotFoundError

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
    def slides(self) -> SlideCollection:
        return SlideCollection(self)

    def page_setup(self) -> dict[str, float]:
        """Slide canvas dimensions in points: `{width, height}`.

        From `Presentation.PageSetup.SlideWidth`/`SlideHeight`, so an agent can
        place shapes relative to the canvas. Points, never EMUs.
        """
        with _com.translate_com_errors():
            ps = self._pres.PageSetup
            return {"width": float(ps.SlideWidth), "height": float(ps.SlideHeight)}

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
            return list(self._pres.SlideMaster.CustomLayouts)
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
        (when it has one); slides without a body just carry their title.
        """
        out: list[dict[str, Any]] = []
        for slide in self.slides:
            bullets: list[str] = []
            try:
                body = slide.placeholder("body")
                bullets = _paragraphs(body.text)
            except AnchorNotFoundError:
                bullets = []
            out.append({"slide": slide.index, "title": slide.title, "bullets": bullets})
        return out

    def anchor_by_id(self, anchor_id: str) -> Anchor:
        """Resolve an `anchor_id` string into an `Anchor`.

        Recognised in v0:
          - `shape:S:N`   — Nth shape (1-based z-order) on slide S
          - `ph:S:KIND`   — placeholder of semantic KIND on slide S
                            (title/ctrtitle/subtitle/body/footer/date/slidenum)
          - `notes:S`     — speaker-notes body of slide S

        `slide:S` is a *container*, not a text anchor — use `deck.slides[S]`.
        `para:`/`cell:` arrive in later stages and are not yet resolvable.

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

        raise AnchorNotFoundError("anchor", anchor_id)

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
                    shape.com.Select()
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
        """`[{name, path, is_active}, ...]` — used by `pptlive status`."""
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
                        "is_active": name == active_name,
                    }
                )
        return out

"""SlideShow — live presentation control (`deck.show`), the v0.6 track.

The most literally "live" surface pptlive offers: drive a running slide show the
way a presenter's clicker does. There is no Word analog — this is pure
PowerPoint.

Object model. A deck's show is started from `Presentation.SlideShowSettings`
(`.Run()` returns a `SlideShowWindow`); once running, `Presentation.SlideShowWindow`
hands back that window and `SlideShowWindow.View` is the controller —
`Next`/`Previous`/`GotoSlide`/`Exit`, plus a read/write `State` whose values are
`PpSlideShowState` (the B/W "blank screen" states live here). Accessing
`SlideShowWindow` when nothing is running raises, so `_window()` treats any
failure as "no show", and the control verbs raise `SlideShowNotRunningError`
(exit 1) rather than a bare COM error.

Politeness note. A slide show *is* the deliberate "take over the user's screen"
operation — like `go_to`, it intentionally moves what the user sees, so these
verbs are **not** wrapped in `deck.edit(...)` (show control isn't a document
mutation and has no undo). Editing *while* a show runs is allowed — the
2026-05-28 spike found a `set_text` succeeds mid-show — so this surface doesn't
block edits; it just drives the presentation alongside them.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from . import _com
from .constants import PpSlideShowRangeType, PpSlideShowState, slide_show_state_name
from .exceptions import SlideNotFoundError, SlideShowNotRunningError

if TYPE_CHECKING:
    from ._presentation import Presentation


class SlideShow:
    """Live slide-show control for one deck (`deck.show`).

    ```
    deck.show.start()            # run from the top
    deck.show.goto(5)            # jump to slide 5
    deck.show.next()             # advance
    deck.show.black()            # blank the screen (B); resume() to come back
    deck.show.state()            # {"running": True, "state": "running", ...}
    deck.show.end()              # exit the show
    ```

    Every control verb returns the post-action `state()` dict, so a caller never
    has to follow up with a separate read. `state()` itself is the only
    side-effect-free verb and the only one that never raises when no show is
    running (it reports `running: false`).
    """

    def __init__(self, deck: Presentation) -> None:
        self._deck = deck

    @property
    def com(self) -> Any:
        """Raw `SlideShowSettings` COM object — the always-available handle.

        For the *running* show, use `window` / `view` (both None when no show is
        running). Escape hatch for show knobs pptlive doesn't wrap (loop,
        narration, advance mode, …).
        """
        with _com.translate_com_errors():
            return self._deck.com.SlideShowSettings

    @property
    def window(self) -> Any | None:
        """The live `SlideShowWindow` COM object, or None if no show is running."""
        return self._window()

    @property
    def view(self) -> Any | None:
        """The live `SlideShowView` COM object, or None if no show is running."""
        win = self._window()
        if win is None:
            return None
        with _com.translate_com_errors():
            return win.View

    def _window(self) -> Any | None:
        """The deck's running `SlideShowWindow`, or None.

        `Presentation.SlideShowWindow` raises (or yields nothing) when no show is
        running, so any failure here means "not running" — we never let it
        bubble as a COM error.
        """
        try:
            win = self._deck.com.SlideShowWindow
        except Exception:
            return None
        return win or None

    def _require_window(self) -> Any:
        win = self._window()
        if win is None:
            raise SlideShowNotRunningError()
        return win

    def is_running(self) -> bool:
        """True iff a slide show is currently running for this deck."""
        return self._window() is not None

    def start(self, *, from_slide: int | None = None) -> dict[str, Any]:
        """Start the slide show (or, if one is already running, keep it).

        With `from_slide` (1-based) the show begins on that slide; otherwise it
        runs the whole deck from the top. If a show is already running this is a
        no-op except that `from_slide`, when given, jumps to that slide — so
        `start()` is safe to call idempotently. Returns the show `state()`.
        Raises `SlideNotFoundError` for an out-of-range `from_slide`.
        """
        if from_slide is not None:
            self._check_slide(from_slide)
        with _com.translate_com_errors():
            existing = self._window()
            if existing is not None:
                if from_slide is not None:
                    existing.View.GotoSlide(int(from_slide))
                return self.state()
            settings = self._deck.com.SlideShowSettings
            if from_slide is not None:
                settings.RangeType = int(PpSlideShowRangeType.SLIDE_RANGE)
                settings.StartingSlide = int(from_slide)
                settings.EndingSlide = len(self._deck.slides)
            settings.Run()
        return self.state()

    def end(self) -> dict[str, Any]:
        """End the slide show. A no-op (not an error) if none is running."""
        win = self._window()
        if win is not None:
            with _com.translate_com_errors():
                win.View.Exit()
        return self.state()

    def next(self) -> dict[str, Any]:
        """Advance to the next build/slide (the clicker's forward press)."""
        with _com.translate_com_errors():
            self._require_window().View.Next()
        return self.state()

    def previous(self) -> dict[str, Any]:
        """Step back to the previous build/slide."""
        with _com.translate_com_errors():
            self._require_window().View.Previous()
        return self.state()

    def goto(self, slide: int) -> dict[str, Any]:
        """Jump the running show to slide `slide` (1-based).

        Raises `SlideNotFoundError` for an out-of-range index, and
        `SlideShowNotRunningError` if no show is running.
        """
        self._check_slide(slide)
        with _com.translate_com_errors():
            self._require_window().View.GotoSlide(int(slide))
        return self.state()

    def black(self) -> dict[str, Any]:
        """Blank the screen to black (the B key). `resume()` returns to the slide."""
        return self._set_state(PpSlideShowState.BLACK_SCREEN)

    def white(self) -> dict[str, Any]:
        """Blank the screen to white (the W key). `resume()` returns to the slide."""
        return self._set_state(PpSlideShowState.WHITE_SCREEN)

    def resume(self) -> dict[str, Any]:
        """Resume from a black/white blank screen back to the running slide."""
        return self._set_state(PpSlideShowState.RUNNING)

    def state(self) -> dict[str, Any]:
        """Report the show's status without changing it — the polite read.

        Always returns a dict: `{running, state, state_code, current_slide,
        position, slide_count}`. When no show is running, `running` is False,
        `state` is `"done"`, and the slide fields are None. `current_slide` is
        the 1-based deck index of the slide on screen; `position` is its place in
        the show sequence (differs from `current_slide` with hidden/custom shows).
        """
        slide_count = len(self._deck.slides)
        win = self._window()
        if win is None:
            return {
                "running": False,
                "state": slide_show_state_name(PpSlideShowState.DONE),
                "state_code": int(PpSlideShowState.DONE),
                "current_slide": None,
                "position": None,
                "slide_count": slide_count,
            }
        with _com.translate_com_errors():
            view = win.View
            state_code = int(view.State)
            current_slide = self._safe_int(lambda: view.Slide.SlideIndex)
            position = self._safe_int(lambda: view.CurrentShowPosition)
        return {
            "running": True,
            "state": slide_show_state_name(state_code),
            "state_code": state_code,
            "current_slide": current_slide,
            "position": position,
            "slide_count": slide_count,
        }

    def _set_state(self, state: PpSlideShowState) -> dict[str, Any]:
        with _com.translate_com_errors():
            self._require_window().View.State = int(state)
        return self.state()

    def _check_slide(self, index: Any) -> None:
        """Validate a 1-based slide index against the deck, else SlideNotFoundError."""
        if isinstance(index, bool) or not isinstance(index, int):
            raise SlideNotFoundError(index)
        if index < 1 or index > len(self._deck.slides):
            raise SlideNotFoundError(index)

    @staticmethod
    def _safe_int(getter: Any) -> int | None:
        try:
            return int(getter())
        except Exception:
            return None

    def __repr__(self) -> str:
        return f"<SlideShow running={self.is_running()!r}>"

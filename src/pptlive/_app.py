"""PowerPoint application wrapper + attach()/connect() context managers.

Note the PowerPoint diff from wordlive: `connect()` has **no `visible=False`
mode**. PowerPoint historically refuses to run invisibly, so the app is always
shown; politeness is about not *moving* the user's view, not about working
hidden. Like wordlive, pptlive never closes PowerPoint on exit — it's the user's
app, even when we launched it.
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import TYPE_CHECKING, Any

from . import _com
from .exceptions import PowerPointNotRunningError

if TYPE_CHECKING:
    from ._presentation import PresentationCollection


class PowerPoint:
    """Handle to a running PowerPoint.Application COM object."""

    def __init__(self, app: Any) -> None:
        self._app = app

    @property
    def com(self) -> Any:
        """Raw Application COM object — escape hatch when pptlive doesn't cover something."""
        return self._app

    @property
    def visible(self) -> bool:
        return bool(self._app.Visible)

    @property
    def presentations(self) -> PresentationCollection:
        from ._presentation import PresentationCollection

        return PresentationCollection(self)

    def viewed_slide_index(self) -> int | None:
        """1-based index of the slide the user is currently looking at, or None.

        None when there's no active window or the active view isn't one where a
        slide is shown (e.g. slide sorter, or a slide show running).
        """
        try:
            return int(self._app.ActiveWindow.View.Slide.SlideIndex)
        except Exception:
            return None

    def __repr__(self) -> str:
        return "<PowerPoint>"


@contextmanager
def attach() -> Iterator[PowerPoint]:
    """Attach to an already-running PowerPoint instance.

    Raises `PowerPointNotRunningError` if no instance is available. Does not
    launch PowerPoint and does not close it on exit.
    """
    with _com.com_apartment():
        app = _com.get_active_powerpoint()
        try:
            yield PowerPoint(app)
        finally:
            del app


@contextmanager
def connect(launch_if_missing: bool = True) -> Iterator[PowerPoint]:
    """Attach to a running PowerPoint, or launch a new one if missing.

    With `launch_if_missing=False` this behaves like `attach()`. There is no
    `visible` parameter — PowerPoint is always visible (see module docstring).
    pptlive never closes PowerPoint on exit, even when it launched the instance:
    the user owns its lifecycle.
    """
    with _com.com_apartment():
        try:
            app = _com.get_active_powerpoint()
        except PowerPointNotRunningError:
            if not launch_if_missing:
                raise
            app = _com.launch_powerpoint()
        try:
            yield PowerPoint(app)
        finally:
            del app

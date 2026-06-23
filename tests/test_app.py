"""Connection entry points: `attach()` and `connect()`.

These wrap the `_com` getters (which tests monkeypatch). `attach()` only ever
talks to an already-running instance; `connect()` adds the launch-if-missing
fallback — the path a fresh user hits when no PowerPoint is open yet.
"""

from __future__ import annotations

from typing import Any

import pytest

import pptlive
from pptlive import _com
from pptlive.exceptions import PowerPointNotRunningError, PresentationNotFoundError


def test_attach_yields_handle_to_running_instance(fake_powerpoint: Any) -> None:
    with pptlive.attach() as handle:
        assert handle.presentations.active is not None


def test_attach_raises_when_not_running(no_powerpoint: None) -> None:
    with pytest.raises(PowerPointNotRunningError):
        with pptlive.attach():
            pass


def test_connect_attaches_when_already_running(
    fake_powerpoint: Any, monkeypatch: pytest.MonkeyPatch
) -> None:
    # If an instance is running, connect() must NOT launch a second one.
    def _should_not_launch() -> Any:
        raise AssertionError("connect() launched despite a running instance")

    monkeypatch.setattr(_com, "launch_powerpoint", _should_not_launch)
    with pptlive.connect() as handle:
        assert handle.presentations.active is not None


def test_connect_launches_when_missing(
    fake_powerpoint: Any, monkeypatch: pytest.MonkeyPatch
) -> None:
    launched = {"called": False}

    def _not_running() -> Any:
        raise PowerPointNotRunningError("no running Microsoft PowerPoint instance found")

    def _launch() -> Any:
        launched["called"] = True
        return fake_powerpoint

    monkeypatch.setattr(_com, "get_active_powerpoint", _not_running)
    monkeypatch.setattr(_com, "launch_powerpoint", _launch)

    with pptlive.connect() as handle:
        assert handle.presentations.active is not None
    assert launched["called"] is True


def test_doc_selector_matches_by_name_then_full_path(
    fake_powerpoint_same_named_decks: Any,
) -> None:
    # Two open decks share the display name "Deck.pptx" but differ by path.
    with pptlive.attach() as handle:
        decks = handle.presentations
        # The bare display name resolves the first match (the common case)...
        assert decks["Deck.pptx"].path == r"C:\\a\\Deck.pptx"
        # ...and the full path disambiguates to the *other* same-named deck.
        assert decks[r"C:\\b\\Deck.pptx"].path == r"C:\\b\\Deck.pptx"


def test_doc_selector_unknown_raises_not_found(fake_powerpoint_same_named_decks: Any) -> None:
    with pptlive.attach() as handle:
        with pytest.raises(PresentationNotFoundError):
            _ = handle.presentations["nope.pptx"]


def test_active_and_list_surface_busy_not_not_found() -> None:
    # A transient busy on the ActivePresentation read is exit 3 (retryable) — it
    # must NOT be collapsed into PresentationNotFoundError (exit 2) / a null
    # active deck. The typed busy passes through translate_com_errors unchanged.
    from types import SimpleNamespace

    from pptlive._presentation import PresentationCollection
    from pptlive.exceptions import PowerPointBusyError

    class _BoomCom:
        Presentations: list = []  # type: ignore[type-arg]

        @property
        def ActivePresentation(self) -> object:
            raise PowerPointBusyError(hresult=0x80010001)

    coll = PresentationCollection(SimpleNamespace(com=_BoomCom()))  # type: ignore[arg-type]
    with pytest.raises(PowerPointBusyError):
        _ = coll.active
    with pytest.raises(PowerPointBusyError):
        coll.list()


def test_connect_no_launch_reraises(no_powerpoint: None) -> None:
    # launch_if_missing=False makes connect() behave like attach().
    with pytest.raises(PowerPointNotRunningError):
        with pptlive.connect(launch_if_missing=False):
            pass


def test_viewed_slide_index_returns_none_when_no_slide_view(
    fake_powerpoint: Any, monkeypatch: pytest.MonkeyPatch
) -> None:
    # No active slide view (slide sorter, or a show running) -> None, not a crash.
    monkeypatch.setattr(fake_powerpoint, "_window", None)
    with pptlive.attach() as handle:
        assert handle.viewed_slide_index() is None


def test_viewed_slide_index_returns_int_when_viewing(fake_powerpoint: Any) -> None:
    with pptlive.attach() as handle:
        assert isinstance(handle.viewed_slide_index(), int)

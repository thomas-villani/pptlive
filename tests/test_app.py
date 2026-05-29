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
from pptlive.exceptions import PowerPointNotRunningError


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

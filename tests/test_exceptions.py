"""COM-error decoding/classification and CLI exit-code mapping."""

from __future__ import annotations

import pptlive.cli.main as cli_main
from pptlive.exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    NoTextFrameError,
    PowerPointBusyError,
    PowerPointNotRunningError,
    PresentationNotFoundError,
    SlideNotFoundError,
    _decode_com_error,
    from_com_error,
)


class _FakeComError(Exception):
    """Stand-in for pywintypes.com_error: just needs a populated `.args`."""


def _com_error(hresult: int, description: str | None = None) -> _FakeComError:
    exc_info = (0, "Microsoft PowerPoint", description, None, 0, hresult) if description else None
    e = _FakeComError()
    e.args = (hresult, "some message", exc_info, None)
    return e


def test_decode_com_error_extracts_hresult_and_description() -> None:
    hresult, description, message = _decode_com_error(_com_error(-2147418111, "Call rejected"))
    assert hresult == -2147418111
    assert description == "Call rejected"
    assert "Call rejected" in message
    assert "0x80010001" in message  # signed -2147418111 rendered as unsigned hex


def test_busy_hresult_classifies_as_powerpoint_busy() -> None:
    exc = from_com_error(_com_error(-2147418111, "Call was rejected by callee."))
    assert isinstance(exc, PowerPointBusyError)
    assert exc.retryable is True
    assert exc.hresult == -2147418111


def test_unknown_hresult_classifies_as_com_error() -> None:
    exc = from_com_error(_com_error(-2147467259, "Unspecified error"))
    assert isinstance(exc, ComError)
    assert not isinstance(exc, PowerPointBusyError)


def test_slide_not_found_is_anchor_not_found() -> None:
    err = SlideNotFoundError(9)
    assert isinstance(err, AnchorNotFoundError)
    assert err.index == 9
    assert "slide:9" in str(err)


def test_no_text_frame_carries_anchor_id() -> None:
    err = NoTextFrameError("shape:3:2")
    assert err.anchor_id == "shape:3:2"
    assert "shape:3:2" in str(err)


def test_exit_codes_match_spec() -> None:
    assert cli_main._exit_for(NoTextFrameError("shape:1:1")) == 6
    assert cli_main._exit_for(AnchorNotFoundError("shape", "shape:9:9")) == 2
    assert cli_main._exit_for(SlideNotFoundError(9)) == 2
    assert cli_main._exit_for(PresentationNotFoundError("x.pptx")) == 2
    assert cli_main._exit_for(AmbiguousMatchError("x", [])) == 5
    assert cli_main._exit_for(PowerPointBusyError()) == 3
    assert cli_main._exit_for(PowerPointNotRunningError()) == 4
    assert cli_main._exit_for(ComError("boom")) == 1

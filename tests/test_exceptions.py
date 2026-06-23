"""COM-error decoding/classification and CLI exit-code mapping."""

from __future__ import annotations

import pytest

import pptlive.cli.main as cli_main
from pptlive import _com
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


def test_rpc_unknown_if_classifies_as_busy() -> None:
    # 0x800706B5 (RPC_S_UNKNOWN_IF) — the chart embedded-Excel transient. Both the
    # unsigned and signed forms must be recognised as a retryable busy error.
    for hresult in (0x800706B5, -2147023179):
        exc = from_com_error(_com_error(hresult, "The interface is unknown."))
        assert isinstance(exc, PowerPointBusyError), hex(hresult & 0xFFFFFFFF)
        assert exc.retryable is True


def test_rpc_call_failed_classifies_as_busy() -> None:
    # 0x800706BE (RPC_S_CALL_FAILED) — a transient failure during the embedded-Excel
    # workbook teardown; retryable (re-runs the idempotent chart-data write).
    for hresult in (0x800706BE, -2147023170):
        exc = from_com_error(_com_error(hresult, "The remote procedure call failed."))
        assert isinstance(exc, PowerPointBusyError), hex(hresult & 0xFFFFFFFF)
        assert exc.retryable is True


def test_rpc_server_unavailable_is_not_busy() -> None:
    # 0x800706BA (RPC_S_SERVER_UNAVAILABLE) is deliberately NOT busy: the embedded
    # server is gone and the whole connection is poisoned for the process, so a
    # retry is futile and would mask a dead connection. It must surface as a plain
    # ComError (exit 1), not a retryable busy.
    for hresult in (0x800706BA, -2147023174):
        exc = from_com_error(_com_error(hresult, "The RPC server is unavailable."))
        assert isinstance(exc, ComError), hex(hresult & 0xFFFFFFFF)
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


def test_classify_is_single_source_for_both_front_ends() -> None:
    # _batch._error_code and cli.main._exit_for both derive from classify(), so
    # the MCP code token and the CLI exit int can't drift apart.
    from pptlive._batch import _error_code
    from pptlive.exceptions import EXIT_CODE_FOR, classify

    for exc in (
        NoTextFrameError("shape:1:1"),
        AnchorNotFoundError("shape", "shape:9:9"),
        SlideNotFoundError(9),
        PresentationNotFoundError("x.pptx"),
        AmbiguousMatchError("x", []),
        PowerPointBusyError(),
        PowerPointNotRunningError(),
        ComError("boom"),
    ):
        code = classify(exc)
        assert _error_code(exc) == code  # MCP token
        assert cli_main._exit_for(exc) == EXIT_CODE_FOR[code]  # CLI exit int


# -- retry_on_busy ----------------------------------------------------------


def test_retry_on_busy_succeeds_after_transient(monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.setattr(_com.time, "sleep", lambda _s: None)  # don't actually wait
    calls = {"n": 0}

    def flaky() -> str:
        calls["n"] += 1
        if calls["n"] < 3:
            raise PowerPointBusyError(hresult=0x800706B5)
        return "ok"

    assert _com.retry_on_busy(flaky, attempts=4, delay=0) == "ok"
    assert calls["n"] == 3  # two failures, then success


def test_retry_on_busy_reraises_after_exhausting(monkeypatch) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.setattr(_com.time, "sleep", lambda _s: None)
    calls = {"n": 0}

    def always_busy() -> None:
        calls["n"] += 1
        raise PowerPointBusyError(hresult=0x800706B5)

    with pytest.raises(PowerPointBusyError):
        _com.retry_on_busy(always_busy, attempts=3, delay=0)
    assert calls["n"] == 3  # tried exactly `attempts` times


def test_retry_on_busy_rejects_nonpositive_attempts() -> None:
    # attempts < 1 is a programming error; raise a clean ValueError rather than
    # relying on an `assert` (stripped under python -O, which would then raise a
    # confusing `raise None` TypeError).
    with pytest.raises(ValueError, match="attempts must be >= 1"):
        _com.retry_on_busy(lambda: "x", attempts=0)


def test_retry_on_busy_passes_through_non_busy() -> None:
    # A real error must surface immediately, not be retried.
    def boom() -> None:
        raise ComError("unspecified", hresult=-2147467259)

    with pytest.raises(ComError):
        _com.retry_on_busy(boom, attempts=5, delay=0)


def test_safe_read_degrades_a_failing_property_to_default() -> None:
    # The defensive-read contract: a property the object can't supply degrades to
    # the default rather than failing the whole structured read.
    def missing() -> str:
        raise AttributeError("no such property")

    assert _com.safe_read(missing, "fallback") == "fallback"


def test_safe_read_propagates_busy() -> None:
    # A genuine busy error is NOT a degradable field — it maps to the retryable
    # exit 3, so it must surface rather than masquerade as a missing value.
    def busy() -> str:
        raise PowerPointBusyError(hresult=0x800706B5)

    with pytest.raises(PowerPointBusyError):
        _com.safe_read(busy, "fallback")

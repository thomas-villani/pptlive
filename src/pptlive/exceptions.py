"""Public exception taxonomy for pptlive.

Ported from wordlive's `exceptions.py` almost verbatim — same COM-error decoding
(`_decode_com_error` / `from_com_error` / `_BUSY_HRESULTS`), same exit-code
mapping shape. The PowerPoint diff is two new members: `NoTextFrameError` (a text
op on a shape with no text frame — exit 6, the one genuinely new code) and
`SlideNotFoundError`, a subclass of `AnchorNotFoundError` so a missing slide
reuses exit 2.
"""

from __future__ import annotations

from typing import Any


class PptliveError(Exception):
    """Base class for all pptlive errors."""


class PowerPointNotRunningError(PptliveError):
    """No running PowerPoint instance is available."""


class PresentationNotFoundError(PptliveError):
    """The requested presentation is not open in PowerPoint."""

    def __init__(self, name: str) -> None:
        super().__init__(f"presentation not found: {name!r}")
        self.name = name


class AnchorNotFoundError(PptliveError):
    """The requested anchor (shape / placeholder / cell / notes) does not exist.

    Covers a missing slide too, via `SlideNotFoundError`, and a zero-match
    `find` (raised with ``kind='find'``). Maps to exit code 2.
    """

    def __init__(self, kind: str, name: str) -> None:
        super().__init__(f"{kind} not found: {name!r}")
        self.kind = kind
        self.name = name


class SlideNotFoundError(AnchorNotFoundError):
    """A slide index is out of range.

    Subclass of `AnchorNotFoundError` so it shares the same exit code (2) and
    so `except AnchorNotFoundError` catches both missing-slide and
    missing-shape errors. Retryable after re-reading `deck.slides.list()`.
    """

    def __init__(self, index: int) -> None:
        super().__init__("slide", f"slide:{index}")
        self.index = index


class LayoutNotFoundError(AnchorNotFoundError):
    """A requested slide layout name/index doesn't exist in the deck.

    Subclass of `AnchorNotFoundError` so it shares exit code 2. Layout names are
    template-dependent (a theme can rename them), so the message lists the deck's
    actual layout names and `available` carries them structured — an agent can
    read them off stderr (or `slide layouts`) and retry with a real name.
    """

    def __init__(self, requested: str, available: list[str]) -> None:
        names = ", ".join(repr(n) for n in available) if available else "(none)"
        # Build the full message first, then hand AnchorNotFoundError the bare
        # name; overwrite args so the available list survives in str(exc).
        super().__init__("layout", requested)
        self.args = (f"layout not found: {requested!r}; available: {names}",)
        self.requested = requested
        self.available = available


class NoTextFrameError(PptliveError):
    """A text operation targeted a shape with no text frame (picture, line, …).

    The one genuinely new code versus wordlive (exit 6). It's common enough —
    an LLM tries to set text on a decorative shape — to deserve a deterministic
    exit code instead of a bare COM failure. Not retryable on the same shape:
    pick a text-bearing anchor (a placeholder or text box) instead.
    """

    def __init__(self, anchor_id: str | None = None) -> None:
        target = f": {anchor_id}" if anchor_id else ""
        super().__init__(f"shape has no text frame{target}")
        self.anchor_id = anchor_id


class AmbiguousMatchError(PptliveError):
    """A find/replace pattern matched more than one occurrence without disambiguation.

    Carries the list of matches so callers (notably LLM drivers) can pick an
    `occurrence` index and retry.
    """

    def __init__(self, find: str, matches: list[dict[str, Any]]) -> None:
        super().__init__(
            f"{len(matches)} matches for {find!r}; pass --all or --occurrence N to disambiguate"
        )
        self.find = find
        self.matches = matches


class PowerPointBusyError(PptliveError):
    """PowerPoint rejected the RPC — a modal dialog, or a slide show is running.

    Retryable in principle; caller decides. While a slide show is running most
    editing calls reject with the same `RPC_E_*` HRESULTs a modal dialog
    produces, so they surface here too — steer toward `deck.show` for live
    presentation control instead.
    """

    def __init__(
        self,
        message: str = "PowerPoint is busy, in a modal dialog, or running a slide show",
        *,
        hresult: int | None = None,
    ) -> None:
        super().__init__(message)
        self.hresult = hresult
        self.retryable = True


class ComError(PptliveError):
    """Generic wrapper for an unclassified pywintypes.com_error."""

    def __init__(
        self, message: str, *, hresult: int | None = None, description: str | None = None
    ) -> None:
        super().__init__(message)
        self.hresult = hresult
        self.description = description


# HRESULTs we recognise as "PowerPoint is momentarily unavailable" rather than a
# real error. Carried over from wordlive verbatim; widen as smoke runs surface
# the slide-show-running rejection codes.
_BUSY_HRESULTS: frozenset[int] = frozenset(
    {
        0x80010001,  # RPC_E_CALL_REJECTED — call rejected by callee (modal dialog, busy)
        0x8001010A,  # RPC_E_SERVERCALL_RETRYLATER — server busy, retry later
        0x80010005,  # RPC_E_SERVERCALL_REJECTED — server rejected the call
        -2147418111,  # signed form of RPC_E_CALL_REJECTED
        -2147417846,  # signed form of RPC_E_SERVERCALL_RETRYLATER
    }
)


def _decode_com_error(exc: Any) -> tuple[int | None, str | None, str]:
    """Pull (hresult, description, readable_message) out of a pywintypes.com_error.

    pywintypes.com_error.args is (hresult, message, exc_info, arg_err) where exc_info,
    when present, is (wcode, source, description, helpfile, helpcontext, scode).
    """
    args: tuple[Any, ...] = getattr(exc, "args", ()) or ()
    hresult: int | None = None
    description: str | None = None
    message = str(exc)

    if len(args) >= 1 and isinstance(args[0], int):
        hresult = args[0]
    if len(args) >= 3 and args[2]:
        exc_info = args[2]
        try:
            description = exc_info[2] if len(exc_info) > 2 else None
            scode = exc_info[5] if len(exc_info) > 5 else None
        except (TypeError, IndexError):
            description, scode = None, None
        if scode is not None and hresult is None:
            hresult = scode

    parts = []
    if description:
        parts.append(description.strip())
    if hresult is not None:
        parts.append(f"HRESULT 0x{hresult & 0xFFFFFFFF:08X}")
    if parts:
        message = " — ".join(parts)
    return hresult, description, message


def from_com_error(exc: Any) -> PptliveError:
    """Classify a pywintypes.com_error into the appropriate pptlive exception."""
    hresult, description, message = _decode_com_error(exc)
    if hresult is not None and hresult in _BUSY_HRESULTS:
        return PowerPointBusyError(message, hresult=hresult)
    return ComError(message, hresult=hresult, description=description)

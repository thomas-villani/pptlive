"""Thin COM helpers — the mockable seam between pptlive and pywin32.

Tests substitute fakes for `get_active_powerpoint` / `launch_powerpoint` via
monkeypatch. Everything else in pptlive only sees duck-typed dispatch objects;
this is the *only* module that imports pywin32. Ported from wordlive's `_com.py`
with the ProgID swapped to `PowerPoint.Application`.

Note the PowerPoint diff: unlike Word, PowerPoint historically refuses to run
invisibly (`Application.Visible = False` raises in most builds), so
`launch_powerpoint` always shows the app. See spec.md §"Visible caveat".
"""

from __future__ import annotations

import threading
import time
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from typing import Any, TypeVar

from .exceptions import PowerPointBusyError, PowerPointNotRunningError, from_com_error

_T = TypeVar("_T")

_POWERPOINT_PROG_ID = "PowerPoint.Application"

#: Per-thread "COM already initialised on this thread" flag. Set once, never
#: cleared — see `com_apartment` for why we deliberately never `CoUninitialize`.
_apartment = threading.local()


@contextmanager
def com_apartment() -> Iterator[None]:
    """STA apartment lifecycle: initialise COM **once per thread**, never uninit.

    This diverges from wordlive's balanced `CoInitialize`/`CoUninitialize` pair,
    and the reason is the MCP server — the first *long-lived* process to drive
    pptlive. wordlive is CLI-only (one-shot processes), so a balanced cycle per
    invocation is fine there. The MCP server, by contrast, re-`attach()`es on
    *every* tool call inside one process; with the balanced pair that means a
    `CoUninitialize` after each call. Repeated `CoUninitialize` on the same
    thread destabilises pythoncom: it drops PowerPoint's automation connection
    (snapping the active window back to **slide 1** — the user-visible "jumps to
    the title slide" bug) and, under repetition, corrupts COM proxy state into a
    hard segfault. Verified live 2026-05-29 (`scripts/undo_test.py` lineage): a
    bare loop of `attach()` cycles segfaults within a few iterations, while a
    single `CoInitialize` + repeated `GetActiveObject` with *no* uninit is rock
    stable and never moves the view.

    So we `CoInitialize` the first time this runs on a given thread and leave the
    apartment open; the OS reclaims it at thread/process exit. One-shot CLI runs
    are unaffected (they init once and exit); the long-lived MCP server now holds
    a single stable apartment across all its tool calls. Tests monkeypatch the
    getters, so they never exercise this against real COM.
    """
    try:
        import pythoncom  # type: ignore[import-not-found]
    except ImportError:
        # Non-Windows or pywin32 missing: yield without initialising. The first
        # real COM call will fail with a clearer error.
        yield
        return

    if not getattr(_apartment, "initialised", False):
        pythoncom.CoInitialize()
        _apartment.initialised = True
    yield


def get_active_powerpoint() -> Any:
    """Return the Application COM object for an already-running PowerPoint, or raise."""
    try:
        from win32com.client import GetActiveObject  # type: ignore[import-not-found]
    except ImportError as e:
        raise PowerPointNotRunningError(
            "pywin32 is not installed; pptlive requires Windows + pywin32"
        ) from e

    try:
        return GetActiveObject(_POWERPOINT_PROG_ID)
    except Exception as e:  # pywintypes.com_error or similar
        raise PowerPointNotRunningError("no running Microsoft PowerPoint instance found") from e


def launch_powerpoint() -> Any:
    """Launch a new PowerPoint instance (always visible) and return its Application.

    PowerPoint must be visible — setting `Visible = False` raises in most
    builds — so there is no hidden mode. We still set `Visible = True`
    explicitly: a freshly `Dispatch`-ed instance can start hidden until the
    first window is shown.
    """
    try:
        from win32com.client import Dispatch  # type: ignore[import-not-found]
    except ImportError as e:
        raise PowerPointNotRunningError(
            "pywin32 is not installed; pptlive requires Windows + pywin32"
        ) from e

    app = Dispatch(_POWERPOINT_PROG_ID)
    try:
        app.Visible = True
    except Exception:
        # Some COM stubs may not let us flip Visible immediately; not fatal.
        pass
    return app


def retry_on_busy(
    operation: Callable[[], _T],
    *,
    attempts: int = 4,
    delay: float = 0.15,
) -> _T:
    """Call `operation`, retrying on a transient `PowerPointBusyError`.

    PowerPoint occasionally rejects an RPC transiently — a modal settling, or a
    chart's embedded-Excel workbook not yet ready right after creation
    (`RPC_S_UNKNOWN_IF` / 0x800706B5). Those map to `PowerPointBusyError`; this
    re-attempts a few times with a short, growing backoff before giving up. Any
    other exception (including a non-busy `ComError`) propagates immediately, so
    real failures still surface fast. `operation` must be idempotent — used for
    the chart-data write, which is a clean rewrite (ClearContents + SetSourceData).
    """
    last: PowerPointBusyError | None = None
    for attempt in range(attempts):
        try:
            return operation()
        except PowerPointBusyError as exc:
            last = exc
            if attempt + 1 < attempts:
                time.sleep(delay * (attempt + 1))
    assert last is not None  # only reached after a busy error
    raise last


def safe_read(fn: Callable[..., Any], default: Any) -> Any:
    """Read a COM property defensively — return `default` if it can't be supplied.

    The best-effort read used by `to_dict` / `read()` dumps: a single property the
    object can't supply (a theme-linked color, an absent axis) degrades to `default`
    rather than failing the whole structured read. Mutations never use this — they
    must surface their failures as typed `PptliveError`s via `translate_com_errors`.

    A genuine `PowerPointBusyError` is **not** swallowed: the taxonomy maps "busy"
    to exit 3 (a transient, retryable state), and silently degrading a field to its
    default would mask that. So a busy error propagates here even though every other
    failure degrades — a degraded field is local, a busy app is the whole read.
    """
    try:
        return fn()
    except PowerPointBusyError:
        raise
    except Exception:  # noqa: BLE001
        return default


@contextmanager
def translate_com_errors() -> Iterator[None]:
    """Translate pywintypes.com_error into pptlive's typed exceptions."""
    try:
        import pywintypes  # type: ignore[import-not-found]

        com_error_type: type = pywintypes.com_error
    except ImportError:
        com_error_type = ()  # type: ignore[assignment]

    try:
        yield
    except Exception as exc:  # noqa: BLE001
        if com_error_type and isinstance(exc, com_error_type):  # type: ignore[arg-type]
            raise from_com_error(exc) from exc
        raise

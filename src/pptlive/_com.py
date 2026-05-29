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

import time
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from typing import Any, TypeVar

from .exceptions import PowerPointBusyError, PowerPointNotRunningError, from_com_error

_T = TypeVar("_T")

_POWERPOINT_PROG_ID = "PowerPoint.Application"


@contextmanager
def com_apartment() -> Iterator[None]:
    """STA apartment lifecycle. Nests safely via pythoncom's reference counting."""
    try:
        import pythoncom  # type: ignore[import-not-found]
    except ImportError:
        # Non-Windows or pywin32 missing: yield without initialising. The first
        # real COM call will fail with a clearer error.
        yield
        return

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


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

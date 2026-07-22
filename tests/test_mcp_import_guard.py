"""The missing-[mcp]-extra experience: a hint, not a traceback.

`pptlive.mcp.server` guards its `mcp` SDK imports so a base `pip install
pptlive` followed by `pptlive-mcp` tells the user to install `pptlive[mcp]`
instead of dumping a raw ModuleNotFoundError. These tests simulate the missing
SDK by poisoning `sys.modules["mcp"]` (None halts the import) and forcing a
fresh import of the guarded modules.
"""

from __future__ import annotations

import importlib
import sys
from typing import Any

import pytest

pytest.importorskip("mcp")


def _hide_mcp_sdk(monkeypatch: Any) -> None:
    for name in [m for m in list(sys.modules) if m == "mcp" or m.startswith("mcp.")]:
        monkeypatch.delitem(sys.modules, name, raising=False)
    monkeypatch.setitem(sys.modules, "mcp", None)
    # Force the guarded modules to re-execute their imports.
    monkeypatch.delitem(sys.modules, "pptlive.mcp.server", raising=False)
    monkeypatch.delitem(sys.modules, "pptlive.mcp.__main__", raising=False)


def test_server_import_raises_install_hint(monkeypatch: Any) -> None:
    _hide_mcp_sdk(monkeypatch)
    with pytest.raises(ImportError, match=r"pptlive\[mcp\]"):
        importlib.import_module("pptlive.mcp.server")


def test_console_entry_exits_1_with_hint(monkeypatch: Any, capsys: Any) -> None:
    _hide_mcp_sdk(monkeypatch)
    main_mod = importlib.import_module("pptlive.mcp.__main__")
    with pytest.raises(SystemExit) as excinfo:
        main_mod.main()
    assert excinfo.value.code == 1
    assert "pptlive[mcp]" in capsys.readouterr().err

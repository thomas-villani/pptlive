"""Spike — prove pptlive's COM seam survives FastMCP's dispatch model.

The MCP stdio server runs an asyncio event loop, and PowerPoint COM objects are
STA-thread-bound — so the question this spike answers, *before* we harden the
server, is: when FastMCP dispatches a tool, does pptlive's per-call
`attach()` (which does `CoInitialize` -> `GetActiveObject` -> `CoUninitialize`)
run cleanly, or does the apartment/thread model blow up?

Finding from reading the SDK (mcp >= 1.x): FastMCP calls a **sync** tool
function *directly* on the event-loop thread (no thread-pool offload — see
`func_metadata.call_fn_with_arg_validation`). So a sync tool's whole
init/work/uninit cycle happens on one consistent thread per call — the same
shape as a one-shot CLI invocation, just repeated in a long-lived process. That
is STA-safe; the only cost is that a COM call briefly blocks the loop, which is
fine for a single user driving PowerPoint serially.

This script confirms that empirically by driving the real `FastMCP.call_tool`
path (the same code the stdio transport runs):

    uv run python scripts/mcp_spike.py

It is read-only and net-zero: the probe tool only *reads* (`presentations.list`,
viewed slide). With PowerPoint closed it should report `powerpoint: not_running`
with a clean typed `PowerPointNotRunningError` — which still proves the COM call
executed on the tool thread without an STA/CoInitialize crash. With a deck open
it returns the live status. Either outcome is a pass; a `pythoncom`/threading
traceback would be the failure.
"""

from __future__ import annotations

import asyncio
import json
import threading
from typing import Any

from mcp.server.fastmcp import FastMCP

import pptlive as pl
from pptlive.exceptions import PowerPointNotRunningError

mcp = FastMCP("pptlive-spike")


@mcp.tool()
def ppt_probe() -> dict[str, Any]:
    """Read-only probe: records the thread it runs on, then attempts attach()."""
    info: dict[str, Any] = {"tool_thread": threading.current_thread().name}
    try:
        with pl.attach() as ppt:
            info["powerpoint"] = "running"
            info["decks"] = [d.get("name") for d in ppt.presentations.list()]
            info["viewed_slide"] = ppt.viewed_slide_index()
    except PowerPointNotRunningError as exc:
        # Reaching this proves CoInitialize + GetActiveObject ran on the tool
        # thread and raised the *expected typed* error — plumbing intact.
        info["powerpoint"] = "not_running"
        info["attach_error"] = str(exc)[:120]
    return info


async def main() -> int:
    findings: dict[str, Any] = {"loop_thread": threading.current_thread().name}
    findings["tools"] = [t.name for t in await mcp.list_tools()]
    try:
        result = await mcp.call_tool("ppt_probe", {})
        findings["call_ok"] = True
        # call_tool returns (content_blocks, structured) or a single value across
        # SDK versions; stringify defensively so the spike is version-agnostic.
        findings["result"] = result
    except Exception as exc:  # noqa: BLE001 — the spike wants to SEE any blow-up
        findings["call_ok"] = False
        findings["error_type"] = type(exc).__name__
        findings["error"] = str(exc)[:300]

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))

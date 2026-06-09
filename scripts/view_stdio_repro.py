"""Definitive view-jump repro: drive the REAL `pptlive-mcp` server over stdio,
exactly as Claude Desktop does — a separate long-lived process, JSON-RPC over
stdio, several sequential tool calls on its own asyncio event loop.

Between calls we read the live viewed-slide index through a short side-channel
`attach()` in THIS process, so we can see whether a server tool call moved the
user's view. Park the deck on a slide first; if `viewed` collapses to 1 after a
call, we've reproduced the regression against the genuine surface.

    uv run python scripts/view_stdio_repro.py
"""

from __future__ import annotations

import asyncio
import json
import sys
from typing import Any

from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

import pptlive as pl

TARGET_SLIDE = 3


def _park(index: int) -> int | None:
    with pl.attach() as ppt:
        try:
            ppt.com.ActiveWindow.View.GotoSlide(index)
        except Exception:
            pass
        return ppt.viewed_slide_index()


def _viewed() -> int | None:
    with pl.attach() as ppt:
        return ppt.viewed_slide_index()


def _last_slide_delete() -> None:
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        deck.slides[len(deck.slides)].delete()


async def main() -> int:
    # Launch the real server the same way the mcpb manifest does in spirit:
    # the installed `pptlive-mcp` console script (pptlive.mcp.__main__:main).
    params = StdioServerParameters(command="uv", args=["run", "pptlive-mcp"])
    trace: list[dict[str, Any]] = []

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()

            parked = _park(TARGET_SLIDE)
            trace.append({"step": "parked", "viewed": parked})

            # 1) read status (should never move the view)
            await session.call_tool("ppt_read", {"op": "status"})
            trace.append({"step": "after read status", "viewed": _viewed()})

            # 2) write to the title (an edit — wrapped in edit())
            with pl.attach() as ppt:
                title = ppt.presentations.active.slides[TARGET_SLIDE].placeholder("title")
                original = title.text
            await session.call_tool(
                "ppt_edit",
                {
                    "op": "write",
                    "anchor_id": f"ph:{TARGET_SLIDE}:title",
                    "text": "stdio repro write",
                },
            )
            trace.append({"step": "after edit write", "viewed": _viewed()})

            # 3) atomic batch: slide_add + write
            await session.call_tool(
                "ppt_batch",
                {
                    "commands": [
                        {"tool": "edit", "op": "slide_add"},
                        {
                            "tool": "edit",
                            "op": "write",
                            "anchor_id": f"ph:{TARGET_SLIDE}:title",
                            "text": "stdio repro batch",
                        },
                    ],
                    "atomic": True,
                    "embed": False,
                },
            )
            trace.append({"step": "after batch (slide_add+write)", "viewed": _viewed()})

            # cleanup: revert title + drop the slide the batch added
            await session.call_tool(
                "ppt_edit",
                {"op": "write", "anchor_id": f"ph:{TARGET_SLIDE}:title", "text": original},
            )
            _last_slide_delete()
            trace.append({"step": "after cleanup", "viewed": _viewed()})

    print(json.dumps(trace, indent=2, default=str))
    jumped = [t for t in trace if t.get("viewed") == 1 and t["step"] != "parked"]
    print(f"\n=== steps where view jumped to slide 1: {jumped or 'none'} ===", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))

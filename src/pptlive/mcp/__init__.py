"""Model Context Protocol (MCP) server for pptlive — optional `pptlive[mcp]` extra.

Exposes pptlive's live-PowerPoint control to MCP clients (Claude Desktop and any
other MCP-compatible agent) as a small, curated set of tools over stdio. Run it
with the `pptlive-mcp` console script, or `python -m pptlive.mcp`.

This package needs the `mcp` SDK (the `pptlive[mcp]` extra): importing it
without the extra raises an ImportError carrying the install hint, and the
`pptlive-mcp` script turns that into a clean exit-1 message. CLI-only users
never load it.

The tools wrap the same public API as the CLI (`attach()` -> `deck` -> anchors),
so the politeness model (preserve the viewed slide + Selection) and atomic-undo
fencing (`deck.edit`) carry over unchanged. See `server.py` for the tool list.
"""

from __future__ import annotations

from .server import build_server, server

__all__ = ["build_server", "server"]

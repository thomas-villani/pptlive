"""Console entry point for the pptlive MCP server (`pptlive-mcp` / `python -m pptlive.mcp`).

Runs the FastMCP server over stdio — the transport Claude Desktop and most MCP
clients launch. See `pptlive.mcp.server` for the tools, and the README's MCP
section for the Claude Desktop config snippet.
"""

from __future__ import annotations

import sys


def main() -> None:
    """Run the stdio MCP server (blocks until the client disconnects)."""
    # Import here, not at module top: a base `pip install pptlive` (no [mcp]
    # extra) must exit with the install hint, not a traceback.
    try:
        from .server import server
    except ImportError as exc:
        if "pptlive[mcp]" not in str(exc):  # a genuine bug, not the missing extra
            raise
        sys.stderr.write(f"{exc}\n")
        raise SystemExit(1) from exc
    server.run()


if __name__ == "__main__":
    main()

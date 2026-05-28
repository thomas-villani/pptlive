"""Console entry point for the pptlive MCP server (`pptlive-mcp` / `python -m pptlive.mcp`).

Runs the FastMCP server over stdio — the transport Claude Desktop and most MCP
clients launch. See `pptlive.mcp.server` for the tools, and the README's MCP
section for the Claude Desktop config snippet.
"""

from __future__ import annotations

from .server import server


def main() -> None:
    """Run the stdio MCP server (blocks until the client disconnects)."""
    server.run()


if __name__ == "__main__":
    main()

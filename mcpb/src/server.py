"""MCPB launcher for the pptlive MCP server (uv runtime).

Thin wrapper over pptlive's ``pptlive-mcp`` entry point
(``pptlive.mcp.__main__:main``). It exists only to give the MCPB ``uv`` runtime
a stable entry file; all the real logic lives in the installed ``pptlive``
package. The server needs no configuration — it drives whatever PowerPoint deck
the user has open — so there are no ``user_config`` env vars to read.
"""

from pptlive.mcp.__main__ import main

if __name__ == "__main__":
    raise SystemExit(main())

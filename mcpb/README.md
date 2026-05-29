# pptlive MCPB bundle

This directory packages pptlive's built-in MCP server as an
[MCPB bundle](https://github.com/modelcontextprotocol/mcpb) (`.mcpb`) for
**one-click install** in Claude Desktop (Settings → Extensions).

It is **not** a new server. It is a thin packaging layer around the existing
`pptlive-mcp` console script (`pptlive.mcp.__main__:main`):

- `manifest.json` — bundle metadata and the `uv` server config. PowerPoint must
  be installed and running; the server drives whatever deck the user has open,
  so there is **no `user_config`** (no workspace folder, no toggles).
- `pyproject.toml` — declares the real published package as the dependency
  (`pptlive[mcp]`). The `uv` runtime resolves and installs it on the user's
  machine at install time, so nothing is bundled in the `.mcpb`.
- `src/server.py` — a stable entry file that just imports and calls `main()`.

**Windows only.** pptlive automates PowerPoint over COM (pywin32), so the
manifest's `compatibility.platforms` is `["win32"]`.

## Installing (for users)

1. Download `pptlive.mcpb` from the
   [latest release](https://github.com/thomas-villani/pptlive/releases/latest).
2. Open **Claude Desktop → Settings → Extensions**.
3. **Drag `pptlive.mcpb` onto the Extensions pane** (or use **Install
   Extension** / **Advanced** to browse for it).
4. The `ppt_read`, `ppt_edit`, `ppt_render`, `ppt_show`, and `ppt_batch` tools
   then appear under the **"+" → Connectors** panel, ready to use on the deck you
   have open in PowerPoint.

Requires a Claude Desktop build with MCPB extension support (late-2025 or newer)
and `uv` available to the host. As an alternative to the bundle, `pptlive
install-mcp` writes the same server entry straight into a client's config file.

## Rebuilding

```bash
npm install -g @anthropic-ai/mcpb   # one-time
mcpb validate manifest.json
mcpb pack . pptlive.mcpb             # from this dir (bare `mcpb pack` names it mcpb.mcpb)
```

## Versioning

The `version` in `manifest.json` and `pyproject.toml` (and the `>=` dependency
pin) must stay in sync with the main package's version in the root
`pyproject.toml`. Bump all three together when releasing.

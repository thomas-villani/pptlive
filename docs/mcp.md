# MCP server

pptlive ships an optional [MCP](https://modelcontextprotocol.io) server so MCP
clients — **Claude Desktop**, Cursor, and other agent hosts — can read and edit
the PowerPoint deck you have open right now, including *seeing* rendered slides.

It talks to the same running PowerPoint instance the CLI and Python API do, over
COM, on Windows. Edits stay polite (your viewed slide and selection are
preserved; each write is a single Ctrl-Z), and failures come back with a stable
error category token.

## Install

The server is an optional extra (it pulls in the official `mcp` SDK):

```
pip install "pptlive[mcp]"

# uv
uv tool install "pptlive[mcp]"
```

## Run

```
pptlive-mcp            # console script (stdio transport)
python -m pptlive.mcp  # equivalent
```

The server speaks MCP over **stdio** — the transport Claude Desktop spawns.
PowerPoint must already be running on the same machine (pptlive *attaches*; it
never launches or closes it, and it can't run PowerPoint hidden).

## Register with Claude Desktop

Add an entry to `claude_desktop_config.json` (Claude Desktop → Settings →
Developer → Edit Config), then restart:

```json
{
  "mcpServers": {
    "pptlive": {
      "command": "pptlive-mcp"
    }
  }
}
```

If `pptlive-mcp` isn't on Claude Desktop's `PATH`, point at the interpreter in
your environment explicitly (note the doubled backslashes in JSON):

```json
{
  "mcpServers": {
    "pptlive": {
      "command": "C:\\Users\\you\\project\\.venv\\Scripts\\python.exe",
      "args": ["-m", "pptlive.mcp"]
    }
  }
}
```

Restart Claude Desktop, open a `.pptx` in PowerPoint, and the `ppt_*` tools
appear.

## Tools

A compact **five-tool dispatch surface** keeps the client's tool list (and its
context cost) lean: each tool takes an `op` (or `command`) argument and routes
to the right verb, so the agent's tool picker sees five definitions instead of
fifteen. They wrap the same Python API, so the politeness model and one-Ctrl-Z
`edit` fencing carry over and reads never move the view.

| Tool | `op`s |
| ---- | ----- |
| `ppt_read` | `status` · `slides` · `outline` · `slide` · `anchor` · `selection` · `table` · `chart` · `layouts` — every read; never moves the view |
| `ppt_edit` | `write` · `format` · `slide_add` / `slide_delete` / `slide_duplicate` / `slide_move` / `set_layout` · `shape_add` / `shape_move` / `shape_resize` / `shape_delete` / `set_alt` · `table_add_row` / `table_delete_row` · `chart_set_type` / `chart_set_data` — every mutation; one Ctrl-Z each |
| `ppt_render` | `slide_image` · `shape_image` (PNGs a vision model can read) · `navigate` (the one deliberate view move) |
| `ppt_show` | live slide show: `state` · `start` · `end` · `next` · `previous` · `goto` · `black` · `white` · `resume` |
| `ppt_batch` | run a **list** of the ops above against one connection — `edit`s fenced into a **single** undo entry (`atomic`), with `stop_on_error` control |

Tables and charts are addressed by their shape's `anchor_id` (a `shape:S:N`);
cells stay `cell:S:N:R:C` anchors you write to with `ppt_edit op="write"`. The
full anchor model (`shape:S:N`, `ph:S:KIND`, `para:S:N:P`, `cell:S:N:R:C`,
`notes:S`, `here:`) is documented under [Concepts](concepts.md#anchor-ids), and
each op's fields mirror the [CLI](cli.md).

### Batches

`ppt_batch` runs a list of commands against one connection — the power tool for
multi-step intents:

```json
{
  "commands": [
    {"tool": "ppt_edit", "op": "slide_add", "params": {"layout": "title_and_content", "index": 4}},
    {"tool": "ppt_edit", "op": "write", "params": {"anchor_id": "ph:4:title", "text": "Q3 Results"}},
    {"tool": "ppt_edit", "op": "write", "params": {"anchor_id": "ph:4:body",  "text": "Revenue up 12%\nChurn down 3%"}}
  ],
  "atomic": true,
  "stop_on_error": true
}
```

- `atomic` (default `true`) — every `edit` command is fenced into a **single**
  undo entry, so the whole batch is one Ctrl-Z. With `atomic=false` each edit is
  its own entry.
- `stop_on_error` (default `true`) — stop at the first failing command. With
  `false`, the batch runs to completion and reports each command's outcome.

It returns `{"ok", "atomic", "count", "results": [...]}`, where each result
carries the same category token (below) that the other tools' errors do.

## Errors

A failed tool call comes back as an MCP `ToolError` whose message is prefixed
with a stable **category token** — the string analog of the CLI's
[exit-code taxonomy](errors.md) — so an agent can branch on the failure mode:

```
AnchorNotFoundError (not_found): shape not found: 'shape:9:9'
```

| token | Meaning | retry? |
| --- | --- | --- |
| `not_found` | anchor / slide / shape / layout / deck missing | no — re-read first |
| `ambiguous` | a fuzzy match hit more than one target | yes — disambiguate |
| `busy` | a modal dialog is open / RPC rejected | **yes** — back off and retry |
| `not_running` | PowerPoint isn't running | no — until it's opened |
| `no_text_frame` | text op on a shape with no text frame | no — pick a text-bearing shape |
| `invalid_args` | bad / missing arguments, unknown op | no — fix the request |
| `error` | other | no |

Inside `ppt_batch` the same tokens are reported per-command instead of aborting
the whole call (when `stop_on_error=false`).

## How it works

Each tool is deliberately **synchronous** and attaches to the running instance
fresh, exactly like a one-shot CLI invocation — it just repeats in a long-lived
process. A 2026 spike confirmed FastMCP calls a sync tool function directly on
its event-loop thread (no thread-pool offload), so every tool's
`CoInitialize → work → CoUninitialize` cycle runs on one consistent thread per
call. That's STA-safe, which is why tools never cache a COM object across calls.
The only cost is that a COM call briefly blocks the event loop — fine for a
single user driving PowerPoint serially.

The server is in-process: it calls the pptlive Python API directly rather than
shelling out, which is also how `ppt_render` returns native image content for a
vision model.

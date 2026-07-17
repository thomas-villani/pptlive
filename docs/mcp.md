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
| `ppt_read` | `status` · `slides` · `outline` · `slide` · `anchor` · `geometry` (slide size + per-shape boxes + overlaps + off-slide) · `selection` · `find` · `table` · `chart` · `smartart` · `comments` · `animations` (a slide's shape animations in play order) · `links` (an anchor's text-run hyperlinks) · `sections` · `headers_footers` (slide-or-master by presence of `slide`) · `theme` · `master` · `layouts` · `text_frame_status` (autofit / wrap / margins / overflow-risk) — every read; never moves the view |
| `ppt_edit` | `write` · `set_paragraphs` (rewrite an anchor as a clean per-paragraph list — the safe bullet path) · `find_replace` · `format` (font + paragraph + shape **fill / line** [+ `fill_transparency` / `line_transparency` alpha] + list bullets in one op; `line_spacing` is a multiple, `line_spacing_points` is exact points) · `text_reset_format` / `shape_reset_layout` (recover a wrecked placeholder) · `slide_add` / `slide_delete` / `slide_duplicate` / `slide_move` / `set_layout` · `shape_add` / `shape_move` / `shape_resize` / `shape_delete` / `shape_order` (z-order) / `set_alt` · `media_add` (insert audio/video narration — autoplay + pace the slide to the clip) / `media_set` (mute / volume / trim an existing clip; trim is `trim_start` / `trim_end`, in seconds) · `shape_group` / `shape_ungroup` / `shape_align` / `shape_distribute` (arrange a set of shapes; `relative_to` = slide or selection) / `shape_add_connector` (a line glued to two shapes, or free-floating) · `shape_set_picture` (re-source a picture in place, keeping geometry / z-order) · `link_set` / `link_remove` (text-**run**-level hyperlinks — a substring or an explicit offset; distinct from the whole-shape `shape_set_hyperlink`) · `shape_gradient_fill` / `shape_picture_fill` / `shape_pattern_fill` (advanced fills) / `shape_set_effect` (shadow/glow/soft-edge/reflection) / `shape_line_style` (dash + arrowheads) · `shape_animate` (entrance / exit effect) / `shape_clear_animations` / `slide_clear_animations` · `shape_set_hyperlink` / `shape_remove_hyperlink` · `slide_set_transition` / `slide_set_background` · `table_add_row` / `table_delete_row` / `table_add_column` / `table_delete_column` · `table_set_fill` / `table_set_border` (cell shading + borders, row/column-wise — `rows` / `cols` are `null` (whole axis) / an int / a list, and the *intersection* is styled) · `chart_set_type` / `chart_set_data` / `chart_recolor_text` · `smartart_set_nodes` / `smartart_recolor_text` / `smartart_format_node` (one node's label, addressed by its depth-first `node_index`) · `comment_add` / `comment_reply` / `comment_delete` · `section_add` / `section_rename` / `section_delete` / `section_move` · `set_headers_footers` (slide override or master default by presence of `slide`) · `theme_set_color` / `theme_set_font` · `master_format_text_style` / `master_format_paragraph_style` / `master_set_background` — every mutation; one Ctrl-Z each |
| `ppt_render` | `slide_image` · `shape_image` · `deck_snapshot` (one image per slide — the whole-deck vision read; `max_dim` caps each slide's long edge, or pass exact `width` / `height`) — PNGs a vision model can read · `deck_pdf` / `save` / `save_as` (explicit output; pptlive never auto-saves) · `export_video` (deck → MP4 via async `CreateVideo`; blocks until done by default, or `wait=false` + poll) / `video_status` · `navigate` (the one deliberate view move) |
| `ppt_show` | live slide show: `state` · `start` · `end` · `next` · `previous` · `goto` · `black` · `white` · `resume` |
| `ppt_batch` | run a **list** of the ops above against one connection — `edit`s fenced into a **single** undo entry (`atomic`), with `stop_on_error` control |

Tables, charts, and SmartArt are addressed by their shape's `anchor_id` (a
`shape:S:N`); cells stay `cell:S:N:R:C` anchors you write to with `ppt_edit
op="write"`. The `theme_*` and `master_*` ops are deck-wide (no anchor) — one
call restyles every inheriting slide. The
full anchor model (`shape:S:N`, `shapeid:S:ID` — the delete-proof handle —
`ph:S:KIND`, `para:S:N:P`, `cell:S:N:R:C`, `notes:S`, `comments:S`, `here:`) is
documented under [Concepts](concepts.md#anchor-ids), and each op's fields mirror
the [CLI](cli.md).

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
- `follow_view` (default on) — when a batch *adds* a slide (`slide_add` /
  `slide_duplicate`), the view is left on the last slide it touched instead of
  snapped back to the pre-batch slide (so building a deck doesn't keep bouncing the
  user to slide 1). Pure-edit batches keep the polite view-restore. Pass
  `follow_view=false` to always restore; a deliberate `navigate` still wins.

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

Each tool is deliberately **synchronous** and re-attaches to the running
instance fresh on every call (a cheap `GetActiveObject`, so it never caches a COM
proxy and stays robust to the user closing/reopening a deck). A 2026 spike
confirmed FastMCP calls a sync tool function directly on its event-loop thread
(no thread-pool offload), so every tool runs on one consistent thread. COM is
`CoInitialize`d **once** for that thread and held open for the life of the
process — *not* torn down per call. (An earlier design re-`CoUninitialize`d after
each call; that repeatedly dropped PowerPoint's automation connection — snapping
its view back to the title slide — and eventually segfaulted, so the apartment is
now kept open for the session. See `_com.com_apartment`.) The only cost is that a
COM call briefly blocks the event loop — fine for a single user driving
PowerPoint serially.

The server is in-process: it calls the pptlive Python API directly rather than
shelling out, which is also how `ppt_render` returns native image content for a
vision model.

# pptlive

Drive a running Microsoft PowerPoint instance from Python — `xlwings`, but for
PowerPoint. Built for both human scripting and LLM agents. Windows-only.

The live-app sibling of [`python-pptx`](https://python-pptx.readthedocs.io/)
(which works the `.pptx` on disk) and the PowerPoint counterpart of
[`wordlive`](../wordlive). Use it when the user already has the deck open and you
want to edit it *live* — no close-the-file, let-the-agent-write, re-open dance.

| Library     | Target                    | Mechanism                |
| ----------- | ------------------------- | ------------------------ |
| python-pptx | a `.pptx` file on disk    | OOXML I/O                |
| **pptlive** | **a running POWERPNT.exe**| **COM automation (pywin32)** |

## Install

```
pip install pptlive

# with the MCP server for Claude Desktop & other MCP agents (see "MCP server")
pip install "pptlive[mcp]"

# add to a project
uv add pptlive
```

(Requires Python 3.10+ and `pywin32` on Windows.)

## Python

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active

    # Reads (structured, side-effect-free)
    slides  = deck.slides.list()        # [{index, id, layout, title, shape_count, has_notes}]
    outline = deck.outline()            # [{slide, title, bullets:[...]}]
    grid    = deck.slides[2].read()     # every shape: anchor_id, name, id, type, geometry, text
    title   = deck.slides[2].placeholder("title").text
    notes   = deck.slides[1].notes.text

    # Polite writes — preserve the user's viewed slide + Selection.
    with deck.edit("Revise the agenda slide"):
        deck.anchor_by_id("ph:2:title").set_text("Agenda")
        deck.anchor_by_id("ph:2:body").set_text("Intro\nDemo\nQ&A")

    # Slide lifecycle — also one Ctrl-Z per edit() block.
    names = deck.layouts()                   # [{index, name}] — what set_layout/add accept
    with deck.edit("Add a results slide"):
        new = deck.slides.add(layout="two_content", index=4)
        deck.slides[7].duplicate()           # copy lands at slide 8
        deck.slides[9].move_to(2)
        deck.slides[4].set_layout("title_and_content")

    # Shapes & geometry — points throughout (pl.units for inches/cm).
    with deck.edit("Lay out the results slide"):
        shapes = deck.slides[4].shapes
        shapes.add_textbox("Revenue up 12%", left=pl.units.inches(1), top=72)
        star = shapes.add_shape("star", left=400, top=120, width=120, height=120)
        shapes.add_picture("logo.png", left=600, top=40)   # embedded, never linked
        deck.slides[4].shapes["Picture 3"].move(top=140)   # absolute, points
        star.delete()

    # Text structure — paragraphs, formatting, bullets (PowerPoint has no
    # named styles, so "styling" is direct font formatting via format_text).
    with deck.edit("Polish the body copy"):
        body = deck.anchor_by_id("ph:4:body")
        body.set_text("Revenue up 12%\nChurn down 3%\nNPS +9")
        body.apply_list("bulleted")                 # bullets on every paragraph
        body.paragraph(2).format_paragraph(indent_level=2, alignment="left")
        body.paragraph(1).format_text(bold=True, size=24, color="#2E74B5")
        body.insert_paragraph_after("Cash runway: 30 months")   # append a bullet

    # Tables — a table is a shape (Shape.has_table); cells are cell:S:N:R:C anchors.
    with deck.edit("Add a metrics table"):
        table = deck.slides[4].shapes.add_table(rows=3, columns=2).table
        table.cell(1, 1).set_text("Metric")
        table.cell(1, 2).set_text("Q3")
        table.add_row(["Revenue", "$4.2M"])          # appends + fills a row
        deck.anchor_by_id("cell:4:5:1:1").format_text(bold=True)   # a Cell is an anchor
    grid = table.read()                              # {slide, shape, rows, columns, cells:[...]}

    # Render — let a vision model *see* the slide it just built (export → read → iterate).
    png = deck.slides[4].export_image(width=1280)    # temp PNG (or pass a path); polite
    #   ...hand `png` to your image tool, look, then revise.

    # Read what the user is looking at, and (opt-in) target it with the here: anchor.
    sel = deck.selection()                           # {type, slide, anchor_id, shapes, ...}
    if sel.anchor_id:
        with deck.edit("Bold the selected text"):
            deck.anchor_by_id("here:").format_text(bold=True)

    # Live slide show — drive the presentation like a clicker (deliberately moves the screen).
    deck.show.start()                                # run from the top
    deck.show.goto(5); deck.show.next(); deck.show.black()   # jump, advance, blank
    deck.show.state()                                # {running, state, current_slide, ...}
    deck.show.end()
```

## Anchors

Addressing is **hierarchical** (slide → shape → text), not a global character
stream — there is no deck-wide `range:`. Anchor ids are colon-separated,
slide-index first:

| anchor_id      | resolves to |
| -------------- | ----------- |
| `shape:S:N`    | Nth shape (1-based z-order) on slide S — the canonical handle |
| `ph:S:KIND`    | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — the LLM-preferred form |
| `para:S:N:P`   | paragraph P (1-based) of shape N on slide S |
| `cell:S:N:R:C` | cell (row R, col C) of the table in shape N on slide S — a `Cell` *is* an anchor, so it takes every text/format verb |
| `notes:S`      | speaker-notes body of slide S |
| `here:`        | whatever the user has selected right now — the shape, or the paragraph holding the text caret (the opt-in way to act on the live selection) |

z-order **drifts** when shapes are added or removed, so `shape:S:N` is resolved
live and never cached; every shape listing also emits `name` (`Shape.Name`) and
`id` (`Shape.Id`, stable across reorder) so you can re-identify after drift.
Steer toward `ph:S:KIND` and `.Name` as the drift-proof forms. `para:S:N:P` and
`cell:S:N:R:C` also resolve live (the paragraph/row count shifts as text or rows
are inserted/deleted).

## CLI

JSON in, JSON out, deterministic exit codes — designed to drop straight into an
LLM tool-use loop. Global flags (`--json`/`--text`, `--doc NAME`) go *before* the
subcommand.

```
pptlive status                                   # open decks, active one, slide in view
pptlive slides                                   # [{index, id, layout, title, shape_count, has_notes}]
pptlive outline                                  # title + body bullets per slide
pptlive slide read 2                             # every shape on slide 2
pptlive shapes --slide 2                         # shapes on slide 2 (anchor_id, name, id, type, geometry)

pptlive read anchor --anchor-id ph:2:title       # read any text anchor (ph:/shape:/notes:)
pptlive read notes --slide 1                     # sugar for --anchor-id notes:1
pptlive write   --anchor-id ph:2:body  --text "Intro\nDemo\nQ&A"
pptlive replace --anchor-id shape:3:1  --text "New text"

pptlive slide layouts                            # the layout names add/set-layout accept
pptlive slide add --layout two_content [--index 4]
pptlive slide duplicate --slide 7
pptlive slide move --slide 9 --to 2
pptlive slide set-layout --slide 4 --layout title_and_content
pptlive slide delete --slide 5
pptlive slide export --slide 2 --out slide2.png [--width 1280] [--format png]  # render to image

pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72
pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120
pptlive shape add --slide 4 --kind picture --path logo.png --left 600 --top 40
pptlive shape add --slide 4 --kind table --rows 3 --cols 2 --left 72 --top 120
pptlive shape move   --anchor-id shape:4:3 --left 100 --top 140
pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200
pptlive shape delete --anchor-id shape:4:3

pptlive paragraphs --anchor-id ph:4:body         # [{anchor_id (para:S:N:P), text, indent_level, bullet}]
pptlive insert --anchor-id para:4:2:3 --text "New bullet" [--before|--after]
pptlive format-paragraph --anchor-id para:4:2:1 --alignment center --indent-level 2
pptlive format-text --anchor-id ph:4:title --bold --size 40 --color "#2E74B5"
pptlive list apply  --anchor-id ph:4:body --type bulleted [--char "•"]
pptlive list remove --anchor-id ph:4:body

pptlive table read --slide 4 --shape 5           # grid of cells, each with its cell:S:N:R:C anchor
pptlive table add-row    --slide 4 --shape 5 --values '["Revenue", "$4.2M"]'
pptlive table delete-row --slide 4 --shape 5 --row 2
pptlive write --anchor-id cell:4:5:1:1 --text "Metric"   # a cell takes write/format-text/...

pptlive selection                                # what the user has selected (-> here:)
pptlive read anchor --anchor-id here:            # read the selected shape/paragraph
pptlive go-to --anchor-id shape:3:1              # deliberate, opt-in view move

pptlive show start [--from 2]                    # run the slide show (deliberately moves the screen)
pptlive show next                                # advance; also: prev, goto --slide N
pptlive show black                               # blank to black (white / resume too)
pptlive show state                               # {running, state, current_slide, ...} (read-only)
pptlive show end
```

Exit codes: `0` ok · `1` other · `2` anchor/slide/shape/presentation not found ·
`3` PowerPoint busy / modal dialog · `4` PowerPoint not running · `5`
ambiguous match · `6` shape has no text frame.

## MCP server

The same live-PowerPoint control, exposed to **Claude Desktop** (and any other
MCP client) as a small set of tools. Install the extra and point your client at
the `pptlive-mcp` stdio server:

```
pip install "pptlive[mcp]"
pptlive-mcp            # stdio MCP server (or: python -m pptlive.mcp)
```

Claude Desktop — add to `claude_desktop_config.json` (Settings → Developer →
Edit Config), then restart:

```json
{
  "mcpServers": {
    "pptlive": { "command": "pptlive-mcp" }
  }
}
```

(If `pptlive-mcp` isn't on the launcher's PATH, use the absolute path to the
script — or `"command": "uv", "args": ["run", "pptlive-mcp"]` from the project.)

It's a **curated ~14-tool** surface (smaller than the full CLI — several tools
take a verb-style `op`/`mode` argument), wrapping the same API, so the politeness
model and one-Ctrl-Z `edit` fencing carry over and reads never move the view:

| tool | what it does |
| ---- | ------------ |
| `ppt_status` / `ppt_slides` / `ppt_outline` | what's open; list slides; the outline |
| `ppt_slide_read` | every shape on a slide (anchor_id, type, geometry, text, has-table) |
| `ppt_read` | read any anchor's text (+ a `paragraphs` breakdown) |
| `ppt_selection` | what the user has selected (→ `here:`) |
| `ppt_write` | set text, or insert a paragraph (`mode`) |
| `ppt_format` | font + paragraph + bullets in one call |
| `ppt_slide_op` | `add` / `delete` / `duplicate` / `move` / `set_layout` / `layouts` |
| `ppt_shape_op` | `add` (textbox/shape/picture/table) / `move` / `resize` / `delete` |
| `ppt_table` | `read` / `add_row` / `delete_row` (cells are `cell:S:N:R:C` anchors) |
| `ppt_export` | render a slide to a PNG a vision model can read |
| `ppt_show` | live slide show: `state` / `start` / `end` / `next` / `previous` / `goto` / `black` / `white` / `resume` |
| `ppt_navigate` | deliberately move the user's view to an anchor |

Tool failures surface as MCP errors carrying a category token — `not_found`,
`ambiguous`, `busy`, `not_running`, `no_text_frame` — the string analog of the
CLI's exit codes, so an agent can branch on them.

## Two things to know

- **Politeness.** By default every operation preserves the slide the user is
  looking at and their shape/text selection. Only verbs that *must* move the
  view say so in their name (`go_to`, `allow_view_move()`).
- **Atomic undo — one Ctrl-Z per block.** PowerPoint has no `UndoRecord`, but it
  groups the COM edits made inside a single `deck.edit(...)` block into one undo
  entry (the scope fences the block with `StartNewUndoEntry`), so a whole block
  reverts with a single Ctrl-Z. The one caveat: there's no explicit "end" fence —
  always wrap mutations in `deck.edit(...)` rather than editing bare, so each
  block stays cleanly self-contained.

## Development

```
uv sync --extra dev
uv run pytest                 # unit tests (fake COM; no PowerPoint needed)
uv run pytest -m smoke        # smoke suite — needs PowerPoint installed
uv run ruff check . && uv run ruff format .
uv run mypy
```

The library targets Python 3.10+ (dev pins 3.13). See `spec.md` for the design
and `IMPLEMENTATION.md` for staged build progress. Windows + COM only.

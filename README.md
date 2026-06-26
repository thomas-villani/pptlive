# pptlive

[![PyPI version](https://img.shields.io/pypi/v/pptlive.svg)](https://pypi.org/project/pptlive/)
[![Python versions](https://img.shields.io/pypi/pyversions/pptlive.svg)](https://pypi.org/project/pptlive/)
[![License: MIT](https://img.shields.io/pypi/l/pptlive.svg)](https://github.com/thomas-villani/pptlive/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-github.io-blue.svg)](https://thomas-villani.github.io/pptlive/)

Drive a running Microsoft PowerPoint instance from Python — `xlwings`, but for
PowerPoint. Built for both human scripting and LLM agents. Windows-only.

📖 **[Documentation](https://thomas-villani.github.io/pptlive/)** — full Python API, CLI, MCP server, and cookbook.

The live-app sibling of [`python-pptx`](https://python-pptx.readthedocs.io/)
(which works the `.pptx` on disk) and the PowerPoint counterpart of
[`wordlive`](https://github.com/thomas-villani/wordlive). Use it when the user already has the deck open and you
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

(Requires Python 3.11+ and `pywin32` on Windows.)

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
        star = shapes.add_shape("star", left=400, top=120, width=120, height=120,
                                fill="#C00000", line="none")     # fill / border on creation
        logo = shapes.add_picture("logo.png", left=600, top=40,   # embedded, never linked
                                  alt_text="Acme logo")           # a drift-proof re-id handle
        deck.slides[4].shapes["Picture 3"].move(top=140)   # absolute, points
        star.set_fill(fill="#1F1F1F", line="#FFFFFF", line_width=2)  # fill ≠ font color
        star.reorder("front")                              # z-order: front/back/forward/backward
        star.delete()

    # Pictures — alt text doubles as a re-identification handle; export one shape for vision.
    logo.set_alt_text("Acme logo (top-right)")           # survives z-order drift
    chart_png = deck.slides[4].shapes["Chart 2"].export_image()   # just that shape, native size

    # Charts — a chart is a shape; its data lives in an embedded Excel workbook.
    with deck.edit("Add a revenue chart"):
        chart = deck.slides[4].shapes.add_chart(
            "column", ["Q1", "Q2", "Q3"], {"Revenue": [10, 20, 30], "Profit": [3, 6, 9]}
        ).chart
        chart.set_type("line")                           # change the kind
        chart.recolor_text("#FFFFFF")                    # all chart text white (dark-theme fix)
    data = chart.read()                                  # {chart_type, categories, series:[...]}

    # SmartArt — a diagram is a shape too; its content is a node tree.
    with deck.edit("Add a process diagram"):
        sa = deck.slides[3].shapes.add_smartart(
            "process", ["Discover", "Design", "Build", "Ship"]   # flat list…
        ).smartart
        sa.set_nodes([{"text": "CEO", "children": ["VP Eng", "VP Sales"]}])  # …or a tree
        sa.recolor_text("#FFFFFF")                       # every node label white
    tree = sa.read()                                     # {layout, nodes:[{text, level, children}]}

    # Text structure — paragraphs, formatting, bullets. (Per-anchor formatting;
    # for deck-wide styling use deck.theme / deck.master below.)
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

    # Deck-wide styling — theme (palette + fonts) and master (text styles +
    # background) restyle every inheriting slide at once. Global + anti-polite,
    # but still one Ctrl-Z; your view doesn't move.
    with deck.edit("Rebrand the deck"):
        deck.theme.set_color("accent1", "#C00000")       # recolor the whole deck
        deck.theme.set_font("major", "Georgia")          # major = headings, minor = body
        deck.master.format_text_style("body", 1, font="Georgia", size=28)
        deck.master.set_background("#1F1F1F")            # solid fill
    palette = deck.theme.read()                          # {colors:{slot:#RRGGBB}, fonts:{major, minor}}

    # Find / replace — fuzzy traversal across shapes, table cells, and notes
    # (no deck-wide character stream, so this is a walk, not a range).
    hits = deck.find("Q3 Reuslts")                       # [{anchor_id, start, length, text, ...}]
    with deck.edit("Fix the typo everywhere"):
        deck.find_replace("Q3 Reuslts", "Q3 Results", all=True)

    # Review comments — slide-anchored at an (x, y) point, and threaded.
    review = deck.slides[2].comments                     # per-slide CommentCollection
    with deck.edit("Leave review notes"):
        c = review.add("Tighten this headline", left=100, top=80)
        c.reply("Agreed — will do")
    roll = deck.comments()                               # deck-wide roll-up {total, slides:[...]}

    # Whole-deck snapshot — one low-res PNG per slide so a vision model can SEE
    # the whole deck cheaply (max_dim caps each slide's long edge). A read — polite.
    snaps = deck.snapshot(max_dim=1000)                  # [Snapshot(slide, image, path), ...]

    # Output tier — explicit, never implicit (pptlive never auto-saves).
    deck.save()                                          # persist to the existing file
    deck.save_as("v2.pptx", overwrite=True)              # write + rebind the working file
    deck.export_pdf("deck.pdf")                          # a read: no rebind, dirty flag kept

    # Media + narrated-video export — build a deck, narrate it, export an MP4.
    with deck.edit("Narrate the deck"):
        deck.slides[1].add_audio("intro.mp3")            # embed; autoplay + pace the slide
        deck.slides[2].add_video("demo.mp4")             # a video clip (stays visible)
    result = deck.export_video("deck.mp4", resolution=1080)   # async CreateVideo; blocks to done
    assert result.ok and result.status == "done"             # result.path is the MP4

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
| `shapeid:S:ID` | shape with stable `Shape.Id` ID on slide S — the **delete-proof** handle (survives a delete/restack that shifts `shape:S:N`) |
| `ph:S:KIND`    | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — the LLM-preferred form |
| `para:S:N:P`   | paragraph P (1-based) of shape N on slide S |
| `cell:S:N:R:C` | cell (row R, col C) of the table in shape N on slide S — a `Cell` *is* an anchor, so it takes every text/format verb |
| `notes:S`      | speaker-notes body of slide S |
| `comments:S`   | the review comments on slide S — a read selector (a container, addressed for reply/delete by `(slide, 1-based index)`) |
| `here:`        | whatever the user has selected right now — the shape, or the paragraph holding the text caret (the opt-in way to act on the live selection) |

z-order **drifts** when shapes are added or removed (or restacked via
`Shape.reorder`), so `shape:S:N` is resolved live and never cached; every shape
listing also emits `name` (`Shape.Name`) and `id` (`Shape.Id`, stable across
reorder). The drift-proof forms are `ph:S:KIND`, `.Name`, and `shapeid:S:ID`.
`para:S:N:P` and `cell:S:N:R:C` also resolve live (the paragraph/row count shifts
as text or rows are inserted/deleted).

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
pptlive find    --text "Q3 Reuslts" [--in slide:2]          # fuzzy locate across shapes/cells/notes
pptlive replace --find "Q3 Reuslts" --text "Q3 Results" --all  # fuzzy replace (or --occurrence N)

pptlive slide layouts                            # the layout names add/set-layout accept
pptlive slide add --layout two_content [--index 4]
pptlive slide duplicate --slide 7
pptlive slide move --slide 9 --to 2
pptlive slide set-layout --slide 4 --layout title_and_content
pptlive slide delete --slide 5
pptlive slide export --slide 2 --out slide2.png [--width 1280] [--format png]  # render to image
pptlive slide geometry 2                         # slide size + shape boxes + overlaps + off-slide (no render)

pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72
pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120
pptlive shape add --slide 4 --kind picture --path logo.png --left 600 --top 40 --alt-text "Acme logo"
pptlive shape add --slide 4 --kind table --rows 3 --cols 2 --left 72 --top 120
pptlive shape move   --anchor-id shape:4:3 --left 100 --top 140
pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200
pptlive shape delete --anchor-id shape:4:3
pptlive shape set-alt --anchor-id shape:4:3 --alt-text "Acme logo"      # drift-proof re-id handle
pptlive shape fill    --anchor-id shape:4:3 --fill "#C00000" --line none --line-width 2  # fill/border
pptlive shape order   --anchor-id shape:4:3 --to front       # z-order: front/back/forward/backward
pptlive shape export  --anchor-id shape:4:3 --out logo.png   # render one shape (native size)
pptlive shape animate --anchor-id shape:4:3 --effect fly_in [--trigger after_previous] [--exit]
pptlive slide animations 4                       # a slide's shape animations in play order

pptlive shape add --slide 4 --kind chart --chart-type column \
    --categories "Q1,Q2,Q3" --series '{"Revenue":[10,20,30]}'
pptlive chart read     --slide 4 --shape 5                    # {chart_type, categories, series}
pptlive chart set-type --slide 4 --shape 5 --chart-type line
pptlive chart set-data --slide 4 --shape 5 --categories "A,B" --series '{"S":[1,2]}'
pptlive chart recolor-text --slide 4 --shape 5 --color "#FFFFFF"   # all chart text (dark-theme fix)

pptlive shape add --slide 3 --kind smartart --smartart-kind process \
    --nodes '["Discover","Design","Build","Ship"]'           # flat list or {text,children} tree
pptlive smartart read      --slide 3 --shape 2               # {layout, nodes:[{text, level, children}]}
pptlive smartart set-nodes --slide 3 --shape 2 --nodes '[{"text":"CEO","children":["Eng","Sales"]}]'
pptlive smartart recolor-text --slide 3 --shape 2 --color "#FFFFFF"   # every node label

pptlive theme  read                              # deck palette (12 slots) + heading/body fonts
pptlive theme  set-color --slot accent1 --color "#C00000"    # recolors the whole deck
pptlive theme  set-font  --which major --name "Georgia"      # major = headings, minor = body
pptlive master read                              # master text styles (title/body/default) + background
pptlive master format-text-style --style body --level 1 --font "Georgia" --size 28
pptlive master set-background --color "#1F1F1F"  # deck-wide; solid fill

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

pptlive comment list [--slide 2]                 # comments on a slide, or the deck-wide roll-up
pptlive comment add   --slide 2 --text "Tighten this" [--left 100 --top 80]
pptlive comment reply  --slide 2 --index 1 --text "Agreed"
pptlive comment delete --slide 2 --index 1       # takes its replies too

pptlive section list                             # named slide spans (deck structure)
pptlive section add --name "Appendix" --before-slide 9   # rename/move/delete too
pptlive slide  set-footer --slide 2 --text "Confidential" # per-slide footer/slide-number/date
pptlive master set-footer --text "Confidential"           # deck-wide default

pptlive snapshot [--slides 2-4] [--max-dim 1000] [--width 1280 --height 720]  # one PNG per slide — the whole-deck vision read
pptlive save                                     # persist to the existing file (explicit)
pptlive save-as v2.pptx [--overwrite]            # write + rebind the working file
pptlive export-pdf deck.pdf                      # a read: PDF without rebinding the working file

pptlive media add --slide 1 --kind audio --path intro.mp3   # narrate (autoplay + pace the slide)
pptlive media add --slide 2 --kind video --path demo.mp4    # insert a video clip
pptlive export-video deck.mp4 --resolution 1080  # deck → MP4 (async CreateVideo; blocks until done)
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

**One-click install (Claude Desktop).** Download `pptlive.mcpb` from the
[latest release](https://github.com/thomas-villani/pptlive/releases/latest) and
drag it onto **Settings → Extensions**. The bundle pulls in pptlive via `uv` on
first run — no separate Python install. (Windows only; see [`mcpb/`](mcpb/).)

**Or let pptlive write the config for you:**

```
pptlive install-mcp                       # → Claude Desktop's claude_desktop_config.json
pptlive install-mcp --client claude-code  # → ./.mcp.json (project-local)
pptlive install-mcp --print               # just print the snippet for any client
```

It registers `uvx --from "pptlive[mcp]" pptlive-mcp` (resolves the published
package — no PATH assumptions). For a local checkout, add `--directory .` to get
`uv run --directory <dir> pptlive-mcp` instead. Restart the client afterward.

To wire it by hand instead, add to `claude_desktop_config.json` (Settings →
Developer → Edit Config) and restart:

```json
{
  "mcpServers": {
    "pptlive": { "command": "uvx", "args": ["--from", "pptlive[mcp]", "pptlive-mcp"] }
  }
}
```

It's a compact **five-tool dispatch surface** — each tool takes an `op` argument
and routes to the right verb, so the agent's tool picker sees five definitions
instead of fifteen. They wrap the same API, so the politeness model and
one-Ctrl-Z `edit` fencing carry over and reads never move the view:

| tool | `op`s |
| ---- | ----- |
| `ppt_read` | `status` · `slides` · `outline` · `slide` · `anchor` · `geometry` (slide size + shape boxes + overlaps + off-slide) · `selection` · `find` · `table` · `chart` · `smartart` · `comments` · `animations` · `sections` · `headers_footers` · `theme` · `master` · `layouts` — every read; never moves the view |
| `ppt_edit` | `write` · `find_replace` · `format` (font + paragraph + shape fill/line + bullets) · `slide_add`/`slide_delete`/`slide_duplicate`/`slide_move`/`set_layout` · `shape_add`/`shape_move`/`shape_resize`/`shape_delete`/`shape_order`/`set_alt` · `media_add` (audio/video narration) · `shape_animate`/`shape_clear_animations`/`slide_clear_animations` · `table_add_row`/`table_delete_row` · `chart_set_type`/`chart_set_data`/`chart_recolor_text` · `smartart_set_nodes`/`smartart_recolor_text` · `comment_add`/`comment_reply`/`comment_delete` · `section_add`/`section_rename`/`section_delete`/`section_move` · `set_headers_footers` · `theme_set_color`/`theme_set_font` · `master_format_text_style`/`master_format_paragraph_style`/`master_set_background` — every mutation; one Ctrl-Z each |
| `ppt_render` | `slide_image` · `shape_image` · `deck_snapshot` (one PNG per slide — the whole-deck vision read; `max_dim` or exact `width`/`height`) · `deck_pdf`/`save`/`save_as` (explicit output) · `export_video`/`video_status` (deck → MP4; async, blocks until done by default) · `navigate` (the one deliberate view move) |
| `ppt_show` | live slide show: `state` · `start` · `end` · `next` · `previous` · `goto` · `black` · `white` · `resume` |
| `ppt_batch` | run a **list** of the ops above against one connection — all `edit`s fenced into a **single** undo entry (`atomic`), with `stop_on_error` control |

Tables, charts, and SmartArt are addressed by their shape's `anchor_id` (a
`shape:S:N` or the delete-proof `shapeid:S:ID`); cells stay `cell:S:N:R:C` anchors
you write to with `ppt_edit op="write"`. The `theme_*`/`master_*` ops are
deck-wide (no anchor).

Tool failures surface as MCP errors carrying a category token — `not_found`,
`ambiguous`, `busy`, `not_running`, `no_text_frame`, `invalid_args` — the string
analog of the CLI's exit codes, so an agent can branch on them. Inside
`ppt_batch` the same tokens are reported per-command instead of aborting.

## For LLM agents — self-bootstrapping

An agent with a shell can orient itself in one command:

```
pptlive llm-help            # the full CLI guide (anchors, every verb, exit codes)
pptlive llm-help --python   # the Python-API guide instead
```

Output is raw Markdown (like `--help`), unaffected by `--json/--text`, so it
drops straight into a model's context. `pptlive --help` points here too.

pptlive ships **two agent skills** — `pptlive-cli` and `pptlive-python` — for
tools that load `SKILL.md` files:

```
pptlive install-skill            # writes both to ./.agents/skills/<name>/SKILL.md
pptlive install-skill --cli      # just one (also --python)
pptlive install-skill --system   # into ~/.agents/skills/ instead
```

For MCP clients, `pptlive install-mcp` (above) registers the server, and the MCP
server also exposes the same guides as `pptlive://guide` /
`pptlive://guide/python` resources.

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

The library targets Python 3.11+ (dev pins 3.13). See `spec.md` for the design
and `IMPLEMENTATION.md` for staged build progress. Windows + COM only.

Full documentation: <https://thomas-villani.github.io/pptlive/>

## A Review From the Other End of the Wire

*by Claude (Opus 4.8), after a live session driving the bundle against an open PowerPoint deck*

I spent a session using pptlive to read, redesign, and stress-test a real deck open on the author's machine — reading its structure, rebuilding a slide from scratch, editing a live chart, and running it as a presentation. Everything I tried, worked. Here's what stands out.

**The read side is genuinely good.** `status` → `outline` → `slide` → `chart` gives a clean descent from "what's open" down to individual shape anchors and live chart data, all without moving the user's view. I could understand a seven-slide deck — its narrative, its palette, its two embedded charts' underlying numbers — before touching anything. The side-effect-free reads are the foundation everything else stands on.

**The anchor model is the right abstraction.** Addressing `ph:7:title`, `shape:7:3`, `para:7:5:2`, `cell:S:N:R:C` as stable handles means edits target exactly what you mean. When I built a four-card layout, I could format the big number and the description line independently because each resolved to its own paragraph anchor.

**`ppt_batch` with atomic undo-grouping is the feature I'd miss most.** A multi-shape redesign collapses into a single Ctrl-Z for the user. And `stop_on_error: false` turned a frustrating debug loop into a single legible report — I could see all nine formatting ops land at once instead of playing failure whack-a-mole.

**The render loop is what changes what's possible.** Once `slide_image` returns a real image content block, I can *see* my own output and iterate. Two moments earned it: I formatted card text white expecting white-on-grey and braced to hunt for a fill op — the render showed the shapes had defaulted to the brand coral and looked correct, so I didn't "fix" a non-problem. Then a live `chart_set_data` edit re-rendered with PowerPoint's *own* engine auto-rescaling the y-axis from 210 to 600 — proof the edit reached the live document, not some shadow model. `shape_image` rounds it out, cropping a single shape to its bounds so I can inspect one card in isolation. Without sight, I'm editing with my eyes closed and narrating confidently. With it, I can be wrong and *catch* it. That gap is the whole game.

**The presentation mode is a real clicker.** `start`, `next`, `previous`, `goto`, `black`/`resume`, `end` — I ran the full sequence and read state back at every step. `goto` jumps anywhere, `black` blanks the screen while remembering the slide underneath, `resume` returns to it, and every op reports running-state and position consistently. It drives the actual fullscreen show, not a simulation of one.

**One sharp lesson worth recording:** the bundle drives a *live* application, which means the model is never the only actor. Earlier in the session I hit a slideshow-navigation error, built a tidy two-part bug theory around it, and was one call from "confirming" it — when the real cause was the user ending the show by hand. We reran the sequence cleanly afterward and every op passed. Anyone building agents against live software should design for that gap between the agent's model and the screen: re-read state, don't trust your own narration, and prefer graceful bounded failures to clever ones.

**Verdict:** pptlive treats PowerPoint as a live, inspectable, scriptable surface instead of a file to overwrite — read without disturbing, edit atomically, *look* and iterate, then present. Every layer I exercised — read, batch edit, render-and-see, single-shape inspection, live chart data, and the presentation clicker — did exactly what it claimed. It's the difference between firing commands into the dark and actually working. I'd reach for it again.

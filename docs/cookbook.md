# Cookbook

End-to-end recipes. Each shows the Python API and, where it helps, the
equivalent CLI. Every mutating recipe is wrapped in `deck.edit(...)`, so it's
polite (the user's viewed slide and selection are preserved) and reverts with a
single Ctrl-Z.

All recipes assume PowerPoint is running with a deck open:

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active
    ...
```

## 1. Build a results slide from scratch

Add a slide, fill its placeholders, and lay out a supporting shape — all under
one undo entry.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Build the Q3 results slide"):
        slide = deck.slides.add(layout="title_and_content", index=4)
        slide.placeholder("title").set_text("Q3 Results")
        slide.placeholder("body").set_text(
            "Revenue up 12%\nChurn down 3%\nNPS +9"
        )
        slide.placeholder("body").apply_list("bulleted")
        slide.notes.set_text("Lead with the revenue number.")
```

CLI:

```bash
pptlive slide layouts                                   # see the names first
pptlive slide add --layout title_and_content --index 4  # -> {"index": 4, ...}
pptlive write --anchor-id ph:4:title --text "Q3 Results"
pptlive write --anchor-id ph:4:body  --text "Revenue up 12%\nChurn down 3%\nNPS +9"
pptlive list apply --anchor-id ph:4:body --type bulleted
pptlive write --anchor-id notes:4 --text "Lead with the revenue number."
```

## 2. Read the deck structure

Side-effect-free reads to orient an agent before it edits.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    deck.slides.list()        # [{index, id, layout, title, shape_count, has_notes}]
    deck.outline()            # [{slide, title, bullets:[...]}]
    deck.page_setup()         # {width, height} in points — the canvas to place on
    deck.slides[2].read()     # every shape: anchor_id, name, id, type, geometry, text
```

The `read()` of a slide is the workhorse: it gives you `name`, the stable
`id`, and (for placeholders) the semantic kind, so you can pick the
drift-proof anchor — `ph:2:title` or `slide.shapes["Title 1"]` — rather than a
volatile `shape:2:N`. See [Concepts → z-order drifts](concepts.md#z-order-drifts-design-around-it).

## 3. Edit and format a placeholder politely

Set text, then apply direct font formatting (PowerPoint has no named paragraph
styles — `format_text` is the analog of "apply style").

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Polish the body copy"):
        body = deck.anchor_by_id("ph:4:body")
        body.set_text("Revenue up 12%\nChurn down 3%\nNPS +9")
        body.apply_list("bulleted")
        body.paragraph(1).format_text(bold=True, size=24, color="#2E74B5")
        body.paragraph(2).format_paragraph(indent_level=2, alignment="left")
        body.insert_paragraph_after("Cash runway: 30 months")
```

The verbs live on the base `Anchor`, so they work the same on a whole-shape
anchor (`ph:4:body`) and on a single `Paragraph` (`body.paragraph(1)`).

## 4. Read and edit a table

A table is a **shape** (`Shape.has_table`); cells are `cell:S:N:R:C` anchors,
and a `Cell` *is* an `Anchor`, so it takes every text/format verb.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Add a metrics table"):
        shape = deck.slides[4].shapes.add_table(rows=3, columns=2)
        table = shape.table
        table.cell(1, 1).set_text("Metric")
        table.cell(1, 2).set_text("Q3")
        table.add_row(["Revenue", "$4.2M"])           # appends + fills a row
        deck.anchor_by_id(f"cell:4:{shape.shape_id}:1:1")  # ...or address it directly
        table.cell(1, 1).format_text(bold=True)

    grid = table.read()        # {slide, shape, rows, columns, cells:[...]}
```

CLI (the table's shape index comes from `slide read` / `shapes`):

```bash
pptlive shape add --slide 4 --kind table --rows 3 --cols 2
pptlive write --anchor-id cell:4:5:1:1 --text "Metric"
pptlive write --anchor-id cell:4:5:1:2 --text "Q3"
pptlive table add-row --slide 4 --shape 5 --values '["Revenue", "$4.2M"]'
pptlive format-text --anchor-id cell:4:5:1:1 --bold
pptlive table read --slide 4 --shape 5
```

## 5. Build, look, iterate (the vision loop)

PowerPoint renders the **live, unsaved** state, so you can build a slide, render
it to PNG, hand it to a vision model, and revise — without ever saving or
re-opening.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Lay out the title slide"):
        deck.anchor_by_id("ph:1:title").set_text("Acme — Q3 Review")

    png = deck.slides[1].export_image(width=1280)   # temp PNG (or pass a path)
    # ...hand `png` to your image tool, look, then come back and adjust.

    one_shape = deck.slides[1].shapes["Title 1"].export_image()  # just that shape
```

`export_image` is polite — it doesn't move the user's view. A slide export
honours `width`/`height` (the other follows the aspect ratio); a *shape* export
is native-size only.

CLI — export then `Read` the file in one step:

```bash
pptlive slide export --slide 1 --width 1280     # prints the temp path
pptlive shape export --anchor-id shape:1:1      # one shape, native size
```

## 6. Act on whatever the user is pointing at

`deck.selection()` reads the user's current selection (resolved to anchors)
without disturbing it. To *act* on it, target the opt-in `here:` anchor — the
one place the politeness model lets you touch the live selection.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    sel = deck.selection()       # SelectionInfo(type, slide, anchor_id, ...)
    if sel.anchor_id:
        print(f"User is on slide {sel.slide}, selection -> {sel.anchor_id}")
        with deck.edit("Bold the selected text"):
            deck.anchor_by_id("here:").format_text(bold=True)
    else:
        print("Nothing selected.")
```

CLI:

```bash
pptlive selection                              # {"type": "text", "anchor_id": "para:2:3:2", ...}
pptlive format-text --anchor-id here: --bold   # act on it
```

## 7. Add and edit a chart

A chart is a shape; its data lives in an embedded Excel workbook. Drive it
through the `Chart` wrapper.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Add a revenue chart"):
        chart = deck.slides[4].shapes.add_chart(
            "column",
            ["Q1", "Q2", "Q3"],
            {"Revenue": [10, 20, 30], "Profit": [3, 6, 9]},
            left=72, top=120, width=480, height=300,
        ).chart
        chart.set_type("line")                 # change the kind

    data = chart.read()        # {chart_type, categories, series:[{name, values}]}
```

To overwrite the data later (the shape index comes from `slide read`):

```bash
pptlive chart set-data --slide 4 --shape 6 \
    --categories "A,B,C" --series '{"S1":[1,2,3],"S2":[4,5,6]}'
```

## 8. Tag a picture so you can re-find it after drift

z-order shifts every time a shape is added, removed, or restacked, so don't lean
on `shape:S:N`. The drift-proof handle is **`shapeid:S:ID`** — the stable
`Shape.Id` printed as `id` in every shape listing, which survives a delete/restack
that would shift the z-order index. Alt text and `.Name` are the other re-find
handles.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Drop in the logo"):
        logo = deck.slides[4].shapes.add_picture(
            "logo.png", left=600, top=40, alt_text="Acme logo (top-right)"
        )
    logo_id = logo.to_dict()["id"]                 # stable Shape.Id

    # Later — after other shapes were added or restacked — re-find it three ways,
    # none of which depend on the live z-order:
    deck.anchor_by_id(f"shapeid:4:{logo_id}").move(top=60)   # the delete-proof handle
    deck.slides[4].shapes["Picture 3"].move(top=60)          # by name
```

## 9. Add and edit a SmartArt diagram

A SmartArt diagram is a shape; its content is a **node tree**. Flat layouts
(`process`, `cycle`, `list`, `pyramid`, `venn`) take any number of top-level
nodes; tree layouts (`hierarchy`, `orgchart`) take a single root with nested
children.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Add a process diagram"):
        sa = deck.slides[3].shapes.add_smartart(
            "process", ["Discover", "Design", "Build", "Ship"],
            left=72, top=160, width=720, height=200,
        ).smartart

    # Reshape it later — strings are leaves, {text, children} nests:
    with deck.edit("Turn it into an org chart"):
        deck.slides[3].shapes[2].smartart.set_nodes(
            [{"text": "CEO", "children": ["VP Eng", "VP Sales", "VP Ops"]}]
        )

    tree = sa.read()      # {layout, layout_id, node_count, nodes:[{text, level, children}]}
```

From the CLI, `--nodes` is the same JSON shape:

```bash
pptlive smartart set-nodes --slide 3 --shape 2 \
    --nodes '["Plan", {"text": "Execute", "children": ["Build", "Test"]}, "Ship"]'
```

## 10. Restyle the whole deck (theme + master)

`format_text` styles one anchor; `deck.theme` and `deck.master` restyle **every
inheriting slide** at once — the palette, the heading/body fonts, the master
text styles, and the background. They're deliberately global and anti-polite,
but each `edit()` block is still one Ctrl-Z and your view doesn't move.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Rebrand the deck"):
        # Theme: the 12-slot palette + heading/body typefaces.
        deck.theme.set_color("accent1", "#C00000")
        deck.theme.set_color("dark1",   "#1F1F1F")
        deck.theme.set_font("major", "Georgia")     # headings
        deck.theme.set_font("minor", "Calibri")     # body

        # Master: text styles (title/body/default × 5 levels) + background.
        deck.master.format_text_style("title", 1, bold=True, size=40)
        deck.master.format_paragraph_style("body", 1, alignment="left", space_after=12)
        deck.master.set_background("#FBFBFB")

    palette = deck.theme.read()   # {colors:{slot:#RRGGBB}, fonts:{major, minor}}
    styles  = deck.master.read()  # {text_styles:{...}, background:{type, color}}
```

```bash
pptlive theme  set-color --slot accent1 --color "#C00000"
pptlive master format-text-style --style title --level 1 --bold --size 40
pptlive master set-background --color "#FBFBFB"
```

## 11. LLM tool-use loop

The CLI is built for this: discover anchors, let the model choose, apply, and
branch on the exit code.

### Tool schema

Expose `pptlive` as a single shell tool whose `args` are the CLI argv:

```json
{
  "name": "pptlive",
  "description": "Drive the live PowerPoint deck. One JSON object on stdout; exit codes signal failure.",
  "input_schema": {
    "type": "object",
    "properties": {"args": {"type": "array", "items": {"type": "string"}}},
    "required": ["args"]
  }
}
```

### Driver loop (sketch)

```python
import subprocess, json

def pptlive(*args):
    p = subprocess.run(["pptlive", *args], capture_output=True, text=True)
    out = json.loads(p.stdout) if p.stdout.strip() else None
    return p.returncode, out

# 1. Discover what's addressable.
_, slides = pptlive("slides")
_, grid   = pptlive("slide", "read", "2")

# 2. The model picks an anchor + new value, returns e.g.:
#    {"anchor_id": "ph:2:title", "text": "Revised Agenda"}

# 3. Apply.
code, result = pptlive("write", "--anchor-id", "ph:2:title", "--text", "Revised Agenda")

# 4. Branch on the exit code.
if code == 2:      # anchor not found — re-read and let the model retry
    ...
elif code == 3:    # PowerPoint busy — back off and retry
    ...
elif code == 6:    # shape has no text frame — pick a text-bearing anchor
    ...
```

For Claude Desktop and other MCP hosts, prefer the [MCP server](mcp.md) — same
control, no shelling out, and it can return rendered slide images as native
image content.

## 12. Presenter-assistant: drive a live slide show

The `show` group deliberately controls what's on screen — a clicker for an
agent. Unlike edits, it's *not* polite (that's the point).

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    deck.show.start()                 # run from the top
    deck.show.goto(5)                 # jump to slide 5
    deck.show.next()                  # advance a build/slide
    deck.show.black()                 # blank to black (then .resume())
    print(deck.show.state())          # {running, state, current_slide, ...}
    deck.show.end()
```

`show.state()` is the safe poll — it never raises and reports `running: false`
when no show is up. The other verbs raise
[`SlideShowNotRunningError`](errors.md) (exit `1` on the CLI) if you call them
with no show running. Editing the deck *during* a running show works (a text
edit mid-show is not rejected).

```bash
pptlive show start --from 1
pptlive show next
pptlive show state
pptlive show end
```

## 13. Work across multiple open decks

`--doc NAME` (CLI) or `ppt.presentations[name]` (Python) targets a specific open
presentation instead of the active one — so you never disturb which deck the
user is focused on.

```python
with pl.attach() as ppt:
    pitch  = ppt.presentations["Pitch.pptx"]
    review = ppt.presentations["Q3 Review.pptx"]

    title = pitch.anchor_by_id("ph:1:title").text
    with review.edit("Copy the pitch title across"):
        review.anchor_by_id("ph:1:title").set_text(title)
```

```bash
pptlive --doc "Q3 Review.pptx" slides
pptlive --doc "Q3 Review.pptx" write --anchor-id ph:1:title --text "Q3 Review"
```

## 14. Find and replace across the deck

There's no deck-wide character stream, so `find` traverses every text frame —
shapes, table cells, and speaker notes — and reports each hit against a
resolvable anchor. Matching is smart-quote / dash / whitespace tolerant, so
text re-typed off a slide still matches. `find_replace` rewrites just the
matched span (run formatting survives) and belongs in an `edit()` block.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    hits = deck.find("Acme")                     # [{anchor_id, start, length, text, context}]
    print(f"{len(hits)} mentions of Acme")

    with deck.edit("Rebrand Acme → Globex"):
        deck.find_replace("Acme", "Globex", all=True)        # every occurrence
        deck.find_replace("teh", "the", occurrence=2)        # only the 2nd hit, deck-wide
        deck.find_replace("Q3 plan", "Q3 forecast", scope="slide:2")   # scoped
```

One match auto-applies; several without `all` / `occurrence` raise
`AmbiguousMatchError` (the candidates ride along); zero matches raise
`AnchorNotFoundError`. `find` itself never raises — a miss is an empty list.

```bash
pptlive find --text "Acme"
pptlive replace --find "Acme" --text "Globex" --all
pptlive replace --find "Q3 plan" --text "Q3 forecast" --in slide:2
```

## 15. Style a shape's fill, border, and z-order

`set_fill` controls a shape's **fill** and **border** — a color, or `"none"` for
transparent / no border. That's distinct from `format_text`'s `color`, which is
the *font* color. `reorder` restacks the shape (`front` / `back` / `forward` /
`backward`); since z-order is what `shape:S:N` indexes, reach for a drift-proof
handle afterwards. `fill` / `line` / `line_width` also ride on creation.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Style the callout"):
        box = deck.slides[2].shapes.add_shape(
            "rounded_rectangle", left=80, top=300, width=300, height=90,
            fill="#C00000", line="none",          # red fill, no border
        )
        box.set_text("Key takeaway")
        box.set_fill(line="#FFFFFF", line_width=2)   # add a white 2pt border
        box.reorder("front")                          # bring it above overlapping shapes
```

```bash
pptlive shape fill  --anchor-id shape:2:6 --fill "#C00000" --line none
pptlive shape order --anchor-id shape:2:6 --to front
```

## 16. Recolor composite text for a dark background

A SmartArt diagram or chart has no text anchor, so `format_text` can't reach its
internal labels. `recolor_text` is the coarse "make every label this color" fix
for a dark theme — it walks each node / every shown chart text element (legend,
both axis tick labels, title, data labels).

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    with deck.edit("Make diagrams legible on the dark master"):
        deck.slides[3].shapes["Diagram 2"].smartart.recolor_text("#FFFFFF")
        deck.slides[4].shapes["Chart 2"].chart.recolor_text("#FFFFFF")
```

## 17. Run a review loop with comments

Comments attach to a **slide** at an `(x, y)` point (not a text range) and are
**threaded**. Adding binds to the signed-in Office account, so a passed
`author` / `initials` is best-effort. There is no resolve/reopen verb (the status
isn't COM-readable on current builds) — delete a thread when it's addressed.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    for slide in deck.slides:
        for c in slide.comments:                 # iterating yields Comment objects
            print(slide.index, c.author, c.text)   # (.list() returns the same as dicts)

    with deck.edit("Leave and answer review notes"):
        note = deck.slides[2].comments.add("Tighten this headline", left=100, top=80)
        note.reply("Done — shortened to five words")
        deck.slides[2].comments[1].delete()      # resolve by deleting (takes its replies)

    roll = deck.comments()                        # deck-wide {total, slides:[...]}
```

```bash
pptlive comment list --slide 2
pptlive comment add  --slide 2 --text "Tighten this headline" --left 100 --top 80
pptlive comment reply  --slide 2 --index 1 --text "Done"
pptlive comment delete --slide 2 --index 1
```

## 18. See the whole deck at once (token-aware vision read)

`deck.snapshot()` renders one low-resolution PNG per slide so a vision model can
*see* every slide cheaply — the "did my styling land across the deck?" read.
`max_dim` caps each slide's long edge in pixels; since every slide shares one
geometry, that's a uniform, predictable per-slide token budget (~1000 px stays
legible). It's a **read** — no `edit()` fence, and the view doesn't move.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    snaps = deck.snapshot(max_dim=1000)           # [Snapshot(slide, png, path), ...]
    for s in snaps:
        ...                                        # hand s.png to your image tool, look

    deck.snapshot("review.png", slides=(2, 4))    # or write to disk: review-s2/3/4.png
```

```bash
pptlive snapshot --slides 2-4 --max-dim 1000 --out review.png
```

## 19. Save, save-as, and export a PDF

pptlive **never auto-saves** — output is always an explicit verb. `save` persists
to the existing file (and raises `UnsavedPresentationError` if the deck has never
been saved — use `save_as` first). `save_as` writes *and rebinds* the working
file (the open deck becomes the new file). `export_pdf` is a **read**: it writes a
PDF without rebinding the working file or clearing its dirty flag.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    if not deck.saved:                            # the dirty flag, also on `status`
        deck.save()                               # -> the existing file
    deck.save_as("Q3 Review v2.pptx", overwrite=True)   # write + rebind
    deck.export_pdf("Q3 Review.pdf")              # deliverable; deck stays bound to the .pptx
```

```bash
pptlive save
pptlive save-as "Q3 Review v2.pptx" --overwrite
pptlive export-pdf "Q3 Review.pdf"
```

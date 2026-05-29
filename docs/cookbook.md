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

z-order shifts every time a shape is added or removed, so don't lean on
`shape:S:N`. Set a picture's **alt text** on create and re-find it by name or
description.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Drop in the logo"):
        logo = deck.slides[4].shapes.add_picture(
            "logo.png", left=600, top=40, alt_text="Acme logo (top-right)"
        )

    # Later — after other shapes were added — re-find it by name, not z-order:
    deck.slides[4].shapes["Picture 3"].move(top=60)
```

## 9. LLM tool-use loop

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

## 10. Presenter-assistant: drive a live slide show

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

## 11. Work across multiple open decks

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

# CLI

The `pptlive` command is a thin, JSON-first wrapper over the
[Python API](python-api.md). Every invocation prints **one JSON object on
stdout** (logs and errors go to stderr) and exits with a deterministic code, so
it drops straight into an LLM tool-use loop. It's the same politeness and
one-Ctrl-Z atomic undo as the library.

```bash
pptlive status
pptlive --text slides              # human-readable instead of JSON
pptlive --doc Pitch.pptx outline   # target a specific open deck
```

## Global flags

Global flags go **before** the subcommand:

| Flag | Default | Meaning |
| ---- | ------- | ------- |
| `--json` / `--text` | `--json` | Output format. `--json` prints one object per call; `--text` prints a human-readable rendering. |
| `--doc NAME` | active deck | Target an open presentation by name instead of the active one. |
| `-h` / `--help` | — | Show help for the command or group. |

## Exit codes

| Exit | Meaning | Retry? |
| ---- | ------- | ------ |
| `0` | success | — |
| `1` | other / unclassified, or no slide show running | depends |
| `2` | anchor / slide / shape / layout / deck not found | after re-reading |
| `3` | PowerPoint busy / modal dialog | **yes**, back off |
| `4` | PowerPoint not running | only if user launches it |
| `5` | ambiguous match | after disambiguating |
| `6` | shape has no text frame | no — pick a text shape |

See [Errors & exit codes](errors.md) for the full mapping.

---

## `status`

List open presentations, which one is active, and the slide currently in view.

```bash
pptlive status
```

```json
{
  "decks": [
    {"name": "Pitch.pptx", "path": "C:\\Users\\you\\Pitch.pptx", "is_active": true}
  ],
  "viewed_slide": 2
}
```

Exits `4` (and emits `{"decks": [], "viewed_slide": null}`) if PowerPoint isn't
running.

## `slides`

One row per slide: index, id, layout, title, shape count, has-notes.

```bash
pptlive slides
```

```json
[
  {"index": 1, "id": 256, "layout": "Title Slide", "title": "Acme Q3", "shape_count": 2, "has_notes": false},
  {"index": 2, "id": 257, "layout": "Title and Content", "title": "Agenda", "shape_count": 2, "has_notes": true}
]
```

## `outline`

Title + body bullets per slide — the Outline-view analog.

```bash
pptlive outline
```

```json
[
  {"slide": 1, "title": "Acme Q3", "bullets": []},
  {"slide": 2, "title": "Agenda", "bullets": ["Intro", "Demo", "Q&A"]}
]
```

## `slide read S`

Every shape on slide `S`: `anchor_id`, name, id, type, placeholder kind,
geometry, alt text, and text.

```bash
pptlive slide read 2
```

```json
{
  "index": 2, "id": 257, "layout": "Title and Content", "title": "Agenda",
  "shapes": [
    {"anchor_id": "shape:2:1", "name": "Title 1", "id": 2, "type": "placeholder",
     "placeholder": "title", "geometry": {"left": 38, "top": 27, "width": 884, "height": 104},
     "alt_text": "", "text": "Agenda"},
    {"anchor_id": "shape:2:2", "name": "Content Placeholder 2", "id": 3, "type": "placeholder",
     "placeholder": "body", "geometry": {"left": 38, "top": 145, "width": 884, "height": 385},
     "alt_text": "", "text": "Intro\rDemo\rQ&A"}
  ]
}
```

## `shapes --slide S`

Just the shape listing for slide `S` (the `shapes` array from `slide read`).

```bash
pptlive shapes --slide 2
```

---

## Reading text — `read`

`read anchor` reads any text anchor; `read notes` is sugar for `notes:S`.

```bash
pptlive read anchor --anchor-id ph:2:title    # placeholder by semantic kind
pptlive read anchor --anchor-id shape:2:2      # shape by z-order
pptlive read anchor --anchor-id para:2:2:1     # one paragraph
pptlive read anchor --anchor-id cell:4:5:1:1   # one table cell
pptlive read anchor --anchor-id here:          # the user's current selection
pptlive read notes --slide 2                   # == --anchor-id notes:2
```

```json
{"anchor_id": "ph:2:title", "kind": "placeholder", "text": "Agenda"}
```

Exits `2` if the anchor doesn't resolve, `6` if it names a shape with no text
frame.

## `write --anchor-id ID --text "…"`

Set the entire text of a text anchor. Preserves the viewed slide; one Ctrl-Z.
Embed `\n` for paragraph breaks.

```bash
pptlive write --anchor-id ph:2:title --text "Agenda"
pptlive write --anchor-id ph:2:body  --text "Intro\nDemo\nQ&A"
pptlive write --anchor-id cell:4:5:1:1 --text "Metric"   # a cell is a text anchor
```

```json
{"ok": true, "anchor_id": "ph:2:title", "kind": "placeholder"}
```

## `replace --anchor-id ID --text "…"`

Replace a text anchor's contents. In the current release this is identical in
effect to `write` (the anchor-addressed form); fuzzy `replace --find OLD --text
NEW` arrives with the find/replace stage.

---

## Slides — the `slide` group

`slide` covers reads plus the lifecycle verbs. Each mutating verb is wrapped in
an `edit()` fence (one Ctrl-Z) and preserves the viewed slide.

### `slide layouts`

List the deck's layout names — the values `add` / `set-layout` accept.

```bash
pptlive slide layouts
```

```json
[{"index": 1, "name": "Title Slide"}, {"index": 2, "name": "Title and Content"}]
```

### `slide add`

Add a slide; defaults to appending and to the `title_and_content` layout.

```bash
pptlive slide add --layout two_content --index 4
```

```json
{"ok": true, "index": 4, "id": 261, "layout": "Two Content"}
```

Exits `2` ([`LayoutNotFoundError`](errors.md)) on an unknown layout name — the
error lists the available ones.

### `slide delete | duplicate | move | set-layout`

```bash
pptlive slide delete    --slide 5
pptlive slide duplicate --slide 7              # copy lands at slide 8
pptlive slide move      --slide 9 --to 2
pptlive slide set-layout --slide 4 --layout title_and_content
```

`duplicate` and `move` report the resulting slide's `index` and `id`; an
out-of-range slide is exit `2`.

### `slide export`

Render a slide to an image so a vision model can *see* it. Renders the current
(unsaved) state; polite. With no `--out`, writes a temp file and prints its
path — so you can export-then-read in one step. Pass one of `--width` /
`--height` and the other follows the aspect ratio.

```bash
pptlive slide export --slide 2 --out slide2.png --width 1280
pptlive slide export --slide 2                          # temp PNG
```

```json
{"ok": true, "slide": 2, "path": "C:\\Users\\you\\AppData\\Local\\Temp\\…png",
 "format": "png", "width": 1280, "height": null}
```

---

## Shapes — the `shape` group

Create and place shapes (geometry in **points**; 1 inch = 72 pt). Each verb is
one Ctrl-Z.

### `shape add`

```bash
# Text box
pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72

# Autoshape (see --shape-type choices)
pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120

# Picture (embedded, never linked) with a drift-proof alt-text handle
pptlive shape add --slide 4 --kind picture --path logo.png --left 600 --top 40 --alt-text "Acme logo"

# Table
pptlive shape add --slide 4 --kind table --rows 3 --cols 2 --left 72 --top 120

# Chart (data is an embedded Excel workbook)
pptlive shape add --slide 4 --kind chart --chart-type column \
    --categories "Q1,Q2,Q3" --series '{"Revenue":[10,20,30],"Profit":[3,6,9]}'
```

```json
{"ok": true, "anchor_id": "shape:4:3", "name": "TextBox 3", "id": 4,
 "type": "textbox", "geometry": {"left": 72, "top": 72, "width": 180, "height": 29},
 "alt_text": "", "text": "Revenue up 12%"}
```

Notes:

- `--kind picture` requires `--path`; `--kind table` requires `--rows`/`--cols`;
  `--kind chart` needs both `--categories` and `--series`.
- `--categories` takes a JSON array or a comma-separated list. `--series` takes
  a JSON object `{"name":[values]}` or an array of `[name,[values]]`.
- A new shape lands at the **top of the z-order** (last slot), so its
  `shape:S:N` is the post-add `Shapes.Count`. A text box created with text
  auto-fits its height, so a requested `--height` is advisory when AutoSize is
  on.

### `shape move | resize | delete`

Address by `shape:S:N` or `ph:S:KIND` (a non-shape anchor like `notes:S` is
exit `2`).

```bash
pptlive shape move   --anchor-id shape:4:3 --left 100 --top 140
pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200
pptlive shape delete --anchor-id shape:4:3
```

`move` and `resize` echo the new `geometry`; each needs at least one
coordinate/dimension.

### `shape set-alt`

Set a shape's alternative text — a description you can re-find the shape by
after z-order drift.

```bash
pptlive shape set-alt --anchor-id shape:4:3 --alt-text "Acme logo (top-right)"
```

### `shape export`

Render a single shape to an image, cropped to its bounds at native pixel size
(no size override — `Shape.Export` doesn't honour one reliably). Polite; temp
file when `--out` is omitted.

```bash
pptlive shape export --anchor-id shape:4:3 --out logo.png
```

---

## Text structure

### `paragraphs --anchor-id ID`

List a shape's paragraphs, each with its `para:S:N:P` anchor, text, indent
level, and bullet.

```bash
pptlive paragraphs --anchor-id ph:4:body
```

```json
[
  {"anchor_id": "para:4:2:1", "text": "Revenue up 12%", "indent_level": 1, "bullet": "bulleted"},
  {"anchor_id": "para:4:2:2", "text": "Churn down 3%",  "indent_level": 1, "bullet": "bulleted"}
]
```

### `insert --anchor-id ID --text "…" [--before | --after]`

Insert a new paragraph relative to any text anchor (`para:` / `ph:` / `shape:` /
`notes:`). Defaults to `--after`.

```bash
pptlive insert --anchor-id para:4:2:3 --text "Cash runway: 30 months" --after
```

### `format-paragraph --anchor-id ID [...]`

Set alignment, spacing, and indent level on a text anchor. At least one option
is required.

```bash
pptlive format-paragraph --anchor-id para:4:2:1 \
    --alignment center --space-before 6 --space-after 6 \
    --line-spacing 1.5 --indent-level 2
```

`--alignment` ∈ `left` / `center` / `right` / `justify`; `--indent-level` is
1–5 (PowerPoint's only paragraph-indent notion).

### `format-text --anchor-id ID [...]`

PowerPoint's analog of "apply style" — it has no named paragraph styles, so
styling is direct font formatting. At least one option is required.

```bash
pptlive format-text --anchor-id ph:4:title --bold --size 40 --color "#2E74B5"
pptlive format-text --anchor-id para:4:2:2 --no-bold --italic --font "Calibri"
```

Toggles: `--bold/--no-bold`, `--italic/--no-italic`, `--underline/--no-underline`.
`--color` is `#RRGGBB`.

### `list apply | remove`

Turn a text anchor's paragraphs into a bulleted/numbered list, or strip the
list formatting.

```bash
pptlive list apply  --anchor-id ph:4:body --type bulleted --char "•"
pptlive list apply  --anchor-id ph:4:body --type numbered
pptlive list remove --anchor-id ph:4:body
```

---

## Tables — the `table` group

A table is a shape; address it by slide + z-order (`--slide S --shape N`).
Cells are `cell:S:N:R:C` anchors you write to with `write`.

### `table read`

```bash
pptlive table read --slide 4 --shape 5
```

```json
{
  "slide": 4, "shape": 5, "anchor_id": "shape:4:5", "rows": 3, "columns": 2,
  "cells": [
    [{"anchor_id": "cell:4:5:1:1", "text": "Metric"}, {"anchor_id": "cell:4:5:1:2", "text": "Q3"}]
  ]
}
```

### `table add-row | delete-row`

```bash
pptlive table add-row    --slide 4 --shape 5 --values '["Revenue", "$4.2M"]'
pptlive table add-row    --slide 4 --shape 5         # blank row
pptlive table delete-row --slide 4 --shape 5 --row 2
```

`--values` is an optional JSON array; the row is filled left-to-right. Both
verbs report the new row count.

---

## Charts — the `chart` group

A chart is a shape; its data lives in an embedded Excel workbook. Address by
`--slide S --shape N`.

```bash
pptlive chart read     --slide 4 --shape 6
pptlive chart set-type --slide 4 --shape 6 --chart-type line
pptlive chart set-data --slide 4 --shape 6 \
    --categories "A,B,C" --series '{"S1":[1,2,3],"S2":[4,5,6]}'
```

```json
{
  "slide": 4, "shape": 6, "anchor_id": "shape:4:6", "chart_type": "column",
  "categories": ["Q1", "Q2", "Q3"],
  "series": [{"name": "Revenue", "values": [10, 20, 30]}]
}
```

`--chart-type` is a friendly name (`column`, `line`, `pie`, …); see
`--help` for the full `--chart-type` choices.

---

## Selection & navigation

### `selection`

Report the user's current selection, resolved to anchors. A polite read — it
doesn't change the selection.

```bash
pptlive selection
```

```json
{"type": "text", "slide": 2, "anchor_id": "para:2:3:2",
 "shapes": [], "paragraph": 2, "text": "Demo"}
```

`type` is `none` / `slides` / `shapes` / `text`. The `anchor_id` is exactly
what `--anchor-id here:` resolves to.

### `go-to --anchor-id ID`

Move the user's view to an anchor's slide — a **deliberate, opt-in** view move
(unlike the polite edit verbs). Selects the target shape by default.

```bash
pptlive go-to --anchor-id shape:3:1
pptlive go-to --anchor-id ph:5:title --no-select
```

---

## Slide show — the `show` group

Drive a running slide show like a presenter's clicker. These **deliberately**
change what's on screen. Every control verb prints the resulting state; they
need a running show (`show next` et al. exit `1` if none is running). `show
state` is the read-only verb and never raises.

```bash
pptlive show start --from 2        # run from slide 2 (default: the top)
pptlive show next                  # advance a build/slide
pptlive show prev
pptlive show goto --slide 5
pptlive show black                 # blank to black
pptlive show white                 # blank to white
pptlive show resume                # un-blank
pptlive show state                 # read-only
pptlive show end
```

```json
{"running": true, "state": "running", "current_slide": 5, "slide_count": 12, "position": 5}
```

An out-of-range `--from` / `--slide` is exit `2`.

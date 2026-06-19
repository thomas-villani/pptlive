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

`read text-frame-status --anchor-id shape:4:3` reports a shape's autofit state
when text looks clipped or overflowing — `autosize` (the autofit mode),
`word_wrap`, the four `margins` (points), and an `overflow_risk` flag
(`"possible"` when autosize is off so text can clip, `"low"` when an autofit mode
is active). All reads; the view never moves.

```json
{"anchor_id": "shape:4:3", "autosize": "none", "word_wrap": true,
 "margins": {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}, "overflow_risk": "possible"}
```

## `find --text "…"`

Fuzzy, deck-wide search. PowerPoint has no document-wide character stream, so
this **traverses** every text frame — shapes, table cells, and speaker notes —
and reports each hit against a resolvable anchor (`shape:S:N`, `cell:S:N:R:C`,
`notes:S`) with a 0-based in-frame offset and a context snippet, in document
order. Matching is smart-quote / dash / whitespace tolerant (and
case-sensitive), so text an LLM re-typed off a slide still matches the original
glyphs. `--in` scopes the search to a `slide:S` or any text anchor.

```bash
pptlive find --text "Q3 revenue"
pptlive find --text "Demo" --in slide:2          # one slide
pptlive find --text "Metric" --in shape:4:5      # one shape / table / notes anchor
```

```json
[{"anchor_id": "notes:1", "start": 12, "length": 10, "text": "Q3 revenue",
  "context": "…recap of Q3 revenue versus plan…"}]
```

Never raises on a miss — zero matches is an **empty array** and exit `0`.

## `write --anchor-id ID --text "…"`

Set the entire text of a text anchor. Preserves the viewed slide; one Ctrl-Z.
Embed `\n` (or `\r`) to start a new paragraph — each line becomes its own
addressable `para:S:N:P`. For a soft line break *within* a paragraph, embed `\v`.

```bash
pptlive write --anchor-id ph:2:title --text "Agenda"
pptlive write --anchor-id ph:2:body  --text "Intro\nDemo\nQ&A"
pptlive write --anchor-id cell:4:5:1:1 --text "Metric"   # a cell is a text anchor
```

```json
{"ok": true, "anchor_id": "ph:2:title", "kind": "placeholder"}
```

## `replace` — whole anchor *or* fuzzy span

Two modes, mutually exclusive:

- `replace --anchor-id ID --text "…"` overwrites a text anchor's whole contents
  — identical in effect to `write`.
- `replace --find OLD --text NEW` runs the same fuzzy traversal as `find` and
  rewrites just the **matched span**, so the rest of the frame keeps its run
  formatting. Scope it with `--in slide:S` (or any anchor). Matches are computed
  once up front and applied in reverse offset order, so a replacement that
  re-contains the search text is safe.

```bash
pptlive replace --anchor-id shape:3:1 --text "New text"     # whole anchor
pptlive replace --find "Acme" --text "Globex" --all         # every occurrence
pptlive replace --find "teh" --text "the" --occurrence 2     # only the 2nd hit
```

One match auto-applies. Several matches without `--all` or `--occurrence` is
exit `5` (ambiguous — the matches are listed so you can disambiguate); zero
matches is exit `2`.

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

Reposition the layout's placeholders **in the same op** with `--placeholders` (a
JSON map of `KIND → {left, top, width, height}` in points, any subset) — the
"body on the left half beside a right panel" case, without an add-then-resize
fix-up. The validated geometry is echoed back:

```bash
pptlive slide add --layout two_content \
    --placeholders '{"body": {"left": 40, "width": 440}}'
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

### `slide geometry N`

A spatial map of slide `N` — the slide size, every shape's bounding `box`, an
`off_slide` flag, and the `overlaps` pairs (biggest first) — so you can catch
overlapping or off-edge shapes **without** a render. Axis-aligned only (rotation
is reported, not accounted for). A read; the view doesn't move.

```bash
pptlive slide geometry 2
```

```json
{"slide": 2, "size": {"width": 960, "height": 540},
 "shapes": [{"shapeid": "shapeid:2:3", "name": "Title 1",
             "box": {"left": 36, "top": 28, "width": 888, "height": 90},
             "off_slide": false}],
 "overlaps": [], "off_slide": []}
```

### `slide animations` · `slide clear-animations`

`slide animations N` lists slide `N`'s shape animations in play order — each row
maps an effect back to its target `shapeid`, so you can see *what* animates *how*
without a render (a read). `slide clear-animations --slide N` removes **every**
animation from a slide and reports how many effects it deleted.

```bash
pptlive slide animations 3
pptlive slide clear-animations --slide 3
```

```json
[{"seq_index": 1, "shapeid": "shapeid:3:4", "shape": "Title 1",
  "effect": "fade", "exit": false, "trigger": "on_click",
  "duration": null, "delay": null}]
```

### `slide headers-footers` · `set-footer` · `slide-number` · `set-date`

Per-slide footer / slide-number / date overrides (the deck-wide defaults live on
the `master` group below — same verb names). `slide headers-footers N` reads the
current settings; the setters mutate (one Ctrl-Z). Setting footer or date **text**
auto-shows that element. A date is either a **fixed** string (`--text`) or an
**auto-updating** format (`--format`, a `PpDateTimeFormat` int) — not both.

```bash
pptlive slide headers-footers 2
pptlive slide set-footer   --slide 2 --text "Acme — Confidential"   # auto-shows
pptlive slide set-footer   --slide 2 --hide                         # just hide it
pptlive slide slide-number --slide 2 --show
pptlive slide set-date     --slide 2 --format 1                     # auto-updating
pptlive slide set-date     --slide 2 --text "June 2026"             # fixed
```

```json
{"ok": true, "slide": 2,
 "headers_footers": {"footer": {"visible": true, "text": "Acme — Confidential"},
                     "slide_number": {"visible": true},
                     "date": {"visible": false, "text": null, "use_format": null}}}
```

A footer / date `text` reads back `null` while that element is hidden — PowerPoint
only exposes the text on a visible element — so a `null` text next to `visible:
false` means "hidden", not "empty".

---

## `snapshot` — see the whole deck cheaply

Render slides to PNG so a vision model can *see* the whole deck at a predictable
token cost. `--max-dim N` caps each slide's **long edge** in pixels (only ever
lowering resolution) — the lever for "render the deck and check my styling
landed" without full-resolution bloat. Because every slide shares one geometry,
the cap is a *uniform* per-slide budget; `~1000` stays legible. Renders the
current (unsaved) state; polite (doesn't move the view).

Select with `--slide N` (one) or `--slides A-B` (an inclusive span); omit both
for the whole deck. With `--out PATH` the PNGs are written (a single slide to
that path, multiple as `<stem>-s<N><suffix>`) and each `path` is reported;
without `--out`, base64 PNG data is returned inline.

For an exact per-slide pixel size instead of the long-edge cap, pass `--width N`
/ `--height N` (one or both; they **override** `--max-dim`, and passing `--max-dim`
together with either is an error). Pixel area — not encoder quality — is what a
vision model is billed on, and `Slide.Export` exposes no JPEG-quality knob, so the
dimensions are the only render-cost lever.

```bash
pptlive snapshot --max-dim 1000                        # whole deck, base64 inline
pptlive snapshot --out deck.png --max-dim 1000         # -> deck-s1.png, deck-s2.png, …
pptlive snapshot --slides 2-4 --max-dim 800            # just slides 2–4
pptlive snapshot --slide 1 --width 1280 --height 720   # exact pixels (overrides --max-dim)
```

```json
{"ok": true, "selector": "all slides", "count": 3, "format": "png",
 "max_dim": 1000, "images": [{"slide": 1, "bytes": 24, "base64": "iVBORw0KG…"}, …]}
```

---

## Save & export — `save` · `save-as` · `export-pdf`

Explicit file output — pptlive **never auto-saves**. `status` shows each deck's
`saved` flag (and flags `(unsaved)` in `--text`), so you can tell there's unsaved
work before deciding to persist it.

- **`save`** — save the deck to its existing file. Exits **1** if the deck has
  never been saved (no path yet) — use `save-as PATH` first. (The guard is
  deliberate: PowerPoint's own `Save` would silently upload a path-less deck to
  your default OneDrive/SharePoint folder.)
- **`save-as PATH [--format pptx] [--overwrite]`** — write a `.pptx` and **rebind**
  the working file to it: afterwards the open deck *is* PATH (its name/path
  follow), matching PowerPoint's Save-As. Refuses to clobber an existing file
  unless `--overwrite`. For PDF, use `export-pdf`.
- **`export-pdf PATH`** — export a pixel-faithful PDF of the deck's current
  (unsaved) state. A **read**: it neither rebinds the working file nor clears its
  dirty flag, so your `.pptx` is untouched. Overwrites an existing PDF.

```bash
pptlive status                              # see which decks have unsaved changes
pptlive save                                # persist in place (needs an existing path)
pptlive save-as C:\out\v2.pptx              # write + rebind the working file
pptlive save-as C:\out\v2.pptx --overwrite  # allow clobbering
pptlive export-pdf C:\out\deck.pdf          # a read — working file untouched
```

```json
{"ok": true, "path": "C:\\out\\deck.pdf"}
```

---

## Shapes — the `shape` group

Create and place shapes (geometry in **points**; 1 inch = 72 pt). Each verb is
one Ctrl-Z.

### `shape add`

```bash
# Text box
pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72

# Autoshape (see --shape-type choices), with a solid fill and no border
pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120 \
    --fill "#1E74B5" --line none

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
 "fill": {"color": null, "visible": true},
 "line": {"color": null, "weight": 1.0, "visible": true},
 "alt_text": "", "text": "Revenue up 12%"}
```

Notes:

- `--kind picture` requires `--path`; `--kind table` requires `--rows`/`--cols`;
  `--kind chart` needs both `--categories` and `--series`.
- `--categories` takes a JSON array or a comma-separated list. `--series` takes
  a JSON object `{"name":[values]}` or an array of `[name,[values]]`.
- `--fill` / `--line` (textbox/shape only) take a `#RRGGBB` hex or `none`
  (transparent fill / no border); `--line-width` is the border weight in points.
- Every shape read carries `fill` and `line` (`{color, visible[, weight]}`); a
  theme/automatic color reads back as `color: null`, never a misleading `#000000`.
- A new shape lands at the **top of the z-order** (last slot), so its
  `shape:S:N` is the post-add `Shapes.Count`. A text box created with text
  auto-fits its height, so a requested `--height` is advisory when AutoSize is
  on.

### `shape move | resize | delete | order`

Address by `shape:S:N`, `shapeid:S:ID`, or `ph:S:KIND` (a non-shape anchor like
`notes:S` is exit `2`).

```bash
pptlive shape move   --anchor-id shape:4:3 --left 100 --top 140
pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200
pptlive shape delete --anchor-id shape:4:3
pptlive shape order  --anchor-id shape:4:3 --to back    # front | back | forward | backward
```

`move` and `resize` echo the new `geometry`; each needs at least one
coordinate/dimension. `order` restacks the shape and echoes its new 1-based
`index` — `--to back` tucks a freshly added background panel *behind* the
existing content (otherwise it lands on top). Note that delete/restack shift the
`shape:S:N` indices of the shapes they pass; re-read after, or address by
`shapeid:S:ID` (below).

### `shape fill`

Set a shape's **fill** and/or **border** (the spatial equivalent of
`format-text`, which is font color). `--fill`/`--line` take a `#RRGGBB` hex or
`none`; `--line-width` is points. `--fill-transparency`/`--line-transparency` are
a `0..1` alpha (0 opaque, 1 fully transparent) — the partial-alpha knob, distinct
from `none` (which hides it entirely). Pass at least one. Echoes the shape's
updated `fill`/`line` (each now carrying `transparency`).

```bash
pptlive shape fill --anchor-id shape:4:3 --fill "#1E74B5" --line none
pptlive shape fill --anchor-id shapeid:4:9 --line "#333333" --line-width 1.5
pptlive shape fill --anchor-id shape:4:3 --fill "#1E74B5" --fill-transparency 0.4
```

### `shape gradient-fill` / `picture-fill` / `pattern-fill`

The non-solid fills (the `fill` read reports a `type` of `gradient`/`picture`/
`patterned` and, for a gradient, the `stops`). Gradients take `--colors` (one =
one-color with `--degree` brightness, two = two-color, three+ = multi-stop with
`--positions` placing the interior stops) **or** a named `--preset`; `--style`
sets the sweep. Pattern fills take a `--pattern` name plus `--fore`/`--back`.

```bash
pptlive shape gradient-fill --anchor-id shape:4:3 --colors "#1a73e8,#ffffff" --style vertical
pptlive shape gradient-fill --anchor-id shape:4:3 --colors "#f00,#0f0,#00f" --positions "0,0.4,1"
pptlive shape gradient-fill --anchor-id shape:4:3 --preset ocean
pptlive shape picture-fill  --anchor-id shape:4:3 --path background.png
pptlive shape pattern-fill  --anchor-id shape:4:3 --pattern percent_50 --fore "#1E74B5" --back "#fff"
```

### `shape effect`

Set a shape's **shadow / glow / soft-edge / reflection** (the read reports an
`effects` field with the active ones). `--shadow`/`--glow` take a JSON object;
`--soft-edge`/`--reflection` take an int preset. Pass `none` to any flag to turn
that effect off.

```bash
pptlive shape effect --anchor-id shape:4:3 \
  --shadow '{"color":"#333333","blur":8,"offset_x":4,"offset_y":4}' \
  --glow '{"color":"#00AAFF","radius":10}' --soft-edge 4 --reflection 5
pptlive shape effect --anchor-id shape:4:3 --shadow none   # remove the shadow
```

### `shape line-style`

Set a shape's line **dash** pattern and/or **arrowheads** (the `line` read reports
`dash` plus `begin_arrow`/`end_arrow` when set). `--dash` is a
`solid`/`dash`/`round_dot`/`dash_dot`/`long_dash`/… name; `--begin-arrow`/
`--end-arrow` are `none`/`triangle`/`open`/`stealth`/`diamond`/`oval`, with
`--begin-arrow-size`/`--end-arrow-size` of `small`/`medium`/`large`. **Arrowheads
apply to lines/connectors only** — a closed shape rejects them (use `--dash`
there). Pass at least one.

```bash
pptlive shape line-style --anchor-id shape:4:3 --dash dash_dot
pptlive shape line-style --anchor-id shape:4:5 --end-arrow triangle --end-arrow-size large
```

### `shape set-alt`

Set a shape's alternative text — a description you can re-find the shape by
after z-order drift.

```bash
pptlive shape set-alt --anchor-id shape:4:3 --alt-text "Acme logo (top-right)"
```

### `shape animate` · `shape clear-animations`

Give a shape an entrance (or, with `--exit`, an exit) animation — the per-shape
sibling of slide transitions, appended to the slide's main animation sequence. A
shape can carry several effects. `--effect` is a curated name
(`appear`/`fade`/`fly_in`/`float_in`/`wipe`/`zoom`/`grow_turn`/`swivel`/`wheel`/
`split`); `--trigger` is `on_click` / `with_previous` / `after_previous`;
`--duration` / `--delay` are seconds. `shape clear-animations` removes just that
shape's effects (vs. the whole-slide `slide clear-animations`).

```bash
pptlive shape animate --anchor-id shape:3:2 --effect fly_in --trigger after_previous
pptlive shape animate --anchor-id shape:3:2 --effect fade --exit          # animate OUT
pptlive shape clear-animations --anchor-id shape:3:2
```

```json
{"ok": true, "anchor_id": "shape:3:2", "shapeid": "shapeid:3:7",
 "animation": {"shapeid": "shapeid:3:7", "shape": "Rectangle 4", "effect": "fly_in",
               "exit": false, "trigger": "after_previous", "duration": null, "delay": null}}
```

### `shapeid:S:ID` — the delete-proof handle

Every shape read emits a stable `id` (`Shape.Id`). Address a shape by it with
`shapeid:S:ID` anywhere an `--anchor-id` is taken. Unlike `shape:S:N` (a z-order
index that shifts when a lower shape is deleted or restacked), the `shapeid`
keeps pointing at the same shape across structural edits — reach for it when a
multi-step batch deletes or reorders shapes it later references.

```bash
pptlive shape fill --anchor-id shapeid:4:9 --fill "#102030"
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
level, bullet, alignment, and effective `font` (`bold`/`italic`/`underline` as
`true`/`false`/`"mixed"`, `size`, `font` name, `color` `#RRGGBB` or `null` for a
theme/automatic color). The `font` values are *rendered* — COM exposes no
"directly set vs inherited" flag (only color distinguishes a literal from a theme
color). Each paragraph also carries `space_before` / `space_after` /
`line_spacing` as `{value, mode}` (where `mode` is `"multiple"` or `"points"` —
see the line-spacing note under `format-paragraph`) and `run_sizes`, the distinct
per-run font sizes — so a stray 5 pt run hiding in an otherwise-18 pt paragraph
shows up as `"run_sizes": [18.0, 5.0]` before it ever renders.

```bash
pptlive paragraphs --anchor-id ph:4:body
```

```json
[
  {"anchor_id": "para:4:2:1", "text": "Revenue up 12%", "indent_level": 1, "bullet": "bulleted",
   "font": {"bold": false, "italic": false, "underline": false, "size": 18.0, "font": "Aptos", "color": "#000000"}},
  {"anchor_id": "para:4:2:2", "text": "Churn down 3%",  "indent_level": 1, "bullet": "bulleted",
   "font": {"bold": false, "italic": false, "underline": false, "size": 18.0, "font": "Aptos", "color": null}}
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

!!! warning "Line spacing has two units"
    `--line-spacing` is a **multiple** (`1.0` single, `1.5`, `2.0`). For an exact
    *point* height use **`--line-spacing-points 24`**. Passing `--line-spacing 24`
    means 24× line height — text shoots off the slide — so a multiple `> 5` is
    **rejected** (exit 1) unless you add `--force`. The same split applies to
    spacing before/after: `--space-before` / `--space-after` are **points**,
    `--space-before-lines` / `--space-after-lines` are **multiples**. Passing both
    the point and the multiple form of the same knob is an error.

### `set-paragraphs --anchor-id ID --json '[...]'`

Rewrite an anchor as a clean per-paragraph list — the **safe** way to author a
bullet list. Each item is a plain string or an object
`{text, list_type?, indent_level?, alignment?, line_spacing?/line_spacing_points?,
size?, bold?, ...}`; one item becomes exactly one addressable `para:S:N:P` (a
newline *inside* an item folds to a soft break), so there's no `\n`-inference and
no separate `list apply` pass. `--file PATH` reads the JSON array from a file.

```bash
pptlive set-paragraphs --anchor-id ph:4:body --json \
  '["Overview", {"text": "Revenue up 12%", "list_type": "bulleted", "indent_level": 1},
                 {"text": "Churn down 3%",  "list_type": "bulleted", "indent_level": 1}]'
```

```json
{"ok": true, "anchor_id": "ph:4:body", "paragraphs": ["para:4:2:1", "para:4:2:2", "para:4:2:3"]}
```

### `reset-format --anchor-id ID` · `shape reset-to-layout --anchor-id ID`

Recover a placeholder that's spiralled into a bad state (giant line spacing, 5 pt
font, off the slide). PowerPoint has no "clear formatting" button, so the two
verbs split the job:

- **`reset-format`** resets paragraph *spacing* to clean defaults (single
  line-spacing, zero before/after) — the only unambiguous reset.
- **`shape reset-to-layout`** restores a placeholder's geometry **and** default
  font size from its layout's matching placeholder (the "5 pt font / shape off the
  slide" fix). It only works on a placeholder anchor.

```bash
pptlive reset-format --anchor-id ph:4:body
pptlive shape reset-to-layout --anchor-id ph:4:body
```

The reliable repair sequence is **`read anchor` → `reset-format` → `shape
reset-to-layout` → `set-paragraphs` → `slide export`** (render and check).

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

### Recolor chart text — `chart recolor-text`

A chart has no text anchor, so `format-text` can't reach its internal text. Use
`chart recolor-text` for the coarse "recolor all of it" move — the fix when the
inherited (black) axis/legend text is invisible on a dark or custom background.

```bash
pptlive chart recolor-text --slide 6 --shape 2 --color "#FFFFFF"
```

```json
{
  "ok": true, "slide": 6, "shape": 2, "anchor_id": "shape:6:2", "color": "#FFFFFF",
  "recolored": ["chart_area", "legend", "category_axis", "value_axis"],
  "series_data_labels": 0
}
```

`recolored` lists the elements actually touched: it only recolors text that's
**shown** (legend/title guarded by their presence; axis tick labels and data
labels best-effort, so a pie chart's absent axes are simply skipped). It never
adds a legend, title, or labels the deck didn't already display.

---

## SmartArt — the `smartart` group

A SmartArt diagram is a shape too; its content is a **node tree**. Address by
`--slide S --shape N`. Create one with `shape add --kind smartart` (with
`--smartart-kind` and optional `--nodes`).

```bash
pptlive shape add --slide 3 --kind smartart --smartart-kind process \
    --nodes '["Discover", "Design", "Build", "Ship"]'
pptlive smartart read      --slide 3 --shape 2
pptlive smartart set-nodes --slide 3 --shape 2 \
    --nodes '["Plan", {"text": "Execute", "children": ["Build", "Test"]}, "Ship"]'
```

```json
{
  "slide": 3, "shape": 2, "anchor_id": "shape:3:2",
  "layout": "process", "layout_id": "urn:microsoft.com/office/officeart/2005/8/layout/process1",
  "node_count": 4,
  "nodes": [{"text": "Discover", "level": 1, "children": []}, ...]
}
```

`--nodes` is a JSON array of plain strings (leaves) and/or `{text, children}`
objects (`children` nests recursively). Flat layouts (`process`, `cycle`,
`list`, `pyramid`, `venn`) take any number of top-level nodes; tree layouts
(`hierarchy`, `orgchart`) take a **single root** with nested children — passing
more than one top-level node to a tree layout is an error (exit 1).

### Recolor node text — `smartart recolor-text`

Node labels live on each node's text frame, with no per-anchor handle, so
`smartart recolor-text` recolors **every** node at once — the fix when the
inherited (black) node text is invisible on a dark or custom background.

```bash
pptlive smartart recolor-text --slide 3 --shape 2 --color "#FFFFFF"
```

```json
{ "ok": true, "slide": 3, "shape": 2, "anchor_id": "shape:3:2",
  "color": "#FFFFFF", "nodes_recolored": 4 }
```

---

## Comments — the `comment` group

Review comments — read a reviewer's notes and reply to them, the "address the
comments" workflow. Comments attach to a **slide** at an `(x, y)` point and are
**threaded**; you address one for reply/delete by `--slide S --index N` (1-based,
from `comment list`). `comment list` with no `--slide` is a deck-wide roll-up.

```bash
pptlive comment list                                  # deck-wide {total, slides:[...]}
pptlive comment list --slide 1                        # one slide's comments + threads
pptlive comment add   --slide 2 --text "Please cite a source."
pptlive comment reply --slide 1 --index 1 --text "Done."
pptlive comment delete --slide 1 --index 1            # removes the comment + its replies
```

```json
[{"index": 1, "author": "Thomas Villani", "initials": "TV",
  "text": "Tighten this headline.", "datetime": "2026-06-07T10:30:00+00:00",
  "left": 12.0, "top": 12.0,
  "replies": [{"index": 1, "author": "Thomas Villani", "text": "Agreed — will do.", ...}]}]
```

A new comment **binds to the signed-in Office account** — the shown author follows
that account, not `--author`/`--initials` (those reach only the legacy fallback used
on a deck that has no existing comment to source an identity from). There is **no
resolve verb**: PowerPoint's COM doesn't expose comment resolution state on current
builds. Add/reply/delete are each one Ctrl-Z and don't move your view.

---

## Theme — the `theme` group

Deck-wide styling: the 12-slot palette and the heading/body typefaces. These are
**global, anti-polite** ops — one change recolors or re-fonts every slide that
inherits the theme (the edit is still one Ctrl-Z, and your *view* doesn't move).

```bash
pptlive theme read                                   # {colors:{slot:#RRGGBB}, fonts:{major, minor}}
pptlive theme set-color --slot accent1 --color "#C00000"
pptlive theme set-font  --which major --name "Georgia"        # major = headings, minor = body
pptlive theme set-font  --which minor --name "Calibri" --script latin
```

`--slot` is one of the 12 palette slots (`dark1`/`dark2`, `light1`/`light2`,
`accent1`…`accent6`, `hyperlink`, `followed_hyperlink`). `--which` is `major`
(headings) or `minor` (body); `--script` is `latin` (default), `east_asian`, or
`complex_script`.

---

## Master — the `master` group

The primary slide master's **text styles** (PowerPoint's nearest "named style"
analog: `title` / `body` / `default`, 5 outline levels each) and its background.
Also global and anti-polite, also one Ctrl-Z.

```bash
pptlive master read                                  # {text_styles:{...}, background:{type, color}}
pptlive master format-text-style      --style body  --level 1 --font "Georgia" --size 28 --color "#333333"
pptlive master format-paragraph-style --style title --level 1 --alignment center --space-after 12
pptlive master set-background --color "#1F1F1F"      # solid fill (v0.9 ships solid only)
```

`format-text-style` mirrors `format-text` (`--bold/--no-bold`, `--italic`,
`--underline`, `--size`, `--font`, `--color`) and `format-paragraph-style`
mirrors `format-paragraph` (`--alignment`, `--space-before`, `--space-after`,
`--line-spacing`) — but applied deck-wide to a `--style` + `--level` instead of
to one anchor. `--level` defaults to `1` (the natural choice for `title`); pass
`--level N` (1–5) for the other outline levels. Each needs at least one
formatting option.

The master also carries the deck's **default** headers / footers — the same four
verbs as the `slide` group (`headers-footers` / `set-footer` / `slide-number` /
`set-date`), but setting the deck-wide default every slide inherits unless it has
its own per-slide override:

```bash
pptlive master headers-footers                       # read the deck defaults
pptlive master set-footer --text "Acme — Confidential"
pptlive master slide-number --show
pptlive master set-date --format 1                    # auto-updating on every slide
```

---

## Sections — the `section` group

PowerPoint **sections** — named spans of slides for organizing a long deck.
Structural edits (no view move), each one Ctrl-Z. Sections are addressed by a
1-based `--section` index.

```bash
pptlive section list
pptlive section add    --name "Appendix" --before-slide 9   # start a span at slide 9
pptlive section add    --name "Backup"                      # append an empty trailing section
pptlive section rename --section 2 --name "Results"
pptlive section move   --section 3 --to 1                   # carries its slides
pptlive section delete --section 3                          # keeps the slides…
pptlive section delete --section 3 --delete-slides          # …unless you say so
```

```json
[{"index": 1, "name": "Intro", "first_slide": 1, "slide_count": 3},
 {"index": 2, "name": "Results", "first_slide": 4, "slide_count": 5}]
```

Two model notes the spike pinned: starting a section with `--before-slide` in
front of a later slide **auto-creates a leading "Default Section"** for the slides
ahead of it, and `delete` keeps the section's slides by default (it just drops the
boundary) — pass `--delete-slides` to remove them too. A section with no slides
reports `first_slide: null`.

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

---

## `exec --script ops.json` — a batch as one Ctrl-Z

Apply a whole batch script against **one** connection in **one** undo entry — the
single-process way to build or restyle a slide without a command per change. The
script is a JSON object `{"label": "...", "ops": [...]}`; each op is one of the
edit/read verbs, named by its MCP op (`write`, `set_paragraphs`, `format`,
`shape_add`, `find_replace`, …) and addressed by `anchor_id` / the same params the
[MCP](mcp.md) ops take. Every op defaults to the `edit` tool.

Each op's keys are **flat** (the op's params sit right alongside `"op"`), exactly
the fields the matching [MCP](mcp.md) op takes — so `set_paragraphs` takes
`paragraphs`, `write` takes `anchor_id` / `text`, and so on:

```json
{
  "label": "Build Q3 slide",
  "ops": [
    {"op": "slide_add", "layout": "title_and_content", "index": 4},
    {"op": "write", "anchor_id": "ph:4:title", "text": "Q3 Results"},
    {"op": "set_paragraphs", "anchor_id": "ph:4:body",
     "paragraphs": [{"text": "Revenue up 12%", "list_type": "bulleted"},
                    {"text": "Churn down 3%",  "list_type": "bulleted"}]}
  ]
}
```

```bash
pptlive exec --script ops.json
```

- The batch is **one Ctrl-Z** — a single automation session, so a partial run
  (some ops applied before a failure) reverts with one undo.
- It **stops at the first failing op** by default — that op's category sets the
  exit code (`2` not-found, `5` ambiguous, …). Pass `--continue` to run every op
  and report each outcome.
- `--no-atomic` fences each op as its own undo entry instead of one.
- **"Follow the work" view policy.** When a batch *adds* a slide
  (`slide_add` / `slide_duplicate`), the view is left on the last slide it touched
  rather than snapped back to the pre-batch slide (so building a deck doesn't keep
  bouncing you to slide 1). Pure-edit batches keep the polite view-restore. Opt out
  with `--no-follow-view` (or the `PPTLIVE_VIEW_FOLLOW=0` env var); a deliberate
  `navigate` op in the batch still wins.

Each result entry carries its `index`, `tool`, `op`, `ok`, and either the op's
`result` payload or an `error` token + `message`:

```json
{"ok": true, "label": "Build Q3 slide", "atomic": true, "count": 3,
 "results": [
   {"index": 0, "tool": "edit", "op": "slide_add",      "ok": true, "result": {"index": 4, "id": 261}},
   {"index": 1, "tool": "edit", "op": "write",          "ok": true, "result": {"ok": true, "anchor_id": "ph:4:title"}},
   {"index": 2, "tool": "edit", "op": "set_paragraphs", "ok": true, "result": {"ok": true, "paragraphs": ["para:4:2:1", "para:4:2:2"]}}
 ]}
```

A malformed script or unknown op is exit `1` (`invalid_args`). `shape:S:N` refs
resolve **live** as each op runs, so address anything you didn't just create by
`ph:S:KIND` / `.Name` / `shapeid:S:ID`.

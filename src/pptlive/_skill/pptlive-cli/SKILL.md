---
name: pptlive-cli
description: Read and edit the Microsoft PowerPoint presentation the user has open right now, from the command line. Inspect structure (slides, outline, shapes, tables, charts, SmartArt), make polite edits (text, formatting, layouts, shapes, pictures, theme/master styling), render a slide or shape to a PNG so a vision model can see the layout, drive a live slide show, and batch changes into a single atomic undo — all JSON-in / JSON-out with deterministic exit codes. Use when the user wants to read, edit, or visually render a .pptx that is currently open in PowerPoint on Windows.
---

# pptlive (CLI)

`pptlive` drives a **running** Microsoft PowerPoint instance over COM (Windows
only). Unlike `python-pptx`, it edits the deck the user has **open right now** —
and politely: their viewed slide and shape/text selection are preserved, and
every `edit` block collapses into a single Ctrl-Z.

Prefer the **CLI**. Every command prints exactly one JSON object on stdout and
returns a deterministic exit code, so you branch on failures without parsing
prose. JSON is the default; `--text` (human-readable), `--json`, and `--doc NAME`
(default: the active deck) are **global flags — put them before the
subcommand**: `pptlive --text outline`, not `pptlive outline --text`.

(For the Python API instead of the CLI, run `pptlive llm-help --python`.)

## First, orient yourself
1. `pptlive status` — confirm PowerPoint is reachable; see open decks + viewed slide.
2. `pptlive slides` — `[{index, id, layout, title, shape_count, has_notes}]`.
3. `pptlive outline` — title + body bullets per slide.
4. `pptlive slide read 2` — every shape on slide 2.
5. `pptlive shapes --slide 2` — shapes with their `anchor_id`, `name`, `id`, type, geometry.

## Anchors — how you address things
Addressing is **hierarchical** (slide → shape → text), slide-index first. There
is no deck-wide `range:`. Pass an anchor as `--anchor-id`:

| anchor_id      | resolves to |
| -------------- | ----------- |
| `shape:S:N`    | Nth shape (1-based z-order) on slide S — the canonical handle |
| `shapeid:S:ID` | shape with stable `Shape.Id` ID on slide S — the **delete-proof** handle (the `id` in any shape listing) |
| `ph:S:KIND`    | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — **prefer this** |
| `para:S:N:P`   | paragraph P (1-based) of shape N on slide S |
| `cell:S:N:R:C` | cell (row R, col C) of the table in shape N on slide S — a cell takes every text/format verb |
| `notes:S`      | speaker-notes body of slide S |
| `comments:S`   | the review comments on slide S — a **read selector** (container, not a text anchor); address one for reply/delete by `--slide S --index N` |
| `here:`        | whatever the user has selected right now (the opt-in way to act on the live selection) |

`body` also matches the generic **content** placeholder (reads back as
`placeholder: object`, e.g. "Content Placeholder N"). A **Two Content** /
**Comparison** layout has *two*, so `ph:S:body` is ambiguous and exits 5 listing
the candidate `shape:S:N` anchors — target each column by `shape:S:N` / `.Name`.

z-order **drifts** as shapes are added, removed, *or restacked* (`shape order`),
so `shape:S:N` is resolved live and never cached; every shape listing also emits
`name` (`Shape.Name`) and `id` (`Shape.Id`, stable across reorder *and* delete) so
you can re-identify after drift. Steer toward `ph:S:KIND`, `.Name`, and
`shapeid:S:ID` (delete-proof) as the drift-proof forms.

## Reading
- `pptlive read anchor --anchor-id ph:2:title` — read any text anchor (`ph:`/`shape:`/`para:`/`cell:`/`notes:`/`here:`).
- `pptlive read notes --slide 1` — sugar for `--anchor-id notes:1`.
- `pptlive paragraphs --anchor-id ph:4:body` — `[{anchor_id (para:S:N:P), text, indent_level, bullet}]`.
- `pptlive table read --slide 4 --shape 5` · `pptlive chart read --slide 4 --shape 5` · `pptlive smartart read --slide 3 --shape 2`.
- `pptlive theme read` · `pptlive master read` — deck-wide palette/fonts and master text styles.
- `pptlive selection` — what the user has selected (resolves to `here:`).
- `pptlive find --text "Q3 revenue" [--in slide:3|shape:3:2|notes:3]` — fuzzy, smart-quote/whitespace-tolerant search across the deck (shapes, table cells, notes). Emits `[{anchor_id, start, length, text, context}]` in document order; empty array (exit 0) on no match.

## Writing — each command is one atomic undo
- `pptlive write --anchor-id ph:2:body --text "Intro\nDemo\nQ&A"` — set a text anchor (`\n`/`\r` = new paragraph, each separately addressable as `para:`; `\v` = soft line break within a paragraph).
- `pptlive replace --anchor-id shape:3:1 --text "New text"` — overwrite a whole anchor.
- `pptlive replace --find "old" --text "new" [--in slide:3] [--all|--occurrence N]` — fuzzy find/replace; rewrites just the matched span (keeps run formatting). One match auto-applies; several without `--all`/`--occurrence` is exit 5 (ambiguous, lists the matches); zero is exit 2.
- `pptlive insert --anchor-id para:4:2:3 --text "New bullet" [--before|--after]` — new paragraph relative to an anchor.
- `pptlive set-paragraphs --anchor-id ph:4:body --json '["First", {"text":"Second","list_type":"bulleted","indent_level":2}]'` — rewrite an anchor as a clean per-paragraph list (strings or `{text, list_type, indent_level, alignment, line_spacing/line_spacing_points, size, bold, ...}`). The **safe** way to author a bullet list — each item is exactly one `para:`, no `\n` inference. `--file PATH` reads the JSON instead.
- `pptlive format-paragraph --anchor-id para:4:2:1 --alignment center --indent-level 2 [--line-spacing 1.5 | --line-spacing-points 24]` — see the line-spacing footgun below; `--line-spacing` is a *multiple*, `--line-spacing-points` is *exact points*.
- `pptlive format-text --anchor-id ph:4:title --bold --size 40 --color "#2E74B5"`.
- `pptlive list apply --anchor-id ph:4:body --type bulleted [--char "•"]` · `pptlive list remove --anchor-id ph:4:body`.
- `pptlive reset-format --anchor-id ph:4:body` — recover a line-spacing spiral (reset paragraph spacing to single + zero before/after). `pptlive shape reset-to-layout --anchor-id ph:4:body` — restore a placeholder's geometry + default font size from its layout (the "5 pt font / off the slide" fix).
- `pptlive read text-frame-status --anchor-id shape:4:3` — autofit/wrap/margin diagnostics (`autosize`, `word_wrap`, `margins`, `overflow_risk`) when text looks clipped.
- `pptlive exec --script ops.json` — apply a whole batch script `{"label": "...", "ops": [{"op": "write", "anchor_id": ..., "text": ...}, ...]}` against one connection as **one Ctrl-Z**. Each op defaults to the `edit` tool (the op names are the MCP `ppt_edit`/`ppt_read`/... ops). Stops at the first failing op (exit code maps to its category) unless `--continue`; `--no-atomic` fences each op separately. The single-process way to build a slide without a command per change.

## PowerPoint text-model gotchas (read before formatting text)
PowerPoint's text model has sharp edges that leak through. The big ones:
- **Line spacing has two units.** `--line-spacing` is a **multiple** (1.0 single, 1.5, 2.0). For an exact *point* height use `--line-spacing-points 24`. Passing `--line-spacing 24` means 24× line height (text shoots off the slide) — so it's **rejected** unless `--force`. Same split for spacing before/after: `--space-before/--space-after` are points, `--space-before-lines/--space-after-lines` are multiples.
- **`\n` is a paragraph, not a soft break.** In `write`, `\n`/`\r` start a new addressable `para:`; `\v` is a soft line break within one paragraph. To author a list reliably, prefer `set-paragraphs` (one item = one bullet) over embedding newlines.
- **Paragraph formatting applies per paragraph, font formatting per run.** A `format-text --size` on a multi-run paragraph may hit only part of it; read `paragraphs` and check `run_sizes` to spot a stray small run.
- **There's no "clear formatting" button.** Re-writing the text does *not* drop run overrides. `reset-format` resets paragraph *spacing* to clean defaults; `shape reset-to-layout` restores a placeholder's geometry + default font. Font size/typeface otherwise need an explicit `format-text`.
- **When text overflows,** read `text-frame-status`: `overflow_risk: "possible"` means autosize is off (text can clip); `"low"` means an autofit mode is active.

### Formatting-field reference
| field (CLI flag) | unit / values | COM mapping | scope |
| --- | --- | --- | --- |
| `--line-spacing` | multiple (1.0, 1.5) | `ParagraphFormat.SpaceWithin` + `LineRuleWithin=msoTrue` | paragraph |
| `--line-spacing-points` | points (exact) | `SpaceWithin` + `LineRuleWithin=msoFalse` | paragraph |
| `--space-before` / `--space-after` | points | `SpaceBefore`/`SpaceAfter` + `LineRuleBefore/After=msoFalse` | paragraph |
| `--space-before-lines` / `--space-after-lines` | multiple | `SpaceBefore`/`SpaceAfter` + `LineRule*=msoTrue` | paragraph |
| `--indent-level` | int 1–5 | `TextRange.IndentLevel` | paragraph |
| `--alignment` | left/center/right/justify/distribute | `ParagraphFormat.Alignment` | paragraph |
| `list apply --type` | bulleted/numbered | `ParagraphFormat.Bullet.{Visible,Type}` | paragraph |
| `--size` | points (warns < 8) | `Font.Size` | run |
| `--bold`/`--italic`/`--underline` | flags | `Font.{Bold,Italic,Underline}` | run |
| `--color` (format-text) | `#RRGGBB` | `Font.Color.RGB` | run |

### Safe patterns
- **Bullet list:** `set-paragraphs --json '[{"text":"A","list_type":"bulleted"},{"text":"B","list_type":"bulleted"}]'` — don't hand-build with `\n` + a separate `list apply`.
- **Repair a wrecked placeholder:** `read anchor` → `reset-format` (spacing) → `shape reset-to-layout` (geometry+font) → `set-paragraphs` (clean text) → `slide export` to verify.

## Slides
- `pptlive slide layouts` — the layout names `add`/`set-layout` accept.
- `pptlive slide add --layout two_content [--index 4]`.
- `pptlive slide duplicate --slide 7` · `pptlive slide move --slide 9 --to 2` · `pptlive slide delete --slide 5`.
- `pptlive slide set-layout --slide 4 --layout title_and_content`.
- `pptlive slide set-transition --slide 4 --effect fade [--duration 0.5] [--advance-after 3] [--on-click/--no-on-click]` — entrance transition (`fade`/`cut`/`dissolve`/`cover_left`/… or `none`); `--advance-after N` auto-advances after N s. Slide reads carry a `transition` dict.
- `pptlive slide set-background --slide 4 --color "#1A2B3C"` (per-slide solid override of the master) · `--follow-master` reverts. Slide reads carry a `background` dict (`{follows_master, type, color}`).
- `pptlive slide export --slide 2 --out slide2.png [--width 1280] [--format png]` — render one slide to an image so a vision model can *see* it.
- `pptlive snapshot [--slide N | --slides A-B] [--out deck.png] [--max-dim 1000]` — render the **whole deck** (one PNG per slide) so you can check styling across every slide cheaply. `--max-dim` caps each slide's long edge (a uniform, predictable per-slide token cost); with `--out` it writes `<stem>-sN<suffix>`, otherwise base64 inline. The "did my restyle land everywhere?" read.

## Save & export (explicit — pptlive never auto-saves)
`status` shows each deck's `saved` flag (and flags `(unsaved)` in `--text`).
- `pptlive save` — save to the existing file. Exits **1** if the deck was never saved (use `save-as` first; the guard stops PowerPoint silently cloud-saving a path-less deck).
- `pptlive save-as PATH [--format pptx] [--overwrite]` — write a `.pptx` and **rebind** the working file to it (the open deck becomes PATH, like Save-As). Refuses to clobber unless `--overwrite`. For PDF use `export-pdf`.
- `pptlive export-pdf PATH` — export a pixel-faithful PDF of the current (unsaved) state. A **read**: no rebind, dirty flag preserved, your `.pptx` untouched. The "hand back a deliverable" path. Overwrites an existing PDF.

## Shapes
- `pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72` (points throughout; 1 in = 72 pt).
- `pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120 --fill "#1E74B5" --line none` (textbox/shape take `--fill`/`--line` = `#RRGGBB` or `none`, `--line-width` pts).
- `pptlive shape add --slide 4 --kind picture --path logo.png --left 600 --top 40 --alt-text "Acme logo"` (embedded, never linked).
- `pptlive shape move --anchor-id shape:4:3 --left 100 --top 140` · `pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200` · `pptlive shape delete --anchor-id shape:4:3`.
- `pptlive shape fill --anchor-id shape:4:3 --fill "#102030" --line none` — shape fill/border (NOT font color; that's `format-text`).
- `pptlive shape order --anchor-id shape:4:3 --to back` — restack (`front`/`back`/`forward`/`backward`); send a new background panel behind existing content.
- `pptlive shape set-alt --anchor-id shape:4:3 --alt-text "Acme logo"` — alt text doubles as a drift-proof re-id handle.
- `pptlive shape set-link --anchor-id shape:4:3 --url https://acme.com` (or `--slide 2` for an in-deck "back to agenda" jump; `--screen-tip "Acme"`). `pptlive shape remove-link --anchor-id shape:4:3`. A link needs no text frame; reads carry a `hyperlink` field (`{address, sub_address}` or null).
- `pptlive shape export --anchor-id shape:4:3 --out logo.png` — render one shape (native size).
- Every shape read carries `fill`/`line` (`{color, visible[, weight]}`) + `hyperlink`; a theme/automatic color is `color: null`. Delete/restack shifts `shape:S:N` — address by `shapeid:S:ID` to survive it.

## Tables, charts, SmartArt
- Tables: `pptlive shape add --slide 4 --kind table --rows 3 --cols 2 --left 72 --top 120`; `pptlive table add-row --slide 4 --shape 5 --values '["Revenue","$4.2M"]'`; `pptlive table delete-row --slide 4 --shape 5 --row 2`; write cells with `pptlive write --anchor-id cell:4:5:1:1 --text "Metric"`.
- Charts (data lives in an embedded Excel workbook): `pptlive shape add --slide 4 --kind chart --chart-type column --categories "Q1,Q2,Q3" --series '{"Revenue":[10,20,30]}'`; `pptlive chart set-type --slide 4 --shape 5 --chart-type line`; `pptlive chart set-data --slide 4 --shape 5 --categories "A,B" --series '{"S":[1,2]}'`.
- SmartArt (content is a node tree): `pptlive shape add --slide 3 --kind smartart --smartart-kind process --nodes '["Discover","Design","Build","Ship"]'`; `pptlive smartart set-nodes --slide 3 --shape 2 --nodes '[{"text":"CEO","children":["Eng","Sales"]}]'`.
- **Recolor composite text** (a chart/SmartArt has no text anchor, so this is the only color path for its internal text): `pptlive chart recolor-text --slide 6 --shape 2 --color "#FFFFFF"` recolors every shown chart text element (legend, axis tick labels, title, data labels); `pptlive smartart recolor-text --slide 3 --shape 2 --color "#FFFFFF"` recolors every node label. The coarse fix when inherited black chart/diagram text goes invisible on a dark (or any custom) background — no rebuild from primitives needed.

## Comments — review thread (the "address the comments" workflow)
Comments attach to a slide and are **threaded**; address one by `--slide S --index N`.
- `pptlive comment list` (deck-wide `{total, slides:[...]}`) or `--slide 1` (one slide + threads).
- `pptlive comment add --slide 2 --text "Cite a source."` · `pptlive comment reply --slide 1 --index 1 --text "Done."` · `pptlive comment delete --slide 1 --index 1`.
- A new comment **binds to the signed-in account** (the passed `--author`/`--initials` only apply to the legacy fallback on a comment-less deck). No resolve verb — comment resolution state isn't COM-readable.

## Theme & master — deck-wide styling
Global and anti-polite, but still one Ctrl-Z; your view doesn't move.
- `pptlive theme set-color --slot accent1 --color "#C00000"` — recolors the whole deck.
- `pptlive theme set-font --which major --name "Georgia"` — major = headings, minor = body.
- `pptlive master format-text-style --style body --level 1 --font "Georgia" --size 28`.
- `pptlive master set-background --color "#1F1F1F"` — deck-wide solid fill.

## View & slide show (deliberate, opt-in screen moves)
- `pptlive go-to --anchor-id shape:3:1` — move the user's view to an anchor's slide.
- `pptlive show start [--from 2]` · `pptlive show next` (also `prev`, `goto --slide N`) · `pptlive show black` (also `white`, `resume`) · `pptlive show state` (read-only) · `pptlive show end`.

## Exit codes — branch on these
| Code | Meaning | Retry? |
| ---- | ------- | ------ |
| 0 | success | — |
| 1 | other / bad input | fix the input |
| 2 | anchor / slide / shape / presentation not found (incl. zero matches) | re-read with `slides`/`shapes`, then retry |
| 3 | PowerPoint busy (a modal dialog is open) | **yes** — back off and retry |
| 4 | PowerPoint not running | only after the user opens PowerPoint |
| 5 | ambiguous match | re-run with a more specific anchor |
| 6 | shape has no text frame | target a shape that holds text |

## Typical workflow
1. `pptlive status` → confirm PowerPoint and the target deck.
2. `pptlive slides` / `outline` / `shapes` → get the anchor ids you need.
3. Edit with the verbs above. Each command is atomic and leaves the user's view + selection untouched.
4. `pptlive slide export …` → render and look, then revise.

Full docs: https://thomas-villani.github.io/pptlive/ · Python API: `pptlive llm-help --python`.

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
| `ph:S:KIND`    | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — **prefer this** |
| `para:S:N:P`   | paragraph P (1-based) of shape N on slide S |
| `cell:S:N:R:C` | cell (row R, col C) of the table in shape N on slide S — a cell takes every text/format verb |
| `notes:S`      | speaker-notes body of slide S |
| `here:`        | whatever the user has selected right now (the opt-in way to act on the live selection) |

z-order **drifts** as shapes are added/removed, so `shape:S:N` is resolved live
and never cached; every shape listing also emits `name` (`Shape.Name`) and `id`
(`Shape.Id`, stable across reorder) so you can re-identify after drift. Steer
toward `ph:S:KIND` and `.Name` as the drift-proof forms.

## Reading
- `pptlive read anchor --anchor-id ph:2:title` — read any text anchor (`ph:`/`shape:`/`para:`/`cell:`/`notes:`/`here:`).
- `pptlive read notes --slide 1` — sugar for `--anchor-id notes:1`.
- `pptlive paragraphs --anchor-id ph:4:body` — `[{anchor_id (para:S:N:P), text, indent_level, bullet}]`.
- `pptlive table read --slide 4 --shape 5` · `pptlive chart read --slide 4 --shape 5` · `pptlive smartart read --slide 3 --shape 2`.
- `pptlive theme read` · `pptlive master read` — deck-wide palette/fonts and master text styles.
- `pptlive selection` — what the user has selected (resolves to `here:`).
- `pptlive find --text "Q3 revenue" [--in slide:3|shape:3:2|notes:3]` — fuzzy, smart-quote/whitespace-tolerant search across the deck (shapes, table cells, notes). Emits `[{anchor_id, start, length, text, context}]` in document order; empty array (exit 0) on no match.

## Writing — each command is one atomic undo
- `pptlive write --anchor-id ph:2:body --text "Intro\nDemo\nQ&A"` — set a text anchor (newlines = paragraphs).
- `pptlive replace --anchor-id shape:3:1 --text "New text"` — overwrite a whole anchor.
- `pptlive replace --find "old" --text "new" [--in slide:3] [--all|--occurrence N]` — fuzzy find/replace; rewrites just the matched span (keeps run formatting). One match auto-applies; several without `--all`/`--occurrence` is exit 5 (ambiguous, lists the matches); zero is exit 2.
- `pptlive insert --anchor-id para:4:2:3 --text "New bullet" [--before|--after]` — new paragraph relative to an anchor.
- `pptlive format-paragraph --anchor-id para:4:2:1 --alignment center --indent-level 2`.
- `pptlive format-text --anchor-id ph:4:title --bold --size 40 --color "#2E74B5"`.
- `pptlive list apply --anchor-id ph:4:body --type bulleted [--char "•"]` · `pptlive list remove --anchor-id ph:4:body`.

## Slides
- `pptlive slide layouts` — the layout names `add`/`set-layout` accept.
- `pptlive slide add --layout two_content [--index 4]`.
- `pptlive slide duplicate --slide 7` · `pptlive slide move --slide 9 --to 2` · `pptlive slide delete --slide 5`.
- `pptlive slide set-layout --slide 4 --layout title_and_content`.
- `pptlive slide export --slide 2 --out slide2.png [--width 1280] [--format png]` — render to image so a vision model can *see* it.

## Shapes
- `pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72` (points throughout; 1 in = 72 pt).
- `pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120`.
- `pptlive shape add --slide 4 --kind picture --path logo.png --left 600 --top 40 --alt-text "Acme logo"` (embedded, never linked).
- `pptlive shape move --anchor-id shape:4:3 --left 100 --top 140` · `pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200` · `pptlive shape delete --anchor-id shape:4:3`.
- `pptlive shape set-alt --anchor-id shape:4:3 --alt-text "Acme logo"` — alt text doubles as a drift-proof re-id handle.
- `pptlive shape export --anchor-id shape:4:3 --out logo.png` — render one shape (native size).

## Tables, charts, SmartArt
- Tables: `pptlive shape add --slide 4 --kind table --rows 3 --cols 2 --left 72 --top 120`; `pptlive table add-row --slide 4 --shape 5 --values '["Revenue","$4.2M"]'`; `pptlive table delete-row --slide 4 --shape 5 --row 2`; write cells with `pptlive write --anchor-id cell:4:5:1:1 --text "Metric"`.
- Charts (data lives in an embedded Excel workbook): `pptlive shape add --slide 4 --kind chart --chart-type column --categories "Q1,Q2,Q3" --series '{"Revenue":[10,20,30]}'`; `pptlive chart set-type --slide 4 --shape 5 --chart-type line`; `pptlive chart set-data --slide 4 --shape 5 --categories "A,B" --series '{"S":[1,2]}'`.
- SmartArt (content is a node tree): `pptlive shape add --slide 3 --kind smartart --smartart-kind process --nodes '["Discover","Design","Build","Ship"]'`; `pptlive smartart set-nodes --slide 3 --shape 2 --nodes '[{"text":"CEO","children":["Eng","Sales"]}]'`.

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

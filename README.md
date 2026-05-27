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
```

## Anchors

Addressing is **hierarchical** (slide → shape → text), not a global character
stream — there is no deck-wide `range:`. Anchor ids are colon-separated,
slide-index first:

| anchor_id      | resolves to |
| -------------- | ----------- |
| `shape:S:N`    | Nth shape (1-based z-order) on slide S — the canonical handle |
| `ph:S:KIND`    | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — the LLM-preferred form |
| `notes:S`      | speaker-notes body of slide S |

z-order **drifts** when shapes are added or removed, so `shape:S:N` is resolved
live and never cached; every shape listing also emits `name` (`Shape.Name`) and
`id` (`Shape.Id`, stable across reorder) so you can re-identify after drift.
Steer toward `ph:S:KIND` and `.Name` as the drift-proof forms.
(`para:S:N:P` and `cell:S:N:R:C` arrive in later stages.)

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

pptlive shape add --slide 4 --kind textbox --text "Revenue up 12%" --left 72 --top 72
pptlive shape add --slide 4 --kind shape --shape-type star --left 400 --top 120 --width 120 --height 120
pptlive shape add --slide 4 --kind picture --path logo.png --left 600 --top 40
pptlive shape move   --anchor-id shape:4:3 --left 100 --top 140
pptlive shape resize --anchor-id shape:4:3 --width 300 --height 200
pptlive shape delete --anchor-id shape:4:3

pptlive go-to --anchor-id shape:3:1              # deliberate, opt-in view move
```

Exit codes: `0` ok · `1` other · `2` anchor/slide/shape/presentation not found ·
`3` PowerPoint busy / slide show running · `4` PowerPoint not running · `5`
ambiguous match · `6` shape has no text frame.

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

---
name: pptlive-python
description: Read and edit the Microsoft PowerPoint presentation the user has open right now, from Python — `import pptlive as pl`. Attach to the running PowerPoint, read structure (slides, shapes, tables, charts, SmartArt, theme/master) as dataclasses/dicts, make polite edits inside `deck.edit("label")` blocks (each one atomic — a single Ctrl-Z), address content with hierarchical anchors, render slides/shapes to images, and drive a live slide show. Use when scripting live PowerPoint from Python on Windows.
---

# pptlive (Python API)

`pptlive` drives a **running** Microsoft PowerPoint instance over COM (Windows
only). Unlike `python-pptx` (which works the `.pptx` on disk), it edits the deck
the user has **open right now** — politely: their viewed slide and shape/text
selection are preserved, and every `deck.edit(...)` block collapses into a single
Ctrl-Z.

(For the command-line interface instead, run `pptlive llm-help`.)

## Attach and read

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active          # or ppt.presentations["Report.pptx"]

    # Reads — structured, side-effect-free, never move the view.
    slides  = deck.slides.list()             # [{index, id, layout, title, shape_count, has_notes}]
    outline = deck.outline()                 # [{slide, title, bullets:[...]}]
    grid    = deck.slides[2].read()          # every shape: anchor_id, name, id, type, geometry, text
    title   = deck.slides[2].placeholder("title").text
    notes   = deck.slides[1].notes.text
    layouts = deck.layouts()                 # [{index, name}] — what set_layout/add accept
```

`deck.anchor_by_id(...)` resolves any anchor id (see **Anchors** below) to an
object that carries the relevant verbs. Every wrapper also exposes a `.com`
escape hatch returning the raw COM object.

## Polite writes — one Ctrl-Z per `edit` block

Wrap **all** mutations in `with deck.edit("label"):`. PowerPoint has no
`UndoRecord`, but the block fences its COM edits into a single undo entry
(`StartNewUndoEntry`) and restores the user's selection on exit. There's no
explicit "end" fence, so always edit inside a block rather than bare.

```python
with deck.edit("Revise the agenda slide"):
    deck.anchor_by_id("ph:2:title").set_text("Agenda")
    deck.anchor_by_id("ph:2:body").set_text("Intro\nDemo\nQ&A")   # \n (or \r) = new paragraph; \v = soft break
```

## Find / replace — fuzzy, deck-wide

There is no deck-wide character stream, so search is a traversal of every text
frame (shapes, table cells, speaker notes). Matching is smart-quote / dash /
whitespace tolerant, so text you re-typed off a slide still matches the original
glyphs. `find` is a read; `find_replace` rewrites only the matched span (run
formatting survives) and belongs in a `deck.edit(...)` block.

```python
hits = deck.find("Q3 revenue")            # [{anchor_id, start, length, text, context}], doc order
hits = deck.find("Demo", scope="slide:2") # scope: "slide:S", any anchor id, a Slide, or an Anchor

with deck.edit("Rename the product"):
    deck.find_replace("Acme", "Globex", all=True)          # every occurrence
    deck.find_replace("teh", "the", occurrence=2)          # only the 2nd match
```

Zero matches raises `AnchorNotFoundError` (exit 2); several matches without
`all`/`occurrence` raises `AmbiguousMatchError` (exit 5, carrying the matches).

## Anchors

Addressing is **hierarchical** (slide → shape → text), slide-index first — no
deck-wide `range:`.

| anchor_id      | resolves to |
| -------------- | ----------- |
| `shape:S:N`    | Nth shape (1-based z-order) on slide S — the canonical handle |
| `shapeid:S:ID` | shape with stable `Shape.Id` ID on slide S — the **delete-proof** handle (`slide.shapes.by_id(ID)`) |
| `ph:S:KIND`    | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — **prefer this** |
| `para:S:N:P`   | paragraph P (1-based) of shape N on slide S |
| `cell:S:N:R:C` | cell (row R, col C) of the table in shape N on slide S — a `Cell` *is* an anchor |
| `notes:S`      | speaker-notes body of slide S |
| `comments:S`   | review comments on slide S — a **container** read via `slide.comments` (not an `Anchor`); one comment is `(slide, 1-based index)` |
| `here:`        | whatever the user has selected right now (opt-in) |

`body` also matches the generic **content** placeholder, which reads back with
`placeholder: "object"` (e.g. "Content Placeholder N"). On a **Two Content** /
**Comparison** layout there are *two* such placeholders, so `ph:S:body` is
ambiguous and raises an error listing the candidate `shape:S:N` anchors — address
each column by its `shape:S:N` (or `.Name`) instead.

z-order **drifts** as shapes are added, removed, *or restacked* (`reorder`), so
`shape:S:N` resolves live and is never cached; listings emit `name` (`Shape.Name`)
and `id` (`Shape.Id`, stable across reorder *and* delete) for re-identification.
Prefer `ph:S:KIND`, `.Name`, and `shapeid:S:ID` — the last keeps pointing at the
same shape across a delete/restack that would renumber `shape:S:N`.

## Slides, shapes, geometry (points throughout — `pl.units` for inches/cm)

```python
with deck.edit("Build the results slide"):
    new = deck.slides.add(layout="two_content", index=4)
    deck.slides[7].duplicate()                       # copy lands at slide 8
    deck.slides[9].move_to(2)
    deck.slides[4].set_layout("title_and_content")

    shapes = deck.slides[4].shapes
    shapes.add_textbox("Revenue up 12%", left=pl.units.inches(1), top=72)
    star = shapes.add_shape("star", left=400, top=120, width=120, height=120, fill="#1E74B5")
    logo = shapes.add_picture("logo.png", left=600, top=40, alt_text="Acme logo")
    deck.slides[4].shapes["Picture 3"].move(top=140)  # by name; absolute, points

    panel = shapes.add_shape("rectangle", left=60, top=60, width=840, height=400)
    panel.set_fill(fill="#102030", line="none")       # fill/border — NOT font color; "none" = transparent
    panel.reorder("back")                             # tuck the panel behind existing content
    star.delete()

logo.set_alt_text("Acme logo (top-right)")            # alt text = drift-proof re-id handle
chart_png = deck.slides[4].shapes["Chart 2"].export_image()   # one shape, native size
```

## Text structure, tables, charts, SmartArt

```python
with deck.edit("Polish the body copy"):
    body = deck.anchor_by_id("ph:4:body")
    body.set_text("Revenue up 12%\nChurn down 3%\nNPS +9")
    body.apply_list("bulleted")
    body.paragraph(2).format_paragraph(indent_level=2, alignment="left")
    body.paragraph(1).format_text(bold=True, size=24, color="#2E74B5")
    body.insert_paragraph_after("Cash runway: 30 months")

with deck.edit("Add a metrics table"):
    table = deck.slides[4].shapes.add_table(rows=3, columns=2).table
    table.cell(1, 1).set_text("Metric")
    table.add_row(["Revenue", "$4.2M"])               # appends + fills a row
    deck.anchor_by_id("cell:4:5:1:1").format_text(bold=True)
grid = table.read()                                   # {slide, shape, rows, columns, cells:[...]}

with deck.edit("Add a revenue chart"):
    chart = deck.slides[4].shapes.add_chart(
        "column", ["Q1", "Q2", "Q3"], {"Revenue": [10, 20, 30], "Profit": [3, 6, 9]}
    ).chart
    chart.set_type("line")
data = chart.read()                                   # {chart_type, categories, series:[...]}

with deck.edit("Add a process diagram"):
    sa = deck.slides[3].shapes.add_smartart("process", ["Discover", "Design", "Build"]).smartart
    sa.set_nodes([{"text": "CEO", "children": ["VP Eng", "VP Sales"]}])   # flat list or tree
tree = sa.read()                                      # {layout, nodes:[{text, level, children}]}

# Recolor composite text — a chart/SmartArt has no text anchor, so format_text
# can't reach it; recolor_text is the only color path. The coarse fix when the
# inherited black text goes invisible on a dark (or any custom) background.
with deck.edit("Make the diagram readable on dark"):
    chart.recolor_text("#FFFFFF")   # every shown chart element: legend/axes/title/data labels
    sa.recolor_text("#FFFFFF")      # every SmartArt node label
```

## Comments — threaded review (read + add/reply/delete)

```python
roll = deck.comments()                          # {total, slides:[{slide, comments:[...]}]}
for c in deck.slides[1].comments:               # per-slide; 1-based, iterable
    print(c.author, c.text, [r.text for r in c.replies])

with deck.edit("Respond to review"):
    c = deck.slides[2].comments.add("Please cite a source.")   # binds to the signed-in account
    deck.slides[1].comments[1].reply("Done.")                  # threaded reply
    deck.slides[1].comments[3].delete()                        # removes the comment + its replies
```

Comments attach to a slide at an `(x, y)` point and are threaded. `add` sources the
modern `Add2` identity off an existing comment, falling back to the legacy
identity-free add on a comment-less deck; the passed `author`/`initials` reach only
that fallback (`Add2` binds to the account). No `resolve()` — comment resolution
state isn't COM-readable.

## Deck-wide styling — theme + master (global, but still one Ctrl-Z)

```python
with deck.edit("Rebrand the deck"):
    deck.theme.set_color("accent1", "#C00000")        # recolor the whole deck
    deck.theme.set_font("major", "Georgia")           # major = headings, minor = body
    deck.master.format_text_style("body", 1, font="Georgia", size=28)
    deck.master.set_background("#1F1F1F")
palette = deck.theme.read()                           # {colors:{slot:#RRGGBB}, fonts:{major, minor}}
```

## Render, selection, slide show

```python
png = deck.slides[4].export_image(width=1280)         # one slide -> temp PNG (or pass a path); polite

# Whole-deck low-res snapshot — see every slide cheaply. max_dim caps each slide's
# long edge (a uniform, predictable per-slide token cost — a model is billed on
# pixel area, not DPI). slides=None (all) | int (one) | (start, end) inclusive.
for snap in deck.snapshot(max_dim=1000):              # -> [Snapshot(slide, image, path), ...]
    review(snap.image)                                # bytes; pass out="deck.png" to also write -sN files

# Save / export — explicit, never implicit (pptlive never auto-saves). save_as
# REBINDS the working file (the deck becomes that path); export_pdf is a READ
# (no rebind, dirty flag preserved — the "hand back a deliverable" path).
if not deck.saved:                                    # deck.saved / deck.path also ride on `status`
    deck.save()                                       # persist in place; UnsavedPresentationError if no path yet
deck.save_as("C:/out/v2.pptx", overwrite=False)       # write .pptx + rebind; refuses to clobber unless overwrite
deck.export_pdf("C:/out/deck.pdf")                    # pixel-faithful PDF; working .pptx untouched

sel = deck.selection()                                # {type, slide, anchor_id, shapes, ...}
if sel.anchor_id:
    with deck.edit("Bold the selected text"):
        deck.anchor_by_id("here:").format_text(bold=True)

deck.go_to(deck.anchor_by_id("shape:3:1"))            # deliberate, opt-in view move

deck.show.start()                                     # run from the top (moves the screen)
deck.show.goto(5); deck.show.next(); deck.show.black()
deck.show.state()                                     # {running, state, current_slide, ...}
deck.show.end()
```

## Errors

All failures raise a `pl.PptliveError` subclass (mirrors the CLI's exit codes):
`AnchorNotFoundError` (slide/shape/anchor missing; `SlideNotFoundError`
subclasses it) · `AmbiguousMatchError` · `NoTextFrameError` (shape holds no text)
· `UnsavedPresentationError` (`save()` on a never-saved deck — use `save_as(path)`)
· `PowerPointBusyError` (a modal dialog is open — back off and retry) ·
`PowerPointNotRunningError` · `PresentationNotFoundError`.

```python
try:
    deck.anchor_by_id("ph:99:title").set_text("…")
except pl.AnchorNotFoundError:
    ...   # re-read with deck.slides.list() / deck.slides[s].read()
```

Full docs: https://thomas-villani.github.io/pptlive/ · CLI: `pptlive llm-help`.

# Concepts

A handful of ideas drive almost every API decision in pptlive. If you
understand these, the rest of the surface follows. They are the same ideas as
its sibling [`wordlive`](https://github.com/thomas-villani/wordlive), re-applied
to PowerPoint's 2-D, multi-slide object model.

## Politeness model

The user is editing — or *presenting* — the same deck as your script. Naïve
automation jumps them to a different slide, stomps their shape selection, and
muddies their undo history. A jump is more jarring in PowerPoint than in Word:
it's a full-screen change of what's on the projector. pptlive refuses to do it
by default.

Every [`deck.edit()`](python-api.md#pptlive.EditScope) scope:

1. Snapshots the user's **viewed slide** (`ActiveWindow.View.Slide.SlideIndex`)
   and their `Selection` (the selected shapes by name, or a collapsed text
   caret).
2. Runs your mutations.
3. Restores the snapshot on the way out.

The captured snapshot is a plain dataclass:

```python
from pptlive import SelectionSnapshot

# Captured at the start of every deck.edit() block.
SelectionSnapshot(slide_index=3, selection_type=2, shape_names=("Title 1",))
```

If you genuinely want to move the user — say, jumping their view to a slide
after building it — opt in explicitly:

```python
with deck.edit("Add a results slide") as scope:
    new = deck.slides.add(layout="title_and_content")
    scope.allow_view_move()        # don't restore the viewed slide on exit
    deck.go_to(new.placeholder("title"))
```

Restoration is best-effort: shape selections round-trip by name; a text
selection collapses to "no selection" rather than guessing an offset that may
no longer exist. If the snapshotted slide was deleted inside the block,
pptlive declines to restore rather than raising.

!!! info "Implementation"
    The snapshot dataclass, plus the `snapshot()` / `restore()` helpers, live in
    [`src/pptlive/_selection.py`](https://github.com/thomas-villani/pptlive/blob/main/src/pptlive/_selection.py).

## Semantic anchors over `Selection`

The PowerPoint object model encourages you to drive everything through the
live `Selection` — the shapes or text the user has highlighted. That's hostile
to both humans (your script fights their clicking) and LLM agents (the
selection is invisible state that changes under you).

pptlive operates on **anchors** instead: addressable handles for slides,
shapes, placeholders, paragraphs, cells, and notes that don't depend on the
selection. Text is written straight through
`Shape.TextFrame.TextRange.Text`, so no edit ever needs to select anything.

A [`Shape`](python-api.md#pptlive.Shape) **is** an
[`Anchor`](python-api.md#pptlive.Anchor) when it has a text frame, and so are
[`Paragraph`](python-api.md#pptlive.Paragraph),
[`Cell`](python-api.md#pptlive.Cell), and
[`Notes`](python-api.md#pptlive.Notes). They share the same verbs:

```python
title = deck.anchor_by_id("ph:3:title")
title.text                          # read
title.set_text("Q3 Results")        # replace
title.format_text(bold=True, size=40, color="#2E74B5")
title.insert_paragraph_after("Subtitle line")
title.com                           # raw COM TextRange — escape hatch
```

In `set_text`, a `\n` (or `\r`) starts a **new paragraph** — each line becomes
its own addressable `para:S:N:P`. For a soft line break *within* a paragraph,
embed `\v` (`pptlive._anchors.SOFT_BREAK`).

Why not Selection-driven? Two reasons:

1. **Idempotent operations are easier to reason about.** "Set the title of
   slide 3 to X" is repeatable; "type X into whatever is selected" is not.
2. **LLM tool use needs stable identifiers.** `ph:3:title` is a stable string;
   the live selection is not.

## Anchor IDs

Word is a linear character stream, so its anchors reduce to a single global
offset. **PowerPoint is a 2-D canvas of discrete objects across an ordered set
of slides** — there is no document-wide character stream and no deck-wide
`range:`. So addressing is *hierarchical* (slide → shape → paragraph), and
anchor IDs are colon-separated with the slide index first:

```
shape:3:2          # 2nd shape (1-based z-order) on slide 3 — the canonical handle
shapeid:3:7        # shape with stable Shape.Id 7 on slide 3 — the delete-proof handle
ph:3:title         # placeholder of semantic KIND on slide 3 (the LLM-preferred form)
para:3:2:1         # paragraph 1 of shape 2 on slide 3
cell:3:5:1:2       # cell (row 1, col 2) of the table in shape 5 on slide 3
notes:3            # speaker-notes body of slide 3
here:              # whatever the user has selected right now (opt-in)
```

`ph:S:KIND` takes a semantic `KIND` of `title`, `ctrtitle`, `subtitle`,
`body`, `footer`, `date`, or `slidenum` — "the title of slide 3" without
caring about z-order. It's the form to prefer in tool-use payloads. `body`
also matches the generic **content** placeholder (which reads back as
`placeholder: "object"`, e.g. `"Content Placeholder 2"`). A **Two Content** /
**Comparison** layout has *two* such placeholders, so `ph:S:body` is ambiguous
and raises an error listing the candidate `shape:S:N` anchors — reach each
column by its `shape:S:N` or `.Name` instead.

The bare `slide:S` form is deliberately **not** an anchor — a whole slide has
no single text range, just like a whole table doesn't. Slide-level verbs
(add shape, set layout, edit notes, duplicate) live on a
[`Slide`](python-api.md#pptlive.Slide) object and the `slide` CLI group, and
you reach a slide with `deck.slides[S]`, not `anchor_by_id`.

`comments:S` is the same kind of read selector — a slide's review comments are a
*container*, not a text anchor. Read them via `deck.slides[S].comments` (or
`deck.comments()` deck-wide); a single comment is addressed for reply/delete by its
`(slide, 1-based index)`, not an anchor id. Comments attach to a slide at an
`(x, y)` point and are **threaded** (`comment.replies` / `comment.reply(text)`).
Adding binds to the signed-in Office account (the modern `Comments.Add2` identity
is sourced off an existing comment, falling back to the legacy identity-free add on
a comment-less deck); there is no resolve verb, since comment resolution state isn't
COM-readable on current builds.

### z-order drifts — design around it

`shape:S:N` uses the 1-based z-order index, which **shifts when shapes are
added, removed, or restacked** (`shape order` / `Shape.reorder`). pptlive
resolves `shape:S:N` *live* on every use and never caches it. To survive drift,
every shape listing also emits:

- `name` — `Shape.Name` (e.g. `"Title 1"`, `"Content Placeholder 2"`),
  usually unique per slide. Look one up with `slide.shapes["Title 1"]`.
- `id` — `Shape.Id`, **stable across reorder *and* delete** for re-identification.
  Address it directly with `shapeid:S:ID` (`slide.shapes.by_id(7)`) — the
  delete-proof handle: deleting a lower shape renumbers every `shape:S:N` above
  it, but the `shapeid` keeps pointing at the same shape.
- `alt_text` — `Shape.AlternativeText`, which you can *set* as a description
  and re-find a picture/diagram by even after drift.

Steer toward `ph:S:KIND`, `.Name`, and `shapeid:S:ID` as the drift-proof forms.
`para:S:N:P` and `cell:S:N:R:C` likewise resolve live, since the
paragraph/row count moves as text or rows are inserted.

These IDs are emitted directly by every read and consumed by
[`deck.anchor_by_id()`](python-api.md#pptlive.Presentation) and every CLI
command that takes `--anchor-id`:

```python
anchor = deck.anchor_by_id("ph:3:title")
anchor.set_text("Updated section title")
```

!!! info "Implementation"
    Resolution is centralised in
    [`Presentation.anchor_by_id`](python-api.md#pptlive.Presentation); see
    [`src/pptlive/_presentation.py`](https://github.com/thomas-villani/pptlive/blob/main/src/pptlive/_presentation.py).

## `EditScope` and atomic undo

`deck.edit("label")` returns an
[`EditScope`](python-api.md#pptlive.EditScope). PowerPoint has no
`Application.UndoRecord` (Word's start/end bracket), but a 2026 spike found it
doesn't need one: **PowerPoint groups consecutive COM edits made within one
automation session into a single undo entry by default**, and
`Application.StartNewUndoEntry()` is a verified *boundary* primitive. So the
scope calls `StartNewUndoEntry()` on entry to fence the block cleanly, and the
whole block reverts with one Ctrl-Z.

```python
with deck.edit("Lay out the results slide"):
    deck.anchor_by_id("ph:4:title").set_text("Q3 Results")
    deck.anchor_by_id("ph:4:body").set_text("Revenue up 12%\nChurn down 3%")
    deck.slides[4].shapes["Chart 2"].move(top=140)

# One Ctrl-Z reverts all three.
```

Two responsibilities are bundled into the same context manager:

1. **Undo fence** — `StartNewUndoEntry()` on entry.
2. **`SelectionSnapshot`** — see [Politeness](#politeness-model).

Two honest caveats:

- **There's no explicit "end" fence.** The block is closed by the *next*
  `edit()` (which re-fences) or by the user's next manual action. So always
  wrap mutations in `deck.edit(...)` rather than editing bare — that's the
  supported path, and it keeps each block cleanly self-contained.
- **Cross-process edits stay separate** (verified): two separate CLI
  invocations each re-fence at their own `edit()` entry, so one Ctrl-Z reverts
  only the most recent invocation's edit.

The scope object exposes one knob:

```python
with deck.edit("Build and reveal") as scope:
    new = deck.slides.add(layout="title_and_content")
    scope.allow_view_move()         # skip the viewed-slide restore
    deck.go_to(new.placeholder("title"))
```

Most code never touches the scope — just `with deck.edit("label"):` and write
your mutations.

!!! info "Implementation"
    [`EditScope`](python-api.md#pptlive.EditScope) lives in
    [`src/pptlive/_edit.py`](https://github.com/thomas-villani/pptlive/blob/main/src/pptlive/_edit.py).

## Geometry is first-class

Word almost never cares where a paragraph sits on the page; PowerPoint
authoring is fundamentally spatial. Every shape carries geometry in
**points** (1 inch = 72 pt — the unit PowerPoint's COM layer uses; EMUs are an
OOXML/`python-pptx` concern and never surface here):

```python
shape.geometry()                       # {left, top, width, height, rotation}
shape.move(left=72, top=120)           # absolute, points
shape.resize(width=300, height=200)
```

Use the [`pl.units`](python-api.md#pptlive.units) helpers rather than
hardcoding multiplications:

```python
import pptlive as pl
shape.move(left=pl.units.inches(1.5), top=pl.units.cm(4))
```

Slide dimensions come from
[`deck.page_setup()`](python-api.md#pptlive.Presentation)
(`SlideWidth` / `SlideHeight`), so an agent can place things relative to the
canvas.

Shapes also carry **fill**, **border**, and **stacking**:

```python
shape.set_fill(fill="#1E74B5", line="none")     # solid fill, no border
shape.set_fill(line="#333333", line_width=1.5)  # just the border
panel.reorder("back")                           # send behind existing content
```

`fill` / `line` take a color (`"#RRGGBB"`, an `(r, g, b)` tuple, or a raw RGB
int) or the string `"none"` (transparent fill / no border); `fill=` and `line=`
also ride on `shapes.add_shape(...)` / `add_textbox(...)`. This is *shape* color
— distinct from [`format_text`](python-api.md#pptlive.Anchor)'s font `color`.
[`reorder`](python-api.md#pptlive.Shape) (`"front"`/`"back"`/`"forward"`/
`"backward"`) is what lets a freshly added background band slide *behind* content
that's already there — otherwise every new shape lands on top. Every shape read
reports `fill` and `line` back (`{color, visible[, weight]}`), with a theme or
automatic color shown honestly as `color: null` rather than a misleading
`#000000`.

### Composite shapes have no text anchor — `recolor_text`

A chart and a SmartArt diagram are single shapes whose text is *internal* — split
across the legend / axes / title / data labels of a chart, or the node tree of a
diagram. There's no `para:`/`ph:` anchor reaching inside, so `format_text` can't
touch it. The coarse fix — the move a dark (or any custom-background) theme needs
when that inherited black text goes invisible — is `recolor_text`:

```python
chart = deck.slides[6].shapes["Chart 2"].chart
diagram = deck.slides[3].shapes["Diagram 1"].smartart
with deck.edit("Make the composites readable on dark"):
    chart.recolor_text("#FFFFFF")     # every shown element: legend, axes, title, data labels
    diagram.recolor_text("#FFFFFF")   # every node label
```

`Chart.recolor_text` recolors only text that's actually **shown** — guarded by
`HasLegend`/`HasTitle`, with axes and data labels best-effort (a pie chart's
absent axes are skipped, never an error) — so it never adds chrome the deck
didn't display. It returns `{..., recolored: [...], series_data_labels: N}` naming
what it touched; `SmartArt.recolor_text` returns `{..., nodes_recolored: N}`. This
is coarse "recolor all text to X" only; per-element targeting and composite-text
*fill* aren't covered — drop to [`.com`](#the-com-escape-hatch) for those.

## Seeing the deck — the token-cost snapshot

An agent that edits over COM wants to *see* the result. A single slide goes
through [`slide.export_image`](python-api.md#pptlive.Slide); the whole deck goes
through [`deck.snapshot`](python-api.md#pptlive.Presentation), which renders one
PNG per slide so a vision model can review every slide at once.

The key knob is **`max_dim`** — a long-edge pixel cap. A vision model is billed on
an image's pixel *area*, not its DPI, and that area depends on geometry — so
capping the long edge gives a *predictable* per-image token budget. And because
every slide in a deck shares one geometry, a single `max_dim` makes that budget
*uniform* across the deck: render the whole thing at `max_dim=1000` and each
slide costs about the same, legible-but-cheap amount. `max_dim` only ever *lowers*
resolution (it's capped at the slide's native 96-DPI size), and `max_dim=None`
renders native.

```python
# "Did my restyle land across every slide?" — cheap whole-deck review.
for snap in deck.snapshot(max_dim=1000):
    show_to_vision_model(snap.image)      # snap.slide, snap.image (bytes), snap.path
```

When you need an exact per-slide size rather than the long-edge cap, pass
`width=` / `height=` instead of `max_dim` (they override it). Pixel area — not
encoder quality — is what a vision model bills on, and `Slide.Export` exposes no
JPEG-quality knob, so the dimensions are the only render-cost lever there is.

Snapshotting is a **read**: the export reflects the current unsaved state but
leaves the viewed slide and Selection untouched, so it needs no `edit()` fence.
(Not to be confused with the *Selection* snapshot above, which `EditScope` uses to
restore the user's view.)

## Saving is explicit — and export is a read

pptlive **never auto-saves**. Every edit lands in the *live* deck (that's the
whole point — it's the file the user has open), but committing those edits to disk
only happens when you ask: `deck.save()`, `deck.save_as(path)`. This is the same
politeness instinct as the view model — the tool doesn't take an irreversible,
outward-facing action (writing the user's file) as a side effect of editing.

Two distinctions matter:

- **`save_as` rebinds; `export_pdf` does not.** `save_as(path)` makes the open
  deck *become* the new file (its `name`/`path` follow), exactly like PowerPoint's
  Save-As — so the user's working file is now `path`. `export_pdf(path)`, by
  contrast, is a **read**: it writes a pixel-faithful PDF but leaves the working
  file, its path, and its dirty flag untouched. Reach for `export_pdf` to "hand
  back a deliverable" without touching what the user is editing.
- **`save()` refuses a path-less deck.** A brand-new deck has no file yet. Calling
  `save()` on it raises `UnsavedPresentationError` rather than guessing a location
  — because PowerPoint's own `Save()` *doesn't* refuse: on a OneDrive/SharePoint
  build it silently uploads to the user's default cloud folder, which is exactly
  the kind of surprise outward action the guard exists to prevent. Use
  `save_as(path)` to choose the destination explicitly.

`deck.saved` (and `deck.path`) surface on every `status` deck row, so an agent can
see there's unsaved work before deciding what to do about it.

## The `.com` escape hatch

pptlive deliberately covers a small surface. When you need something it
doesn't, every wrapper exposes the raw COM object via `.com`:

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active

    # Anything pptlive covers, use the pptlive API.
    with deck.edit("Rotate the logo"):
        # Anything it doesn't, drop to COM.
        deck.com.Slides(3).Shapes("Logo").Rotation = 15
```

`ppt.com`, `deck.com`, `slide.com`, `shape.com`, and `anchor.com` all return
the underlying pywin32 dispatch object. Treat this as a forward-compatibility
seam: as pptlive grows, today's COM call may become tomorrow's high-level
helper, but the escape hatch is permanent.

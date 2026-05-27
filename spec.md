# pptlive — Spec (working draft)

> Working name: `pptlive`. Final name TBD (see Open Questions).
> Status: sketch — the PowerPoint sibling of [`wordlive`](./spec.md), written to
> seed a separate repo. Reuses wordlive's anchors-over-`Selection`,
> politeness-first, structured-I/O, LLM-first-CLI design wholesale, and diverges
> only where PowerPoint's object model forces it to. Read `spec.md` first; this
> document is mostly *the diff*.

## Overview

A small Python library + CLI for **driving a running Microsoft PowerPoint
instance** from Python, designed for both human scripting and LLM agents.
Sibling to (not part of) `python-pptx`:

| Library       | Target                                  | Mechanism                |
| ------------- | --------------------------------------- | ------------------------ |
| python-pptx   | `.pptx` file on disk                    | OOXML I/O                |
| **pptlive**   | **Running POWERPNT.exe**                | **COM automation (pywin32)** |

Windows-only by nature. Use it when the user already has the deck open and you
want to edit it *live* — or when you want an LLM agent to build/revise slides
alongside a human inside the same PowerPoint session.

## Motivation

The thing that prompted this: today, an agent editing a presentation works the
file with `python-pptx`, the user has to **close the deck, let the agent write,
then re-open to see what changed**, then close again for the next edit. That
round-trip is slow, error-prone (PowerPoint holds the file lock; a save while
the agent is mid-write corrupts or conflicts), and — the user's word —
*unintuitive*. You can't watch a slide take shape; you can't leave the deck open
on a second monitor and narrate changes.

The same four reasons that justify wordlive justify pptlive:

1. **xlwings exists; nothing equivalent for PowerPoint does.** Driving
   PowerPoint from Python today means raw pywin32: no type hints, magic integer
   constants (`MsoShapeType`, `PpPlaceholderType`, `PpSlideLayout`), late-bound
   string lookups, STA threading footguns.
2. **LLM agents need a small, semantic surface.** The full PowerPoint object
   model is hopeless for a model. "Set the title of slide 3", "add a bullet to
   the body of the agenda slide", "move the chart down 40 points" is what tool
   use wants.
3. **File-level libraries can't help when the user is presenting or editing.**
   If PowerPoint has the deck open, `python-pptx` can't safely touch it on disk
   — and the user sees nothing until they re-open. COM is the only live path.
4. **"Polite" editing is a real engineering problem.** Naïve scripts jump the
   user to a different slide, stomp their shape selection, and steal focus. A
   wrapper that handles this once is broadly valuable.

## Non-Goals

- **Cross-platform.** COM is Windows-only; we don't pretend otherwise.
- **Cloud co-authoring / PowerPoint on the web.** That's Microsoft Graph — a
  different stack.
- **Full PowerPoint object-model coverage.** We expose what common slide edits
  need; raw COM stays accessible via `.com` as an escape hatch.
- **Replacing `python-pptx`.** Different surface (live app vs. file), different
  problem. Generate decks from scratch in a batch job? Use `python-pptx`. Edit
  the deck the user is looking at? Use pptlive.
- **Slide rendering / thumbnail export as a core feature.** `Slide.Export(...)`
  to PNG is one `.com` call away; we may surface it later but it isn't the point.
- **Designer / AI layout suggestions.** Out of scope.

## Design Principles

Inherited from wordlive verbatim, with one principle weakened by PowerPoint's
object model (see #3 — read it, it's the big one):

1. **Politeness first.** Default behaviour preserves the user's **viewed slide**,
   shape/text `Selection`, and focus. Operations that *must* change what the
   user is looking at say so in their name (`go_to`, `allow_view_move()`).
   Changing the active slide is the PowerPoint equivalent of stomping the
   cursor — and more jarring, because it's a full-screen jump.
2. **Semantic anchors over `Selection`.** Operations target slides, shapes,
   placeholders, table cells, or notes — never the live `Selection` unless
   explicitly requested. Text is set through `Shape.TextFrame.TextRange.Text`
   directly, so no edit needs to select anything.
3. **Atomic undo — available, via a different mechanism than Word.** Word has
   `Application.UndoRecord` (`StartCustomRecord` / `EndCustomRecord`, added in
   Word 2010). **PowerPoint has no equivalent bracket** — but the 2026-05-26
   spike (`scripts/undo_test.py`) found it doesn't need one: PowerPoint **groups
   consecutive COM edits made within one automation session into a single undo
   entry by default**, and `Application.StartNewUndoEntry()` is a verified
   *boundary* primitive (it starts a fresh entry; subsequent edits accumulate
   into it). So pptlive's `edit()` scope is **both** a view/Selection-preservation
   scope **and** an atomic-undo scope: it calls `StartNewUndoEntry()` on entry to
   fence the block, and the whole block reverts with one Ctrl-Z — near-parity with
   wordlive. Two honest caveats: there is no explicit "end" fence (the block is
   closed by the next `edit()` or by the user's next manual action), so always
   wrap mutations in `deck.edit(...)`. (Cross-*process* edits — separate CLI
   invocations — were also verified to stay distinct undo entries, since each
   invocation re-fences at its own `edit()` entry.)
4. **Structured I/O.** Reads return dataclasses/dicts; CLI emits one JSON object
   per invocation; exit codes are deterministic. No string scraping.
5. **Escape hatch always available.** Every wrapper exposes `.com` for the raw
   COM object (`Presentation`, `Slide`, `Shape`, `TextRange`).
6. **Synchronous core, optional event hooks.** COM is STA; we don't fight it.

## The anchor model — where PowerPoint forces a redesign

This is the heart of the diff, so it gets its own section.

**Word is a linear character stream.** Its anchors all reduce to a `Range` over
absolute UTF-16 offsets: `para:N`, `heading:N`, `bookmark:NAME`, and the generic
`range:START-END` that `find()` emits and `replace` consumes. One global
coordinate space.

**PowerPoint is a 2-D canvas of discrete objects across an ordered set of
slides.** There is no document-wide character stream — text lives inside a
`TextRange`, inside a `Shape`, inside a `Slide`. So addressing is *hierarchical*
(slide → shape → paragraph) rather than a single offset, and there is no
`range:START-END` analog at the deck level. Offsets are only meaningful *within
one shape's text frame*.

### Anchor-id grammar

Colon-separated, slide index first, mirroring how wordlive extended `bookmark:X`
to the 3-part `table:N:R:C` for cells:

| anchor_id            | Resolves to                                                        | Notes |
| -------------------- | ------------------------------------------------------------------ | ----- |
| `slide:S`            | Slide S (1-based)                                                  | A **container**, not a text anchor — like `table:N` in wordlive. Addressed via `doc.slides[S]` / the `slide` CLI group, **not** `anchor_by_id`. |
| `shape:S:N`          | Nth shape (1-based z-order) on slide S                             | The canonical, machine-stable handle. An `Anchor` if it has a text frame. |
| `ph:S:KIND`          | The placeholder of semantic KIND on slide S                       | KIND ∈ `title`, `ctrtitle`, `subtitle`, `body`, `footer`, `date`, `slidenum`. **The LLM-preferred form** — "the title of slide 3" without caring about z-order. Resolves to the underlying shape. |
| `para:S:N:P`         | Paragraph P (1-based) in shape N on slide S                        | An `Anchor` over that paragraph's `TextRange`. The `para:N` analog, two levels deeper. |
| `cell:S:N:R:C`       | Cell (row R, col C) of the table in shape N on slide S            | Shape N must satisfy `HasTable`. The `table:N:R:C` analog. |
| `notes:S`            | The speaker-notes text of slide S                                  | An `Anchor` over the notes-page body placeholder's `TextRange`. |

Design decisions (mirroring how wordlive resolved the same tensions):

- **z-order int vs. name.** `shape:S:N` uses the 1-based z-order index as the
  canonical id (like `heading:N` is index-canonical even though you can look up
  by text). Shape `.Name` ("Title 1", "Content Placeholder 2") is usually
  unique-per-slide and more durable than z-order, so the Python API *and* every
  listing expose it (`slide.shapes["Title 1"]`), and `ph:S:KIND` covers the
  common placeholders semantically. We keep the `shape:` scheme int-only to
  avoid the "is the last token an index or a name?" ambiguity wordlive
  deliberately designed out of `table:`.
- **`slide:S` is not an `Anchor`.** A whole slide has no single text range, just
  like a whole table doesn't. Slide-level verbs (set layout, add shape, edit
  notes, duplicate) live on a `Slide` object and the `slide` CLI group, exactly
  as table-level verbs (`add_row`) live on `Table` / the `table` group.
- **No deck-wide `range:`.** `find()` returns hits as `para:S:N:P` plus an
  in-shape character offset, not a global `range:`. A within-shape
  `range:S:N:START-END` form is **deferred** until a real mid-paragraph-edit
  workflow needs it (same reasoning that kept wordlive's inline insert deferred).

### Geometry is first-class (it isn't in Word)

Word almost never cares where a paragraph sits on the page; PowerPoint authoring
is *fundamentally* spatial. Every shape anchor therefore carries geometry, in
**points** (1 inch = 72 pt — the unit PowerPoint's COM layer uses; EMUs are an
OOXML/`python-pptx` concern and never surface here):

```python
shape.geometry()                       # {left, top, width, height, rotation}
shape.move(left=72, top=120)           # absolute, points
shape.resize(width=300, height=200)
```

Slide dimensions come from `doc.page_setup()` (`SlideWidth`/`SlideHeight`), so an
agent can place things relative to the canvas.

## Architecture

### Attachment model

```python
import pptlive

with pptlive.attach() as ppt:           # GetActiveObject("PowerPoint.Application")
    deck = ppt.presentations.active     # or ppt.presentations["Pitch.pptx"]
    ...

with pptlive.connect(launch_if_missing=True) as ppt:
    ...
```

- `GetActiveObject("PowerPoint.Application")` first; falls back to `Dispatch`
  when `launch_if_missing=True`.
- Context manager handles `CoInitialize` / `CoUninitialize`; releases COM refs
  on exit; **never** closes PowerPoint (it's the user's app).
- **`Visible` caveat.** Unlike Word/Excel, PowerPoint historically refuses to
  run invisibly — `Application.Visible = False` raises in most builds. So
  `connect()` has no `visible=False` mode; the app is always shown. *(Spike
  item — confirm current behaviour.)*

### Threading

Same as wordlive: STA, `CoInitialize` on context entry; event callbacks (if/when
added) pumped on a worker thread and marshalled to a callback registry.

### Error taxonomy

| Exception                  | Meaning                                                       | Exit |
| -------------------------- | ------------------------------------------------------------- | ---- |
| `PptliveError`             | base class                                                    | —    |
| `PowerPointNotRunningError`| no instance, `launch_if_missing=False`                        | 4    |
| `PresentationNotFoundError`| named or active presentation missing                          | 2    |
| `AnchorNotFoundError`      | shape / placeholder / cell / notes anchor absent; zero `find` matches | 2 |
| `SlideNotFoundError`       | slide index out of range (subclass of `AnchorNotFoundError`)  | 2    |
| `NoTextFrameError`         | text op on a shape with no text frame (e.g. a picture/line)   | 6    |
| `AmbiguousMatchError`      | fuzzy `find_replace` matched >1 without disambiguation        | 5    |
| `PowerPointBusyError`      | RPC rejected — modal dialog, or a slide show is running       | 3    |
| `ComError`                 | generic wrap of `pywintypes.com_error` with decoded HRESULT   | 1    |

Reuses wordlive's `_decode_com_error` / `from_com_error` and `_BUSY_HRESULTS`
set verbatim. `NoTextFrameError` is the one genuinely new code — it's common
enough (an LLM tries to set text on a decorative shape) to deserve a
deterministic exit code instead of a bare COM failure.

## Python API Sketch

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active

    # --- READS (structured, side-effect-free) ---
    slides   = deck.slides.list()          # [{index, layout, title, shape_count, has_notes}]
    outline  = deck.outline()              # [{slide, title, bullets:[...]}] — the Outline-view analog
    grid     = deck.slides[3].read()       # every shape on slide 3 (anchor_id, name, type, geometry, text)
    title    = deck.slides[3].placeholder("title").text
    notes    = deck.slides[3].notes.text

    # --- ANCHORS (cheap, lazy) ---
    body  = deck.anchor_by_id("ph:3:body")     # placeholder by semantic kind
    chart = deck.slides[3].shapes["Chart 2"]   # shape by name
    cell  = deck.anchor_by_id("cell:5:1:2:3")  # table cell

    # --- POLITE WRITES (preserves viewed slide + Selection) ---
    with deck.edit("Revise agenda slide"):     # view/Selection scope + one Ctrl-Z (see principles)
        deck.anchor_by_id("ph:2:title").set_text("Agenda")
        deck.anchor_by_id("ph:2:body").set_text("Intro\nDemo\nQ&A")
        chart.move(top=140)

    # --- SLIDE LIFECYCLE (no Word analog) ---
    deck.slides.add(layout="title_and_content", index=4)
    deck.slides[7].duplicate()
    deck.slides[9].move_to(2)
    deck.slides[4].set_layout("two_content")

    # --- LIVE SLIDE SHOW CONTROL (the most literally "live" surface) ---
    deck.show.start()
    deck.show.goto(5); deck.show.next(); deck.show.black()

    # --- EXPLICITLY MOVE THE USER (rare, opt-in) ---
    deck.go_to("shape:3:2")                     # jump the user's view to a shape

    # --- RAW ESCAPE HATCH ---
    deck.com.Slides(3).Shapes(2).TextFrame.TextRange.Font.Bold = True
```

### Key abstractions

- **`PowerPoint`** — application handle (`attach` / `connect`).
- **`Presentation`** — wraps a `Presentation` COM object; exposes `slides`,
  `outline()`, `page_setup()`, `find()` / `find_replace()`, `anchor_by_id()`,
  `edit()`, `show`, `go_to()`.
- **`SlideCollection` / `Slide`** — `slides[S]`, iteration, `add` / `delete` /
  `duplicate` / `move_to` / `set_layout`; a `Slide` exposes `shapes`,
  `placeholder(kind)`, `notes`, `read()`.
- **`ShapeCollection` / `Shape`** — `shapes[N]` (z-order) or `shapes["Name"]`;
  `add_textbox` / `add_picture` / `add_table` / `add_shape`. A `Shape` **is** an
  `Anchor` when it has a text frame (inherits `text` / `set_text` /
  `insert_paragraph_*`), plus geometry verbs (`move`, `resize`, `geometry()`).
- **`Anchor`** — base type for text-bearing handles: `Shape` (with text frame),
  `Paragraph`, `Cell`, `Notes`. Same verb set as wordlive (`text`, `set_text`,
  `apply_style`, `format_paragraph`, list verbs), minus anything Word-only.
- **`EditScope`** — context manager from `deck.edit(label)`. Snapshots and
  restores the **viewed slide index** + `Selection`, and fences a single undo
  entry via `StartNewUndoEntry()` on entry (no Word-style `UndoRecord` exists,
  but PowerPoint groups the in-session block into one Ctrl-Z). `allow_view_move()`
  opts out of the view-restore, the way `allow_cursor_move()` does in wordlive.
- **`SlideShow`** — `deck.show`: `start`, `end`, `next`, `previous`, `goto(n)`,
  `black()`, `white()`, `state()`. Wraps `SlideShowSettings.Run()` /
  `SlideShowWindow.View`.

## CLI Sketch

LLM-first: JSON in, JSON out, deterministic exit codes. `--doc NAME` selects a
presentation; default is the active one. `--json` (default) or `--text`.

```
pptlive status                                  # open decks, active one, current slide in view
pptlive slides                                  # [{index, layout, title, shape_count, has_notes}]
pptlive outline                                 # title + body bullets per slide (Outline-view analog)
pptlive slide read S                            # every shape on slide S: anchor_id, name, type, geometry, text
pptlive shapes --slide S                        # list shapes on slide S

pptlive read --anchor-id ph:3:title             # read any text anchor
pptlive read notes --slide 3
pptlive write --anchor-id ph:2:body --text "..."   # set_text on a text anchor
pptlive replace --anchor-id shape:3:2 --text "..."
pptlive replace --find OLD --text NEW [--in slide:3|shape:3:2] [--all|--occurrence N]
pptlive find --text "..." [--in slide:3]        # hits as para:S:N:P + in-shape offset
pptlive insert --anchor-id ph:2:body --text "..." [--before|--after] [--style "..."]

pptlive slide add [--layout two_content] [--index 4]
pptlive slide delete --slide S
pptlive slide duplicate --slide S
pptlive slide move --slide S --to N
pptlive slide set-layout --slide S --layout title_and_content

pptlive shape add --slide S --kind textbox|picture|table|shape [geometry + payload opts]
pptlive shape move   --anchor-id shape:S:N [--left N] [--top N]
pptlive shape resize --anchor-id shape:S:N [--width N] [--height N]
pptlive shape delete --anchor-id shape:S:N

pptlive table read --slide S --shape N          # table cells as cell:S:N:R:C anchors
pptlive style apply --anchor-id ph:2:body --name "..."
pptlive format-paragraph --anchor-id para:2:1:3 [--alignment center] [--space-after 6] ...

pptlive show start|end|next|prev|black|white|state
pptlive show goto --slide S

pptlive go-to --anchor-id shape:3:2             # deliberate, opt-in view move
pptlive exec --script ops.json                  # batch ops, applied as one Ctrl-Z (see below)
```

Conventions (same contract as wordlive):

- Exit codes: `0` ok, `2` anchor/slide/shape/presentation not found (incl. zero
  `find` matches), `3` PowerPoint-busy / show running, `4` PowerPoint-not-running,
  `5` ambiguous match, `6` shape-has-no-text-frame, `1` other.
- One JSON object on stdout per invocation; logs to stderr.

Wiring it as an LLM tool stays trivial:

```json
{ "tool": "pptlive", "args": ["write", "--anchor-id", "ph:3:title", "--text", "Q3 Results"] }
```

### `exec` ops

Same batch-script shape as wordlive (`{"label": ..., "ops": [...]}`). Because an
`exec` run is a single automation session, **the whole batch collapses to one
undo entry** — the run fences it with `StartNewUndoEntry`, then PowerPoint groups
the rest, so a 5-op script is one Ctrl-Z, not five. That matches wordlive's atomic
batch even without an `UndoRecord`. Failure semantics match wordlive: if op K
fails, ops 1..K-1 are already applied (and sit in that one undo entry, so a single
Ctrl-Z reverts the partial batch); we report the failing index and re-raise so the
exit code maps correctly. There is no `"tracked": true` key (PowerPoint has no
Track Changes).

Op kinds (slide/shape-shaped rather than paragraph-shaped):

```json
{
  "label": "Build the Q3 results slide",
  "ops": [
    {"op": "add_slide", "layout": "title_and_content", "index": 4},
    {"op": "set_text", "anchor_id": "ph:4:title", "text": "Q3 Results"},
    {"op": "set_text", "anchor_id": "ph:4:body",  "text": "Revenue up 12%\nChurn down 3%"},
    {"op": "add_shape", "slide": 4, "kind": "picture", "path": "chart.png",
       "left": 400, "top": 120, "width": 300, "height": 200},
    {"op": "set_notes", "slide": 4, "text": "Lead with the revenue number."},
    {"op": "move_shape", "anchor_id": "shape:4:3", "top": 140}
  ]
}
```

Proposed op set: `add_slide`, `delete_slide`, `duplicate_slide`, `move_slide`,
`set_layout`, `set_text`, `insert_paragraph`, `find_replace`, `set_notes`,
`add_shape`, `move_shape`, `resize_shape`, `delete_shape`, `set_cell`,
`apply_style`, `format_paragraph`, `apply_list`.

## Key Technical Concerns

- **Undo — resolved by the v0 spike, not the limitation we feared.** We expected
  no `UndoRecord` to mean per-op undo. The 2026-05-26 spike (`scripts/undo_test.py`)
  showed otherwise: PowerPoint groups consecutive in-session COM edits into one
  undo entry by default, and `StartNewUndoEntry()` is a boundary primitive. So
  `edit()` fences each block as one Ctrl-Z and an `exec` batch is one undo entry —
  no snapshot-by-duplicate or save-point machinery needed. Two residual concerns:
  1. **No explicit "end" fence.** The block is closed only by the next `edit()`
     (which re-fences) or a user action — so mutations made *outside* an `edit()`
     block can bleed into an adjacent entry. The library always wraps writes in
     `edit()`; document that bare `set_text` is the unsupported path.
  2. **Cross-process isolation — verified.** Two separate CLI-style invocations
     (each its own `attach()` + `edit()` fence) edited the same slide; one Ctrl-Z
     reverted only the second invocation's edit, so separate invocations stay
     distinct undo entries — each re-fences at its own `edit()` entry.
- **View / Selection preservation.** Snapshot `ActiveWindow.View.Slide.SlideIndex`
  (the slide the user is looking at in Normal view) and `ActiveWindow.Selection`
  (type + shape/text range) on enter; restore on exit unless `allow_view_move()`
  was called. Restoring a `Selection` is fiddlier than Word's single cursor —
  `Selection` can be `ppSelectionNone/Slides/Shapes/Text`. *(Spike: confirm a
  shape-range Selection round-trips cleanly; if not, restore only the viewed
  slide and `Unselect()` the rest.)*
- **PowerPoint must be visible.** See the attachment-model caveat; affects how
  "polite" the tool can be (you cannot do work in a hidden window).
- **Modal-dialog / slideshow busy state.** COM calls fail with the same
  `RPC_E_*` HRESULTs wordlive already classifies as busy. Additionally, while a
  slide show is running, most editing calls reject — surface as
  `PowerPointBusyError` (retryable, exit 3), and steer agents to the `show`
  group for live-presentation control instead.
- **`Slides.Add` vs. `Slides.AddSlide`.** Legacy `Add(Index, Layout)` takes a
  `PpSlideLayout` enum (limited, deprecated); modern `AddSlide(Index,
  CustomLayout)` needs a `CustomLayout` object pulled from the deck's
  design/master. pptlive's `slides.add(layout="...")` maps friendly names →
  the right `CustomLayout` from `Presentation.SlideMaster.CustomLayouts`,
  falling back to `Add` only if needed. *(Spike: layout-name → CustomLayout
  resolution across templates with renamed layouts.)*
- **Notes access.** Speaker notes live in the notes-page body placeholder
  (`Slide.NotesPage.Shapes.Placeholders` — the body, conventionally index 2).
  *(Spike: confirm the placeholder index/type is stable across templates;
  resolve by `PlaceholderFormat.Type == ppPlaceholderBody`, not a hard index.)*
- **Magic constants.** Ship a typed enum module like wordlive's `constants.py`:
  `MsoShapeType`, `MsoAutoShapeType`, `PpPlaceholderType`, `PpSlideLayout`,
  `PpSelectionType`, `PpSlideShowState`, `MsoTextOrientation`. Friendly string
  aliases (`"title"`, `"two_content"`, `"textbox"`) coerce to the right int the
  way wordlive's alignment names do.
- **Units.** Points throughout the COM layer (1 in = 72 pt). Offer
  `pl.units.inches(1.5)` / `pl.units.cm(4)` helpers so agents needn't hardcode
  multiplications. Never expose EMUs.

## Roadmap (suggested staging)

Ordered by **LLM-agent leverage**, mirroring `feature-plan.md`.

- **v0 — close the live-edit loop.** `attach` / `connect` / context manager;
  `slides` / `outline` / `slide read` reads; `Shape`-as-`Anchor` with
  `text` / `set_text`; `ph:S:KIND` + `shape:S:N` + `notes:S` resolution;
  `edit()` as a view/Selection scope + atomic-undo fence (`StartNewUndoEntry`);
  CLI `status`, `slides`, `outline`, `read`, `write`, `replace`; typed exceptions;
  constants enum. **Spike first (done 2026-05-26):** UndoRecord absence (→ found
  default in-session grouping + `StartNewUndoEntry` boundary instead), notes
  placeholder resolution, Selection round-trip, `Visible` behaviour — all
  confirmed via `scripts/spike.py` + `scripts/undo_test.py`.
- **v0.1 — slide lifecycle.** `slide add / delete / duplicate / move / set-layout`;
  layout-name → `CustomLayout` mapping. The first genuinely no-Word-analog track.
- **v0.2 — shapes & geometry.** `shape add` (textbox/picture/shape), `move` /
  `resize` / `delete`, `geometry()`; the spatial surface that makes
  slide-building possible.
- **v0.3 — text structure.** `para:S:N:P` anchors, `insert` before/after in a
  text frame, `apply_style`, `format_paragraph`, list verbs (PowerPoint bullets
  via `TextRange.ParagraphFormat`/`IndentLevel` — its own fiddly lift, like
  wordlive's lists).
- **v0.4 — tables.** `add_table`, `cell:S:N:R:C` anchors, `table read`. Cells
  are `Anchor`s, exactly like wordlive.
- **v0.5 — live slide show control.** The `show` group (`start/end/next/prev/
  goto/black/white/state`). The most literally "live" capability and the one
  with the cleanest payoff for a presenter-assistant agent.
- **v0.6 — pictures & charts.** `add_picture` (embed, never link), alt text as
  the LLM-readable re-identification handle (carried over from wordlive v0.8);
  `add_chart` with an embedded-Excel data spike (wordlive v0.10 reasoning
  applies); image **extraction** for vision models via the per-shape
  approach (wordlive v0.9).
- **v0.7+ — defer.** Event sinks (`SlideShowNextSlide`, `WindowSelectionChange`),
  async wrapper, slide/thumbnail export, transitions & animations, master/layout
  authoring.

## Open Questions

1. **Name.** Working name `pptlive` (direct wordlive parallel). Alternatives:
   `pptwings` (xlwings echo), `slidelive`, `decklive`, `livepptx`. Decide before
   first commit — affects package, import, and CLI binary names.
2. ~~**The undo gap.** Which mitigation becomes the documented default?~~
   **Settled by the v0 spike (2026-05-26): no mitigation needed.** PowerPoint
   groups in-session COM edits into one undo entry by default and
   `StartNewUndoEntry()` fences boundaries, so `edit()` and `exec` are already
   atomic (one Ctrl-Z). Cross-process isolation (separate CLI calls = separate
   entries) is also verified — see Key Technical Concerns.
3. **`shape:` addressing stability.** z-order index is canonical but shifts when
   shapes are added/removed mid-batch. Do listings also emit a stabler handle
   (shape `.Name`, or a synthesized id), and should `exec` re-resolve by name
   between ops? Word's `para:N` had the same drift; PowerPoint's is worse because
   `add_shape` is common.
4. **Multi-presentation scope.** First-class multi-deck support, or
   single-active-deck with explicit `--doc` naming otherwise (wordlive's
   unresolved Q4)?
5. **Slide identity.** Index (`slide:3`) is what users say, but PowerPoint also
   has stable `SlideID`. Expose both? Re-resolve `exec` slide refs by `SlideID`
   to survive reordering mid-batch?
6. **Test strategy.** COM tests need PowerPoint installed. Same answer as
   wordlive: a small smoke suite on a Windows runner with PowerPoint, plus a
   mockable wrapper layer for unit-testing the politeness/anchor logic against a
   `fake_powerpoint` fixture.

## Inspirations / Prior Art

- **wordlive** — the direct sibling; this spec is its diff. Reuse anchors,
  `EditScope` shape, error decoding, CLI contract, and `exec` design.
- **xlwings** — API ergonomics for COM-driven Office automation.
- **python-pptx** — the file-side counterpart (the python-docx analog). Shape
  abstractions (shapes, placeholders, paragraphs) similarly so users moving
  between them feel at home.
- **"Talk to Your Slides"** (arXiv 2505.11604) — prior art on the
  live-COM-plus-LLM combination for slide editing; worth reading for what an
  agent actually wants to do once the surface exists.
```

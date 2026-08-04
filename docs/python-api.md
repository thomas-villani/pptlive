# Python API

Every entry on this page is generated from the docstrings in the
[`pptlive`](https://github.com/thomas-villani/pptlive/tree/main/src/pptlive)
package, so it stays in sync with the code. If something looks thin, the fix
is in the source docstring, not here.

The public surface is small on purpose. Three rough layers:

- **Connect** — [`attach`](#pptlive.attach) / [`connect`](#pptlive.connect)
  return a [`PowerPoint`](#pptlive.PowerPoint) handle.
- **Address** — [`Presentation`](#pptlive.Presentation) exposes
  [`slides`](#pptlive.SlideCollection), each [`Slide`](#pptlive.Slide) exposes
  [`shapes`](#pptlive.ShapeCollection), and
  [`anchor_by_id`](#pptlive.Presentation) resolves the hierarchical anchor
  scheme (`shape:S:N`, `ph:S:KIND`, `para:S:N:P`, `cell:S:N:R:C`, `notes:S`,
  `here:`).
- **Mutate** — wrap writes in [`Presentation.edit()`](#pptlive.Presentation) →
  [`EditScope`](#pptlive.EditScope) for atomic undo and view/selection
  preservation.

See [Concepts](concepts.md) for the *why* behind these shapes.

---

## Connecting to PowerPoint

::: pptlive.attach

::: pptlive.connect

::: pptlive.PowerPoint

## Presentations

::: pptlive.Presentation

::: pptlive.PresentationCollection

## Slides

`Presentation.slides` is a [`SlideCollection`](#pptlive.SlideCollection). Index
a slide by 1-based position (`deck.slides[3]`), iterate it, or use the
lifecycle verbs (`add` / `delete` / `duplicate` / `move_to` / `set_layout`). A
[`Slide`](#pptlive.Slide) exposes `shapes`, `placeholder(kind)`, `notes`,
`read()`, and `export_image(...)`.

::: pptlive.SlideCollection

::: pptlive.Slide

## Shapes & geometry

`Slide.shapes` is a [`ShapeCollection`](#pptlive.ShapeCollection) — index by
1-based z-order (`shapes[2]`) or by name (`shapes["Title 1"]`), and create with
`add_textbox` / `add_shape` / `add_picture` / `add_table` / `add_chart`. A
[`Shape`](#pptlive.Shape) **is** an [`Anchor`](#pptlive.Anchor) when it has a
text frame (so it inherits `text` / `set_text` / `format_text` / the list and
paragraph verbs), and always carries geometry (`move`, `resize`, `geometry()`)
in **points**, plus `alt_text` / `set_alt_text` and per-shape
`export_image(...)`. Every shape also carries a stable `shapeid` (`shapeid:S:ID`,
the delete-proof handle) alongside its z-order `anchor_id`.
[`ShapeById`](#pptlive.ShapeById) is the handle you get back from that
`shapeid` — it resolves by `Shape.Id` on every access, so it keeps pointing at
the same shape across a delete/restack that would shift a `shape:S:N` index.

A **picture** can be cropped. `Shape.crop(*, left=, right=, top=, bottom=)` is
the raw primitive — points off each named edge, 1:1 with PowerPoint's
`PictureFormat.Crop*` — but note that cropping *shrinks the shape box* rather
than letterboxing inside it. `Shape.crop_to_fit(*, left=, top=, width=, height=,
fit="cover")` is the verb that reconciles an aspect mismatch for you:
`fit="cover"` fills the box exactly and centre-crops the overflow (the
full-bleed move — and it keeps the picture *on* the slide, so
`Slide.geometry_report()`'s `off_slide` flag stays trustworthy), while
`fit="contain"` shrinks the whole picture to fit and centres it, letterboxed.
Each box argument defaults to the picture's current geometry, and any existing
crop is cleared first, so re-fitting never compounds. Both return
`{crop, geometry}` (plus `fit`), and a cropped picture carries its `crop` in
shape reads.

A shape can also animate: `Shape.animate(effect="fade", *, trigger="on_click",
duration=None, delay=None, exit=False)` appends a whole-shape entrance (or, with
`exit=True`, exit) effect to the slide's main sequence, and
`Shape.clear_animations()` removes just that shape's effects. Read them back per
slide with [`Slide.animations()`](#pptlive.Slide) (ordered rows, each mapped to its
target by `shapeid`) and wipe a whole slide with `Slide.clear_animations()`. A
slide's spatial layout is available without a render via
[`Slide.geometry_report()`](#pptlive.Slide) (slide size + per-shape boxes +
overlaps + off-slide flags).

::: pptlive.ShapeCollection

::: pptlive.Shape

::: pptlive.ShapeById

::: pptlive.PlaceholderShape

## Anchors

Every text-bearing handle subclasses [`Anchor`](#pptlive.Anchor) and shares the
same verbs — `text`, `set_text`, `insert_paragraph_before/after`,
`format_text`, `format_paragraph`, `apply_list` / `remove_list` — so the same
calls work uniformly on a whole shape, one paragraph, a table cell, or a
slide's notes. PowerPoint has no named paragraph styles, so "styling" is direct
font formatting via `format_text` (bold / italic / underline / size / font /
color). A paragraph read's `font` block also reports `color_source`
(`"direct"` / `"theme"` / `"mixed"`) and `theme_color` (the inherited slot when
themed), so you can tell a run color *set on the run* from one *cascaded from the
theme* — the one place PowerPoint exposes that direct-vs-inherited distinction.
`Anchor.set_paragraphs([...])` is the one-pass authoring path: each item is a
string or a `{"text", ...}` dict that becomes exactly one addressable `para:`.
Its item keys cover **both** paragraph and font formatting — the full table is in
the [`set_paragraphs`](#pptlive.Anchor.set_paragraphs) docstring below — so
following it with a per-paragraph `format_text` loop is pure extra COM
round-trips. An unknown key raises `ValueError` naming the valid ones rather than
being silently ignored.

`Shape.text_frame_status()` returns a [`TextFrameStatus`](#pptlive.TextFrameStatus)
— autosize mode / word-wrap / vertical anchor / margins / a coarse `overflow_risk`
flag — so an agent can see a text box heading for a "formatting spiral" before it
clips, without a render. `Shape.set_text_frame(autosize=, word_wrap=,
vertical_anchor=, margins=, margin_left=, …)` is the **setter** half: it takes the
same knobs and returns the resulting status. The two that matter most for precise
layout are `autosize="none"` — a new text box grows to fit its text, so a `height`
you set is advisory until autofit is off — and `margins=0`, since PowerPoint's
0.1 in (7.2 pt) inner margins silently eat padding math. `add_textbox` and
`add_shape` accept the same arguments, so a box can be created already pinned.

::: pptlive.Anchor

::: pptlive.Paragraph

::: pptlive.ParagraphCollection

::: pptlive.Notes

::: pptlive.TextFrameStatus

## Tables

A table is a **shape on a slide** (`Shape.has_table` / `Shape.table`), not a
deck-scoped collection. Reach a table through its shape
(`slide.shapes[N].table`) and address its cells as `cell:S:N:R:C`. A
[`Cell`](#pptlive.Cell) *is* an [`Anchor`](#pptlive.Anchor), so
`doc.anchor_by_id("cell:4:5:1:1")` returns a handle that works with `set_text`,
`format_text`, and `format_paragraph` like any other anchor.

::: pptlive.Table

::: pptlive.Cell

## Charts

A chart is also a shape (`Shape.has_chart` / `Shape.chart`); its data lives in
an **embedded Excel workbook**. [`Chart`](#pptlive.Chart) reads the chart type,
categories, and series, and writes them back with `set_type` / `set_data`.

::: pptlive.Chart

## SmartArt

A SmartArt diagram is a shape too (`Shape.has_smartart` / `Shape.smartart`); its
content is a tree of nodes. [`SmartArt`](#pptlive.SmartArt) reads the layout kind
and the nested node tree, and replaces it with `set_nodes` — a flat list of
strings, or `{text, children}` mappings that nest. Create one via
`shapes.add_smartart(kind, nodes)`.

::: pptlive.SmartArt

## Theme & master — deck-wide styling

Where `format_text` styles one anchor, [`deck.theme`](#pptlive.Theme) and
[`deck.master`](#pptlive.Master) restyle the **whole deck** by editing what every
slide inherits. `Theme` is the 12-slot palette plus the heading/body typefaces;
`Master` is the primary slide master's text styles (`title` / `body` /
`default`, 5 levels each) and background. These are deliberately global and
anti-polite — one call recolors or re-fonts every inheriting slide — so wrap them
in `deck.edit()` for the one-Ctrl-Z fence (the user's view doesn't move).

::: pptlive.Theme

::: pptlive.Master

## Deck structure — sections & headers/footers

[`deck.sections`](#pptlive.SectionCollection) is the deck's named slide spans —
`list()` returns `{index, name, first_slide, slide_count}` rows, and
`add(name, *, before_slide=None)` / `rename` / `delete(*, delete_slides=False)` /
`move` edit them by 1-based section index. [`HeadersFooters`](#pptlive.HeadersFooters)
is a shared wrapper mounted at two scopes — `slide.headers_footers` (a per-slide
override) and `deck.master.headers_footers` (the deck-wide default every slide
inherits) — with `read()` plus `set_footer` / `set_slide_number` / `set_date`.
A footer / date text reads back as `None` while that element is hidden (PowerPoint
only exposes the text on a visible element), and setting text auto-shows it.

::: pptlive.SectionCollection

::: pptlive.HeadersFooters

## Review comments

`slide.comments` is a per-slide [`CommentCollection`](#pptlive.CommentCollection)
(1-based, `add` / `list` / iterate / index); `deck.comments()` is the deck-wide
roll-up (`{total, slides: [...]}`). Comments attach to a **slide** at an
`(x, y)` point (not a text range) and are **threaded** —
[`Comment.replies`](#pptlive.Comment) / `Comment.reply(text)` walk and extend a
thread. Adding a comment needs the signed-in Office-account identity: `add`
lifts it off any existing comment via the modern `Comments.Add2`, falling back
to the legacy identity-free `Comments.Add` on a comment-less deck; a reply
lifts identity off its parent. Two honest caveats: `Add2` **binds to the
signed-in account** (a passed `author` / `initials` is best-effort — even the
legacy `Add` may ignore them on a modern build), and there is **no
resolve/reopen verb** (`.Status` / `.Resolved` aren't COM-readable on current
builds) — delete a thread once it's addressed instead.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    with deck.edit("Leave and answer review notes"):
        note = deck.slides[2].comments.add("Tighten this headline", left=100, top=80)
        note.reply("Done — shortened to five words")
        deck.slides[2].comments[1].delete()      # resolve by deleting (takes its replies)
```

::: pptlive.Comment

::: pptlive.CommentCollection

## Rendering

[`slide.export_image`](#pptlive.Slide) renders one slide to an image;
[`deck.snapshot`](#pptlive.Presentation) renders the whole deck (or a slide
selection) to one PNG per slide so a vision model can *see* every slide cheaply.
Its `max_dim` long-edge pixel cap gives a predictable, uniform per-slide token
budget (a model is billed on pixel area, not DPI); pass exact `width` / `height`
instead for a fixed per-slide size (they override `max_dim`, and passing both
forms is a `ValueError`). Both are reads — they reflect the current unsaved state
but leave the viewed slide and Selection untouched. Each rendered slide comes back
as a `Snapshot`.

::: pptlive.Snapshot

## Saving & export

Three **explicit, never-implicit** verbs on [`Presentation`](#pptlive.Presentation)
(pptlive never auto-saves): `deck.save()` persists to the existing file;
`deck.save_as(path, *, fmt="pptx", overwrite=False)` writes a `.pptx` and **rebinds**
the working file to it (the open deck becomes that file, like PowerPoint's Save-As),
refusing to clobber unless `overwrite=True`; and `deck.export_pdf(path)` writes a
pixel-faithful PDF as a **read** — unlike `save_as` it neither rebinds the working
file nor clears its dirty flag, so your `.pptx` is untouched. `deck.saved` (the
`Presentation.Saved` dirty flag) and `deck.path` ride on every `status` deck row so
an agent can see unsaved state. `save()` on a never-saved deck raises
[`UnsavedPresentationError`](#pptlive.UnsavedPresentationError) rather than letting
PowerPoint silently route the file to a default cloud folder.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    if not deck.saved:
        deck.save()                       # persist in place (must already have a path)
    deck.save_as("C:/out/v2.pptx")        # write + rebind the working file
    deck.export_pdf("C:/out/deck.pdf")    # a read — working file untouched
```

## Media & narrated-video export

The "build a deck, narrate it, export a video" path.
[`slide.add_audio(path)`](#pptlive.Slide) / [`slide.add_video(path)`](#pptlive.Slide)
insert an audio/video clip (embedded by default; `link=True` keeps the file on disk).
`autoplay` plays the clip on slide entry, `hide_icon` hides the audio icon while idle
(audio only), and `pace_slide` auto-advances the slide to the clip's length — so an
exported video paces itself to the narration. Each shape read carries a `media` dict
(`{type, length_s, start_s, end_s, muted, volume, autoplay}` — `start_s`/`end_s` are
the trim window in seconds) and `has_media`. `Shape.set_media_playback(muted=,
volume=, start=, end=)` sets those playback options on an existing clip.

[`deck.export_video(path)`](#pptlive.Presentation) exports the deck to an MP4 via
PowerPoint's async `CreateVideo`. Like `export_pdf` it is a **read** (no rebind, dirty
flag preserved). It **blocks by default**, polling to completion and returning a
[`VideoExportResult`](#pptlive.VideoExportResult); pass `wait=False` to return the
in-flight status immediately and poll [`deck.video_status()`](#pptlive.Presentation)
until it reports `done`. A failed or timed-out encode raises
[`VideoExportError`](#pptlive.VideoExportError).

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    with deck.edit("Narrate the deck"):
        deck.slides[1].add_audio("intro.mp3")     # autoplay + pace the slide (defaults)
        deck.slides[2].add_video("demo.mp4")      # stays visible; same knobs
    result = deck.export_video("C:/out/deck.mp4", resolution=1080)
    assert result.ok and result.status == "done"  # result.path is the written MP4
```

::: pptlive.VideoExportResult

## Slide show

[`deck.show`](#pptlive.SlideShow) drives a running slide show like a presenter's
clicker — `start`, `end`, `next`, `previous`, `goto(n)`, `black()` / `white()`
/ `resume()`, and the read-only `state()`. Unlike the polite edit verbs, these
deliberately drive what's on screen, so `show` is **not** wrapped in `edit()`.

::: pptlive.SlideShow

## Editing & selection

`deck.edit(label)` returns an [`EditScope`](#pptlive.EditScope) — the
view/selection-preservation and atomic-undo scope. `deck.selection()` reads the
user's current [`SelectionInfo`](#pptlive.SelectionInfo) (resolved to anchors)
without perturbing it; act on it by targeting the opt-in `here:` anchor.

::: pptlive.EditScope

::: pptlive.SelectionInfo

::: pptlive.SelectionSnapshot

## Units

Geometry is in points throughout (1 inch = 72 pt). These helpers convert so you
needn't hardcode multiplications; EMUs never surface.

::: pptlive.units

## Constants

Typed `IntEnum`s for the `Mso*` / `Pp*` / `Xl*` magic constants, plus
friendly-string coercers (`"title"`, `"two_content"`, `"star"`, `"column"`)
that map names to the right int the way an LLM would phrase them.

::: pptlive.constants

## Exceptions

::: pptlive.PptliveError

::: pptlive.PowerPointNotRunningError

::: pptlive.PresentationNotFoundError

::: pptlive.AnchorNotFoundError

::: pptlive.SlideNotFoundError

::: pptlive.LayoutNotFoundError

::: pptlive.NoTextFrameError

::: pptlive.UnsavedPresentationError

::: pptlive.VideoExportError

::: pptlive.SlideShowNotRunningError

::: pptlive.AmbiguousMatchError

::: pptlive.ReplaceVerificationError

::: pptlive.PowerPointBusyError

::: pptlive.ComError

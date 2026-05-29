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
`export_image(...)`.

::: pptlive.ShapeCollection

::: pptlive.Shape

::: pptlive.PlaceholderShape

## Anchors

Every text-bearing handle subclasses [`Anchor`](#pptlive.Anchor) and shares the
same verbs — `text`, `set_text`, `insert_paragraph_before/after`,
`format_text`, `format_paragraph`, `apply_list` / `remove_list` — so the same
calls work uniformly on a whole shape, one paragraph, a table cell, or a
slide's notes. PowerPoint has no named paragraph styles, so "styling" is direct
font formatting via `format_text` (bold / italic / underline / size / font /
color).

::: pptlive.Anchor

::: pptlive.Paragraph

::: pptlive.ParagraphCollection

::: pptlive.Notes

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

::: pptlive.SlideShowNotRunningError

::: pptlive.AmbiguousMatchError

::: pptlive.PowerPointBusyError

::: pptlive.ComError

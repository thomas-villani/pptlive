# pptlive — implementation tracker

Staged build plan + progress, ordered by **LLM-agent leverage** (the same
ordering principle as wordlive's `feature-plan.md`). The design lives in
`spec.md`; this file is the checklist. Update statuses as work lands and record
resolved open questions inline (strike them through, link the commit).

**Status legend:** `[ ]` not started · `[~]` in progress · `[x]` shipped.

> **Bootstrap + v0 + v0.1 + v0.2 + v0.3 + v0.4 + v0.5 + v0.6 + v0.7 (pictures &
> charts) + v0.8 (SmartArt) + v0.9 (master/theme styling) + the MCP server have
> landed** (fake-COM unit tests green: `ruff`, `mypy`, `pytest` all pass; 374 tests;
> v0.7 verified live 2026-05-28 via `scripts/picture_spike.py` + `scripts/chart_spike.py`;
> v0.8 verified live 2026-05-28 via `scripts/smartart_spike.py` + a shipped-wrapper run;
> v0.9 verified live 2026-05-29 via `scripts/master_spike.py`, net-zero).
> The library is usable as an LLM tool two ways now — the JSON **CLI** and an optional
> **MCP server** (`pptlive[mcp]` → `pptlive-mcp`, five `op`-dispatch tools over stdio
> for Claude Desktop & other MCP agents; see the *MCP server* section below). It drives
> **live slide-show control** (`deck.show`: `start/end/next/prev/goto/black/white/
> resume/state` over `SlideShowSettings.Run()` / `SlideShowWindow.View`) — **verified
> live 2026-05-28** via `scripts/show_spike.py` (net-zero; every wrapper behaved as
> coded, and the spike **overturned the spec's busy-during-show assumption** — see
> v0.6 below). It drives the **slide
> lifecycle** (`slide add/delete/duplicate/move/set-layout` + layout-name
> resolution) — verified live 2026-05-26 via `scripts/layout_spike.py` —
> **shapes & geometry** (`shape add` textbox/autoshape/picture + `move/resize/
> delete`) — verified live 2026-05-27 via `scripts/shape_spike.py` — and **text
> structure** (`para:S:N:P` anchors, insert, paragraph/font formatting, bullets) —
> designed from a live COM probe and verified via `scripts/text_spike.py` — and
> **slide render + live selection** (`slide export` → PNG so a vision model can
> *see* the slide; `selection` / `here:` over `ActiveWindow.Selection`) — verified
> live 2026-05-27 via `scripts/render_select_spike.py` — and **tables**
> (`add_table`, `cell:S:N:R:C` cell anchors, `table read`/`add-row`/`delete-row`) —
> verified live 2026-05-28 via `scripts/table_spike.py` (all net-zero; the
> v0.1–v0.5 sections record the findings). The four **Spike**
> items below were
> **verified against real COM on 2026-05-26** (PowerPoint desktop, a 3-slide
> deck). Items #2/#3/#4 confirmed as specced. **#1 overturned the spec's
> headline assumption:** PowerPoint *does* group consecutive in-session COM edits
> into a single undo entry by default, and `Application.StartNewUndoEntry()` is a
> verified *boundary* primitive — so `deck.edit()` blocks are atomic (one
> Ctrl-Z). **Implemented:** `EditScope.__enter__` now fences with
> `StartNewUndoEntry()`, the "no atomic undo" docs are corrected throughout, and
> two unit tests pin the fence. Cross-*process* isolation is also **verified**
> (separate CLI invocations stay distinct undo entries — each re-fences at its
> own `edit()` entry). Still open: CI and v0.1+. The spike harnesses live at
> `scripts/spike.py` and `scripts/undo_test.py`.
>
> **Post-v0.9 tiers have also landed** (each with its own SHIPPED section below):
> **v1.0** fuzzy find / replace, the **v1.2** shape-styling round (fill / border,
> z-order, the delete-proof `shapeid:S:ID` anchor, SmartArt / chart
> `recolor_text`), **v1.3** threaded review comments, and the **v1.1** output tier
> (whole-deck snapshot + explicit `save` / `save_as` / PDF export). All fake-COM
> unit tests stay green (`ruff`, `mypy`, `pytest`).

---

## Spike first (resolve against real PowerPoint before hardening)

These four unknowns shape the v0 design. Confirm each on a Windows box with
PowerPoint, write a one-line finding here, and let it inform the code.

- [x] **UndoRecord absence → REFRAMED: in-session edits already group.** Confirm
  there is *no* undo-grouping primitive (`Application.UndoRecord` or newer). *Assumed
  (now wrong):* none exists and each mutation is its own Ctrl-Z. *Finding
  (2026-05-26, interactive Ctrl-Z test via `scripts/undo_test.py`):*
  `Application.UndoRecord` **is absent** (`AttributeError`) — Word's block API
  doesn't exist. **But it isn't needed:** PowerPoint **groups consecutive COM
  edits made within one automation session into a single undo entry by default.**
  Verified three ways: (a) 2 edits + `StartNewUndoEntry` → 1 Ctrl-Z reverts both;
  (b) 2 edits, *no* `StartNewUndoEntry` → 1 Ctrl-Z **still** reverts both (so the
  call isn't the cause); (c) `[edit1, edit2] · StartNewUndoEntry() · [edit3]` → the
  1st Ctrl-Z reverts **only** edit 3, the 2nd reverts edits 1+2 together. So
  `Application.StartNewUndoEntry()` is a **verified boundary primitive** (it
  *splits* the undo stream), not a grouping one. **Design consequence:** a
  `deck.edit()` block — all edits in one `attach()`/process — is *already*
  effectively atomic (one Ctrl-Z); `EditScope.__enter__` now calls
  `StartNewUndoEntry()` so each block is a clean, self-contained undo entry (no
  bleed into adjacent edits or prior user actions). The "no atomic undo" docs
  were revised accordingly. **Cross-process also verified (2026-05-26):** two
  separate CLI-style invocations (each its own `attach()` + `edit()` fence) edited
  the same slide; one Ctrl-Z reverted only the *second* process's edit, so
  separate invocations stay distinct undo entries (`xproc1`/`xproc2` phases).
  (CLAUDE.md's "no atomic undo, confirm nothing snuck in" guidance is superseded.)
- [x] **`Visible` behaviour.** Confirm `Application.Visible = False` raises (so
  `connect()` ships no `visible=False`). *Assumed:* it raises — `connect()` has
  no `visible` param and `_com.launch_powerpoint()` always sets `Visible = True`.
  *Finding (2026-05-26):* **Confirmed — it raises.** `Application.Visible = False`
  → `com_error 0x80020009` ("Exception occurred."), inner SCODE `0x80048240`,
  message *"Application.Visible : Invalid request. Hiding the application window is
  not allowed."* Original value was `-1` (msoTrue) and was restored cleanly. No
  hidden mode is possible; `connect()`/`launch_powerpoint()` are correct as built.
- [x] **Selection round-trip.** Snapshot `ActiveWindow.View.Slide.SlideIndex` +
  `ActiveWindow.Selection` and restore it. Confirm a shape-range Selection
  round-trips cleanly; if not, fall back to restoring only the viewed slide and
  `Unselect()`. *Built:* `_selection.snapshot/restore` capture the viewed slide
  + selected shape *names* and re-select by name on restore, already falling
  back to `Unselect()` on any failure. *Finding (2026-05-26):* **Confirmed —
  round-trips cleanly.** `Shapes.Range(["Title 1"]).Select()` set `Selection.Type`
  → `SHAPES(2)` and `ShapeRange` yielded exactly `["Title 1"]`. The name-based
  restore mechanism works against real COM. (Text selections — `Type==TEXT(3)` —
  still collapse to `Unselect()` on restore, as designed: not faithfully
  round-trippable in v0.)
- [x] **Notes placeholder resolution.** Confirm the notes body resolves by
  `PlaceholderFormat.Type == ppPlaceholderBody` (not a hard index) across
  templates. *Built:* `_anchors.Notes` resolves the notes-page body by
  `PlaceholderFormat.Type == BODY`, not by index. *Finding (2026-05-26):*
  **Confirmed.** On all three slides the notes page held two placeholders —
  index 1 = `TITLE(1)` (no text frame), index 2 = `BODY(2)` (has text frame).
  Resolving by **type** (not a hard index) found the body every time, and the
  body placeholder exists with a text frame even when the slide has no notes
  text (so `Notes.text` returns `""` rather than raising). `_anchors.Notes` is
  correct as built.

(Layout-name → `CustomLayout` mapping and the undo-mitigation default are spiked
in their own stages below — v0.1 and the Open Questions section.)

---

## Bootstrap (pre-v0 scaffolding)

- [x] `pyproject.toml` — real metadata, deps (`click`, `pywin32`), `pptlive`
  script, ruff/mypy/pytest config. *(done in the scoping pass)*
- [x] `CLAUDE.md`, `IMPLEMENTATION.md`. *(this pass)*
- [x] `README.md` — short version of the wordlive README, PowerPoint-flavored.
- [x] Package skeleton under `src/pptlive/` per the CLAUDE.md module layout.
- [x] `_com.py` seam: `com_apartment`, `get_active_powerpoint`,
  `launch_powerpoint`, `translate_com_errors` (ported from wordlive `_com.py`,
  `Word.Application` → `PowerPoint.Application`; `launch` is visible-only).
- [x] `exceptions.py`: ported wordlive's taxonomy, renamed `Word*` →
  `PowerPoint*`/`Pptlive*`, added `NoTextFrameError` (exit 6) and
  `SlideNotFoundError(AnchorNotFoundError)`. `_decode_com_error` /
  `from_com_error` / `_BUSY_HRESULTS` reused verbatim.
- [x] `tests/conftest.py`: a `fake_powerpoint` fixture — a connected object
  graph (app → presentations → slides → shapes → text frames → placeholders →
  notes), plus `no_powerpoint`, `real_powerpoint`, and `ppt`/`deck` helpers.
- [x] `uv sync --extra dev`; `uv run pytest` (68 passed), `ruff`, `mypy` green.
- [ ] CI: a Windows runner for the smoke suite + a cross-OS unit-test job
  (cross-cutting; can trail v0).

---

## v0 — close the live-edit loop

The minimum that makes pptlive usable as an LLM tool: attach, read structure,
read/set text on the common anchors, polite view/Selection scope, the JSON CLI.

- [x] `attach()` / `connect(launch_if_missing=True)` + context manager
  (`CoInitialize`/`CoUninitialize`; never closes PowerPoint). No `visible=False`.
- [x] `Presentation` + `PresentationCollection` (`.active`, `["Name.pptx"]`).
- [x] Reads: `deck.slides.list()`, `deck.outline()`, `deck.slides[S].read()`,
  `deck.page_setup()` (SlideWidth/SlideHeight in points). Shape listings emit
  `anchor_id` (z-order) **plus `name` and `id` (`Shape.Id`)** for drift-proof
  re-identification.
- [x] `Shape` as `Anchor` (when it has a text frame): `text` / `set_text`. Also
  ships geometry reads (`geometry()`) + `move`/`resize` (used by `read()` now;
  CLI verbs land in v0.2).
- [x] Anchor resolution for `ph:S:KIND`, `shape:S:N`, `notes:S` via
  `deck.anchor_by_id(...)`; `Slide.placeholder(kind)`, `Slide.notes`,
  `slide.shapes["Name"]`. `PlaceholderShape` re-resolves by kind (drift-proof).
- [x] `EditScope` via `deck.edit(label)` — snapshot/restore viewed slide +
  Selection; `allow_view_move()`. **No UndoRecord.** Plus `deck.go_to(...)`.
- [x] `NoTextFrameError` raised when a text op hits a frameless shape (picture/line).
- [x] `constants.py` starter enums: `MsoShapeType`, `PpPlaceholderType`,
  `PpSelectionType`, `MsoTriState`, `PpViewType` (+ friendly-name coercers).
- [x] CLI: `status`, `slides`, `outline`, `slide read S`, `shapes --slide S`,
  `read anchor --anchor-id …` / `read notes --slide S`, `write --anchor-id …`,
  `replace --anchor-id …`, `go-to`. Exit codes wired (0/1/2/3/4/5/6).
- [x] Docs say plainly: **no atomic undo** — each mutation is its own Ctrl-Z.

---

## v0.1 — slide lifecycle (first no-Word-analog track) — SHIPPED

- [x] `slides.add(layout=…, index=…)`, `Slide.delete`, `duplicate`, `move_to(n)`,
  `set_layout(name)`. The verbs only mutate; wrap in `deck.edit(label)` (as the
  CLI does) for view preservation + a one-Ctrl-Z fence. `add`/`duplicate`/`move_to`
  return the resulting `Slide`; `add` defaults to appending; `move_to` returns the
  same slide at its new `index`.
- [x] Layout-name → `CustomLayout` mapping from
  `Presentation.SlideMaster.CustomLayouts` (`Presentation._resolve_layout` +
  `constants.match_layout_name`); modern `Slides.AddSlide(Index, CustomLayout)`,
  legacy `Slides.Add(Index, PpSlideLayout)` only when a deck exposes no custom
  layouts. **Spike RESOLVED (verified live 2026-05-26, `scripts/layout_spike.py`):**
  names match case/separator-insensitively against the deck's *real* layout names
  first (so renamed-layout templates resolve by their actual name), then a small
  friendly-alias table for the standard Office set; an unknown name raises
  `LayoutNotFoundError` (exit 2) **listing the available names**, and
  `deck.layouts()` / `slide layouts` make them discoverable up front. The test
  deck reported 11 layouts (9 standard + two vertical-text variants);
  `AddSlide`/`Slide.Duplicate` (1-based `SlideRange`)/`MoveTo(toPos)`/`CustomLayout =`
  /`Delete` all behaved as coded, net-zero. *Honest caveat:* a **bulk COM
  enumerator** (`list(deck.com.Slides)`, used by `SlideCollection.__iter__`) was
  observed to *transiently* yield a stale slide handle once, right after a
  move/duplicate across separate `edit()` blocks — not reproducible on retest and
  a non-issue for the CLI (each verb is its own process). Resolve slides **by
  index** in tight post-mutation loops if you hit it; revisit iteration hardening
  only if smoke runs show it reliably.
- [x] CLI `slide add|delete|duplicate|move|set-layout`, plus `slide layouts`
  (discovery read). Exit codes reuse the v0 mapping (2 = slide/layout not found).

## v0.2 — shapes & geometry (makes slide-building possible) — SHIPPED

- [x] `ShapeCollection.add_textbox(text, …)` / `add_shape(shape_type, …)` /
  `add_picture(path, …)`; `Shape.delete()`. (`move`/`resize`/`geometry()` already
  shipped in v0.) All geometry in **points**. The verbs only mutate; wrap in
  `deck.edit(label)` (as the CLI does) for view preservation + a one-Ctrl-Z fence.
  New shapes append at the **top of the z-order** (last slot), so the returned
  `Shape` is addressed by the post-add `Shapes.Count` (`ShapeCollection._added()`).
  `add_shape` takes a friendly name (`constants.autoshape_type_for`,
  `MsoAutoShapeType`) or a raw int; `add_picture` **embeds** (never links) and
  raises `FileNotFoundError` for a missing path.
- [x] `units.inches()` / `units.cm()` / `mm()` / `points()` helpers (already in
  `units.py` since bootstrap; never expose EMUs).
- [x] CLI `shape add|move|resize|delete` (the new `shape` group; `shapes --slide S`
  stays the listing). `shape add --kind textbox|shape|picture` with
  `--shape-type` (a `click.Choice` over `constants.AUTOSHAPE_CHOICES`),
  `--text`/`--path`, and `--left/--top/--width/--height`. move/resize/delete take
  `--anchor-id` (`shape:S:N` or `ph:S:KIND`); a non-shape anchor (e.g. `notes:S`)
  is exit 2. **Spike RESOLVED (verified live 2026-05-27, `scripts/shape_spike.py`,
  net-zero):** `AddTextbox(Orientation, L, T, W, H)`, `AddShape(Type, L, T, W, H)`,
  and `AddPicture(FileName, LinkToFile=msoFalse, SaveWithDocument=msoTrue, L, T, W,
  H)` all add + return a shape; a new shape lands at the last z-order index (text
  round-tripped through it); the shipped autoshape ints are correct
  (`rectangle`=1, `oval`=9, `right_arrow`=33, `five_point_star`=92); `move`/
  `resize`/`Delete` behave; a picture has no text frame. *Honest caveat:* a text
  box created **with** text auto-fits its height to the content (PowerPoint's
  default AutoSize) — requesting `height=72` for "MARKER-TB" came back ~29 pt. The
  `left`/`top`/`width` you pass are honored; height is advisory when AutoSize is on.

## v0.3 — text structure — SHIPPED

- [x] `para:S:N:P` anchors (`Paragraph` over `TextRange.Paragraphs(P, 1)`),
  plus `Shape.paragraphs` / `Shape.paragraph(p)` and the structured
  `paragraphs.list()`. Resolves live (the paragraph count drifts as text is
  inserted/deleted); out-of-range `P` is exit 2, a frameless shape is
  `NoTextFrameError` (exit 6).
- [x] `insert_paragraph_before/after`, `format_text`, `format_paragraph`,
  `apply_list`/`remove_list` live on the **base `Anchor`** (act on
  `self._text_range()`), so they work on a whole-shape anchor *and* on one
  `Paragraph`. **`apply_style` reframed → `format_text`:** PowerPoint has no
  named paragraph styles (no `Presentation.Styles` analog), so "styling" is
  direct font formatting (bold/italic/underline/size/font/color). Indent is
  `IndentLevel` (1-5), PowerPoint's only paragraph-indent notion (there is no
  points-based `LeftIndent` on `ParagraphFormat`).
- [x] CLI `paragraphs`, `insert` (`--before/--after`), `format-paragraph`
  (alignment/spacing/indent-level), `format-text` (the `style apply` reframe),
  and the `list` group (`apply`/`remove`). Exit codes reuse the mapping.
- [x] **Spike RESOLVED (designed from a live COM probe 2026-05-27, then verified
  via `scripts/text_spike.py`, net-zero).** Findings that shaped the code:
  `TextRange.Paragraphs(P, 1)` is 1-based and a non-final paragraph's `.Text`
  **includes its trailing `\r`** (the last one doesn't); assigning
  `Paragraphs(P,1).Text` **preserves the paragraph break** (no Word-style
  trailing-mark gymnastics needed); `InsertBefore(text+"\r")` cleanly prepends a
  paragraph, while `InsertAfter` needs **end-detection** (append `text+"\r"` for a
  non-final paragraph, prepend `"\r"+text` for the final one) — both verified to
  land a clean paragraph; alignment ints are left=1/center=2/right=3/justify=4;
  `SpaceBefore`/`SpaceAfter` are points and `SpaceWithin` is the line-spacing
  multiple; `Bullet.Visible/Type/Character` and `Font.Bold/Italic/Underline/Size/
  Name/Color.RGB` round-trip; `Font.Color.RGB` is R-low-byte (`#FF0000` -> 255);
  `IndentLevel` (1-5) works on both a textbox and a body placeholder. The fake's
  paragraph model reproduces the char-splice behavior exactly, so the unit tests
  are faithful.

## v0.4 — render & live selection (the vision loop + cursor analog) — SHIPPED

The highest agent-leverage surface left: let a vision model *see* the slide it
just built (export → `Read` the PNG → iterate), and let the agent read — and,
opt-in, target — what the user is currently looking at (wordlive's cursor/`here`
marker, re-applied to PowerPoint's 2-D Selection). **Decisions made:** `slide
export` writes to a temp PNG when `--out` is omitted (so export-then-`Read` is one
step), and the `here:` anchor shipped alongside the `selection` read (it reuses
the same Type→anchor map). Verified end-to-end on a live deck via the wrapper-level
`scripts/render_select_spike.py` (net-zero): native dims 1280×720, requested
640×480, aspect-fill 1920→1080, the export reflects an unsaved edit, it's polite
(view+selection unchanged), and `here:` resolved to both `shape:2:3` and
`para:2:3:2` ("Demo").

- [x] **Slide render.** `Slide.export_image(path, *, width=None, height=None,
  fmt="png") -> Path` over `Slide.Export(FileName, FilterName, ScaleWidth,
  ScaleHeight)`; `Presentation.export_images(dir)` for the whole deck. CLI
  `pptlive slide export --slide S --out PATH [--width N] [--height N]` → JSON
  `{path, width, height, format}` (default `--out` to a temp PNG so the agent can
  export-then-`Read` in a single step). **Spike DONE (2026-05-27,
  `scripts/render_select_spike.py`, net-zero):** `Export(path, "PNG")` writes a
  valid PNG; default dims = the slide's native pixel size (a 960×540 pt 16:9 slide
  → **1280×720 px**, i.e. 96 DPI); `ScaleWidth`/`ScaleHeight` are honored (pixels);
  the export **captures unsaved live edits** (re-rendering "BEFORE"→"AFTER" text
  gave different bytes — it renders the in-memory state, not last-saved, which is
  exactly the iterate loop); it is **polite** (viewed slide + Selection unchanged
  before/after); a **relative path is a footgun** — it lands in PowerPoint's
  Documents dir (OneDrive-redirected on this box), *not* the process CWD, so the
  wrapper must `os.path.abspath` the target before calling `Export`.
- [x] **Live selection read.** `deck.selection()` / `pptlive selection` → the
  current `ActiveWindow.Selection` resolved to our anchor vocabulary (the read
  complements `status`, which already reports the viewed slide). **Spike DONE:**
  `Selection.Type` maps cleanly — NONE(0)→nothing, SHAPES(2)→`shape:S:N` per shape
  (z-order recovered from the stable `Shape.Id`), TEXT(3)→host shape + recovered
  paragraph index → `para:S:N:P` (selecting para 2 of `Intro\rDemo\rQ&A` gave
  `start_char` 7 → index 2, text `"Demo\r"`). SLIDES(1) is a sorter/thumbnail-pane
  state, **not** reproducible in Normal view — map it from `Selection.SlideRange`
  only when it actually occurs. Reading the selection doesn't perturb it.
- [x] **`here:` targetable anchor (opt-in).** `anchor_by_id("here:")` resolves
  live to the selected shape/paragraph — the explicit opt-in the politeness model
  reserves ("never target the Selection unless explicitly asked"). Defer if the
  read alone covers the workflow; the resolution reuses the read's Type→anchor map.

## v0.5 — tables — SHIPPED

A table is a **shape on a slide** (not a doc-scoped collection as in wordlive):
it satisfies `Shape.HasTable` and exposes the grid via `Shape.Table`. So there's
no deck-wide `tables` collection — a table is reached through its shape
(`slide.shapes[N].table`) or the `cell:S:N:R:C` anchor, and a `Cell` *is* an
`Anchor` (it targets the cell's own text frame, so it inherits every text/format
verb with no special-casing). Verified end-to-end on a live deck via the
wrapper-level `scripts/table_spike.py` (net-zero).

- [x] **`add_table` + the `HasTable` gate.** `ShapeCollection.add_table(rows,
  columns, *, left/top/width/height)` over `Shapes.AddTable`, returning the table
  `Shape` (last z-order). `Shape.has_table` / `Shape.table`; `shape_to_dict` now
  emits `has_table` so a listing reveals which shapes are tables. **Spike RESOLVED
  (2026-05-28, `scripts/table_spike.py`, net-zero): the headline finding is that
  `AddTable` can return a shape whose `Type` reports placeholder (14), not table
  (19)** — when it fills a content placeholder — so `HasTable` is the *only*
  reliable gate (the wrapper never checks `Type`). `Shape.table` raises
  `AnchorNotFoundError` (kind `"table"`, exit 2) on a non-table shape.
- [x] **`cell:S:N:R:C` anchors (`Cell` *is* an `Anchor`).** `Table.cell(r, c)` →
  `Cell` over `Table.Cell(r, c).Shape.TextFrame.TextRange`; inherits `set_text`/
  `format_text`/`format_paragraph`/`apply_list`/`insert_paragraph_*` unchanged.
  Resolved through `Presentation.anchor_by_id("cell:S:N:R:C")`. Bounds-checked
  (out-of-range row/col → `AnchorNotFoundError`, since live COM raises on a bad
  cell). Cell text is a plain text frame (paragraphs split by `\r`, multi-line
  round-trips — no Word end-of-cell markers to strip).
- [x] **`Table.read` / `grid` + row edits.** `read()` emits `{slide, shape,
  anchor_id, rows, columns, cells}` with each cell carrying its `cell:S:N:R:C`
  anchor; `grid()` is the row-major text. `add_row(values=None)` appends one row
  (`Rows.Add()`) and fills it left-to-right; `delete_row(index)` removes it
  (bounds-checked). Verified live: append grew 3→4 and the new row was
  addressable/fillable; delete shrank back.
- [x] **CLI.** `shape add --kind table --rows R --cols C`, and the `table` group:
  `table read --slide S --shape N`, `table add-row --slide S --shape N [--values
  JSON]`, `table delete-row --slide S --shape N --row R`. Exit codes reuse the
  mapping (2 = no table at that shape / out-of-range cell).

## MCP server — Claude Desktop & other MCP agents — SHIPPED

A second front-end alongside the JSON CLI: an optional **Model Context Protocol**
server so Claude Desktop (and any MCP client) can drive the live deck directly.
Ships as the `pptlive[mcp]` extra (the official `mcp` SDK; folded into `dev` so
the suite exercises it), with a `pptlive-mcp` console script (`python -m
pptlive.mcp`) running **FastMCP over stdio**. Lives in `src/pptlive/mcp/`
(`server.py` + `__main__.py`); `tests/test_mcp.py` (62 tests, `importorskip`'d)
drives the tool functions against the same `fake_powerpoint` deck.

- [x] **Five-tool dispatch surface (not 1:1 with the CLI).** Each tool takes an
  `op` and routes to a verb, so the agent's tool picker sees five definitions
  instead of fifteen: `ppt_read`
  (`status|slides|outline|slide|anchor|selection|table|chart|layouts` — every
  read), `ppt_edit`
  (`write|format|slide_add|slide_delete|slide_duplicate|slide_move|set_layout|shape_add|shape_move|shape_resize|shape_delete|set_alt|table_add_row|table_delete_row|chart_set_type|chart_set_data`
  — every mutation), `ppt_render` (`slide_image|shape_image|navigate`), `ppt_show`
  (`state|start|end|next|previous|goto|black|white|resume`), and `ppt_batch` (a
  list of the above ops over one connection; `atomic` fences every `edit` into a
  single undo entry). Each op's logic lives in a `_<tool>_core(handle, op, params)`
  helper that does no `attach()` of its own — the public tool wraps it in
  `attach()` (+ an `edit()` fence for `ppt_edit`), and `ppt_batch` reuses the very
  same cores across one shared `attach()`. Tables/charts are addressed by their
  shape's `anchor_id` (`shape:S:N`), not separate slide+shape ints. So the
  politeness model + one-Ctrl-Z fence carry over for free, and reads never move
  the view.
- [x] **Errors mirror the CLI exit-code taxonomy.** A `PptliveError` is re-raised
  as an MCP `ToolError` whose message carries a stable category token (`not_found`
  / `ambiguous` / `busy` / `not_running` / `no_text_frame` / `invalid_args` /
  `error`) — the string analog of the CLI's exit codes — so the agent can branch
  on failure. Inside `ppt_batch` the same tokens are reported per-command.
- [x] **Spike RESOLVED (2026-05-28, `scripts/mcp_spike.py`).** The one genuine new
  risk was COM under the server's event loop: PowerPoint COM is STA-thread-bound.
  Finding: FastMCP calls a **sync** tool function *directly on its event-loop
  thread* (no thread-pool offload — confirmed in the SDK's
  `func_metadata.call_fn_with_arg_validation`, and empirically: `loop_thread ==
  tool_thread == MainThread` driving the real `call_tool` path). So each tool's
  `attach()` runs on one consistent thread. The original design re-`CoUninitialize`d
  per call on the assumption that the balanced cycle was STA-safe — **which proved
  wrong** (diagnosed 2026-05-29): repeated `CoUninitialize` on the long-lived
  event-loop thread dropped PowerPoint's automation connection (snapping its view
  to slide 1 — the "jumps back to the title slide" report) and eventually
  segfaulted. Fix: `com_apartment()` now initialises COM **once per thread and
  never uninitialises** (the OS reclaims it at process exit). Tools still
  re-`attach()` per call (cheap `GetActiveObject`, no cached proxy) but the
  apartment stays open for the session. The full stdio path was also exercised
  end-to-end (a `stdio_client` spawned `pptlive-mcp`, initialized, and listed all
  13 tools).

## v0.6 — live slide show control (the most literally "live" surface) — SHIPPED (fake-COM; live spike pending)

The most literally "live" surface: drive a running slide show like a presenter's
clicker. No Word analog — pure PowerPoint. `deck.show` is **not** wrapped in
`edit()` (a show deliberately takes over the screen, like `go_to`, and show
control has no undo); the inverse interaction — edits *during* a show — is the
busy one. Lives in `_show.py` (`SlideShow`), exposed as `Presentation.show`.

- [x] **`deck.show` wrapper.** `start(*, from_slide=None)` (idempotent: a second
  `start()` keeps the running show, but `from_slide` jumps it), `end()` (no-op if
  none running), `next()`/`previous()`, `goto(n)`, `black()`/`white()`/`resume()`
  (the B/W blank-screen states via `View.State`, a `PpSlideShowState`), `state()`
  (the only side-effect-free verb, and the only one that never raises when no show
  is running — it reports `running: false`), and `is_running()`. Every control
  verb returns the post-action `state()` dict. `_window()` treats any failure to
  reach `Presentation.SlideShowWindow` as "not running"; the control verbs then
  raise the new `SlideShowNotRunningError` (exit 1, a precondition failure — not a
  missing anchor). Out-of-range `goto`/`from_slide` → `SlideNotFoundError` (exit 2).
- [x] **Editing-during-show — spec assumption OVERTURNED (2026-05-28 spike).** The
  spec assumed "while a slide show is running, most editing calls reject → surface
  as `PowerPointBusyError` (exit 3)". **It doesn't.** The spike ran a real show and
  did a `set_text` on a slide-1 placeholder mid-show: it **succeeded**
  (`edit_succeeded_during_show: true`), raised nothing, and round-tripped (text set,
  then restored). So a text edit during a running show is *not* rejected and there
  is **no busy HRESULT to add** to `_BUSY_HRESULTS` from this path. This is the
  third spec assumption a spike has corrected (after undo grouping and `Visible`).
  *Honest scope:* only a `TextRange.Text` set was probed; a structural op (add/
  delete slide) during a show wasn't — if one is ever seen to reject, classify it
  then. The `PowerPointBusyError` docs were softened accordingly (it stays the home
  for genuine modal-dialog `RPC_E_*` rejections; "a show blocks edits" is no longer
  claimed).
- [x] **CLI** `show start|end|next|prev|goto|black|white|resume|state` (the new
  `show` group; `show start --from N`, `show goto --slide N`). Each prints the
  resulting state JSON; exit codes reuse the mapping (1 = no show running, 2 =
  out-of-range slide).
- [x] **MCP** `ppt_show` (op = `state|start|end|next|previous|goto|black|white|
  resume`, optional `slide`).
- [x] **Live spike RESOLVED (2026-05-28, `scripts/show_spike.py`, net-zero).** On a
  real 4-slide deck: `state()` reported `done`/null before start; `start()` →
  running on slide 1; `next`/`previous`/`goto(last)`/`goto(first)` tracked both
  `current_slide` and `position` exactly; `black`/`white`/`resume` round-tripped
  `View.State` (3 → 4 → 1); `end()` exited and `state()` went back to `done`. The
  show **exits at the end of the run by design** (the spike's `finally` calls
  `end()`, so it never strands the user in presentation mode). `net_zero_ok: true`
  (slide count unchanged; the busy-probe text was restored). Plus the
  editing-during-show finding above.

## v0.7 — pictures & charts — SHIPPED

`add_picture` (embed, never link) already shipped in v0.2; v0.7a added the two
picture-polish tracks around it — **alt text as the re-identification handle**
and **per-shape image extraction** — and v0.7b adds **charts** (`add_chart` +
the `Chart` wrapper over the embedded-Excel data).

- [x] **Alt text as the LLM re-identification handle (wordlive v0.8 pattern).**
  `Shape.alt_text` (read) + `Shape.set_alt_text` (write) over
  `Shape.AlternativeText`; `add_picture(..., alt_text=)` sets it on create; every
  shape listing now emits `alt_text`, so an agent can tag a picture/diagram with a
  description and re-find it after z-order drift without leaning on the volatile
  `shape:S:N`. CLI `shape set-alt` + `--alt-text` on `shape add`; MCP `ppt_edit`
  op `set_alt` (+ `alt_text` on op `shape_add`).
- [x] **Image extraction for vision models (per-shape `Shape.Export`; wordlive
  v0.9 pattern).** `Shape.export_image(path=None, *, fmt="png")` — the per-shape
  complement to v0.4's whole-slide render, cropped to the shape's rendered bounds
  at **native pixel size**. Wraps `Shape.Export(PathName, Filter)` where `Filter`
  is the new `PpShapeFormat` **int** enum (`shape_image_filter_for`), *not*
  `Slide.Export`'s string FilterName (and the raster set is narrower — no TIFF).
  Temp-file default + relative→absolute path, like `Slide.export_image`, but
  **no output-size override** (see the spike finding). CLI `shape export`; MCP
  `ppt_render` op `shape_image`.
- [x] **Live spike RESOLVED (2026-05-28, `scripts/picture_spike.py`, net-zero) —
  spec assumption OVERTURNED.** All three probes ran on a live deck: alt-text
  round-trip (`""` → set → read-back → in-listing → restored), `add_picture` with
  `alt_text` (a 2×2 PNG embedded at 1.5 pt, alt text set, then deleted), and
  per-shape export. **The headline finding:** `Shape.Export`'s ScaleWidth/
  ScaleHeight do **not** map to output pixels the way `Slide.Export`'s do.
  Native export is reliable (a 720 pt-wide shape on a 960 pt slide → **960 px**),
  but requesting 400×300 (raw via `.com`) gave **399×241** — width roughly
  tracked, height didn't, aspect wasn't preserved. So `export_image` ships
  **native-only**; a size override would have been a misleading promise. (This is
  the 4th spec/assumption a spike has corrected — after undo grouping, `Visible`,
  and editing-during-show — and specifically overturns the v0.4-symmetry
  expectation that the slide export's pixel semantics carry to shapes.)
  Fake-COM-tested too (277 tests green: ruff/mypy/pytest).
- [x] **`add_chart` + the `Chart` wrapper (v0.7b; wordlive v0.10 reasoning).**
  `ShapeCollection.add_chart(chart_type, categories=None, series=None, *,
  geometry)` over `Shapes.AddChart2`, returning the chart `Shape` (last z-order);
  `Shape.has_chart` (the gate, like `has_table`) / `Shape.chart`. A chart's data
  lives in an **embedded Excel workbook**, driven by `Chart` (`_charts.py`):
  `read()` (chart type + categories + series), `set_type()`, `set_data(categories,
  series)`. `chart_type` is a friendly name (`XlChartType` + `chart_type_for`/
  `CHART_TYPE_CHOICES`). CLI `shape add --kind chart` (+ `--chart-type`/
  `--categories`/`--series`) and the `chart` group (`read`/`set-type`/`set-data`);
  MCP `ppt_read` op `chart` + `ppt_edit` ops `chart_set_type`/`chart_set_data` +
  `kind="chart"` on `ppt_edit` op `shape_add`.
- [x] **Chart spike RESOLVED (2026-05-28, `scripts/chart_spike.py`, net-zero) —
  two non-obvious COM findings.** The exploratory pass found: (1) **`SetSourceData`
  takes a STRING range** (`"Sheet1!$A$1:$C$4"`), not a `Range` object — the Range
  form raised `E_FAIL`; (2) **`SetSourceData` dissolves the default Excel Table
  (ListObject)** — so relying on `ListObject.Resize` breaks the *second* write
  (`DISP_E_BADINDEX`). The shipping `set_data` therefore uses
  `ChartData.Activate()` → `UsedRange.ClearContents()` → write corner/series-names/
  categories/values → `SetSourceData(string)` → `Workbook.Close()`, with **no
  ListObject** — verified live across first-write, shrink (2×1), and grow (4×3)
  with no stale data, plus `ChartType` round-trip and a clean workbook close.
  `AddChart2` reports `Type=chart` here but `HasChart` is the reliable gate (the
  table lesson). The regression spike now drives the *shipped* wrappers end-to-end
  (net-zero). Fake-COM-tested too (303 tests green: ruff/mypy/pytest).

## v0.8 — SmartArt — SHIPPED

Never specced (it appears nowhere in the original `spec.md` — genuinely net-new),
but it fits the established shape-gate + wrapper pattern *exactly* (the
`_charts.py` / tables mold): a friendly-name-keyed diagram you populate from a
flat list or a tree and can **read back to reconstruct the data** — the loop the
user wanted. **Exploratory spike RESOLVED (2026-05-28, `scripts/smartart_spike.py`,
net-zero)** drove raw COM to design the surface; the findings:

- **Layout identity — map friendly names → stable URN `.Id`.**
  `Application.SmartArtLayouts` had **159** installed; the collection *index
  drifts* but each layout's `.Id` is a stable URN
  (`urn:microsoft.com/office/officeart/2005/8/layout/<seg>`). All 7 core layouts
  resolve by trailing segment: `list1` (Vertical Box List), `process1` (Basic
  Process), `cycle1` (Text Cycle), `hierarchy1`, `orgChart1` (Organization Chart),
  `pyramid1` (Basic Pyramid), `venn1` (Basic Venn). So a `smartart_layout_for`
  table maps friendly name → URN and resolves live (parallels `chart_type_for` /
  the layout-name resolver). Start with these 7; widen on demand.
- **The gate is `Shape.HasSmartArt`** (not `Type`). `AddSmartArt` reports
  `Type == msoSmartArt (24)` here, but follow the table/chart lesson and gate on
  `Has*`. A non-SmartArt shape (textbox, `Type 17`) reports `HasSmartArt == False`.
- **Create:** `Shapes.AddSmartArt(SmartArtLayout, Left, Top, Width, Height)` →
  the SmartArt `Shape` (geometry in points, like every other `add_*`).
- **Populate (flat) — solid.** `SmartArt.Nodes`: `.Add()` / `Item(i).Delete()`
  to size the top-level list to the caller's item count, then
  `Item(i).TextFrame2.TextRange.Text = …` (note **TextFrame2**, not `TextFrame`).
  Flat round-trip verified (`['Discover','Design','Build','Ship']`). Default node
  counts vary per layout (list/process/venn/pyramid = 3, cycle = 5), so the
  wrapper *must* size to the item list rather than assume.
- **Populate (tree) — gotchas, resolved.** (a) Tree layouts ship a **pre-built
  skeleton** (`hierarchy1` defaults to `Nodes.Count==1` but `AllNodes==6`;
  `orgChart1` → 1/5 — the root already has empty placeholder children), so the
  wrapper must **clear to one empty root first** (delete top-level extras, then
  strip the root's descendants). (b) `node.Nodes.Add()` adds a **sibling**, not a
  child — true nesting needs `node.AddNode(msoSmartArtNodeBelow=5, type)`. With
  that recipe a clean `CEO → VP Eng → Eng Lead`, `VP Sales` tree built and read
  back correctly (`nests_correctly: true`). (c) **`SmartArt.Nodes.Add()` is a
  no-op on a tree layout** (orgChart/hierarchy cap at a single top-level root) but
  *does* grow flat layouts — so "multiple top-level nodes" is a flat-layout
  capability; tree layouts take one root + `AddNode` children.
- **Two things that do NOT round-trip (deferred, not shipped in the first cut):**
  (1) **node type** — a node created with `type=msoSmartArtNodeAssistant=4` reads
  back as `Type==1` (default), so assistant nodes are write-only and unverifiable;
  left out of v1. (2) Size-override on the shape isn't a SmartArt concern (geometry
  is the normal shape geometry).
- **Read back / reconstruct:** recurse `SmartArt.Nodes` capturing
  `TextFrame2.TextRange.Text` + `.Level` + children (and `SmartArt.Layout.Id` →
  friendly kind), or flat via `.AllNodes`. Round-trips for text + structure.

Build (mirrors v0.7b charts) — all landed; fake-COM unit-tested (340 tests green:
ruff/mypy/pytest) and **verified live 2026-05-28** (shipped wrappers driven
end-to-end against a real deck, net-zero: flat process round-trips, orgChart
builds CEO→VP(level 2)→AE(level 3), the gate/listing report `has_smartart`, and a
tree layout rejects multiple top-level roots):

- [x] **`_smartart.py` — `SmartArt` wrapper + `smartart_layout_for` URN table.**
  `Shape.has_smartart` (the gate) / `Shape.smartart`; `SmartArt.read()` (layout +
  nested `{text, level, children}`), `set_nodes(nodes)` (flat list *or* nested
  tree — clear-skeleton-then-build, `AddNode(BELOW)` for children, with a
  kind-based pre-check that rejects multiple roots on tree layouts before any COM).
- [x] **`ShapeCollection.add_smartart(kind, nodes=None, *, geometry)`** over
  `Shapes.AddSmartArt`, returning the SmartArt `Shape` (last z-order); `kind` is a
  friendly name resolved to the live `SmartArtLayout` by URN
  (`_resolve_smartart_layout`, exit 2 if the layout isn't installed).
  `shape_to_dict` emits `has_smartart`.
- [x] **Constants:** `MsoSmartArtNodePosition` enum + `SMARTART_CHOICES` /
  `SMARTART_TREE_KINDS` / `smartart_layout_for` (friendly → URN segment) +
  `smartart_layout_name` (URN → friendly, for read-back). (`MsoSmartArtNodeType`
  deferred with assistant.)
- [x] **CLI** `shape add --kind smartart` (+ `--smartart-kind` / `--nodes` JSON)
  and a `smartart` group (`read` / `set-nodes`). Exit codes reuse the mapping
  (2 = unknown layout / no SmartArt at that shape).
- [x] **MCP** `ppt_read` op `smartart` + `ppt_edit` op `smartart_set_nodes`
  (+ `kind="smartart"` on op `shape_add`, with `smartart_kind` / `nodes`).
- [ ] **Hardening spike — deferred follow-ups** (the v1 surface is verified; these
  remain): orgChart assistant/branch nodes (the write-only `type` issue), layout
  availability across Office versions (the 7 are core/stable on this build), and
  widening past the 7 core layouts on demand.

## v0.9 — master / theme styling — SHIPPED

"Overall / master styles" — the deck-wide counterpart to v0.3's per-run
`format_text`. These are **global, anti-polite** authoring ops (they restyle the
whole deck at once — the opposite of the per-anchor model), so they get their own
two surfaces — `deck.theme` and `deck.master` — *not* a fold into `format_text`.
They still mutate, so they go through `deck.edit()` for the one-Ctrl-Z fence (the
view doesn't move, so restore is a no-op). Lives in `_theme.py` (`Theme` +
`Master`, both bound to the `Presentation`, reaching the primary
`SlideMaster`). **All four sub-areas built and verified live 2026-05-29
(`scripts/master_spike.py`, net-zero — every write round-tripped):**

- [x] **`_theme.Theme` (`deck.theme`) — palette + typefaces.** `read()` →
  `{colors:{12 named slots}, fonts:{major, minor}}`; `set_color(slot, color)` over
  `Theme.ThemeColorScheme.Colors(1–12).RGB`; `set_font(which, name, *, script)` over
  `ThemeFontScheme.MajorFont`/`MinorFont`. The RGB long is the same R-low-byte form
  as `Font.Color.RGB`, so `parse_color`/`color_hex` are reused. **Fonts go through
  `.Item(1=Latin/2=EastAsian/3=ComplexScript)` — the late-bound `.Latin` accessor
  raises `AttributeError`.**
- [x] **`_theme.Master` (`deck.master`) — text styles + background.** `read()` →
  `{text_styles:{title/body/default:{levels:[…5…]}}, background:{type, color}}`;
  `format_text_style(style, level, …)` + `format_paragraph_style(style, level, …)`
  over `TextStyles(ppTitleStyle=2/ppBodyStyle=3/ppDefaultStyle=1).Levels(1–5)`;
  `set_background(color)` (solid fill). The text-style verbs drive the **same** COM
  `Font`/`ParagraphFormat` objects as `Anchor.format_text`/`format_paragraph`, so
  the application logic was lifted into shared `_anchors.apply_font` /
  `apply_paragraph_format` helpers and reused verbatim.
- [x] **Constants:** `PpTextStyleType` + `text_style_for`/`TEXT_STYLE_CHOICES`;
  `MsoThemeColorSchemeIndex` + `theme_color_for`/`THEME_COLOR_CHOICES` (12 slots,
  `hlink`/`folhlink` aliases); `theme_font_slot_for`/`theme_font_script_for` +
  `THEME_FONT_SLOTS`/`THEME_FONT_SCRIPT_CHOICES`.
- [x] **CLI** the `theme` group (`read`/`set-color`/`set-font`) and the `master`
  group (`read`/`format-text-style`/`format-paragraph-style`/`set-background`).
- [x] **MCP** `ppt_read` ops `theme`/`master` + `ppt_edit` ops `theme_set_color`/
  `theme_set_font`/`master_format_text_style`/`master_format_paragraph_style`/
  `master_set_background`.
- [x] **Spike RESOLVED (2026-05-29, `scripts/master_spike.py`, net-zero).** Drove
  the shipped wrappers on a live deck; theme color, theme font, master body-L1 font
  + size, **and** the master background all round-tripped, then restored to their
  captured originals (`net_zero_ok: true`). Open questions resolved: **(a)
  multi-master** — `Presentation.Designs.Count == 1` on this deck; v0.9 targets the
  *primary* `SlideMaster`, and `deck.master` staying the primary keeps the API
  non-breaking if a `deck.masters` collection is added later. **(b) undo
  granularity** — master edits carry the same `StartNewUndoEntry` fence as content
  edits (the one-Ctrl-Z reversion is an interactive check, not auto-asserted).
  **(c) background** — the master's background was a plain solid, so the round-trip
  ran and reverted exactly; the spike *skips* the destructive solid write when the
  background isn't already a plain solid (to stay net-zero).

**Deferred (not built):** multi-master / per-`Design` styling; per-layout
(`CustomLayouts(i).Background`) backgrounds; non-solid background fills
(gradient/picture); the East-Asian/Complex-Script theme fonts beyond the `--script`
opt-in; legacy `.ppt` theme-object behavior.

## v1.0 — find / replace — SHIPPED

Closes the last wordlive surface-parity gap (the one specced-but-unbuilt module).
PowerPoint has no deck-wide character stream, so `find` is a **traversal** of
slides × shapes → each text frame, table cells, and speaker notes; there is no
`range:` anchor. Each hit is reported against a resolvable text anchor
(`shape:S:N`, `cell:S:N:R:C`, `notes:S`) with a 0-based in-frame offset.

- [x] **`_findreplace.py`** — wordlive's pure fuzzy-matching core ported **almost
  verbatim** (`_normalize` / NFKC + smart-quote/dash/NBSP folds + whitespace
  collapse, `find_matches`, `Match`). OS-independent and
  unit-tested against the fake. Fold-table keys are written as `\u`/`\x` escapes
  so the exotic code points survive a source round-trip.
- [x] **`Presentation.find(text, *, scope=None)`** → `[{anchor_id, start, length,
  text, context}]` in document order; `Presentation.find_replace(find, replace, *,
  scope=None, all=False, occurrence=None)` → applied list. `scope` accepts a
  `slide:S`/anchor-id string, a `Slide`, an `Anchor`, or `None` (whole deck).
  Matching is **case-sensitive** (NFKC doesn't lowercase), like wordlive.
- [x] **Replacement writes through `TextRange.Characters(start+1, length)`** — only
  the matched span changes, so the rest of the frame keeps its run formatting
  (the PowerPoint analog of wordlive's `Range(start,end).Text=`). Matches in one
  frame are applied in **reverse** offset order so earlier offsets stay valid.
  Because matches are computed once up front from the original text (not via a
  loop over native `.Replace`), the **offset-drift hazard the spike flagged**
  (a replacement that re-contains the search text) cannot occur.
- [x] **Errors:** zero matches → `AnchorNotFoundError("find", …)` (exit 2);
  multi-match without `all`/`occurrence` → `AmbiguousMatchError` (exit 5, carries
  the matches). `find` itself never raises on zero — it returns `[]`.
- [x] **CLI:** `find --text … [--in SCOPE]` (emits the match array; empty array /
  exit 0 on no match) and `replace --find OLD --text NEW [--in SCOPE]
  [--all|--occurrence N]` (the existing `replace --anchor-id` whole-anchor form is
  unchanged; the two modes are mutually exclusive).
- [x] **MCP:** `ppt_read` op `find` (+ `text`/`scope` args); `ppt_edit` op
  `find_replace` (+ `find`/`scope`/`replace_all`/`occurrence` args). Both flow
  through `ppt_batch` via the shared `_read_core`/`_edit_core`.
- [x] **Spike RESOLVED (2026-06-07, `scripts/findreplace_spike.py`, net-zero)** —
  pinned the COM behaviours the design rests on (empty-match `None` sentinel,
  1-based `TextRange` offsets, `.Replace` is first-only, the drift hazard, and
  notes/cell reach). The Python-side matching sidesteps native `.Replace` entirely
  and adds the fuzzy matching native Find lacks. See `roadmap.md` §v1.0.

**Deferred:** a within-shape `range:S:N:START-END` anchor (until a real
mid-paragraph workflow needs it). *(The CLI `exec` batch verb shipped 2026-06-10 —
see the `exec` batch ops section below.)*

## v1.2 — shape styling: fill / border, z-order, shapeid, composite recolor — SHIPPED

The styling round that closed the PPTLIVE-007…010 field-guide gaps — the
shape-level visual verbs and the delete-proof anchor that earlier rounds asked
for. Resolves spec Open Q #3 (symbolic `exec` binding stays deferred).

- [x] **Shape fill / border (PPTLIVE-007/008)** — `Shape.set_fill(fill=/line=/
  line_width=)` sets the fill and border (a color, or `"none"` for transparent /
  no border) — distinct from `format_text`'s font `color`; `fill=`/`line=`/
  `line_width=` also ride on `add_shape`/`add_textbox`. Every shape listing now
  carries `fill`/`line` (`{color, visible[, weight]}`), guarded by the same
  theme-sentinel `color_hex_or_none` as font color (so a theme-linked fill reads
  back `None`, never a wrong `#000000`).
- [x] **Z-order (PPTLIVE-008)** — `Shape.reorder("front"|"back"|"forward"|
  "backward")` restacks via `Shape.ZOrder` (`MsoZOrderCmd`). Because z-order is
  what `shape:S:N` indexes, this is why that anchor is resolved live and never
  cached.
- [x] **`shapeid:S:ID` anchor (PPTLIVE-010)** — `ShapeById` resolves by stable
  `Shape.Id` (emitted as `id` in every shape listing), the **delete-proof** handle
  that survives a delete/restack which shifts `shape:S:N`. The drift-proof forms
  are now `ph:S:KIND`, `.Name`, and `shapeid:S:ID`.
- [x] **Composite-text recolor (PPTLIVE-009)** — a SmartArt diagram / chart has no
  text anchor, so `format_text` can't reach its internal labels.
  `SmartArt.recolor_text(color)` walks `AllNodes`; `Chart.recolor_text(color)`
  sets every shown chart text element (legend / both axis tick labels / title /
  per-series data labels) plus the `ChartArea` default. Coarse "recolor all text
  to X" only (the dark-theme fix), guarded by `HasLegend`/`HasTitle`/best-effort
  axes. Deferred: composite-text *fill* and per-element targeting.
- [x] **Library + CLI (`shape fill`, `shape order`, `shapeid:` everywhere a shape
  anchor is accepted, `chart recolor-text`, `smartart recolor-text`) + MCP
  (`ppt_edit` `format` fill keys / `shape_order` / `chart_recolor_text` /
  `smartart_recolor_text`).** Unit-tested in `tests/test_styling.py` +
  `tests/test_cli.py` + `tests/test_mcp.py`. See `roadmap.md`.

## v1.3 — review loop: comments — SHIPPED

The highest-leverage roadmap tier and the one whose COM risk was fully burned down
by `scripts/comments_spike.py` (2026-06-07, net-zero) before any build. Comments are
the PowerPoint diff from wordlive's range-anchored `_comments.py`: they attach to a
**slide** at an `(x, y)` point, are **threaded**, and binding an author needs the
signed-in Office-account identity.

- [x] **`_comments.py`** — `Comment` (`author`/`author_initials`/`text`/`datetime`
  (ISO)/`left`/`top`, threaded `replies`, `reply`, `delete`, `to_dict`) + per-slide
  `CommentCollection` (`Slide.comments`; 1-based `[i]`/iter, `add`/`list`).
  `Presentation.comments()` is the deck-wide roll-up `{total, slides:[...]}`.
- [x] **Identity** — `add` lifts the modern `Comments.Add2(Left, Top, Author,
  Initials, Text, ProviderID, UserID)` identity off any existing comment
  (`_discover_identity`), falling back to the legacy identity-free `Comments.Add` on
  a comment-less deck; a `reply` lifts identity straight off its parent. Honest
  caveats: `Add2` binds to the account (the passed author/initials only reach the
  legacy path), and `Comment.Status`/`.Resolved` aren't COM-readable, so **no
  resolve verb** ships. Both verified in the spike.
- [x] **CLI** — a `comment` group: `list [--slide S]` (per-slide / deck roll-up),
  `add --slide --text [--left --top --author --initials]`, `reply --slide --index
  --text`, `delete --slide --index`, all through `deck.edit(...)` (one Ctrl-Z).
- [x] **MCP** — `ppt_read` op `comments` (`slide?`); `ppt_edit` ops `comment_add` /
  `comment_reply` / `comment_delete`. Flow through `ppt_batch` via `_read_core` /
  `_edit_core`.
- [x] **Tests** — fake `Comments` collection in conftest (threaded, identity-bound,
  legacy-`Add` fallback path); `tests/test_comments.py` + `tests/test_mcp.py`
  coverage. See `roadmap.md` §v1.3.

**Deferred:** resolve/reopen (not COM-readable); account-identity sourcing for a
comment-less deck (the legacy `Add` covers it; option (a) needs its own micro-spike);
mentions / rich-text comment bodies.

## v1.1 — output: deck snapshot (low-res whole-deck render) — SHIPPED

The first cut of the v1.1 output tier, and the cheapest high-leverage feature left.
Ported from wordlive's `_snapshot.py` (the 2026-06-09 commit) but **shorter**: where
wordlive routes `ExportAsFixedFormat` → PDF → PyMuPDF, PowerPoint's `Slide.Export`
already renders a sized PNG (verified v0.4), so a deck snapshot needs no PDF detour
and no new dependency. The token lever is `max_dim`, a long-edge pixel cap — a vision
model is billed on pixel *area* (not DPI), so capping the long edge is a predictable,
and (since every slide shares one geometry) *uniform*, per-slide budget.

- [x] **`_snapshot.py`** — `Snapshot(slide, image, path)` dataclass; `_capped_dims`
  (long-edge cap at the 96-DPI native scale, never upscales); `render` (per slide →
  `Slide.export_image` to a temp, read bytes, unlink), `build_snapshots` (write
  files: single → `out`, multiple → `<stem>-sN<suffix>`), `snapshot` (compose).
- [x] **`Presentation.snapshot(out=None, *, slides=None, fmt="png", max_dim=None)`**
  — `slides` = `None` (all) / `int` (one) / `(start, end)` inclusive. A read; no
  `edit()` fence. The folder-based `export_images` (v0.4) stays for bulk-to-disk.
- [x] **CLI** — `snapshot [--slide N | --slides A-B] [--out PATH] [--max-dim N]
  [--format]`: path per slide with `--out`, base64 inline otherwise.
- [x] **MCP** — `ppt_render` op `deck_snapshot` (`{slides?, max_dim?, fmt?}`) returns
  one "slide N" label + image block per slide (embed default `max_dim` ~1000),
  reusing `_render_reply`; paths only in `structuredContent` (no base64
  double-encode). Flows through `ppt_batch` via `_render_core`.
- [x] **Tests** — `tests/test_snapshot.py` (cap math, selection, file placement,
  CLI path-vs-base64) + `tests/test_mcp.py` `deck_snapshot` coverage. The fake
  `Slide.Export` already writes a dim-encoding stub PNG, so no conftest change.
  See `roadmap.md` §v1.1.

**Still open in v1.1:** a `jpg`-quality / per-slide size knob on the snapshot
(the snapshot lever is `max_dim` only).

## v1.1 — output: save & PDF export — SHIPPED

Completes the output tier: the explicit, never-implicit file-output verbs. Three
spike findings shaped it (see `scripts/save_export_spike.py`,
`export_pdf_argforms_spike.py`, `saveas_pdf_dirty_spike.py`):
(1) `ExportAsFixedFormat` won't marshal under the late-bound `_com` dispatch —
*every* arg form (named, positional, `…2`) raises `TypeError: Python instance can
not be converted to a COM object` — so PDF rides `SaveAs(path, ppSaveAsPDF=32)`;
(2) `SaveAs`-to-PDF writes a faithful PDF *without* rebinding the working file or
clearing its dirty flag (a true export); (3) PowerPoint's `Save()` does **not**
raise on a never-saved deck (it silently cloud-saves on OneDrive/SharePoint
builds), so the never-saved guard lives in Python on an empty `Presentation.Path`.

- [x] **`constants.py`** — `PpSaveAsFileType` (`OPEN_XML_PRESENTATION=24`,
  `PDF=32`) + `save_format_for(fmt)` (resolves `"pptx"`; rejects `"pdf"` → pointer
  to `export_pdf`).
- [x] **`exceptions.py`** — `UnsavedPresentationError(PptliveError)` (exit 1),
  exported from `__init__`.
- [x] **`Presentation`** — `saved` property (`bool(Presentation.Saved)`);
  `save()` (guards empty `.Path` → `UnsavedPresentationError`, else `Save()`);
  `save_as(path, *, fmt="pptx", overwrite=False)` (`SaveAs(abspath, 24)`, rebinds,
  `FileExistsError` on clobber); `export_pdf(path)` (`SaveAs(abspath, 32)`, a read).
  `PresentationCollection.list()` now emits `saved` (so `status` shows it).
- [x] **CLI** — top-level `save`, `save-as PATH [--format pptx] [--overwrite]`,
  `export-pdf PATH` (positional path, mirrors wordlive); `_fmt_status` flags
  `(unsaved)`. `FileExistsError` → clean stderr + exit 1.
- [x] **MCP** — `ppt_render` ops `save` / `save_as` / `deck_pdf` (+ `overwrite`
  param). `_render_reply` now only embeds image-format paths, so a PDF/pptx `path`
  rides in `structuredContent` only — never mis-encoded as an inline image (matters
  in `ppt_batch`, which feeds all render results through `_render_reply`).
- [x] **Tests** — `tests/test_save_export.py` (library: saved/save/save_as/
  export_pdf rebind-vs-read, never-saved guard, overwrite; CLI: all three verbs +
  status). `tests/test_mcp.py` extended (deck_pdf no-embed, save/save_as, batch
  PDF-doesn't-break-embedding). Fake `_FakePresentation` gained `Path`/`Saved`/
  `Save`/`SaveAs` mirroring the verified COM behavior. Live wrapper smoke confirmed
  the whole path net-zero on a throwaway deck.

## v1.6 — text-model reliability & safe authoring — SHIPPED (2026-06-10)

Source: the gpt-5.4 MCP-session review (`docs/reviews/gpt-5.4-review.md`, 2026-06-10).
No new object-model coverage — this hardens the *existing* text/format surface
against the PowerPoint sharp edges that leak through. The reviewer's catastrophe:
`line_spacing: 24` expecting *24 pt* gave **24× line spacing** (text off the slide).

- [x] **Spike RESOLVED (2026-06-10, `scripts/text_model_spike.py`, net-zero).**
  Three findings pin the build:
  1. **LineRule is the unit selector.** `ParagraphFormat.SpaceWithin` stores a bare
     number; the paired `LineRuleWithin` bool picks its unit — `msoTrue (-1)` ⇒
     **multiple/lines**, `msoFalse (0)` ⇒ **points**. `LineRuleBefore`/`LineRuleAfter`
     pair the same way with `SpaceBefore`/`SpaceAfter`. All three `LineRule*` flags
     **set cleanly and read back** (so a read can report the mode). Today's code never
     touches them — it writes `SpaceWithin` only, leaving the unit at whatever the
     paragraph already had. That's the footgun.
  2. **No true "clear formatting" primitive — re-setting `.Text` does NOT drop run
     overrides** (a vandalised 5 pt/bold range read back 5 pt/bold after re-setting
     the same text). So `reset_format` must **re-apply defaults explicitly**, not hope
     for inheritance. The matching `CustomLayout` placeholder *is* fully readable
     (geometry `Left/Top/Width/Height` + `TextFrame.TextRange.Font.Size`, matched by
     `PlaceholderFormat.Type`) — the body placeholder's layout default size (28 pt)
     equalled the live baseline, so the layout placeholder is the source of truth for
     both `reset_to_layout` (geometry) and `reset_format` (default size).
  3. **Autofit reads.** `TextFrame2.AutoSize` returns a clean `MsoAutoSize` int
     (classic `TextFrame.AutoSize` returns mixed `-2`, so prefer TextFrame2);
     `WordWrap` + the four `Margin*` (points) read off classic `TextFrame`.
     `Font.AutoScale` does **not** exist on this build → no direct shrink-% signal, so
     overflow-risk is a coarse mode-derived heuristic, not a measured extent.
- [x] **`line_spacing` disambiguation + guardrail.** `format_paragraph` keeps
  `line_spacing` (multiple → `SpaceWithin` + `LineRuleWithin=msoTrue`) and adds
  `line_spacing_points` (→ `msoFalse`); `space_before`/`space_after` now set
  `LineRuleBefore/After=msoFalse` (honest points) with `space_before_lines`/
  `space_after_lines` companions; the shared `apply_paragraph_format` carries all of
  it. Passing both forms of a pair → `ValueError`; a `line_spacing` multiple `> 5`
  (`LINE_SPACING_MULTIPLE_MAX`) → `ValueError` unless `force=True`. CLI gains
  `--line-spacing-points`/`--space-before/after-lines`/`--force`; MCP `format` gains
  the same. `cli/main._run` + MCP `_mcp_errors` now map a library `ValueError` to a
  clean exit 1 / `invalid_args` (instead of a traceback / 500).
- [x] **Extended paragraph read diagnostics.** `paragraph_to_dict` gains
  `space_before`/`space_after`/`line_spacing` as `{value, mode}` (mode off the paired
  `LineRule*` via `_spacing_dict`) + `run_sizes` (distinct per-run font sizes via
  `_run_sizes`, the mixed-run tell). Flows through `ppt_read` `anchor` automatically.
- [x] **`set_paragraphs([...])`** — `Anchor.set_paragraphs(items)` (str or
  `{text, list_type?, indent_level?, alignment?, line_spacing?, size?, ...}`); one item
  = one addressable `para:` (a newline inside an item folds to a soft break), per-item
  keys forwarded to `format_paragraph`/`format_text`/`apply_list`. CLI `set-paragraphs
  --json/--file`; MCP `ppt_edit` `set_paragraphs`.
- [x] **Recovery verbs.** `Anchor.reset_format()` resets paragraph *spacing* to clean
  single-spaced defaults (the only unambiguous reset — no COM "clear formatting"
  exists). `Shape.reset_to_layout()` restores a placeholder's geometry + default font
  size from the matching `CustomLayout` placeholder. CLI `reset-format` / `shape
  reset-to-layout`; MCP `text_reset_format` / `shape_reset_layout`.
- [x] **Text-frame / autofit diagnostics (read)** — `Shape.text_frame_status()` →
  `TextFrameStatus(autosize, word_wrap, margins, overflow_risk)` (autosize off
  `TextFrame2`; `overflow_risk` coarse + mode-derived). CLI `read text-frame-status`;
  MCP `ppt_read` `text_frame_status`; `TextFrameStatus` + `MsoAutoSize`/`autosize_name`
  added. Non-fatal `warnings` array on `format` edits (tiny font, big forced multiple,
  list on a single soft-break paragraph).
- [x] **Docs** — "PowerPoint text-model gotchas" section + formatting-field reference
  table + safe patterns added to both `_skill` guides.

## v1.0+ — defer

- [ ] Event sinks (`SlideShowNextSlide`, `WindowSelectionChange`); async wrapper.
- [ ] Transitions & animations; full layout authoring (add/rename `CustomLayouts`,
  place placeholders). (Whole-slide `Slide.Export` was promoted to v0.4; master /
  theme *styling* was promoted to v0.9 above.)

---

## `exec` batch ops — SHIPPED (2026-06-10)

**Shipped** as the CLI `exec --script ops.json` verb, on a **fastmcp-free dispatch
seam extracted from the MCP server**: `pptlive/_batch.py` now holds the four op
enums, the handler registries, every `_<tool>_*` handler, the `_<tool>_core`
dispatchers, and `run_batch(handle, deck, commands, *, atomic, stop_on_error,
label)`. `mcp/server.py` shrank to FastMCP tool wrappers over that seam (image
embedding + `_mcp_errors` mapping stay there); `cli/commands.py` imports `run_batch`
for `exec`. Invalid args now raise the native `BatchOpError(PptliveError)` (instead
of fastmcp's `ToolError`) so the base CLI never needs the `[mcp]` extra — the MCP
server maps `BatchOpError` → `ToolError`, the CLI → exit 1, and a failed op's
category token maps to the CLI exit code. `exec` defaults each op to the `edit`
tool, stops at the first failure unless `--continue`, and `--no-atomic` fences each
op separately. The op *names* are the live MCP `ppt_edit`/`ppt_read`/… ops (not the
older proposed list below). Net behaviour of `ppt_batch` is unchanged (it now calls
`run_batch`); the full `tests/test_mcp.py` is the regression net.

Same `{"label", "ops":[…]}` shape as wordlive. **The batch is one undo entry** —
an `exec` run is a single automation session, so it fences with
`StartNewUndoEntry` on entry and PowerPoint groups the rest: a 5-op script is one
Ctrl-Z (see Spike #1 / Open Q #2). On failure at op K, ops 1..K-1 are already
applied and sit in that one entry, so a single Ctrl-Z reverts the partial batch
(report the failing index, re-raise so the exit code maps). No `"tracked"` key
(PowerPoint has no Track Changes). Proposed op set: `add_slide`,
`delete_slide`, `duplicate_slide`, `move_slide`, `set_layout`, `set_text`,
`insert_paragraph`, `find_replace`, `set_notes`, `add_shape`, `move_shape`,
`resize_shape`, `delete_shape`, `set_cell`, `apply_style`, `format_paragraph`,
`apply_list`.

`shape:S:N` refs in a script resolve **live at the moment each op runs** — an
`add_shape` can shift later z-order indices, so address anything you didn't just
create by `ph:S:KIND` or `.Name`. Symbolic binding (`add_shape "as": "label"` →
`shape:@label`) is deferred (Open Q #3).

---

## Cross-cutting (any release)

- [ ] **HRESULT coverage** — start from wordlive's `_BUSY_HRESULTS`; widen as
  real `com_error`s show up in smoke runs (add the slide-show-running rejection).
- [ ] **Smoke fixtures** — a real `.pptx` checked in with known slides /
  placeholders / a table / notes, so smoke tests have a stable target.
- [ ] **`\n`-in-live-reads smoke spike** (carried over from the retired
  `REFACTOR-PLAN.md`). Confirm whether a raw `TextRange.Text` read ever contains a
  bare `\n` and, if so, whether COM `.Paragraphs()` treats it as a break. If it
  does, `_selection.read_selection`'s `\r`-only paragraph count would undercount and
  should add `\n` (but never `\v`, which is an explicit *soft* break). A live check,
  not a code edit — the current `\r`-only count is correct for `para:` addressing as
  far as the unit tests and prior spikes show.
- [x] **Docs** — MkDocs Material site under `docs/` (mirrors wordlive's setup:
  `mkdocs.yml`, `docs` extra in `pyproject.toml`, mkdocstrings autodoc of the
  public surface). Pages: Home (README include), Getting started, Concepts,
  Python API, CLI, MCP server, Cookbook (11 end-to-end recipes), Errors & exit
  codes, Design. `spec.md` stays the canonical design doc. Builds clean under
  `uv run --extra docs mkdocs build --strict`.
- [x] **Agent skills + self-bootstrapping** — pptlive ships **two** skills
  (`pptlive-cli` + `pptlive-python`) under `src/pptlive/_skill/<name>/SKILL.md`,
  loaded via `_guide.py` (bundled in the wheel by hatchling's default data-file
  inclusion). CLI: `llm-help [--python]` dumps a skill body to stdout (raw
  Markdown, like `--help`, which now points at it); `install-skill
  [--cli|--python] [--system] [--force]` writes them to `.agents/skills/`. The
  MCP server exposes them as `pptlive://guide` / `pptlive://guide/python`
  resources + server `instructions`.
- [x] **MCP install** — `install-mcp [--client claude-desktop|claude-code]
  [--directory DIR] [--config PATH] [--print] [--force]` merges an `mcpServers`
  entry (`uvx --from "pptlive[mcp]" pptlive-mcp`, or `uv run --directory DIR`
  for a checkout) into the client config. Plus a one-click `.mcpb` bundle under
  `mcpb/` (manifest validated; `uv` runtime; Windows-only `compatibility`),
  built with `@anthropic-ai/mcpb pack` at release. **Note:** keep the version in
  `mcpb/manifest.json` + `mcpb/pyproject.toml` in sync with the root
  `pyproject.toml` on each bump (no bumpversion wiring yet).

---

## Open questions (from spec.md — resolve, don't guess)

1. ~~**Name.**~~ **Resolved: `pptlive`.** The whole Bootstrap + v0 tree
   (`pyproject` name, `src/pptlive/`, `pptlive` script, imports) commits to it
   and ships green. Revisit only if the user prefers
   `pptwings`/`slidelive`/`decklive`/`livepptx` before the first public release.
2. ~~**The undo gap — biggest call.**~~ ~~**Resolved: per-op undo + honest
   docs.**~~ **RE-RESOLVED (2026-05-26 spike): `edit()` blocks are atomic.** The
   per-op-undo premise was *wrong* — see Spike #1. PowerPoint groups consecutive
   in-session COM edits into one undo entry by default, and `StartNewUndoEntry()`
   is a verified boundary primitive. So the documented model becomes: **a
   `deck.edit(label)` block is one Ctrl-Z** (it should call `StartNewUndoEntry()`
   on entry to fence the block cleanly), and `exec` scripts likewise collapse to a
   single undo entry. **Done:** (a) `EditScope.__enter__` calls
   `app.StartNewUndoEntry()` best-effort; (b) the "no atomic undo" wording is
   rewritten in `README.md`, `spec.md`, `_edit.py`, `__init__.py`, `_anchors.py`,
   the CLI help, and `CLAUDE.md`; (c) two unit tests pin the fence; (d) the
   cross-process caveat is **verified** — separate CLI invocations stay distinct
   undo entries (`xproc1`/`xproc2`). Nothing left open here.
3. ~~**`shape:` addressing stability.** z-order drifts when shapes are
   added/removed.~~ **Resolved: honest + defer.** `shape:S:N` stays int-only and
   canonical, resolved **live** at each op (no caching). Every shape listing also
   emits `name` and `id` (`Shape.Id`, stable across reorder) for re-identification,
   and docs steer agents to `ph:S:KIND` / `.Name` as the drift-proof forms. The
   symbolic-binding mechanism (`add_shape "as": "label"` → `shape:@label`
   re-resolved by `Shape.Id`) is **deferred** until a real create-then-edit batch
   workflow needs it.
4. **Multi-presentation scope.** First-class multi-deck, or single-active-deck
   with explicit `--doc` naming (wordlive's unresolved Q4)?
5. **Slide identity.** *Partly settled:* every `slides.list()` / `slide.read()`
   row now emits both the 1-based `index` and the stable `id` (`SlideID`), and
   `Slide.id` exposes it. Still open: whether `exec` re-resolves slide refs by
   `SlideID` to survive mid-batch reordering (deferred with `exec` itself).
6. **Test strategy.** Smoke suite on a Windows+PowerPoint runner + the mockable
   `fake_powerpoint` layer for unit-testing politeness/anchor logic. (Same answer
   as wordlive — treat as decided unless revisited.)

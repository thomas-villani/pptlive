# pptlive — guide for Claude

`pptlive` drives a **running** Microsoft PowerPoint instance from Python over COM
(pywin32) — `xlwings`, but for PowerPoint, and built for LLM agents. Windows-only.
It is the sibling of `../wordlive`: the same design, re-applied to PowerPoint's
2-D object model.

- **`spec.md`** is the design doc. It is written deliberately as *the diff
  against wordlive* — read wordlive's `spec.md` first if a section assumes it.
- **`IMPLEMENTATION.md`** tracks staged build progress (the analog of wordlive's
  `feature-plan.md`). Update it as milestones land.

## `../wordlive` is the reference implementation

pptlive copies wordlive's structure, error taxonomy, `EditScope` shape, CLI
contract, `_com` seam, and test approach **almost verbatim**. Before designing
or building anything, open the equivalent wordlive module
(`../wordlive/src/wordlive/`), then apply *only* the PowerPoint diff from
`spec.md`. Matching wordlive's established conventions matters more than
inventing new ones — when in doubt, do what wordlive does.

## Toolchain — use `uv` for everything

Never call `pip` or a bare `python`; always go through `uv`.

```
uv sync --extra dev                       # install deps + dev tools into .venv
uv run pytest                             # unit tests (fake COM; no PowerPoint needed)
uv run pytest -m smoke                    # smoke suite — needs PowerPoint installed
uv run ruff check . && uv run ruff format .
uv run mypy
uv run pptlive status                     # exercise the CLI
```

Dev pins Python 3.13 (`.python-version`); the **library targets 3.11+**. wordlive
is being bumped to the same floor in lockstep, so the sibling parity holds — 3.11+
syntax (`StrEnum`, `assert_never`, …) is fair game. `ruff` and `mypy` are
configured for `py311` in `pyproject.toml`.

## Module layout (built through v0.9 — see IMPLEMENTATION.md)

```
src/pptlive/
  __init__.py        public surface + __all__
  constants.py       typed IntEnums for Mso*/Pp*/Xl* magic constants (+ friendly string aliases)
  exceptions.py      PptliveError taxonomy; _decode_com_error / from_com_error / _BUSY_HRESULTS
  units.py           points / inches() / cm() helpers — never expose EMUs
  _com.py            the ONLY pywin32 seam: com_apartment, get_active_powerpoint,
                     launch_powerpoint, translate_com_errors, retry_on_busy
                     (tests monkeypatch the getters)
  _app.py            PowerPoint handle + attach() / connect()
  _presentation.py   Presentation (the wordlive Document analog) + PresentationCollection
  _slides.py         SlideCollection / Slide  (add [+placeholder geometry]/delete/duplicate/move_to/set_layout, notes, read(), geometry_report(), animations()/clear_animations())
  _shapes.py         ShapeCollection / Shape / ShapeById  (a Shape IS an Anchor when it has a text frame; geometry + fill/line + z-order + animate() + set_picture() re-source verbs)
  _anchors.py        Anchor base + Paragraph, Cell, Notes  (set_paragraphs / reset_format + line-spacing unit knobs [v1.6])
  _tables.py         Table / Cell  (a table is a shape; cell:S:N:R:C anchors; cell fill + borders, row/column-wise; add/delete row+column)  [v0.5/v-next]
  _charts.py         Chart       (a chart is a shape; data via embedded Excel)         [v0.7]
  _smartart.py       SmartArt    (a diagram is a shape; node tree read/set_nodes; recolor_text + per-node format_node)  [v0.8/v-next]
  _theme.py          Theme + Master  (deck-wide palette/fonts/text-styles/background + headers_footers)  [v0.9]
  _sections.py       SectionCollection (deck.sections; named slide spans, 1-based index)  [v0.6.0]
  _headersfooters.py HeadersFooters (shared; Slide.headers_footers override + Master.headers_footers default)  [v0.6.0]
  _findreplace.py    fuzzy match core (find_matches/normalize); find()/find_replace() on Presentation [v1.0]
  _comments.py       Comment / CommentCollection (slide.comments; threaded, identity-bound add/reply) [v1.3]
  _snapshot.py       Snapshot + deck.snapshot() — whole-deck low-res PNGs, max_dim token cap [v1.1]
  _batch.py          fastmcp-free dispatch seam: op StrEnums + handler registries + _<tool>_core
                     dispatchers + run_batch(); imported by BOTH cli `exec` and mcp/server [v1.0/v1.6]
  _selection.py      viewed-slide + Selection snapshot/restore
  _edit.py           EditScope — view/Selection preservation + atomic undo via StartNewUndoEntry (see below)
  _show.py           SlideShow control (deck.show)
  _guide.py          loads the bundled SKILL.md guides (cli/python); shared by CLI + MCP
  _skill/pptlive-cli/SKILL.md, _skill/pptlive-python/SKILL.md   the two agent skills
  cli/{__init__,__main__,main,commands}.py   + llm-help / install-skill / install-mcp
  mcp/{__init__,__main__,server}.py   thin FastMCP wrappers over _batch.py: five op-dispatch
                     tools (ppt_read/edit/render/show/batch) + image embed + pptlive://guide; pptlive[mcp]
mcpb/                one-click `.mcpb` bundle (manifest.json, pyproject.toml, src/server.py)
tests/conftest.py    fake_powerpoint fixture (MagicMock COM), no_powerpoint, real_powerpoint
```

`find()` / `find_replace()` (`_findreplace.py`) shipped in v1.0 — fuzzy traversal
search across shapes / table cells / notes; library + CLI (`find`, `replace
--find`) + MCP (`ppt_read` find, `ppt_edit` find_replace).

**Review comments (`_comments.py`) shipped in v1.3.** `slide.comments` is a
per-slide `CommentCollection` (1-based, `add`/`list`/`[i]`); `deck.comments()` is
the deck-wide roll-up. Comments attach to a **slide** at an `(x, y)` point (not a
text range) and are **threaded** (`Comment.replies` / `Comment.reply`). Adding
needs the signed-in Office-account identity (`ProviderID`/`UserID`): `add` lifts
it off any existing comment via the modern `Comments.Add2`, falling back to the
legacy identity-free `Comments.Add` on a comment-less deck; a reply lifts identity
off its parent. Two honest caveats baked in: `Add2` **binds to the account** (the
passed `author`/`initials` are best-effort — even the legacy `Add` binds to the
signed-in account on a modern build, so they may be ignored), and there is **no
resolve/reopen verb** — `Comment.Status`/`.Resolved` are not COM-readable on
current builds. Library + CLI (`comment list/add/reply/delete`) + MCP (`ppt_read`
`comments`; `ppt_edit` `comment_add`/`comment_reply`/`comment_delete`).

**Text-model reliability (v1.6) + the `exec` CLI (v1.0) shipped 2026-06-10.** v1.6
hardens the *existing* text/format surface (no new object-model coverage) against
PowerPoint's sharp edges — the headline footgun being `line_spacing`: it is a
**multiple**, so `line_spacing=24` meant 24× line height (text off the slide).
`format_paragraph` now keeps `line_spacing` (multiple → `SpaceWithin` +
`LineRuleWithin=msoTrue`) and adds `line_spacing_points` (→ `msoFalse`); the
`space_before`/`space_after` points-intent is made honest (sets
`LineRuleBefore/After=msoFalse`) with `space_before_lines`/`space_after_lines`
companions; passing both forms of a pair is a `ValueError`, as is a `line_spacing`
multiple `> 5` unless `force=True`. New verbs: `Anchor.set_paragraphs(items)` (one
item = one addressable `para:`, the safe bullet-list path), `Anchor.reset_format()`
(reset paragraph *spacing* to clean defaults — the only unambiguous reset, since
PowerPoint exposes no "clear formatting"), `Shape.reset_to_layout()` (restore a
placeholder's geometry + default font size from its `CustomLayout` placeholder),
and `Shape.text_frame_status()` → `TextFrameStatus` (autosize/wrap/margins/
overflow-risk). `paragraph_to_dict` gained `space_before`/`space_after`/
`line_spacing` as `{value, mode}` + `run_sizes` (the mixed-run tell); edits return a
non-fatal `warnings` array (tiny font, big forced multiple, list on a soft-break
paragraph). A library `ValueError` now maps to a clean CLI exit 1 / MCP
`invalid_args` instead of a traceback. The **`exec` CLI verb** (the last
specced-but-unbuilt item) applies a `{"label", "ops":[…]}` script as one Ctrl-Z; it
runs on **`pptlive/_batch.py`** — a fastmcp-free dispatch seam (the op enums,
handler registries, `_<tool>_core` dispatchers, and `run_batch`) extracted from
`mcp/server.py` so the base CLI never needs the `[mcp]` extra. Invalid args raise
the native `BatchOpError` (the MCP server maps it to `ToolError`, the CLI to exit
1). All four front-ends + both SKILL guides updated.

**Deck snapshot (`_snapshot.py`) shipped in v1.1.** `deck.snapshot(out=None, *,
slides=None, fmt="png", max_dim=None)` renders slides to PNG so a vision model can
*see* the whole deck cheaply — the token-cost-aware read. The lever is `max_dim`,
a **long-edge pixel cap**: a model is billed on pixel *area* (not DPI), so capping
the long edge gives a predictable per-slide budget, and since every slide shares
one geometry that budget is *uniform* across the deck (~1000 px stays legible).
It's the PowerPoint analog of wordlive's snapshot but **shorter** — `Slide.Export`
already renders a sized PNG, so there's no PDF/PyMuPDF detour and no new dependency
(it reuses `Slide.export_image`). Returns one `Snapshot(slide, image, path)` per
slide; `slides` is `None` (all) | `int` (one) | `(start, end)` inclusive; with
`out` it writes files (single → that path, multiple → `<stem>-sN<suffix>`). A
**read** — the export leaves the viewed slide + Selection untouched, so no
`deck.edit()` fence. Library + CLI (`snapshot --slide/--slides/--out/--max-dim`) +
MCP (`ppt_render` op `deck_snapshot`, returns a "slide N" label + image block per
slide). The folder-based `deck.export_images` (v0.4) stays for bulk-to-disk.

**Save + PDF export (v1.1) completed the output tier (2026-06-09).** Three
explicit, never-implicit verbs on `Presentation` (pptlive never auto-saves):
`deck.save()` (persist to the existing file), `deck.save_as(path, *, fmt="pptx",
overwrite=False)` (write + **rebind** the working file), and `deck.export_pdf(path)`
(write a PDF — a **read**: no rebind, dirty flag preserved). Plus a `deck.saved`
dirty-flag property, surfaced (with `path`) on every `status` deck row. Three
spike findings shaped the design: (1) `ExportAsFixedFormat` — the nominal PDF API
— **won't marshal under the late-bound `_com` dispatch** (a trailing object-typed
param raises `TypeError`), so PDF rides `SaveAs(path, ppSaveAsPDF=32)`, which
produces a faithful PDF *without* rebinding the working file or touching its dirty
flag (a true export); (2) PowerPoint's `Save()` does **not** raise on a never-saved
deck — on a OneDrive/SharePoint build it silently uploads to the user's default
cloud folder — so `save()` guards in Python on an empty `Presentation.Path` and
raises `UnsavedPresentationError` (exit 1) instead; (3) `save_as` to `.pptx` (24)
*does* rebind the open deck to the new file (matching PowerPoint's Save-As).
`save_as` refuses to clobber unless `overwrite=True` (`FileExistsError`, surfaced
clean at CLI/MCP). Library + CLI (`save`, `save-as PATH [--format/--overwrite]`,
`export-pdf PATH`) + MCP (`ppt_render` ops `save`/`save_as`/`deck_pdf`).
Constants: `PpSaveAsFileType` (`OPEN_XML_PRESENTATION=24`, `PDF=32`) +
`save_format_for`.

**Authoring-feedback round (2026-06-18) — four fixes from a live Claude Desktop
session.** All driven by the "build a deck while watching PowerPoint" workflow:
(1) **"Follow the work" view policy** — the long-standing "jumps back to slide 1
after a batch" report was the politeness *view-restore* firing as designed, then
*cascading*: every batch restored to the pre-batch slide, so slide 1 became a fixed
point. Fix lives in `run_batch` (`_batch.py`): when an atomic batch **adds** a slide
(`slide_add`/`slide_duplicate`), it leaves the view on the last slide it touched
(via the existing `EditScope.allow_view_move()` opt-out + an explicit `go_to`)
instead of snapping back; pure-edit batches keep the polite restore, and a
deliberate `navigate`/`show` still wins. Configurable: default on, off via
`PPTLIVE_VIEW_FOLLOW=0` (env), MCP `ppt_batch(follow_view=False)`, or CLI
`exec --no-follow-view`. (2) **`geometry` read** — `Slide.geometry_report()` returns
the slide size + each shape's bounding `box` + `off_slide` flag + the `overlaps`
pairs (biggest first), so an agent catches overlaps / off-edge shapes *without* a
render. Axis-aligned only (rotation reported, not accounted for). MCP `ppt_read
op="geometry"`, CLI `slide geometry N` (named `geometry`, **not** `layout`, to avoid
colliding with `layouts`/`set-layout`). (3) **`shapeid` everywhere** — every shape
read (`shape_to_dict`) *and* mutation return now echoes the restack-proof
`shapeid:S:ID` (new `Shape.shapeid`) next to `anchor_id`, so a chained edit survives
the z-order drift it causes (the reported `shape_order` footgun). (4) **Placeholder
geometry on `slide_add`** — optional `placeholders={KIND: {left,top,width,height}}`
(points, any subset) repositions the layout's placeholders in the same op (the
"body on the left half beside a right panel" case), killing the add-then-resize
fix-up; validated pre-COM (clean `ValueError`/`invalid_args`), echoes the resulting
geometry. All four wired library + CLI + MCP + both SKILL guides. Still open from
that session: **direct-vs-inherited font color** in `text_frame_status` (no general
COM direct-vs-theme flag exists — needs a live spike first).

**Shape animations (v0.10) shipped 2026-06-18 — the v1.5 long tail deferred at
v0.4.0.** The sibling of slide transitions: whole-shape entrance/exit effects over
`Slide.TimeLine.MainSequence.AddEffect`. `Shape.animate(effect="fade", *,
trigger="on_click", duration=None, delay=None, exit=False)` appends one effect;
`exit=True` makes the shape animate **out** (the "disappear" ask — the same
`MsoAnimEffect` ids serve entrance and exit, the `Effect.Exit` flag is the only
difference). `Slide.animations()` reads the `MainSequence` back as ordered
`{seq_index, shapeid, shape, effect, exit, trigger, duration, delay}` rows (each
mapped to its target by the drift-proof `shapeid:S:ID` off `Effect.Shape.Id`), and
is folded into `Slide.read()`. `Slide.clear_animations(anchor=None)` wipes the whole
slide or one shape's effects (tail-first delete so live indices don't shift);
`Shape.clear_animations()` delegates to it. Curated `MsoAnimEffect`
(appear/fade/fly_in/float_in/wipe/zoom/grow_turn/swivel/wheel/split) +
`MsoAnimTriggerType` (on_click/with_previous/after_previous) follow the
`entry_effect_for` pattern with raw-int passthrough. A confirmation spike
(`scripts/animation_curated_spike.py`, net-zero) verified the full curated set
round-trips `EffectType` and that `Timing.TriggerDelayTime` (the `delay` knob)
round-trips before hardening. Library + CLI (`shape animate`/`shape
clear-animations`/`slide animations`/`slide clear-animations`) + MCP (`ppt_edit`
`shape_animate`/`shape_clear_animations`/`slide_clear_animations`; `ppt_read`
`animations`) + both SKILL guides. Deferred: per-paragraph levels, motion paths,
`EffectParameters`, in-sequence reordering.

**Deck-structure & feedback batch (v0.6.0) shipped 2026-06-18 — four roadmap
items de-risked together by `scripts/batch2_spike.py` (net-zero).** (1) **Sections**
(`_sections.py`, `deck.sections`) over `Presentation.SectionProperties`: `list`/`add(
name, before_slide=)`/`rename`/`delete(*, delete_slides=False)`/`move`, 1-based
section index. The spike pinned the model — `AddBeforeSlide` starts a span at a slide
and **auto-creates a leading "Default Section"** when it's the first section in front
of a later slide; `Delete` keeps slides unless told otherwise. (2) **Headers/footers**
(`_headersfooters.py`, a shared `HeadersFooters`) mounted at `Slide.headers_footers`
(per-slide override) and `Master.headers_footers` (deck-wide default), with
`read`/`set_footer`/`set_slide_number`/`set_date`; the spike footgun is handled —
`Footer.Text`/`DateAndTime.UseFormat` only read while the element is **visible**
(hidden → "Invalid request"), so reads are guarded (null when hidden) and setting text
auto-shows. (3) **Direct-vs-inherited font color** — the open Claude Desktop ask is
answered: `ColorFormat.Type` distinguishes a run color *set on the run* (`rgb`) from
one *cascaded from the theme/master* (`scheme`), so `font_to_dict` now emits
`color_source` (`direct`/`theme`/`mixed`) + `theme_color` (the inherited slot). (4)
**Snapshot size override** — `deck.snapshot(width=, height=)` for exact per-slide
pixels (overrides `max_dim`; both is a `ValueError`); the spike confirmed **JPEG
quality is not COM-exposable** on `Slide.Export`, so pixel dimensions stay the only
render-cost lever. All wired CLI (`section …`, `slide`/`master` `headers-footers`/
`set-footer`/`slide-number`/`set-date`, `snapshot --width/--height`) + MCP (`ppt_read`
`sections`/`headers_footers`; `ppt_edit` `section_*`/`set_headers_footers`;
`ppt_render` `deck_snapshot` width/height) + both SKILL guides. New constants
`MsoColorType`/`color_source_name`/`theme_color_name`. Still open: per-element
color-source on `text_frame_status`, friendly `PpDateTimeFormat` names.

**Table cell styling (v-next) — the "post-creation edit surface" round.** Motivated
by an audit of what can be *restyled* after creation (vs. wordlive's delete-and-
recreate habit). The headline finding **overturns a prior assumption**: a
`scripts/cell_style_spike.py` probe proved PowerPoint's COM *does* expose cell
fill (`Cell.Shape.Fill`) and borders (`Cell.Borders(index)`) — both round-trip — so
table-cell shading is **not** a COM dead-end after all. Built on that:
`Table.set_fill(fill, *, rows=None, cols=None, transparency=None)` (solid cell
shading, or `fill="none"` to clear) and `Table.set_border(*, color=None, weight=None,
dash=None, edges="all", rows=None, cols=None, visible=None)`, both **row/column-wise**
— `rows`/`cols` are `None` (whole axis) | int (one) | list (several) and the
*intersection* is styled (so `rows=1` shades the header row, `cols=2` a column, both
`None` the whole table). `Cell.set_fill`/`Cell.set_border` are thin per-cell
delegations. The `Borders(index)` edge order (1 top / 2 left / 3 bottom / 4 right /
5 diagonal-down / 6 diagonal-up) was pinned **visually** via
`scripts/cell_border_map_spike.py` (distinct colors per index → exported PNG), and
lives as `constants.border_edges_for` (`"all"` = the four sides; diagonals opt-in by
name). `Table.read()` cells now echo `fill` (theme-sentinel-guarded → `None`, never a
wrong `#000000`, like shape fills) — but note the live caveat: a default table
**style** writes a real per-cell banded RGB, so a fresh untouched cell reads back
that style color, not `None` (no COM flag separates a direct cell fill from a
style-cascaded one — it's OOXML-only; the read reports the effective rendered fill). Wired library + CLI (`table set-fill`/`set-border`
with `--rows/--cols` taking `1` or `1,3`) + MCP (`ppt_edit` `table_set_fill`/
`table_set_border`) + both SKILL guides. `set_fill` reuses the shape `apply_shape_fill`
helper; colors/edges/indices all validate before any COM mutation. Still open on the
table surface: cell merge state (no COM read property — OOXML-only) and per-cell
text-direction/margins.

**Re-source a picture in place (v-next) — the next post-creation edit-surface
item.** `Shape.set_picture(path, *, alt_text=None)` swaps an existing **picture's**
image without the wordlive delete-then-recreate dance, **preserving position / size /
rotation / name / alt text / z-order slot**. A `scripts/set_picture_spike.py` probe
(pixel-sampling the exported shape — PowerPoint exports a solid image as a *1-bit
palette PNG*, which the spike's decoder handles) settled the mechanism: PowerPoint's
COM has **no in-place image swap** for a picture shape — `Fill.UserPicture` set
`Fill.Type=6` (msoFillPicture) but the rendered pixel stayed the old color, i.e. it
fills *behind* the unchanged raster (so `set_picture_fill` is the wrong verb for a
real picture; it stays for autoshape fills). So `set_picture` is a delete + re-insert
that copies everything addressable; the spike pinned the three pitfalls the wrapper
handles: pictures default to **locked aspect** (copy the box with `LockAspectRatio`
off, else width/height snap to the new image's ratio), the old delete **drifts
z-order** (re-resolve the new shape by stable `Shape.Id`, never an index), and the old
**z-order slot** is restored by send-to-back-then-step-forward (`BringForward`=2, not
3). Two honest caveats baked in: the picture gets a **new `Shape.Id`** (so it returns
a fresh `ShapeById` handle — the old wrapper is spent, like after `delete()`), and
**animations / hyperlinks / crop / picture adjustments** bound to the old picture are
**not** carried over. Non-picture shape → `ValueError` (pointing at `set_picture_fill`);
missing file → `FileNotFoundError` (both before any COM mutation). Wired library
(`Shape.set_picture`, `_shapes.is_picture`/`replace_picture`) + CLI (`shape set-picture
--path [--alt-text]`) + MCP (`ppt_edit` `shape_set_picture`, echoes the new `shapeid` +
`geometry`) + both SKILL guides. Still open on the edit surface: shape-type swap.

**Table columns + SmartArt per-node text (v-next) — two edit-surface gaps closed
together.** A single net-zero spike (`scripts/table_col_smartart_node_spike.py`)
pinned both mechanisms against live PowerPoint. (1) **Table columns** complete the
row/column symmetry: `Table.add_column(values=None, *, before=None)` (over
`Columns.Add()` to append / `Columns.Add(n)` to insert before column `n`; `values`
fill top-to-bottom) and `Table.delete_column(index)` (over `Columns(n).Delete()`),
both mirroring `add_row`/`delete_row` exactly — same `AnchorNotFoundError` (kind
`"table column"`) on out-of-range, same `ValueError` refusing to delete the last
column (PowerPoint has no zero-column table). (2) **SmartArt per-node text** —
`SmartArt.format_node(index, *, bold/italic/underline/size/font/color)` formats one
node's label, the per-node companion to the coarse `recolor_text`. The spike proved
**`AllNodes` enumerates in exactly depth-first order**, so `read()` now stamps each
node with a 1-based `node_index` (its depth-first / `AllNodes` position) that
`format_node` takes as the address. A node's text lives on `TextFrame2`, whose
`Font2` differs from the classic `Font` (`_anchors.apply_font`): color is on
`Font.Fill.ForeColor.RGB` and underline is the `UnderlineStyle` enum
(`constants.MsoTextUnderlineType`), handled by a separate `_apply_node_font` helper;
`color` validates and `index` bounds-checks (→ `AnchorNotFoundError` kind
`"smartart node"`) before any COM. Wired library + CLI (`table add-column`/
`delete-column`, `smartart format-node`) + MCP (`ppt_edit` `table_add_column`/
`table_delete_column`/`smartart_format_node`) + both SKILL guides. Still open on the
edit surface: shape-type swap, table cell merge, SmartArt node *fill* / structural
node add-delete.

Agent skills shipped as **two** guides (`pptlive-cli` + `pptlive-python`), not
wordlive's single one — `llm-help [--python]` dumps one, `install-skill` writes
them to `.agents/skills/`, and `install-mcp` / the `mcpb/` bundle wire up MCP.

## Conventions (inherited from wordlive — keep them)

1. **`.com` escape hatch on every wrapper.** Each wrapper (`PowerPoint`,
   `Presentation`, `Slide`, `Shape`, `Anchor`, …) exposes a `.com` property
   returning the raw COM object. We never block a caller because we haven't
   wrapped something.
2. **`_com.py` is the only place that touches pywin32.** Everything else sees
   duck-typed dispatch objects. This is the mockable seam: tests monkeypatch
   `_com.get_active_powerpoint` / `_com.launch_powerpoint` to inject a fake.
   Don't `import win32com` anywhere else.
3. **Wrap COM calls in `with _com.translate_com_errors():`** so
   `pywintypes.com_error` becomes a typed `PptliveError`. Reuse wordlive's
   `_decode_com_error` / `from_com_error` / `_BUSY_HRESULTS` verbatim.
4. **`from __future__ import annotations` at the top of every module.**
5. **Points throughout, never EMUs.** PowerPoint's COM layer is points
   (1 in = 72 pt). EMUs are a `python-pptx`/OOXML concern and must never surface.
   Provide `pl.units.inches()` / `pl.units.cm()` helpers.
6. **Structured I/O.** Reads return dataclasses/dicts. The CLI prints exactly
   **one JSON object on stdout** per invocation (logs to stderr), with
   deterministic exit codes. No string scraping. Global flags (`--json`/`--text`,
   `--doc NAME`) go *before* the subcommand.
7. **Constants are typed `IntEnum`s, added only as a feature needs them.** Don't
   pre-populate. Friendly string aliases (`"title"`, `"two_content"`,
   `"textbox"`) coerce to the right int the way wordlive's alignment names do.

## The three things PowerPoint changes vs. Word — read before coding

1. **Atomic undo works after all — just differently. PowerPoint has no
   `UndoRecord`.** Word's `EditScope` brackets a block with
   `Application.UndoRecord` (start/end); PowerPoint has no such bracket. **But the
   2026-05-26 spike (`scripts/undo_test.py`) found PowerPoint already groups
   consecutive in-session COM edits into one undo entry by default**, and
   `Application.StartNewUndoEntry()` is a verified *boundary* primitive. So
   `edit()` calls `StartNewUndoEntry()` on entry to fence the block, and the whole
   block is **one Ctrl-Z** — near-parity with wordlive. Honest caveat: there's no
   explicit "end" fence (the next `edit()` or a user action closes it), so always
   wrap mutations in `deck.edit(...)`. Cross-*process* edits (separate CLI calls)
   are verified to stay distinct undo entries — each re-fences at its own `edit()`
   entry. (`_edit.py` is where the fence lives.)
2. **PowerPoint must be visible.** `Application.Visible = False` raises in most
   builds, so `connect()` has no `visible=False` mode. Politeness is about *not
   moving the user's view*, not about working hidden.
3. **Anchors are hierarchical (slide → shape → paragraph), not a global offset.**
   There is no document-wide character stream and no deck-wide `range:`. Anchor
   ids are colon-separated, slide-index first:

   | anchor_id        | resolves to |
   | ---------------- | ----------- |
   | `slide:S`        | slide S (1-based) — a **container**, not a text anchor; use `deck.slides[S]`, not `anchor_by_id` |
   | `shape:S:N`      | Nth shape (1-based z-order) on slide S — canonical handle; an `Anchor` if it has a text frame |
   | `shapeid:S:ID`   | shape with stable `Shape.Id` ID on slide S — the **delete-proof** handle (the `id` in every shape listing); survives a delete/restack that shifts `shape:S:N` |
   | `ph:S:KIND`      | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — the LLM-preferred form |
   | `para:S:N:P`     | paragraph P in shape N on slide S |
   | `cell:S:N:R:C`   | cell (row R, col C) of the table in shape N on slide S |
   | `notes:S`        | speaker-notes body of slide S |
   | `comments:S`     | the review comments on slide S — a **read selector** (like `slide:S`, a container, not a text `Anchor`); read via `deck.slides[S].comments`, address one for reply/delete by `(slide, 1-based index)` |

   `body` also aliases the generic **content** placeholder (`PpPlaceholderType.OBJECT`,
   reads back as `placeholder: "object"`). When a KIND matches **more than one**
   placeholder on a slide (the two content bodies of a Two Content / Comparison
   layout), `_find_placeholder` raises `AmbiguousMatchError` (exit 5) listing the
   candidate `shape:S:N` anchors rather than silently picking the first — reach each
   one by `shape:S:N`/`.Name`.

   `shape:` is int-only (z-order) to avoid index-vs-name ambiguity; expose shape
   `.Name` separately (`slide.shapes["Title 1"]`). **z-order drifts** when shapes
   are added/removed *or restacked* (`Shape.reorder`): resolve `shape:S:N` **live**
   (never cache it), and have every shape listing emit `name` + `id` (`Shape.Id`,
   stable across reorder). The drift-proof forms are `ph:S:KIND`, `.Name`, and
   `shapeid:S:ID` (`ShapeById`, resolves by stable `Shape.Id` — the delete-proof
   handle that PPTLIVE-010 asked for; built in the pt3 styling round). (Resolved
   Open Q #3 — symbolic `exec` binding deferred.)

   **Shape styling (pt3 round):** `Shape.set_fill(fill=/line=/line_width=)` sets the
   fill / border (a color, or `"none"` for transparent / no border) — distinct from
   `format_text`'s font `color`; `fill=`/`line=` also ride on `add_shape`/`add_textbox`.
   `Shape.reorder("front"|"back"|"forward"|"backward")` restacks via `Shape.ZOrder`.
   Every shape listing now carries `fill`/`line` (`{color, visible[, weight]}`), with
   the same theme-sentinel guard as font color (`color_hex_or_none` → `None`, never a
   wrong `#000000`).

   **Composite-text recolor (PPTLIVE-009, resolved):** a SmartArt diagram / chart has
   no text anchor, so `format_text` can't reach its internal labels. `SmartArt.recolor_text(color)`
   walks `AllNodes` setting each node's `TextFrame2.TextRange.Font.Fill.ForeColor.RGB`;
   `Chart.recolor_text(color)` sets every **shown** chart text element — legend / both
   axis tick labels / title / per-series data labels (classic chart `Font.Color` is the
   RGB long directly, not a `ColorFormat`; `DataLabels` is a method) plus the `ChartArea`
   global default. Coarse "recolor all text to X" only (the dark-theme fix); guarded by
   `HasLegend`/`HasTitle` and best-effort axes (`HasAxis` is an Excel-ism PowerPoint's
   chart COM rejects — a pie's absent axes are skipped, not an error). CLI `chart
   recolor-text` / `smartart recolor-text`; MCP `chart_recolor_text` / `smartart_recolor_text`.
   Still deferred: composite-text *fill* (node-shape / series fill) and per-element
   (vs. whole-shape) targeting.

## Politeness model (the whole point)

Default behavior preserves the user's **viewed slide**, **shape/text Selection**,
and focus. Operations that *must* move what the user sees say so in their name
(`go_to`, `show.goto`, `allow_view_move()`). Never target the live `Selection`
unless explicitly asked — write text through `Shape.TextFrame.TextRange.Text`
directly so no edit needs to select anything.

## Error taxonomy → exit codes

`PptliveError` base; reuse wordlive's COM-error decoding. Exit codes:
`0` ok · `1` other · `2` anchor/slide/shape/presentation not found (incl. zero
`find` matches) · `3` PowerPoint busy / slide show running · `4` PowerPoint not
running · `5` ambiguous match · `6` shape has no text frame (`NoTextFrameError`,
the one genuinely new code). `SlideNotFoundError` subclasses
`AnchorNotFoundError` (so it reuses exit 2).

## Testing

- Unit tests run against a **`fake_powerpoint`** MagicMock fixture in
  `tests/conftest.py` (model it on wordlive's `fake_word`) — no PowerPoint
  needed, runs on any OS in CI. This is where the politeness/anchor logic is
  proven.
- Live COM behavior goes behind `@pytest.mark.smoke`, skipped by default
  (`addopts = -m 'not smoke'`); the `real_powerpoint` fixture skips if PowerPoint
  isn't reachable. Run with `uv run pytest -m smoke` on a Windows box.
- Resolve spike questions (see IMPLEMENTATION.md) against real PowerPoint before
  hardening the corresponding code.

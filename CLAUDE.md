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
  _slides.py         SlideCollection / Slide  (add/delete/duplicate/move_to/set_layout, notes, read())
  _shapes.py         ShapeCollection / Shape / ShapeById  (a Shape IS an Anchor when it has a text frame; geometry + fill/line + z-order verbs)
  _anchors.py        Anchor base + Paragraph, Cell, Notes
  _tables.py         Table / Cell  (a table is a shape; cell:S:N:R:C anchors)         [v0.5]
  _charts.py         Chart       (a chart is a shape; data via embedded Excel)         [v0.7]
  _smartart.py       SmartArt    (a diagram is a shape; node tree read/set_nodes)      [v0.8]
  _theme.py          Theme + Master  (deck-wide palette/fonts/text-styles/background)  [v0.9]
  _findreplace.py    fuzzy match core (find_matches/normalize); find()/find_replace() on Presentation [v1.0]
  _comments.py       Comment / CommentCollection (slide.comments; threaded, identity-bound add/reply) [v1.3]
  _snapshot.py       Snapshot + deck.snapshot() — whole-deck low-res PNGs, max_dim token cap [v1.1]
  _selection.py      viewed-slide + Selection snapshot/restore
  _edit.py           EditScope — view/Selection preservation + atomic undo via StartNewUndoEntry (see below)
  _show.py           SlideShow control (deck.show)
  _guide.py          loads the bundled SKILL.md guides (cli/python); shared by CLI + MCP
  _skill/pptlive-cli/SKILL.md, _skill/pptlive-python/SKILL.md   the two agent skills
  cli/{__init__,__main__,main,commands}.py   + llm-help / install-skill / install-mcp
  mcp/{__init__,__main__,server}.py   five op-dispatch tools (ppt_read/edit/render/show/batch)
                     + pptlive://guide resources; pptlive[mcp]
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
`comments`; `ppt_edit` `comment_add`/`comment_reply`/`comment_delete`). Still in
`spec.md` but unbuilt: the standalone CLI `exec` batch verb (MCP `ppt_batch`
covers batch).

**Deck snapshot (`_snapshot.py`) shipped in v1.1.** `deck.snapshot(out=None, *,
slides=None, fmt="png", max_dim=None)` renders slides to PNG so a vision model can
*see* the whole deck cheaply — the token-cost-aware read. The lever is `max_dim`,
a **long-edge pixel cap**: a model is billed on pixel *area* (not DPI), so capping
the long edge gives a predictable per-slide budget, and since every slide shares
one geometry that budget is *uniform* across the deck (~1000 px stays legible).
It's the PowerPoint analog of wordlive's snapshot but **shorter** — `Slide.Export`
already renders a sized PNG, so there's no PDF/PyMuPDF detour and no new dependency
(it reuses `Slide.export_image`). Returns one `Snapshot(slide, png, path)` per
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
`save_format_for`. Still in `spec.md` but unbuilt: the standalone CLI `exec` verb.

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

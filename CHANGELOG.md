# Changelog

All notable changes to **pptlive** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **`write` (`set_text`) now treats `\n` as a real paragraph break** (PPTLIVE-001).
  An LLM building a bullet body with `"a\nb\nc"` previously got **one** paragraph
  full of soft line breaks (`<a:br>`), so the lines were not individually
  addressable as `para:S:N:P`. `\n` / `\r\n` / `\r` are now all normalized to a
  paragraph break, so each line is its own addressable paragraph. A within-paragraph
  soft line break is still available — embed `\v` (`pptlive._anchors.SOFT_BREAK`).
  Docs across the MCP `write` op, the CLI `--text` help, and both SKILL guides were
  corrected (they previously mislabeled `\n` as "paragraphs").

### Fixed

- **CLI `shape add --kind shape --text X` now applies the text.** It was silently
  dropped (only the `textbox` branch passed `--text` through), while the MCP
  `shape_add` op already set it — a CLI/MCP drift. Autoshapes created with `--text`
  now carry that text.
- **MCP `ppt_render` `save_as` now honors a `save_format` argument** (default
  `"pptx"`), matching the CLI's `save-as --format`. It previously hard-coded the
  format and always reported `"pptx"`. An unrecognized format surfaces as
  `invalid_args` rather than an unclassified error.
- **Ambiguous fuzzy `replace` now follows the standard CLI failure contract.** It
  used to print a JSON error object on **stdout** *and* exit 5, unlike every other
  failure (which writes only to stderr). The contract is now uniform: stdout JSON
  on success only; the actionable "N matches — set occurrence/all" hint goes to
  stderr with exit 5. The MCP path is unchanged (it still returns the structured
  `matches` for an agent to retry on).
- **Theme/master color reads no longer emit a wrong `#000000` for a theme-linked
  color.** `Master` text-style and background color reads now route through the
  same `color_hex_or_none` theme-sentinel guard that font and shape fill/line reads
  already used, so an inherited/automatic color reads back as `null`, not black.

### Added

- **Save & PDF export — the explicit file-output verbs** (v1.1, completing the
  output tier). Three never-implicit verbs on `Presentation` (pptlive never
  auto-saves): `deck.save()` persists to the existing file; `deck.save_as(path, *,
  fmt="pptx", overwrite=False)` writes a `.pptx` and **rebinds** the working file to
  it (the open deck becomes that file, matching PowerPoint's Save-As), refusing to
  clobber unless `overwrite=True`; `deck.export_pdf(path)` writes a pixel-faithful
  PDF and is a **read** — it neither rebinds the working file nor clears its dirty
  flag, so your `.pptx` is untouched. A `deck.saved` dirty-flag property joins
  `path` on every `status` deck row (`{name, path, saved, is_active}`), so an agent
  sees unsaved state before deciding to save. `save()` on a never-saved deck raises
  `UnsavedPresentationError` (exit 1) rather than letting PowerPoint silently route
  the file to the user's default cloud folder (a verified `Save()` behavior on
  OneDrive/SharePoint builds). PDF goes through `SaveAs(path, ppSaveAsPDF=32)`:
  `ExportAsFixedFormat` is the nominal API but won't marshal under pptlive's
  late-bound COM dispatch, and `SaveAs`-to-PDF produces the same faithful PDF as a
  pure export. CLI `save` / `save-as PATH [--format pptx] [--overwrite]` /
  `export-pdf PATH`; MCP `ppt_render` ops `save` / `save_as` / `deck_pdf`. New
  `PpSaveAsFileType` constant (`OPEN_XML_PRESENTATION=24`, `PDF=32`).
- **Deck snapshot — whole-deck low-resolution render for vision models** (v1.1).
  `deck.snapshot(out=None, *, slides=None, fmt="png", max_dim=None)` renders slides
  to PNG so a vision model can *see* the whole deck cheaply — the token-cost-aware
  read. The lever is `max_dim`, a **long-edge pixel cap** (only ever lowering
  resolution): a model is billed on an image's pixel *area*, not its DPI, so capping
  the long edge gives a predictable per-slide token budget — and because every slide
  shares one geometry, that budget is *uniform* across the deck (~1000 px stays
  legible for "did my styling land"). The PowerPoint analog of wordlive's snapshot
  but shorter: `Slide.Export` already renders a sized PNG, so there's no PDF/PyMuPDF
  detour and no new dependency. Returns one `Snapshot(slide, png, path)` per slide;
  `slides` is `None` (all) / an `int` (one) / a `(start, end)` inclusive span; with
  `out` it writes files (single → that path, multiple → `<stem>-sN<suffix>`),
  otherwise the bytes ride in `.png`. A read — leaves the viewed slide and Selection
  untouched (no `edit()` fence). CLI `snapshot --slide/--slides/--out/--max-dim
  --format` (path per slide with `--out`, base64 inline otherwise); MCP `ppt_render`
  op `deck_snapshot` (`{slides?, max_dim?, fmt?}`) returns one "slide N" label +
  image block per slide, defaulting `max_dim` to ~1000 px when embedding.
- **Review comments — read + add/reply/delete** (v1.3, the review loop).
  PowerPoint's review-comment channel, across library + CLI + MCP. Comments attach
  to a **slide** at an `(x, y)` point (not a text range) and are **threaded**:
  `slide.comments` is a per-slide `CommentCollection` (1-based; `add`/`list`/`[i]`,
  each comment carrying `author`/`initials`/`text`/`datetime`/`left`/`top` and its
  `replies`), `Comment.reply(text)` appends to the thread, `Comment.delete()` removes
  it, and `deck.comments()` is the deck-wide roll-up. Adding binds to the signed-in
  Office account: `add` sources the modern `Comments.Add2` `ProviderID`/`UserID` off
  any existing comment and falls back to the legacy identity-free `Comments.Add` on a
  comment-less deck (the passed `author`/`initials` are best-effort — `Add2` binds to
  the signed-in account, and on a modern Office build even the legacy `Add` does, so
  they may be ignored). MCP `ppt_read` op `comments`; `ppt_edit` ops `comment_add` /
  `comment_reply` / `comment_delete`. CLI `comment list/add/reply/delete`. No
  resolve/reopen verb — `Comment.Status`/`.Resolved` are not COM-readable on current
  builds (documented).
- **Shape fill & border color** (PPTLIVE-007). `Shape.set_fill(fill=/line=/line_width=)`
  sets a shape's solid fill and/or border (a `#RRGGBB` color, an `(r, g, b)` tuple,
  a raw RGB int, or `"none"` for transparent fill / no border) — the spatial
  complement to `format_text`'s *font* color. `fill=`/`line=`/`line_width=` also ride
  on `add_shape`/`add_textbox`. Exposed as MCP `ppt_edit` op `format`
  (`fill_color`/`line_color`/`line_width`) and `shape_add`, and CLI `shape fill` +
  `shape add --fill/--line/--line-width`. Every shape read now reports `fill` and
  `line` (`{color, visible[, weight]}`), with the same theme-sentinel guard as font
  color (`color_hex_or_none` → `null`, never a wrong `#000000`).
- **Shape z-order control** (PPTLIVE-008). `Shape.reorder("front"|"back"|"forward"|
  "backward")` restacks a shape via `Shape.ZOrder` and returns its new 1-based
  position — so a freshly added background panel can slide *behind* existing content
  (previously every new shape landed on top, forcing a destructive delete-and-rebuild).
  MCP `ppt_edit` op `shape_order` (`order=`); CLI `shape order --to`.
- **`shapeid:S:ID` — a delete-proof shape anchor** (PPTLIVE-010). `slide.shapes.by_id(ID)`
  / `anchor_by_id("shapeid:S:ID")` resolves a shape by its stable `Shape.Id` (the `id`
  already in every shape listing). Unlike `shape:S:N` — a z-order index that shifts
  down when a lower shape is deleted or restacked — a `shapeid` keeps pointing at the
  same shape across structural edits. Resolves live, so it also survives reorder.
- **Composite-text recolor for SmartArt & charts** (PPTLIVE-009). A SmartArt diagram
  / chart has no text anchor, so `format_text` couldn't reach its internal labels —
  on a dark (or any custom-background) theme the inherited black node / axis / legend
  text went invisible with no in-place fix, forcing a rebuild from primitives.
  `SmartArt.recolor_text(color)` recolors every node label; `Chart.recolor_text(color)`
  recolors every **shown** chart text element (legend, both axis tick labels, title,
  per-series data labels) plus the `ChartArea` global default. Coarse "recolor all text
  to X" only, and only what's already displayed (guarded by `HasLegend`/`HasTitle`;
  axes/data-labels best-effort, so a pie chart's absent axes are skipped). MCP `ppt_edit`
  ops `chart_recolor_text` / `smartart_recolor_text`; CLI `chart recolor-text` /
  `smartart recolor-text`. Composite-text *fill* and per-element targeting remain
  deferred — drop to `.com` for those.
- **Placeholder ambiguity guard** (PPTLIVE-004). On Two Content / Comparison
  layouts (two generic `object` content placeholders), `ph:S:body` used to silently
  resolve to the *first* one. It now raises `AmbiguousMatchError` (exit 5 / MCP
  `ambiguous`) listing the candidate `shape:S:N` anchors, consistent with
  `find_replace`'s guard. A more-preferred placeholder type still wins over a
  less-preferred one (so a real `body` beats a generic `object`); only an *equal*
  best-rank tie is ambiguous.
- **Richer effective font in `ppt_read` op `anchor`** (PPTLIVE-003). Each paragraph
  now carries a `font` block — `bold`/`italic`/`underline` as `true`/`false`/
  `"mixed"` (the `msoTriStateMixed` signal `is_true` used to discard), `size`,
  `font` name, and `color` (`#RRGGBB`, or `null` for an inherited theme/automatic
  color). These are *effective* (rendered) values; COM exposes no general per-run
  "directly set vs inherited" flag (only color distinguishes a literal RGB from a
  theme color) — documented honestly, validated by `scripts/inherit_probe.py`.
- **`PPTLIVE_VIEW_DEBUG` env var** traces what `snapshot`/`restore` capture (with
  the thread name) to stderr — a zero-overhead diagnostic for "view jumps to slide
  1" reports in MCP hosts we can't attach a debugger to.

### Fixed

- **A deliberate `navigate` / `show` inside an atomic `ppt_batch` is no longer
  snapped back** to the pre-batch slide on scope exit (the batch's single
  `EditScope` now opts out of the view restore once a view-moving command runs).
  Standalone `ppt_render navigate` was already correct (no enclosing scope).
- **`find` `context` snippet** now renders paragraph/line separators as visible
  glyphs (`⏎` / `↵`) instead of flattening them to spaces (PPTLIVE-006); offsets
  are unaffected.

### Changed (ergonomics)

- **`master_format_text_style` / `master_format_paragraph_style` `level` now
  defaults to `1`** (library, CLI, and MCP) — the natural choice for `title`, which
  has a single level. Previously omitting it was an error.
- **Ambiguity error wording is surface-neutral** (PPTLIVE-005): it names both the
  MCP params (`occurrence=N` / `replace_all=true`) and the CLI flags
  (`--occurrence` / `--all`), instead of only the CLI flags.
- **Every MCP tool description now contains "PowerPoint"** plus its action verbs
  (PPTLIVE-002), so a `tool_search("powerpoint")` surfaces all five tools (it used
  to find only `ppt_edit` / `ppt_batch`).

### Docs

- Documented the generic `object` content placeholder kind and the `body`→`object`
  alias (CLAUDE.md anchor table + both SKILL guides), and the chart series ordering
  rule (insertion order; bar charts render bottom-to-top by Excel convention — not
  a reorder).
- Documented shape fill/border, z-order, and the `shapeid:S:ID` handle across the
  docs site (`concepts.md`, `cli.md`), both SKILL guides, and CLAUDE.md. Noted the
  one styling gap **not** yet closed: SmartArt-node and chart-internal **text color**
  remain unaddressable (PPTLIVE-009) — recolor needs rebuilding those composites from
  primitives for now.

> **Note on the recurring "view jumps to slide 1" report:** the fix that landed in
> 0.1.3 (COM apartment held open) is intact, and the current source preserves the
> view under every tested path (in-process and the real `pptlive-mcp` stdio server
> — see `scripts/view_repro.py` / `scripts/view_stdio_repro.py`). If a Claude
> Desktop install still snaps to slide 1 on every action, it is running a **stale
> bundle environment** predating 0.1.3 — `uv cache clean pptlive` and reinstall the
> extension (a version bump forces a fresh resolve).

## [0.2.0] — 2026-06-08

### Added

- **Fuzzy find / replace across the deck — the last wordlive surface-parity gap.**
  `find` and `find_replace` are now live on the library (`Presentation.find` /
  `find_replace`), the CLI (`find`, `replace --find`), and MCP (`ppt_read` op
  `find`, `ppt_edit` op `find_replace`; both also work inside `ppt_batch`).
  PowerPoint has no deck-wide character stream, so search is a **traversal** of
  every text frame — shapes, table cells, and speaker notes — and each hit is
  reported against a resolvable text anchor (`shape:S:N`, `cell:S:N:R:C`,
  `notes:S`) with a 0-based in-frame offset, plus a context snippet. `scope` (CLI
  `--in`) restricts the search to a `slide:S` or any text anchor.
  - Matching reuses wordlive's fuzzy core (NFKC + smart-quote / dash / NBSP folds
    + whitespace collapse), so text an LLM re-typed off a slide still matches the
    original glyphs; it is case-sensitive, like wordlive.
  - Replacement writes through `TextRange.Characters`, so only the matched span
    changes and the rest of the frame keeps its run formatting. Matches are
    computed once up front (not via a re-scanning native `.Replace` loop), which
    sidesteps both the first-only and the offset-drift hazards a replacement that
    re-contains the search text would otherwise trigger.
  - One match auto-applies; several without `--all` / `--occurrence` raise
    `AmbiguousMatchError` (exit 5, listing the matches); zero matches raise
    `AnchorNotFoundError` (exit 2). `find` itself never raises on zero — it
    returns an empty list. The pre-existing `replace --anchor-id` whole-anchor
    form is unchanged.
  - Grounded by a live, net-zero COM spike (`scripts/findreplace_spike.py`).

## [0.1.3] — 2026-06-04

### Fixed

- **MCP server no longer jumps the user's view to the title slide (and no longer
  crashes).** `com_apartment()` previously did a balanced
  `CoInitialize`/`CoUninitialize` on every `attach()`. That is harmless for
  one-shot CLI processes, but the long-lived MCP server re-`attach()`es on every
  tool call, firing `CoUninitialize` repeatedly on its event-loop thread. That
  destabilises pythoncom: it drops PowerPoint's automation connection — snapping
  the active window back to slide 1 — and, under repetition, corrupts COM proxy
  state into a hard segfault (reproduced within ~6 `attach()` cycles). COM is now
  initialised once per thread and never uninitialised (the OS reclaims it at
  thread/process exit), so the server holds one stable apartment across all its
  tool calls. Verified end-to-end: 12 real MCP tool calls keep the view fixed and
  no longer crash.

### Docs

- README: added a hands-on review of pptlive driven from Claude Desktop.

## [0.1.2] — 2026-05-29

### Added

- **`ppt_render` returns rendered images inline so remote MCP hosts can see
  them.** Image-producing ops (`slide_image` / `shape_image`, and the same
  commands inside `ppt_batch`) now return the pixels *through* the MCP call as a
  base64 `ImageContent` block, not just a filesystem path — so a hosted client
  (e.g. claude.ai talking to a local bundle) whose model runs in a separate
  sandbox can still complete the render → look → iterate loop. Both the inline
  image *and* the structured `path`/metadata are returned, so a co-located
  filesystem tool still has the path. The image is encoded exactly once (the
  structured content carries only the small path dict), verified to survive
  FastMCP's inferred output-schema validation. `slide_image` defaults to ~1024 px
  on the long edge to keep text-heavy slides cheap (override with `width`/`height`,
  or `embed=False` for path-only).

### CI

- Bumped the release workflow's GitHub Actions to their Node 24 majors.

## [0.1.1] — 2026-05-29

### Added

- Added an MIT `LICENSE` file and declared the license in the package metadata.

## [0.1.0] — 2026-05-29

Initial public release. `pptlive` drives a **running** Microsoft PowerPoint
instance from Python over COM (pywin32) — *xlwings, but for PowerPoint*, and
built for LLM agents. It is the PowerPoint sibling of
[`wordlive`](https://github.com/thomas-villani/wordlive), copying its structure,
error taxonomy, `EditScope` shape, CLI contract, `_com` seam, and test approach.

### Added

- **Live editing over COM**, with a politeness model that preserves the user's
  viewed slide, shape/text selection, and focus by default; only verbs that must
  move the view (`go_to`, `show.goto`, `allow_view_move()`) say so in their name.
- **Atomic undo** — `deck.edit(...)` fences a block with `StartNewUndoEntry()` so
  the whole block is one Ctrl-Z.
- **Slide lifecycle** — add / delete / duplicate / move / set-layout, with layout
  resolution.
- **Shapes & geometry** — add textbox / autoshape / picture; move / resize / delete.
- **Text structure** — paragraph anchors, insert, paragraph/font formatting, bullets.
- **Hierarchical anchors** — `slide:S`, `shape:S:N`, `ph:S:KIND`, `para:S:N:P`,
  `cell:S:N:R:C`, `notes:S` (slide-index-first, resolved live).
- **Render & live selection** — slide/shape export to PNG, selection read, and the
  `here:` anchor.
- **Tables** — `add_table`, `cell:S:N:R:C` anchors, table read / add-row / delete-row.
- **Live slide show control** — `deck.show`.
- **Pictures** — alt text and per-shape image export.
- **Charts** — `add_chart` and the `Chart` wrapper (data via embedded Excel).
- **SmartArt** — generate diagrams and read nodes back to reconstruct the tree.
- **Theme & master styling** — deck-wide palette, fonts, text styles, background.
- **CLI** — one JSON object per invocation on stdout, deterministic exit codes,
  plus `llm-help`, `install-skill`, and `install-mcp`.
- **MCP server** (`pptlive[mcp]`) — five op-dispatch tools
  (`ppt_read` / `ppt_edit` / `ppt_render` / `ppt_show` / `ppt_batch`) and
  `pptlive://guide` resources, for Claude Desktop and other MCP clients.
- **Agent skills** — two bundled guides (`pptlive-cli` + `pptlive-python`).
- **One-click `.mcpb` bundle** for installing the MCP server.
- **Docs site** — MkDocs Material, published to GitHub Pages on push to `main`.
- **Release automation** — `bump-my-version` syncs the root and MCPB bundle
  versions; a `v*` tag publishes to PyPI via trusted publishing.

[Unreleased]: https://github.com/thomas-villani/pptlive/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/thomas-villani/pptlive/compare/v0.1.3...v0.2.0
[0.1.3]: https://github.com/thomas-villani/pptlive/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/thomas-villani/pptlive/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/thomas-villani/pptlive/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/thomas-villani/pptlive/releases/tag/v0.1.0

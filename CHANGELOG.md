# Changelog

All notable changes to **pptlive** are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/thomas-villani/pptlive/compare/v0.1.3...HEAD
[0.1.3]: https://github.com/thomas-villani/pptlive/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/thomas-villani/pptlive/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/thomas-villani/pptlive/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/thomas-villani/pptlive/releases/tag/v0.1.0

# Design

This page gives the condensed rationale. The canonical, longer design document
is [`spec.md`](https://github.com/thomas-villani/pptlive/blob/main/spec.md) in
the repo root — written deliberately as *the diff against
[`wordlive`](https://github.com/thomas-villani/wordlive)* — and the staged
roadmap is in
[`IMPLEMENTATION.md`](https://github.com/thomas-villani/pptlive/blob/main/IMPLEMENTATION.md).

## Why pptlive exists

There is no good Python library for **driving a live Microsoft PowerPoint
session**. The options today are:

| Library         | Target                              | Mechanism             |
| --------------- | ----------------------------------- | --------------------- |
| `python-pptx`   | `.pptx` file on disk                | OOXML I/O             |
| **`pptlive`**   | **Running `POWERPNT.exe`**          | **COM (`pywin32`)**   |

File-side libraries can't help when the user has the deck open — PowerPoint
holds the lock, and any change you make on disk is invisible until they close
and re-open. The round-trip (close the deck, let the agent write, re-open to
see it, close again) is slow and error-prone. COM is the only live path. And
raw `pywin32` is brutally LLM-hostile: magic integer constants (`MsoShapeType`,
`PpPlaceholderType`, `XlChartType`), untyped late-bound dispatch, modal-dialog
footguns, STA threading rules.

`xlwings` exists for Excel; `wordlive` for Word. pptlive is the equivalent for
PowerPoint, with one extra goal: be *first-class* for LLM tool use, not
retrofitted.

## Design principles

The four principles, in priority order:

1. **Politeness first.** Default behaviour preserves the user's **viewed
   slide**, shape/text `Selection`, and focus. They keep editing — or
   presenting — alongside your script. Operations that *must* move the screen
   say so explicitly ([`deck.go_to(...)`](python-api.md#pptlive.Presentation),
   the [`deck.show`](python-api.md#pptlive.SlideShow) verbs,
   [`scope.allow_view_move()`](python-api.md#pptlive.EditScope)). Jumping the
   active slide is the PowerPoint equivalent of stomping the cursor — and more
   jarring, because it's a full-screen change.
2. **Semantic anchors over `Selection`.** Operations target hierarchical named
   handles — `ph:S:KIND`, `shape:S:N`, `para:S:N:P`, `cell:S:N:R:C`, `notes:S`
   — not the live selection. Anchors are stable enough to address in JSON and
   visible to an LLM as strings; the selection is neither.
3. **Atomic undo.** Every [`deck.edit()`](python-api.md#pptlive.Presentation)
   block fences a single undo entry via `StartNewUndoEntry`, so one Ctrl-Z
   reverts the whole intent — even though PowerPoint has no Word-style
   `UndoRecord`. See [Concepts](concepts.md#editscope-and-atomic-undo).
4. **Structured I/O.** Reads return dataclasses / dicts; the CLI emits one JSON
   object per invocation; exit codes are deterministic. No string scraping
   anywhere. See the [Errors page](errors.md#cli-exit-codes) for the contract.

Underlying all four: an **escape hatch**. Every wrapper exposes `.com`. When
pptlive doesn't cover something, drop to raw COM rather than giving up.

## The three things PowerPoint changes vs. Word

pptlive copies wordlive's structure almost verbatim. Three places where
PowerPoint's object model forced a divergence:

1. **Atomic undo, by a different mechanism.** Word brackets a block with
   `Application.UndoRecord`. PowerPoint has no such bracket — but it groups
   consecutive in-session COM edits into one undo entry by default, and
   `StartNewUndoEntry()` is a verified boundary primitive. So `edit()` fences
   on entry and the block is one Ctrl-Z. (One of three spec assumptions a live
   spike corrected.)
2. **PowerPoint must be visible.** `Application.Visible = False` raises in most
   builds (verified: *"Hiding the application window is not allowed"*). So
   `connect()` has no hidden mode — politeness is about *not moving the user's
   view*, not about working unseen.
3. **Anchors are hierarchical, not a global offset.** There is no
   document-wide character stream and no deck-wide `range:`. Addressing is
   slide → shape → paragraph, slide-index first. z-order drifts as shapes are
   added/removed, so `shape:S:N` is always resolved live and listings emit a
   stable `Shape.Id`, `Shape.Name`, and `alt_text` for re-identification.

## What's out of scope

- **Cross-platform support.** COM is Windows-only. We don't pretend otherwise.
- **Cloud co-authoring / PowerPoint on the web.** Microsoft Graph is a
  different stack and a different problem.
- **Full PowerPoint object-model coverage.** Anything we don't cover is one
  `.com` access away.
- **Replacing `python-pptx`.** Different surface, different problem. Generate a
  deck from scratch in a batch job? Use `python-pptx`. Edit the deck the user
  is looking at? Use pptlive.
- **Designer / AI layout suggestions.** Out of scope.

## Architecture at a glance

```
your code / LLM
       │
       ▼
┌───────────────────────────────────────────────────┐
│  pptlive public API                               │
│    attach / connect  →  PowerPoint                │
│                          │                        │
│                          ▼                        │
│                      Presentation                 │
│                          │                        │
│            ┌─────────────┼─────────────┐          │
│            ▼             ▼             ▼          │
│         slides        anchor_by_id   show         │
│            │             │                        │
│            ▼             ▼                         │
│          Slide  →  shapes → Shape ── is-a ──┐     │
│                              │              ▼     │
│                              ▼          Anchor    │
│                       Table / Chart   (text, set_text,
│                                        format_text, …) │
└───────────────────────────────────────────────────┘
                          │
                          ▼
            EditScope (StartNewUndoEntry + SelectionSnapshot)
                          │
                          ▼
     pywin32  →  PowerPoint.Application (COM, STA-threaded)
```

The library is intentionally flat: ~15 modules, no plugin system, no hierarchy
beyond PowerPoint → Presentation → Slide → Shape → Anchor. `_com.py` is the
only module that touches pywin32, which is what makes the whole surface
unit-testable against a fake-COM fixture with no PowerPoint installed.

## What comes next

The roadmap lives in
[`IMPLEMENTATION.md`](https://github.com/thomas-villani/pptlive/blob/main/IMPLEMENTATION.md).
The current release covers the politeness / anchors / `EditScope` core, the
LLM-first CLI, the slide lifecycle, shapes & geometry, text structure
(paragraphs, font formatting, bullets), tables (cells as `cell:S:N:R:C`
anchors), slide render + live selection (`here:`), pictures (alt text + per-shape
export), charts (embedded-Excel data), and live slide-show control
(`deck.show`). It also ships an optional [MCP server](mcp.md) for Claude Desktop
and other MCP clients. Deferred: event sinks
(`SlideShowNextSlide`, `WindowSelectionChange`), an async wrapper, transitions &
animations, and master/layout authoring.

## Full design document

For the unabridged version — the original motivation, the error taxonomy in
more detail, the rejected alternatives, the resolved spikes, and the open
questions — see
[`spec.md`](https://github.com/thomas-villani/pptlive/blob/main/spec.md) in the
repo root.

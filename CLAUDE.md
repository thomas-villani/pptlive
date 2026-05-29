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

Dev pins Python 3.13 (`.python-version`), but the **library targets 3.10+** to
match wordlive — do not use 3.11+ syntax. `ruff` and `mypy` are configured for
`py310` in `pyproject.toml`.

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
  _shapes.py         ShapeCollection / Shape  (a Shape IS an Anchor when it has a text frame; geometry verbs)
  _anchors.py        Anchor base + Paragraph, Cell, Notes
  _tables.py         Table / Cell  (a table is a shape; cell:S:N:R:C anchors)         [v0.5]
  _charts.py         Chart       (a chart is a shape; data via embedded Excel)         [v0.7]
  _smartart.py       SmartArt    (a diagram is a shape; node tree read/set_nodes)      [v0.8]
  _theme.py          Theme + Master  (deck-wide palette/fonts/text-styles/background)  [v0.9]
  _selection.py      viewed-slide + Selection snapshot/restore
  _edit.py           EditScope — view/Selection preservation + atomic undo via StartNewUndoEntry (see below)
  _show.py           SlideShow control (deck.show)
  cli/{__init__,__main__,main,commands}.py
  mcp/{__init__,__main__,server}.py   five op-dispatch tools (ppt_read/edit/render/show/batch); pptlive[mcp]
tests/conftest.py    fake_powerpoint fixture (MagicMock COM), no_powerpoint, real_powerpoint
```

Not yet built (still in `spec.md`): `_findreplace.py` (`find()` / `find_replace()`)
and `_skill/SKILL.md` (the LLM-facing CLI reference).

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
   | `ph:S:KIND`      | placeholder of semantic KIND (`title`/`ctrtitle`/`subtitle`/`body`/`footer`/`date`/`slidenum`) — the LLM-preferred form |
   | `para:S:N:P`     | paragraph P in shape N on slide S |
   | `cell:S:N:R:C`   | cell (row R, col C) of the table in shape N on slide S |
   | `notes:S`        | speaker-notes body of slide S |

   `shape:` is int-only (z-order) to avoid index-vs-name ambiguity; expose shape
   `.Name` separately (`slide.shapes["Title 1"]`). **z-order drifts** when shapes
   are added/removed: resolve `shape:S:N` **live** (never cache it), and have
   every shape listing emit `name` + `id` (`Shape.Id`, stable across reorder) so
   an agent can re-identify after drift. Steer toward `ph:S:KIND` and `.Name` as
   the drift-proof forms; document the hazard honestly rather than building
   re-resolution machinery. (Resolved Open Q #3 — symbolic `exec` binding deferred.)

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

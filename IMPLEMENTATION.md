# pptlive — implementation tracker

Staged build plan + progress, ordered by **LLM-agent leverage** (the same
ordering principle as wordlive's `feature-plan.md`). The design lives in
`spec.md`; this file is the checklist. Update statuses as work lands and record
resolved open questions inline (strike them through, link the commit).

**Status legend:** `[ ]` not started · `[~]` in progress · `[x]` shipped.

> **Bootstrap + v0 + v0.1 have landed** (fake-COM unit tests green: `ruff`,
> `mypy`, `pytest` all pass; 92 tests). The library is usable as an LLM tool
> against a real, already-running PowerPoint, and now drives the **slide
> lifecycle** (`slide add/delete/duplicate/move/set-layout` + layout-name
> resolution) — verified live 2026-05-26 via `scripts/layout_spike.py` (net-zero;
> the v0.1 section records the findings). The four **Spike** items below were
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

## v0.2 — shapes & geometry (makes slide-building possible)

- [ ] `shape add` (textbox/picture/shape), `Shape.move(left, top)` /
  `resize(width, height)` / `delete()` / `geometry()` — all in **points**.
- [ ] `units.inches()` / `units.cm()` helpers.
- [ ] CLI `shape add|move|resize|delete`.

## v0.3 — text structure

- [ ] `para:S:N:P` anchors (`Paragraph` over a `TextRange`).
- [ ] `insert` before/after within a text frame; `apply_style`;
  `format_paragraph`; list verbs (`TextRange.ParagraphFormat` / `IndentLevel`).
- [ ] CLI `insert`, `style apply`, `format-paragraph`, list commands.

## v0.4 — tables

- [ ] `add_table`; `cell:S:N:R:C` anchors (`Cell` *is* an `Anchor`);
  `table read --slide S --shape N`. Shape must satisfy `HasTable`.

## v0.5 — live slide show control (the most literally "live" surface)

- [ ] `deck.show`: `start`/`end`/`next`/`previous`/`goto(n)`/`black()`/`white()`/
  `state()` over `SlideShowSettings.Run()` / `SlideShowWindow.View`.
- [ ] While a show runs, editing calls reject → surface as `PowerPointBusyError`
  (exit 3) and steer agents to the `show` group.
- [ ] CLI `show start|end|next|prev|black|white|state|goto`.

## v0.6 — pictures & charts

- [ ] `add_picture` (embed, never link); alt text as the LLM re-identification
  handle (wordlive v0.8 pattern).
- [ ] Image **extraction** for vision models (per-shape; wordlive v0.9 pattern).
- [ ] `add_chart` with an embedded-Excel data spike (wordlive v0.10 reasoning).

## v0.7+ — defer

- [ ] Event sinks (`SlideShowNextSlide`, `WindowSelectionChange`); async wrapper.
- [ ] Slide / thumbnail export (`Slide.Export`); transitions & animations;
  master/layout authoring.

---

## `exec` batch ops (lands with v0.1+)

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
- [ ] **Docs** — keep `spec.md` as design; a `cookbook.md` of end-to-end
  LLM-tool examples is likely more useful than API reference at this stage.
- [ ] **`SKILL.md`** — port wordlive's agent-facing CLI reference once v0 CLI is
  stable; add an `install-skill` command.

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

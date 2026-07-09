# pptlive — roadmap (remaining work)

Audit date: 2026-06-18. This is the single forward-looking roadmap — **only the
work still left to do**. The historical per-tier plan and spike findings for
already-shipped work now live inline in `CHANGELOG.md` and `IMPLEMENTATION.md`
(this file replaced the old post-v0.9 `roadmap.md`, which was removed when the two
were consolidated).

Audit baseline (as of the v0.6 release prep):

- Shipped since the last roadmap audit: **sections**, **headers/footers**,
  **whole-shape animations** (entrance/exit), **direct-vs-inherited font color**,
  and the **snapshot pixel-size override** — so those large surfaces are no longer
  open (see CHANGELOG `[Unreleased]` / `IMPLEMENTATION.md` §v0.10–v0.11).
- Fake-COM/unit coverage is healthy: `uv run pytest -q` passes (815 tests).
- Source audit found no shipped implementation yet for the remaining large
  surfaces below (media insertion/video export, run-level hyperlinks, tags,
  document properties, OLE, grouping, connectors, etc.).

**Status legend:** `[ ]` not started · `[~]` spiked / partially available · `[x]`
shipped. Spike-first remains the rule: confirm live PowerPoint COM behaviour,
record the finding, then harden library + CLI + MCP + tests together.

---

## Priority tiers

| Tier | Theme | Status | Why it remains |
| ---- | ----- | ------ | -------------- |
| **Linter / regularizer** | Consistency audit + one-pass autofix (`deck.lint()`/`regularize()`) — the wordlive linter, re-applied | `[ ]` design ✅ · proofing spiked ✅ | **Highest-leverage next feature.** Pure composition over shipped verbs; new work is `format_info()` + the rule engine. Design in `spec-linter.md`. See its own section below. |
| **v1.4-rest** | Navigation & structure: sections, headers/footers, run-level hyperlinks | `[x]` | Sections + headers/footers shipped (v0.6); **text-run-level hyperlinks shipped 2026-06-25** — tier complete. |
| **v1.5-rest** | Animations | `[x]` main cut · `[ ]` long tail | Whole-shape entrance/exit shipped (v0.10); per-paragraph levels / motion paths / reordering remain. |
| **v1.7** | Media + narrated-video export | `[x]` | SHIPPED 2026-06-25 — insert audio/video narration, self-time slides, export MP4 (async `CreateVideo`). Only the media long tail (trimming, bookmarks, recorded narration) remains. |
| **Opportunistic** | Tables/charts/SmartArt/arrangement/tags/metadata/OLE | mixed | **Arrangement (group/align/distribute/connectors) shipped 2026-06-25.** Pull the rest on demand. |
| **Deferred** | Async/events, full layout authoring, deep theme/master follow-ups | `[ ]` | Real but lower leverage or larger architectural lift. |

---

## Linter / regularizer — the consistency audit + one-pass autofix

The highest-leverage feature still open, and (like the wordlive linter it ports)
**pure composition** — `deck.lint()` audits a deck for presentation-quality defects,
`deck.regularize()` autofixes the mechanical ones in one atomic-undo pass, both over
verbs pptlive already ships. Full design: **`spec-linter.md`**; staged checklist:
**`IMPLEMENTATION.md` → "Linter / regularizer — PLANNED"**. The one-paragraph why:
the last hour before a deck ships is spent on objective, mechanical, already-scriptable
fixes (headers all one font/size, shapes lined up, no empty bullets, numeric columns
right-aligned, copyright/confidential present, slide numbers, slide size) — exactly what
an agent should own.

**The design diff from wordlive** (why this is a real port, not a copy): PowerPoint has
~no named paragraph styles, so consistency isn't Word's "direct override fighting the
style." It's three PowerPoint-native primitives — **P2 mode/dominant across peers** (all
titles alike; the headline), **P3 spatial regularity** (alignment/geometry, on the
shipped `geometry_report()`), and weakly **P1 placeholder-vs-master cascade**. Rules stay
`consistency` / `structural` / `policy` (profile-driven), and the `Finding`/`regularize`
engine + `adds_content` gate port near-verbatim from `_linting.py`.

- [ ] **Build the foundation + the two clusters the user named first** — `format_info()`
  read, the peer-mode helper, then `title-font-consistent` (P2) and `edge-alignment` /
  `shape-off-slide` (P3), then the `regularize` loop + idempotency test. Then P4 text
  (empty bullets, table numerics), P5 deck (notices, slide numbers, slide size) +
  profiles. Wire Python / CLI / **exec op** / MCP + both SKILL guides. Findings anchor by
  the drift-proof `shapeid:S:ID`.

- [~] **Proofing (`spelling`) — SPIKED 2026-07-08 (`scripts/proofing_spike.py`), a GO
  via a borrowed engine; later batch.** PowerPoint has a spell checker but exposes **no
  way to call it** — no `Application.CheckSpelling`/`GetSpellingSuggestions`, no
  `TextRange.SpellingErrors`. **The path: borrow a hidden `Word.Application` over COM**
  (Word's checker works on bare strings, runs `Visible=False`, and is near-ubiquitous
  beside PowerPoint). We tokenize each frame ourselves → **exact `(shapeid, start,
  length)` spans + suggestions**; both live test-deck typos caught. Build musts:
  - **Fail gracefully if Word isn't installed** (rare for a PowerPoint user, but real):
    a failed `Dispatch("Word.Application")` makes `lint` **skip the `spelling` cluster and
    say so in the report** — never a crash; only a `--rules spelling` that explicitly
    asked may error with a "needs Word" message.
  - **One scratch invisible doc per pass** (`Documents.Add()` — `GetSpellingSuggestions`
    needs a document; `CheckSpelling` doesn't), reused for all tokens, closed in `finally`.
  - **`_com.py` owns the only `Dispatch("Word.Application")`** (`word_speller()` helper),
    keeping the pywin32-seam rule intact; natural home for a future `deck.proofing()`.

---

## v1.4-rest — navigation & deck structure

Shape-level hyperlinks/actions are shipped. The remaining v1.4 work is structural.

- [x] **Sections.** SHIPPED 2026-06-18 (v0.6.0) — `deck.sections` (`SectionCollection`):
  `list`/`add(name, before_slide=)`/`rename`/`delete(*, delete_slides=False)`/`move`,
  library + CLI (`section …`) + MCP + tests. Spike `scripts/batch2_spike.py` (net-zero)
  pinned the index semantics (AddBeforeSlide auto-creates a leading "Default Section";
  Delete keeps slides). See IMPLEMENTATION.md §v0.11.
  - **COM surface:** `Presentation.SectionProperties`: `.Count`, `.Name(i)`,
    `.SlidesCount(i)`, `.AddSection(index, name)`, `.Rename`, `.Delete`, `.Move`.
  - **Wrapper shape:** `deck.sections` / `deck.sections()` read returning ordered
    `{index, name, first_slide, slide_count}` rows; add / rename / delete / move
    verbs.
  - **Front-ends:** CLI `section list|add|rename|delete|move`; MCP read/edit ops.
  - **Notes:** structural, no view move. Spike exact index semantics before build.

- [x] **Headers / footers / slide numbers / date.** SHIPPED 2026-06-18 (v0.6.0) —
  shared `HeadersFooters` at `Slide.headers_footers` (override) + `Master.headers_footers`
  (default): `read`/`set_footer`/`set_slide_number`/`set_date`, library + CLI + MCP +
  tests. Spike pinned the footgun (Text/UseFormat only read while visible → guarded).
  See IMPLEMENTATION.md §v0.11.
  - **COM surface:** `Slide.HeadersFooters`, `SlideMaster.HeadersFooters`:
    `.Footer.Text` / `.Visible`, `.SlideNumber.Visible`,
    `.DateAndTime.Format` / `.UseFormat` / `.Visible`.
  - **Wrapper shape:** deck/master-level defaults plus per-slide overrides, mirroring
    the existing master-vs-slide background split.
  - **Front-ends:** CLI `deck set-footer`, `slide set-footer`, `deck slide-numbers`
    or equivalent; MCP edit ops and read fields.
  - **Spike:** exact inheritance behaviour and whether applying to all slides uses a
    master setting, per-slide loop, or both.

- [x] **Text-run-level hyperlinks.** SHIPPED 2026-06-25 — `Anchor.set_link`/
  `remove_link`/`links` on the `Anchor` base (so `Shape`/`Paragraph`/`Cell`/`Notes`
  all carry them) over `TextRange.Characters(...).ActionSettings(ppMouseClick)
  .Hyperlink`. Address a span by substring `text=` or explicit `start`/`length`;
  destination is a `url=` or in-deck `slide=` jump. Library + CLI (`link
  set`/`remove`/`list`) + MCP (`ppt_edit` `link_set`/`link_remove`, `ppt_read`
  `links`) + both SKILL guides + tests. Spike `scripts/run_link_spike.py`
  (net-zero) confirmed the COM round-trips and that linking splits the runs. See
  IMPLEMENTATION.md §v-next.

---

## v1.5-rest — animations

Slide transitions are shipped. Animations were spiked live and the common cut is
de-risked; only the build remains.

- [x] **Whole-shape entrance / exit animations.** SHIPPED 2026-06-18 (v0.10) —
  `Shape.animate(effect, *, trigger, duration, delay, exit)` +
  `slide.animations()` + `Shape.clear_animations()` / `slide.clear_animations()`,
  curated `MsoAnimEffect`/`MsoAnimTriggerType`, library + CLI + MCP + tests + both
  SKILL guides. Confirmation spike `scripts/animation_curated_spike.py` (net-zero)
  verified the full curated effect set + the `delay` knob before hardening. The
  "shape disappear" line below lands as the `exit=True` flag. See IMPLEMENTATION.md.
  - **Spike status:** confirmed 2026-06-11 (`scripts/animation_spike.py`, net-zero).
    `Slide.TimeLine.MainSequence.AddEffect(Shape, EffectId, Level=0, Trigger)`
    round-trips `EffectType`, `Exit`, `.Shape.Id` / `.Shape.Name`, and timing fields.
  - **COM surface:** `Slide.TimeLine.MainSequence`, `Sequence.AddEffect`,
    `Effect.Delete`, `Effect.Timing.Duration`, `.TriggerType`, `.TriggerDelayTime`,
    `.Exit`.
  - **Wrapper shape:**
    - `Shape.animate(effect="fade"|"appear"|raw_int, *, trigger="on_click"|"with_previous"|"after_previous", duration=None, delay=None, exit=False)`
    - `slide.animations()` read mapping effects back to `shapeid:S:ID` / shape name.
    - `Shape.clear_animations()` and/or `slide.clear_animations(anchor=None)`.
  - **Constants:** curated `MsoAnimEffect` subset (`appear`, `fade`, maybe a few
    common emphasis effects) + raw-int passthrough; `MsoAnimTriggerType`.
  - **Front-ends:** CLI `shape animate`, `shape clear-animations`, `slide animations`;
    MCP read/edit ops.

- [x] Should add shape disappear as well — shipped as `animate(..., exit=True)`.

- [ ] **Animation long tail, only after the common cut lands.**
  - Per-paragraph animation levels (`Level`).
  - Motion paths.
  - `EffectParameters` and other effect-specific knobs.
  - Reordering animations in a sequence.

---

## v1.7 — media + narrated-video export — SHIPPED 2026-06-25

**SHIPPED** (library + CLI + MCP + both SKILL guides + tests; re-verified live,
net-zero). `Slide.add_audio` / `add_video` (over `Shapes.AddMediaObject2`, with
`autoplay` / `hide_icon` / `pace_slide` reusing `Slide.set_transition`),
`Shape.has_media` / `Shape.media` reads, and `deck.export_video(...)` /
`deck.video_status()` (over the async `Presentation.CreateVideo` /
`CreateVideoStatus` — blocking by default, `wait=False` + poll for non-blocking).
New `PpMediaType` / `PpMediaTaskStatus` constants + `VideoExportError`. CLI `media`
group + `export-video` / `video-status`; MCP `ppt_edit` `media_add`, `ppt_render`
`export_video` / `video_status`. See CHANGELOG `[Unreleased]`. The original
2026-06-11 spike (`scripts/media_video_spike.py`) proved the chain end-to-end with
no new dependency; what remains below is the deferred long tail.

- [x] **Insert media: `slide.add_audio(path)` / `slide.add_video(path)`.**
  - **Spike status:** confirmed via `Shapes.AddMediaObject2`.
  - **COM surface:** `Shapes.AddMediaObject2(FileName, LinkToFile=False,
    SaveWithDocument=True, Left, Top, Width, Height)`; read `Shape.MediaType`,
    `Shape.MediaFormat.Length`, `Muted`, `Volume`, `StartPoint`, `EndPoint`.
  - **Wrapper shape:** `Slide.add_audio(path, *, left=0, top=0, width=..., height=...,
    link=False, autoplay=True, hide_icon=True, pace_slide=True)` and video sibling.
  - **Notes:** resolve paths to absolute paths; document embed-vs-link file-size
    tradeoff.

- [x] **Auto-play + per-slide pacing.**
  - **COM surface:** `Shape.AnimationSettings.PlaySettings.PlayOnEntry`,
    `HideWhileNotPlaying`; reuse `SlideShowTransition.AdvanceOnTime` /
    `.AdvanceTime` for pacing.
  - **Wrapper shape:** `pace_slide=True` reads `MediaFormat.Length` and sets slide
    auto-advance to the clip duration.

- [x] **Export video: `deck.export_video(path)`.**
  - **Spike status:** `Presentation.CreateVideo(...)` marshals and produced a real
    MP4; it is async.
  - **COM surface:** `Presentation.CreateVideo(FileName, UseTimingsAndNarrations,
    DefaultSlideDuration, VertResolution, FramesPerSecond, Quality)` and
    `Presentation.CreateVideoStatus` (`None` / `Queued` / `InProgress` / `Done` /
    `Failed`).
  - **Wrapper shape:** `deck.export_video(path, *, use_timings=True,
    default_slide_duration=5, resolution=720, fps=30, quality=85, wait=False,
    timeout=None)` returning a task/status object; blocking `wait=True` convenience.
  - **Front-ends:** CLI `media add ...`, `export-video ...`; MCP `ppt_edit`
    `media_add` and `ppt_render` `export_video` / `video_status`.
  - **Build questions:** map failed encodes cleanly to errors; decide polling shape
    for MCP non-wait calls; document native recorded narration as a separate path not
    used by this tier. *(Resolved: failed/timeout → `VideoExportError`; `export_video`
    blocks by default with `wait=False` + `video_status()` for non-blocking polling.)*

- [ ] **Media long tail (deferred from the v1.7 cut).**
  - Trimming (`MediaFormat.StartPoint` / `EndPoint`), bookmarks, poster frames,
    volume/mute setters, video styling.
  - Native recorded narration (`Slide.NotesPage` / `RecordNarration`) — a separate
    capture path, not the file-insertion path shipped here.
  - WMV export (`SaveAs` 37) — MP4 via `CreateVideo` is the one path shipped.

---

## Output polish

The core output tier shipped (`snapshot`, save/save-as, PDF). Remaining small
polish:

- [x] **Snapshot/export image knobs.** SHIPPED 2026-06-18 (v0.6.0) — `deck.snapshot(
  width=, height=)` exact per-slide pixels (overrides `max_dim`), library + CLI
  (`--width`/`--height`) + MCP. JPEG quality **dropped**: the spike confirmed it's not
  COM-exposable on `Slide.Export` (no quality param / `ExportBitmapResolution` /
  `ExportConfiguration`), and pixel area — not encoder quality — is what a vision model
  bills on. `max_dim` long-edge cap stays the default token lever.

---

## Shape styling follow-ups

v1.2 styling is complete for practical 2-D shape appearance. These are long-tail
visual controls, not blockers.

- [ ] **3-D effects.** `Shape.ThreeD` / `SetThreeDFormat` preset, depth, bevels,
  rotation/material/lighting. Prior spike showed this is the genuine effects long
  tail; build only if requested by a workflow.

- [ ] **Fine-grained line geometry.** Per-side/per-corner line formatting and any
  shape-specific geometry controls not covered by current fill/line/effects verbs.

- [ ] **Non-solid per-slide / master / layout backgrounds.** Reuse the shipped
  advanced fill helpers for gradient / picture backgrounds, but spike exact support
  on `Slide.Background`, `SlideMaster.Background`, and `CustomLayout.Background`.

---

## Opportunistic — pull in on demand

- [ ] **Deeper tables.**
  - Merge/split (`Cell.Merge` / `.Split`).
  - Cell fill and borders (`Cell.Shape.Fill`, `Cell.Borders`).
  - Column width / row height.
  - Built-in table styles (`Table.ApplyStyle(styleId)`) and header/banding flags
    (`.FirstRow`, `.HorizBanding`).
  - Note: current reads intentionally do not expose merge state because PowerPoint
    COM lacks a clean merge-state read; spike before promising stable merged-region
    reads.

- [ ] **Deeper charts.**
  - Title / legend / axis content and geometry (`Chart.Axes(...)`, `.HasTitle`,
    title text, legend position, axis labels).
  - Per-element text targeting rather than coarse `Chart.recolor_text(color)`.
  - Series fill / line styling and data-label styling.

- [ ] **Deeper SmartArt.**
  - Node-shape fill and per-node styling.
  - Org-chart assistant / branch node handling from the v0.8 follow-ups.
  - Layout availability differences across Office versions and expanding beyond the
    current core layouts.

- [x] **Shape arrangement beyond z-order.** SHIPPED 2026-06-25 —
  `ShapeCollection.group`/`align`/`distribute`/`add_connector` + `Shape.ungroup`,
  library + CLI (`shape group`/`ungroup`/`align`/`distribute`/`connect`) + MCP +
  both SKILL guides + tests. Connectors ship in the **full shape-attached** form
  (`begin=`/`end=` glue via `BeginConnect`/`EndConnect` + `RerouteConnections`) with
  a geometry fallback. Spike `scripts/arrangement_spike.py` (net-zero) pinned the
  model: group gets a **new** `Shape.Id` (members keep theirs); **ungroup keeps the
  members' original ids** (no identity churn); `RerouteConnections` reassigns the
  connection sites (requested sites are advisory). See IMPLEMENTATION.md §v-next.
  - Still open (long tail): merged-region group reads, connection-site selection
    UX, and align/distribute relative-to-margins.

- [ ] **Durable file-persisted shape tags.**
  - **COM surface:** `Shape.Tags.Add(name, value)`, `.Item(name)`, delete/read.
  - Complements shipped `shapeid:S:ID`: `Shape.Id` is stable within a session, while
    tags can persist through save/reopen and support cross-session re-identification.

- [ ] **Document properties / metadata.**
  - **COM surface:** `Presentation.BuiltInDocumentProperties` and
    `CustomDocumentProperties`.
  - Read/write title, author, subject, keywords, custom fields.

- [ ] **OLE / other embeds.**
  - **COM surface:** `Shapes.AddOLEObject`.
  - Niche; build when a workflow needs embedded files/objects.

- [ ] **Opaque deck targeting: `doc_id`.**
  - `--doc` already matches display name or full path. Still open: a stable opaque
    `doc_id` returned by `status` and accepted everywhere, including never-saved
    decks where `FullName` is only a bare display name.

- [ ] **Batch symbolic bindings.**
  - `exec` shipped. Still open: `add_shape` / other creation ops returning a symbolic
    label (`"as": "hero_box"`) that later ops can reference as `shape:@hero_box` in
    the same batch.

- [ ] **First threaded comment identity micro-spike.**
  - Comments shipped with legacy fallback for a comment-less deck. If first-comment
    modern threading matters, spike whether the signed-in Office identity can be read
    directly rather than discovered from an existing comment.

---

## Deferred / architectural

- [ ] **Event sinks / async wrapper.**
  - `Application` events such as `WindowSelectionChange`, `SlideShowNextSlide`,
    `PresentationCloseFinal`, `SlideSelectionChanged`.
  - Requires STA event pumping, callback registry, and an async/threaded model. Build
    only when a reactive presenter-assistant workflow is concrete.

- [ ] **Theme/master follow-ups.**
  - Multi-master / per-`Design` styling (`Presentation.Designs`).
  - Per-layout backgrounds (`CustomLayouts(i).Background`).
  - East-Asian / Complex-Script theme fonts beyond the current opt-in.
  - Legacy `.ppt` behaviour.

- [ ] **Full layout authoring.**
  - Add/rename `CustomLayouts`, create/place placeholders programmatically.
  - Rare for agent workflows because templates usually pre-exist; deep COM.

---

## Cross-cutting project work

- [ ] **CI for tests, not just docs/release.**
  - Add a cross-OS unit-test/lint/typecheck workflow.
  - Longer-term: Windows + installed PowerPoint runner for smoke tests, if a reliable
    runner is available.

- [ ] **Checked-in smoke fixture deck.**
  - Add a stable `.pptx` fixture with known slides, placeholders, table, notes,
    comment, hyperlink, styled shape, transition/background, etc.
  - Current live smoke tests create fresh decks, but there is no checked-in fixture.

- [ ] **HRESULT coverage.**
  - Keep widening `_BUSY_HRESULTS` as genuine modal-dialog / busy COM errors appear.
  - Do not add slideshow-running as a busy case; live testing overturned that old
    assumption.

- [ ] **Release/version sync automation.**
  - Release CI now verifies root `pyproject.toml`, `mcpb/manifest.json`, and
    `mcpb/pyproject.toml` agree. Still open: automate the version bump itself so the
    MCP bundle cannot drift before tagging.

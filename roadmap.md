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
| **v1.4-rest** | Navigation & structure: sections, headers/footers | `[x]` main cut · `[ ]` run-level hyperlinks | Sections + headers/footers shipped (v0.6); only text-run-level hyperlinks remain. |
| **v1.5-rest** | Animations | `[x]` main cut · `[ ]` long tail | Whole-shape entrance/exit shipped (v0.10); per-paragraph levels / motion paths / reordering remain. |
| **v1.7** | Media + narrated-video export | `[~]` | Spike proved the high-ceiling workflow: insert narration, self-time slides, export MP4. **The largest single remaining tier.** |
| **Opportunistic** | Tables/charts/SmartArt/arrangement/tags/metadata/OLE | mixed | Pull on demand when a workflow needs it. |
| **Deferred** | Async/events, full layout authoring, deep theme/master follow-ups | `[ ]` | Real but lower leverage or larger architectural lift. |

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

- [ ] **Text-run-level hyperlinks (deferred from the hyperlink cut).**
  - Shape-level click actions are done. If a workflow needs linked words inside a
    textbox, spike `TextRange.ActionSettings` / run-level hyperlink behaviour and
    add a scoped verb over `para:` ranges.

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

## v1.7 — media + narrated-video export

This is the highest-ceiling remaining tier: an agent can build a deck, add
per-slide narration, and export a finished video. The 2026-06-11 spike proved the
chain works end-to-end with no new dependency; implementation is still absent.

- [~] **Insert media: `slide.add_audio(path)` / `slide.add_video(path)`.**
  - **Spike status:** confirmed via `Shapes.AddMediaObject2`.
  - **COM surface:** `Shapes.AddMediaObject2(FileName, LinkToFile=False,
    SaveWithDocument=True, Left, Top, Width, Height)`; read `Shape.MediaType`,
    `Shape.MediaFormat.Length`, `Muted`, `Volume`, `StartPoint`, `EndPoint`.
  - **Wrapper shape:** `Slide.add_audio(path, *, left=0, top=0, width=..., height=...,
    link=False, autoplay=True, hide_icon=True, pace_slide=True)` and video sibling.
  - **Notes:** resolve paths to absolute paths; document embed-vs-link file-size
    tradeoff.

- [~] **Auto-play + per-slide pacing.**
  - **COM surface:** `Shape.AnimationSettings.PlaySettings.PlayOnEntry`,
    `HideWhileNotPlaying`; reuse `SlideShowTransition.AdvanceOnTime` /
    `.AdvanceTime` for pacing.
  - **Wrapper shape:** `pace_slide=True` reads `MediaFormat.Length` and sets slide
    auto-advance to the clip duration.

- [~] **Export video: `deck.export_video(path)`.**
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
    used by this tier.

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

- [ ] **Shape arrangement beyond z-order.**
  - Group / ungroup (`ShapeRange.Group`, `.Ungroup`).
  - Align / distribute.
  - Connectors (`Shapes.AddConnector`, `ConnectorFormat.BeginConnect(shape, site)`).
  - Note: grouping changes shape indices; prefer `shapeid:S:ID` and document any
    identity churn from group/ungroup.

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

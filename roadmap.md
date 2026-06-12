# pptlive — roadmap (post-v0.9)

Forward-looking companion to `IMPLEMENTATION.md` (which tracks **shipped** work)
and `spec.md` (the canonical design). This file is the **future**: the features
still worth building, ordered by the same principle the whole project uses —
**LLM-agent leverage** — and gated by **what the Windows PowerPoint COM API
actually exposes**. Each item names its COM surface, calls the spike risk
honestly, and sketches the wrapper shape so it slots into the established
`Has*`-gate / friendly-name / `.com`-escape-hatch / `deck.edit()`-fence mold.

> **Where we are.** v0 → v0.9 + the MCP server have shipped (see
> `IMPLEMENTATION.md`): attach, reads, text, slide lifecycle, shapes & geometry,
> tables, charts, SmartArt, render, live selection, slide-show control, and
> theme/master styling — across the JSON CLI **and** the five-tool MCP surface.
> The object model is well covered for **structure and content**.
>
> **Shipped since this roadmap was written (2026-06-07 → -09):** **v1.0
> find/replace** (fuzzy, deck-wide); a round of **MCP test-feedback fixes** (`\n`
> → real paragraphs, placeholder-ambiguity guard, richer formatting reads,
> view-preservation hardening); the first **shape-styling** cut — **`set_fill`**
> (solid fill + border colour/weight, `"none"` = transparent), **`reorder`**
> (z-order front/back/forward/backward), and the delete-proof **`shapeid:S:ID`**
> anchor; and **composite-text recolor** (`SmartArt.recolor_text` /
> `Chart.recolor_text` — the only colour path for diagram-node / chart-element
> text); and **v1.3 review comments** — threaded read + add/reply/delete
> (`slide.comments` / `deck.comments()`), the review loop that makes pptlive feel
> indispensable on a shared deck. So the v1.2 styling gate is **partway open**, the
> `shape:S:N` drift hazard now has a stable-handle answer, and the collaboration
> loop is in. **And v1.1 is complete (2026-06-09):** `deck.snapshot()` — a
> whole-deck low-resolution render (one PNG per slide, `max_dim` long-edge cap) so a
> vision model can *see* the whole deck at a predictable, uniform per-slide token
> cost — plus the **save & PDF-export** tier: `deck.save()` / `save_as()` (explicit,
> never-implicit; `save_as` rebinds the working file) and `deck.export_pdf()` (a
> read — no rebind, dirty flag preserved), with a `deck.saved` dirty flag on every
> `status` row. PDF rides `SaveAs(…, ppSaveAsPDF=32)` since `ExportAsFixedFormat`
> won't marshal under the late-bound `_com` dispatch.
>
> What's still thin is the rest of **appearance and behaviour** (gradients,
> effects, motion, navigation).
>
> **Shipped 2026-06-10:** the **v1.6 text-model reliability** tier (line-spacing
> points-vs-multiple + guardrail, `set_paragraphs`, `reset_format` /
> `reset_to_layout`, `text_frame_status` + edit warnings, richer paragraph
> diagnostics, docs) **and** the last specced-but-unbuilt item — the CLI **`exec`**
> verb, on a fastmcp-free `_batch.py` dispatch seam extracted from the MCP server.
>
> **Shipped 2026-06-10 (v0.4.0) — the cross-tier quick-wins release:** the cheapest
> high-leverage cut from three *different* open tiers, each spiked-first and shipped
> across all four front-ends. **Shape hyperlinks** (v1.4 navigation —
> `Shape.set_hyperlink(url=/slide=)` + `remove_hyperlink`, shape-level click action,
> URL or in-deck slide-jump; `hyperlink` on every shape read). **Slide transitions**
> (v1.5 motion — `Slide.set_transition(effect, duration, advance_after)` over a curated
> `PpEntryEffect` map; `transition` on every slide read; animations still deferred).
> **Per-slide background** (v1.2 styling — `Slide.set_background(color)` /
> `follow_master_background()`, the per-slide override of the master; `background` on
> every slide read). So v1.2/v1.4/v1.5 are each **partway open** now, with their
> remaining cuts (effects/gradients, sections/headers-footers, animations) still ahead.
>
> **Shipped 2026-06-12 (v0.5.0) — v1.2 styling completion.** The advanced fills
> (**gradient / picture / pattern**) and shape **effects** (**shadow / glow /
> soft-edge / reflection**) cuts, both spiked 2026-06-11 and now built across all
> four front-ends as dedicated explicit verbs (`Shape.set_gradient_fill` /
> `set_picture_fill` / `set_pattern_fill` / `set_effect`; CLI `shape gradient-fill`/
> `picture-fill`/`pattern-fill`/`effect`; MCP `ppt_edit` `shape_gradient_fill`/
> `shape_picture_fill`/`shape_pattern_fill`/`shape_set_effect`). Every shape read now
> carries a `fill.type` discriminator (+ gradient `stops` / pattern detail) and an
> `effects` field. Live net-zero confirmed (multi-stop reads back sorted; preset
> "ocean" → a 4-stop ramp; all four effects round-trip). **So v1.2 styling is now
> complete** — the only fill deferrals left are partial-alpha `.Transparency`, line
> `.DashStyle`/arrowheads, and the 3-D effect long tail.
>
> **Spiked 2026-06-11 (planning round, no code shipped — four net-zero spikes).**
> De-risked four deferred tiers ahead of building them; every COM behaviour pinned
> on the live deck (findings inline in the sections below):
> - **v1.2 advanced fills** (`fill_advanced_spike.py`) — two/one-colour + preset
>   gradients, multi-stop `GradientStops` (via legacy `Insert`, since `Insert2`
>   won't marshal), absolute-path picture fill, and pattern fill **all work**.
> - **v1.2 effects** (`effects_spike.py`) — shadow / glow / soft-edge / reflection
>   **all round-trip** (no write-only hazard); 3-D is the long tail.
> - **v1.5 animations** (`animation_spike.py`) — `AddEffect` round-trips
>   `EffectType` / `Exit` / `Timing` and maps each effect to its `Shape.Id`/`.Name`,
>   so the whole-shape "fade this in" / "show this" cut (incl. exit + read) is ready.
> - **Media → narrated-video export** (`media_video_spike.py`) — the **"gold
>   idea"**: insert audio (`AddMediaObject2`) + auto-play + per-slide pacing, then
>   `Presentation.CreateVideo` (it **marshals**, unlike PDF's `ExportAsFixedFormat`)
>   produced a real MP4 end-to-end with **no new dependency**. Candidate for its own tier.

**Status legend:** `[ ]` not started · `[~]` in progress · `[x]` shipped.
Spike-first remains the rule: confirm each COM behaviour on a live deck, write a
one-line finding, *then* harden.

---

## Tiering (read this first)

| Tier | Theme | Why now | COM risk |
| ---- | ----- | ------- | -------- |
| **v1.0** | **find / replace + `exec` CLI** | Last wordlive-parity gap; deck-wide search is table-stakes for "change X everywhere" | Low — `TextRange.Find/Replace` exist |
| **v1.1** | **Output: save & PDF/image export** *(SHIPPED 2026-06-09 — `deck.snapshot()` low-res whole-deck render + `save`/`save_as`/`export_pdf`)* | Trivial COM, huge practical payoff ("export the deck to PDF"); the one thing every agent eventually wants | Low — PDF via `SaveAs(…, ppSaveAsPDF)` (`ExportAsFixedFormat` won't marshal late-bound) |
| **v1.2** | **Shape styling — fill / line / effects** *(COMPLETE: solid fill/line + z-order + per-slide bg, **gradient/picture/pattern + shadow/glow/soft-edge/reflection shipped v0.5.0**)* | Biggest *authoring* gap: agents can place a shape but can't colour it; blocks good-looking decks | **Done** — all cuts shipped; live net-zero confirmed |
| **v1.3** | **Review loop — comments** *(SHIPPED 2026-06-09 — read + add/reply/delete, threaded; resolve-state not COM-readable)* | "Address the reviewer's comments" is a killer workflow; read is side-effect-free & polite | **Low** — read (incl. threads), add & reply all verified live; comment-less-deck identity solved via legacy-`Add` fallback |
| **v1.4** | **Navigation & structure — hyperlinks, sections, headers/footers** | Makes multi-slide decks navigable and organized | Low |
| **v1.5** | **Motion — transitions & animations** *(transitions shipped; animations spiked 2026-06-11, whole-shape cut ready)* | Polish; transitions are trivial, animations are the long tail | **Low-med** — spike confirmed `AddEffect` round-trips; only per-paragraph/motion-path stays fiddly |
| **v1.6** | **Text-model reliability — safe formatting, diagnostics, recovery** | Hardens the authoring loop agents *already* use; the gpt-5.4 review's top ask | Low — mostly `LineRule*` bools + reads; the reset primitive needs a spike |
| **v1.7** | **Automated presentation development — media + narrated-video export** *(spiked 2026-06-11, end-to-end MP4 proven)* | The "gold idea": an agent builds a deck, narrates it, and hands back a **video** — the highest-ceiling capability, and uniquely a *live-app* one | **Low** — spike confirmed `AddMediaObject2` + `CreateVideo` both marshal; no new dependency |
| **opportunistic** | deeper tables/charts, arrangement, tags, metadata, OLE embeds | Pull in on demand when a workflow needs it | varies |
| **deferred** | events / async; v0.8/v0.9 follow-ups | Real but lower leverage | — |

The split below each tier follows the IMPLEMENTATION.md house style: wrapper +
constants + CLI + MCP + spike, all four front-ends moving together.

---

## v1.0 — find / replace + `exec` CLI (close the parity gap)

The only **specced-but-unbuilt** module from `spec.md`. wordlive has `find()` /
`find_replace()`; pptlive has the anchor grammar to express hits
(`para:S:N:P` + in-shape offset, per spec) but no implementation. This is the
last thing keeping pptlive from full wordlive surface parity.

- [x] **`_findreplace.py` — `Presentation.find()` / `find_replace()` — BUILT
  (2026-06-07).** Shipped across library + CLI (`find`, `replace --find`) + MCP
  (`ppt_read` op `find`, `ppt_edit` op `find_replace`); see `IMPLEMENTATION.md`
  §v1.0. The build ports wordlive's **fuzzy Python core** verbatim and writes
  replacements through `TextRange.Characters` — this both adds the smart-quote /
  whitespace tolerance native `.Find` lacks *and* sidesteps the native-`.Replace`
  first-only + offset-drift traps the spike below flagged (matches are computed
  once up front, not via a re-scanning loop). PowerPoint
  has **no document-wide character stream** (the whole reason the anchor model is
  hierarchical), so search is a *traversal*: iterate `Slides × Shapes`, descend
  into `TextFrame.TextRange`, table cells (`Cell.Shape.TextFrame`), and notes.
  - **COM surface:** PowerPoint's `TextRange` exposes **`.Find(FindWhat, After,
    MatchCase, WholeWords)`** and **`.Replace(FindWhat, ReplaceWhat, After,
    MatchCase, WholeWords)`** (both return a `TextRange` of the match, or a
    falsy/empty range when nothing matches — *spike to confirm the empty-match
    sentinel*; it differs from Word's `Found` boolean). Loop `.Replace` to get
    every occurrence in a frame.
  - **Wrapper shape:** `find(text, *, in_=None, match_case=False, whole_words=False)`
    → list of `{anchor_id: "para:S:N:P", start_char, length, text, context}`,
    scoped by `in_` (`slide:S` / `shape:S:N` / whole deck). `find_replace(find,
    repl, *, in_, all_/occurrence, match_case, whole_words)` → count replaced.
    Zero matches is **exit 2** (an `AnchorNotFoundError`, per the established
    taxonomy); a fuzzy multi-match without `--all`/`--occurrence` is
    `AmbiguousMatchError` (exit 5) — both already wired.
  - **Politeness:** search is a read (no view move). `find_replace` goes through
    `deck.edit()` for the one-Ctrl-Z fence.
  - **Spike RESOLVED (2026-06-07, `scripts/findreplace_spike.py`, net-zero).**
    Every uncertain behaviour pinned on a live deck:
    - **Empty-match sentinel:** `TextRange.Find` returns **`None`** (VBA
      `Nothing`) on no match — a clean loop stop condition, not an empty range
      and not a raise.
    - **Offsets are 1-based.** `Find("beta")` in `"alpha beta …"` → `Start=7`,
      `Length=4`. Iterating with `Find(what, After=cur.Start+cur.Length)` walked
      all three `"alpha"`s cleanly (`Start` = 1, 12, 24) — `After` advancement is
      the iteration primitive.
    - **`.Replace` is FIRST-ONLY** (not replace-all, contrary to a common
      belief): one call replaced only the first `"alpha"` and **returned the
      replaced range**. Replace-all = **loop `Replace` until it returns `None`**
      (re-fetch the `TextRange` each iteration — a returned range is a *live*
      reference that re-reads mutated text). Verified: a loop turned
      `"alpha beta alpha gamma alpha"` → `"X beta X gamma X"` in 3 iterations.
    - **Offset-drift hazard CONFIRMED.** When the replacement re-contains the
      search text (`"alpha"` → `"alpha_X"`), a naive replace-until-`None` loop
      **spins forever** (the second call re-matched inside the replacement). So
      `find_replace` must advance `After` past each replacement, *not* restart
      from 0. This is the one real correctness trap.
    - **Reach:** notes text is found directly (`notes.com.Find("needle")` → hit),
      confirming notes are an ordinary `TextRange` the traversal visits. Table
      **cell** text frames are likewise ordinary `TextRange`s (proven in v0.5),
      so `Find` works on them — but the spike's own table probe was inconclusive
      because `AddTable` on a content-placeholder layout filled the placeholder
      (the known v0.5 z-order quirk), not a Find limitation. Grouped-shape /
      SmartArt / chart-text reach is determined by *our traversal depth*, not by
      `Find` — decide it when the traversal is written.
- [x] **`exec --script ops.json` at the CLI — SHIPPED (2026-06-10).** The MCP
  `_*_core` dispatch + the batch loop were **extracted into a fastmcp-free
  `pptlive/_batch.py`** (the four op enums, registries, every handler, the cores,
  and a new `run_batch`); `mcp/server.py` now wraps that seam and `cli/commands.py`
  imports it for `exec` — so a single process applies a `{"label", "ops":[…]}`
  script as one Ctrl-Z without the optional `[mcp]` extra. Invalid args raise the
  native `BatchOpError` (the MCP server maps it to `ToolError`, the CLI to exit 1);
  a failing op's category maps to the CLI exit code. `--continue` / `--no-atomic`
  knobs. No new COM — pure plumbing reuse. (Symbolic `shape:@label` binding stays
  deferred, Open Q #3.)

---

## v1.1 — output: save & export (PDF / images)

The cheapest high-leverage feature left. An agent that has built a deck live
almost always wants to **hand back an artifact** — "export this to PDF",
"give me PNGs of every slide". Whole-slide PNG export already shipped (v0.4);
this adds the deck-level outputs that are one COM call away.

> **Philosophy note (be honest):** pptlive is the *live-app* sibling of
> `python-pptx`; file I/O is deliberately not its centre of gravity (a Non-Goal
> in `spec.md`). But **export** ≠ authoring-on-disk — it's a read-only snapshot
> of the live deck, the same spirit as `Slide.Export`. Saving the *working file*
> is the more debatable one; gate it behind an explicit verb and never auto-save.

- [x] **`deck.export_pdf(path)` — SHIPPED (2026-06-09).** A pixel-faithful PDF of
  the deck's current (unsaved) state, and a **read**: it neither rebinds the
  working file nor clears its dirty flag, so the `.pptx` is untouched (verified
  live). **Spike finding:** `ExportAsFixedFormat` — the nominal API — won't marshal
  under pptlive's late-bound `_com` dispatch (a trailing object-typed param raises
  `TypeError` for *every* arg form, named or positional), so PDF goes through
  `SaveAs(path, ppSaveAsPDF=32)`, which produces the same faithful PDF as a pure
  export. CLI `export-pdf PATH`; MCP `ppt_render` op `deck_pdf`.
- [x] **Whole-deck image render — SHIPPED as `deck.snapshot()` (2026-06-09).**
  The multi-slide complement to v0.4's single-slide PNG, reframed around LLM token
  cost (ported from wordlive's snapshot): renders slides to PNG with a `max_dim`
  **long-edge cap** so a vision model can *see* the whole deck at a predictable,
  uniform per-slide budget. `deck.snapshot(out=None, *, slides=None, fmt, max_dim)`
  returns one `Snapshot(slide, image, path)` per slide (`slides` = all / one int /
  inclusive span); CLI `snapshot --slide/--slides/--out/--max-dim`; MCP `ppt_render`
  op `deck_snapshot` (one "slide N" label + image block per slide). Built on
  `Slide.export_image` (no PDF/PyMuPDF, no new dependency). The earlier folder-based
  `deck.export_images` (v0.4) stays for bulk-to-disk. **Still open:** a `jpg`-quality
  knob and per-slide `width`/`height` overrides (snapshot's lever is `max_dim`
  only).
- [x] **`deck.save()` / `save_as(path)` — SHIPPED (2026-06-09).** Explicit-only,
  never implicit. `save()` persists to the existing file; `save_as(path, *,
  fmt="pptx", overwrite=False)` writes a `.pptx` and **rebinds** the working file
  to it (the open deck becomes that file, like PowerPoint's Save-As), refusing to
  clobber unless `overwrite=True`. `deck.saved` (the `Presentation.Saved` dirty
  flag) joins `path` on every `status` deck row. **Spike correction:** `Save` on a
  never-saved deck does **not** raise — on a OneDrive/SharePoint build it silently
  uploads to the user's default cloud folder — so `save()` guards in Python on an
  empty `Presentation.Path` and raises `UnsavedPresentationError` (exit 1) instead.
  CLI `save` / `save-as PATH [--format/--overwrite]`; MCP `ppt_render` ops
  `save` / `save_as`.

---

## v1.2 — shape styling: fill / line / effects

The biggest **authoring** gap. Today an agent can *place* a shape and set its
*text*, but cannot control its **fill, outline, or effects** — so it can't make
a slide actually look designed. This is the per-shape visual counterpart to
v0.3's `format_text` and v0.9's theme palette.

- [~] **`Shape.set_fill(...)` — SOLID FILL + BORDER SHIPPED (2026-06-08,
  PPTLIVE-007).** A single `set_fill(fill=, line=, line_width=)` (not the separate
  `set_fill`/`set_line` originally sketched) sets the shape's **solid fill** and/or
  **border** — a `#RRGGBB` / `(r,g,b)` / raw-int colour or `"none"` (transparent
  fill / no border, via `Fill.Visible = msoFalse`). `fill=`/`line=`/`line_width=`
  also ride on `add_shape`/`add_textbox`; every shape read now reports `fill`/`line`
  (`{color, visible[, weight]}`) with the `color_hex_or_none` theme-sentinel guard.
  Wired library + CLI (`shape fill`, `shape add --fill/--line/--line-width`) + MCP
  (`format` `fill_color`/`line_color`/`line_width`, `shape_add`). Reuses
  `parse_color`/`color_hex` (R-low-byte RGB long). **Still open:** `.Transparency`
  (partial alpha) and line `.DashStyle` / arrowheads — and the gradient / picture /
  pattern cut, now **spiked & de-risked** (next bullet).
- [x] **Gradient / picture / pattern fills — SHIPPED (v0.5.0, 2026-06-12).**
  Dedicated `Shape.set_gradient_fill` / `set_picture_fill` / `set_pattern_fill`
  verbs + a `fill.type` discriminator on every shape read; live net-zero confirmed
  (multi-stop reads back sorted, preset "ocean" → a 4-stop ramp). See
  `IMPLEMENTATION.md` §v1.2-advanced. The spike findings that shaped it:
  All three deferred fill types work over the same `MsoFillFormat` surface the solid
  cut uses:
  - **Gradients:** `Fill.TwoColorGradient(style, variant)` then `ForeColor.RGB` /
    `BackColor.RGB` (`Type`→3 gradient, `GradientColorType`→2, exactly 2 stops at
    0.0/1.0); `Fill.OneColorGradient(style, variant, degree)` (`GradientColorType`→1,
    and the **only** type where `GradientDegree` reads back — it *raises* on read for
    the other two, so guard it); `Fill.PresetGradient(style, variant, presetType)`
    (`GradientColorType`→3, emits the preset's multi-stop ramp). `GradientStyle` /
    `GradientVariant` round-trip.
  - **GradientStops (multi-stop):** fully **readable** — `.Count` + each stop's
    `.Position` (0..1), `.Transparency`, `.Color.RGB`. Writing: **`Insert2(...)` FAILS**
    ("index out of bounds") under late-bound dispatch — but the **legacy `Insert(rgb,
    position)` works** (it *appends* at the end of the collection regardless of
    position, so read stops back **sorted by `.Position`**). Same modern-variant-flaky /
    legacy-reliable pattern as v1.0's `Replace`/`Add2`. So multi-stop is achievable.
  - **Picture fill:** `Fill.UserPicture(absPath)` → `Type`→6 (msoFillPicture),
    `TextureType`→2. **Relative path FAILS** (`ERROR_FILE_NOT_FOUND`) — the predicted
    `Export` footgun recurs, so the wrapper must `os.path.abspath` first.
  - **Pattern fill:** `Fill.Patterned(MsoPatternType)` + `ForeColor`/`BackColor`
    → `Type`→2 (msoFillPatterned), `Pattern` reads back the int.
  - **Build:** one `set_fill` extension — `fill="grad:#aaa,#bbb[@style]"` /
    `fill="pic:PATH"` / `fill="pat:NAME,#fg,#bg"` style, or explicit kwargs; read emits
    a `fill.type` discriminator (`solid`/`gradient`/`picture`/`pattern`) + the stops.
    Constants: `MsoFillType`, `MsoGradientStyle`, `MsoPresetGradientType`, `MsoPatternType`.
- [x] **Effects (second cut) — SHIPPED (v0.5.0, 2026-06-12).**
  `Shape.set_effect(shadow=/glow=/soft_edge=/reflection=)` + an `effects` field on
  every shape read (active effects only); live net-zero confirmed all four
  round-trip. See `IMPLEMENTATION.md` §v1.2-advanced. The spike findings: the
  roadmap's feared "write-only / non-round-tripping" hazard does **not** materialize
  for the common effects:
  - **Shadow** (`Shape.Shadow`): the **individual-property path** round-trips cleanly —
    `.Visible`, `.ForeColor.RGB`, `.Transparency`, `.Blur`, `.Size`, `.OffsetX/Y`, and
    `.Style` (→2 outer). Caveat: setting individual props pushes the **`.Type` preset
    read-back to `-2` (msoShadowMixed)** — so read `.Style` + the props, not `.Type`;
    and default (invisible) `Transparency` reads the `-2147483648` sentinel (guard it).
  - **Glow** (`Shape.Glow`): clean round-trip — `.Color.RGB`, `.Radius` (0 = off),
    `.Transparency`.
  - **SoftEdge** (`Shape.SoftEdge`): clean — `.Type` preset (0..6) **and** the derived
    `.Radius` both read back.
  - **Reflection** (`Shape.Reflection`): clean — `.Type` preset (0..9) reads back.
  - **ThreeD** (`Shape.ThreeD`): `SetThreeDFormat(preset)` + `.Depth`/bevels set and
    mostly read back (`BevelTopType` didn't honor a post-preset override — minor); the
    genuine **long tail**, defer past shadow/glow/soft-edge/reflection.
  - **Build:** `Shape.set_effect(shadow=/glow=/soft_edge=/reflection=)` (each a small
    friendly dict or `"none"`), every shape read emits the present effects. Start with
    shadow (the common ask). Constants: `MsoShadowType`/`MsoShadowStyle`, glow/soft-edge/
    reflection are radius/preset ints.
- [x] **Per-slide background — SHIPPED (v0.4.0).** `Slide.FollowMasterBackground
  = msoFalse` + `Slide.Background.Fill` — the per-slide override of v0.9's master
  background (which is deck-wide). Same `MsoFillFormat` surface as shape fill, so it
  falls out of the same helper. **Spike CONFIRMED (2026-06-10,
  `scripts/slide_background_spike.py`, net-zero):** `FollowMasterBackground` defaults
  to `msoTrue` (-1); setting it `msoFalse` (0) then `Slide.Background.Fill.Solid()` +
  `.ForeColor.RGB = rgb` sets the slide's own solid colour (read back exact, `Fill.Type=1`
  solid); **re-setting `FollowMasterBackground = msoTrue` cleanly reverts** to the master
  background (the revert verb). Solid only this cut (gradient/picture deferred).
- **Constants:** `MsoFillType`, `MsoGradientStyle`/`MsoPresetGradientType`,
  `MsoLineDashStyle`, `MsoPatternType`, `MsoShadowType` — added as each verb
  needs them (don't pre-populate, per convention #7).
- **CLI/MCP:** the solid cut shipped as CLI **`shape fill`** + `shape add
  --fill/--line/--line-width` and MCP `format` (`fill_color`/`line_color`/
  `line_width`) — *not* the originally-sketched `format-shape` /
  `shape_set_fill`/`shape_set_line` ops. The remaining `--fill-gradient`/
  `--line-dash`/`--shadow` knobs land with their cuts above. All through
  `deck.edit()`.
- **Spike — DONE (2026-06-11).** Both open questions answered: gradient stops are
  fiddly because **`Insert2` won't marshal** — use the legacy `Insert(color,
  position)` and read stops back **sorted by `.Position`**; and `UserPicture`
  **does** require an absolute path (relative raises `ERROR_FILE_NOT_FOUND`). So the
  build can ship **solid + line + two-colour + preset + multi-stop (via `Insert`) +
  picture + pattern** in one cut — the only deferral left is partial-alpha
  `.Transparency` and line `.DashStyle`/arrowheads.

---

## v1.3 — review loop: comments — SHIPPED (2026-06-09)

"Read the reviewer's comments and address them" is one of the highest-value
agent workflows on a real, shared deck — and **reading is side-effect-free and
perfectly polite** (no view move, no edit fence). This is the feature most
likely to make pptlive feel *indispensable* on a working deck rather than a
clean-room one.

**Built** (`_comments.py`, see `IMPLEMENTATION.md` §v1.3): `slide.comments` (per-slide,
1-based, `add`/`reply`/`delete`/`list`, threaded), `deck.comments()` deck-wide roll-up;
CLI `comment list/add/reply/delete`; MCP `ppt_read` `comments` + `ppt_edit`
`comment_add`/`comment_reply`/`comment_delete`. Identity for the modern `Add2` is
sourced off an existing comment, with a legacy identity-free `Add` fallback on a
comment-less deck (resolving the one open question below). **Not built** (COM doesn't
expose it): resolve/reopen — `Comment.Status`/`.Resolved` are "no longer supported by
this version".

- [x] **`Slide.comments` (read) — SHIPPED; COM sees UI comments incl. threads.**
  `Slide.Comments` collection: per comment `.Author`, `.AuthorInitials`,
  `.Text`, `.DateTime`, `.Left`/`.Top` (anchor position). Surface as `comments:S`
  reads → `[{author, text, datetime, anchor}]`, plus a deck-wide roll-up.
  **High leverage, zero risk to the user's state.** **Decisive test PASSED
  (2026-06-07, `scripts/comments_spike.py`):** a comment authored in the
  PowerPoint **Review-tab UI** (the modern path) read back cleanly through the
  COM walk — `Author`, `AuthorInitials`, `Text`, tz-aware `DateTime`,
  `Left`/`Top`. So v1.3 read ships **full coverage**, not the feared
  legacy-only-with-caveat. (Build the wrapper; this is no longer a risk.)
- [x] **Add / reply / delete — SHIPPED, threaded write works.** Both verified live on a
  temp slide (net-zero), using identity keys discovered from a real comment.
  - **Add a comment:** `Slide.Comments.Add2(Left, Top, Author, AuthorInitials,
    Text, ProviderID, UserID)` — the 7-arg modern form. (The 5-arg call fails
    `DISP_E_PARAMNOTOPTIONAL`; `ProviderID`/`UserID` are mandatory. Note the
    **capital-D** spelling — `ProviderId`/`UserId` don't exist.)
  - **Reply to a comment:** `Comment.Replies.Add2(Left, Top, Author,
    AuthorInitials, Text, ProviderID, UserID)` — same 7-arg shape on the
    `.Replies` collection. Verified: `parent_replies_count` went to 1 and the
    reply read back. Replies inherit the parent's anchor `Left`/`Top`.
  - **Identity source — the one real design question.** `ProviderID`/`UserID`
    are the signed-in Office account (`ProviderID="AD"`,
    `UserID="S::user@domain::<guid>"` on this box). The wrapper **lifts them off
    any existing comment in the deck** (`_discover_identity`). **RESOLVED at build
    (2026-06-09): option (b)** — on a deck with *zero* comments to source from,
    `add` falls back to legacy `Comments.Add(Left, Top, Author, Initials, Text)`
    (no IDs needed). Option (a) (reading the app/account identity directly, for a
    *threaded* first comment) stays deferred behind its own micro-spike.
  - **Authorship can't be spoofed (caveat to document).** `Add2` **ignores the
    passed `Author`/`AuthorInitials` and binds the comment to the account behind
    `UserID`** — we passed `"Spike Author"`/`"SA"` and it read back as the
    signed-in user (`"Thomas Villani"`/`"TV"`). `Text` *is* honored. So an
    agent-authored comment is correctly attributed to the human's account, not a
    fake name — arguably the right behaviour, but call it out.
- **HEADLINE QUESTION RESOLVED (2026-06-07).** A comment authored in the modern
  PowerPoint **Review-tab UI** read back through the COM `Slide.Comments` walk
  with every field intact, **and its threaded reply read back** by recursing
  `.Replies` (`replies_count: 1`, full reply text/author/datetime). So **UI
  comments — including threads — are fully COM-visible**: v1.3 ships full read
  coverage, not legacy-only-with-caveat.
- **Unsupported on this build (don't rely on):** `Comment.AuthorIndex`
  (`"no longer supported by this version"`), `.Status`, `.Resolved` — so
  comment *resolution state* is not COM-readable here; a "resolve comment" verb
  would need its own spike (likely not exposed).
- **CLI/MCP (SHIPPED):** CLI `comment list [--slide S]` (per-slide / deck-wide
  roll-up, replies nested into a thread tree), `comment add/reply/delete`; MCP
  `ppt_read` op `comments` and `ppt_edit` ops `comment_add` / `comment_reply` /
  `comment_delete` (identity sourced per the design note above).

---

## v1.4 — navigation & structure: hyperlinks, sections, headers/footers

The connective tissue of a real deck. None of this is in the object model yet
and all of it is low-risk COM.

- [x] **Hyperlinks & actions — SHAPE-LEVEL CUT SHIPPED (v0.4.0).**
  `Shape.ActionSettings(ppMouseClick).Hyperlink` (`.Address` for URLs/files,
  `.SubAddress` for "jump to slide N"). Lets an agent build clickable navigation,
  agenda links, "back to TOC" buttons. `set_hyperlink(anchor, *, url=None,
  slide=None)`; reads emit any existing link per shape. **Spike CONFIRMED
  (2026-06-10, `scripts/hyperlink_spike.py`, net-zero):** setting
  `ActionSettings(1).Hyperlink.Address = url` **auto-flips `.Action` to
  `ppActionHyperlink` (7)**; `Hyperlink.Delete()` reverts `.Action` to 0 and
  `.Address` to `""` (an un-linked shape reads `.Address` as `""`, not a raise).
  **`SubAddress` slide-jump format = `"<SlideID>,<index>,<title>"`** (the
  PowerPoint-UI form); it round-trips, and a *bare index* string is auto-canonicalized
  by PowerPoint to the full form (`"1"` read back `"256,1,Slide 1"`) — the wrapper
  builds the explicit `SlideID,index,title` form for determinism. *Text-run-level
  `TextRange.ActionSettings` deferred (shape-level only this cut).*
- [ ] **Sections.** `Presentation.SectionProperties`: `.Count`, `.Name(i)`,
  `.SlidesCount(i)`, `.AddSection(index, name)`, `.Rename`, `.Delete`,
  `.Move`. Organizing large decks — a `deck.sections` read + add/rename/move
  verbs. Pure structural, no view move.
- [ ] **Headers / footers / slide numbers / date.** `Slide.HeadersFooters` and
  `SlideMaster.HeadersFooters` (`.Footer.Text`/`.Visible`, `.SlideNumber.Visible`,
  `.DateAndTime.Format/.UseFormat`). Common request ("add slide numbers", "put
  the date in the footer"). Per-slide vs deck-wide mirrors the v0.9
  anchor-vs-master split.
- **CLI/MCP:** folded into the existing groups (`shape set-link`, a new `section`
  group, `deck set-footer`). Spike: `SubAddress` slide-reference string format
  (it's an index-or-SlideID encoding — confirm the exact form COM expects).

---

## v1.5 — motion: transitions & animations

Pure PowerPoint, no Word analog (like slide-show control). **Transitions are
trivial; animations are the genuine long tail** — split accordingly.

- [x] **Slide transitions — SHIPPED (v0.4.0).** `Slide.SlideShowTransition`:
  `.EntryEffect` (`PpEntryEffect`, a large but flat enum), `.Duration`,
  `.AdvanceOnTime`/`.AdvanceTime` (auto-advance), `.AdvanceOnClick`.
  `slide.set_transition(effect, *, duration, advance_after)`; reads round-trip
  cleanly. Friendly names → `PpEntryEffect` (the `chart_type_for` pattern, a
  **curated** common subset + raw-int passthrough). **Spike CONFIRMED (2026-06-10,
  `scripts/transition_spike.py`, net-zero):** `EntryEffect` write/read round-trips for
  the documented cut / blinds / checkerboard / cover / dissolve / fade / uncover
  families (e.g. cut=257, blinds_horizontal=769, cover_left=1281, dissolve=1537,
  fade=1793, uncover_left=2049) — **but PowerPoint *validates*: the "wipe" family
  ints (3329-3332) are rejected with "this enumeration value is not valid for
  transitions"** (so the spike's first-pass guessed labels were wrong; the curated
  set is restricted to round-trip-verified families, and the raw-int passthrough is
  the escape hatch for anything exotic). `Duration` round-trips; **auto-advance needs
  BOTH `AdvanceOnTime=msoTrue` AND `AdvanceTime=<seconds>`** (default
  `AdvanceOnClick=msoTrue`/`AdvanceOnTime=msoFalse`), so `advance_after=N` sets
  `AdvanceOnTime=-1` + `AdvanceTime=N` together. *Animations (next item) stay deferred.*
- [~] **Animations — SPIKE CONFIRMED (2026-06-11, `scripts/animation_spike.py`,
  net-zero); the whole-shape cut round-trips, ready to build.** The open question
  ("does `AddEffect` round-trip, or is it write-only like SmartArt assistant nodes?")
  is **answered: it round-trips fully** — no write-only hazard for the common asks.
  - **Add:** `Slide.TimeLine.MainSequence.AddEffect(Shape, EffectId, Level=0,
    Trigger)`. The two headline asks: **"fade this in"** = `AddEffect(shape,
    msoAnimEffectFade=10, 0, onClick=1)`; **"show this"** = `AddEffect(shape,
    msoAnimEffectAppear=1, 0, trigger)`. `MsoAnimTriggerType`: onPageClick=1,
    withPrevious=2, afterPrevious=3.
  - **Read back (all clean):** the returned `Effect` exposes `.EffectType` (fade=10 /
    appear=1 survive), `.Exit`, `.Shape.Id` **+ `.Shape.Name`** (maps each effect back
    to its shape — the key to a `slide.animations` read: *"Rectangle 1 fades in
    after previous"*), and `.Timing.Duration` / `.TriggerType` / `.TriggerDelayTime` /
    `.Speed` — all round-trip. `MainSequence.Count` + `MainSequence(i)` iterate the
    whole sequence; `Effect.Delete()` drops the count.
  - **Tune + exit:** `Effect.Timing.Duration = 2.0` and `.TriggerType` round-trip; an
    **exit** effect is the *same* enum + **`Effect.Exit = msoTrue`** (reads back -1) —
    so entrance *and* exit ("fade this out") fall out of one verb.
  - **Build:** `Shape.animate(effect="fade"|"appear"|..., *, trigger="on_click"|
    "with_previous"|"after_previous", duration=, exit=False)` + `slide.animations`
    read + `Shape.clear_animations()`. Curated `MsoAnimEffect` subset (the
    `chart_type_for`/transition pattern) + raw-int passthrough. **Defer** the long tail:
    per-paragraph `Level`, motion paths, and `EffectParameters` (untested — likely the
    one genuinely fiddly corner).

---

## v1.6 — text-model reliability & safe authoring — SHIPPED (2026-06-10)

> **Shipped (2026-06-10).** All items below landed — `line_spacing`
> disambiguation + guardrail, `set_paragraphs`, `reset_format` /
> `reset_to_layout`, `text_frame_status` + edit `warnings`, extended paragraph
> diagnostics, and the docs. See `IMPLEMENTATION.md` §v1.6. The checkboxes below are
> retained for the design rationale.

Source: the **gpt-5.4 MCP-session review** (`docs/reviews/gpt-5.4-review.md`, 2026-06-10).
Unlike v1.2–v1.5, this tier adds **no new object-model coverage** — it hardens the
*existing* text/formatting surface against the PowerPoint sharp edges that leak
through the abstraction. The reviewer's verdict: *"the architecture feels right;
the pain wasn't missing functionality — it was PowerPoint's text model leaking
through."* That makes it the **highest-leverage open tier**: it makes the loop
agents *already* use trustworthy, rather than adding surface they don't yet have.

> **Already shipped, so de-scoped from the reviewer's list.** Two of the
> reviewer's top items predate or partly exist in the current build:
> - **`\n` → real paragraphs** (PPTLIVE-001). The reviewer's "multi-line text
>   collapses into one paragraph with line breaks" was fixed before this review:
>   `set_text`/`write` normalize `\n`/`\r\n`/`\r` to paragraph breaks, with `\v`
>   (`SOFT_BREAK`) the explicit within-paragraph soft break. The reviewer's run
>   almost certainly predates the fix.
> - **Richer paragraph reads (first cut).** `paragraph_to_dict` already emits
>   `bullet` / `indent_level` / `alignment` and the effective `font` sub-dict — the
>   reviewer's "expose per-paragraph diagnostics" ask, partially in hand (extended
>   by the read item below).
> - **`--doc` disambiguation by full path** (2026-06-10) — a down payment on the
>   reviewer's "canonical doc id" ask (the opaque-handle version is below).

- [x] **Disambiguate `line_spacing` — the headline footgun.** Today
  `format_paragraph(line_spacing=)` writes `ParagraphFormat.SpaceWithin`, which is a
  **multiple** (`1.0` single, `1.5`, `2.0`). The reviewer passed `24` expecting *24
  pt* and got 24× spacing — text off the slide, unrecoverable without a rewrite.
  PowerPoint models the two readings with a companion flag:
  `ParagraphFormat.LineRuleWithin` (`msoTrue` = the value is in **lines/multiple**;
  `msoFalse` = the value is in **points**), and likewise `LineRuleBefore` /
  `LineRuleAfter` pair with `SpaceBefore` / `SpaceAfter`. **Design:** add an explicit
  mode rather than overload one number — either `line_spacing_mode: "multiple" |
  "exact_points"` + `line_spacing_value`, or the PowerPoint-native pair `line_spacing`
  (multiple, *unchanged meaning*) + `line_spacing_points` (sets
  `LineRuleWithin=msoFalse`). Keep the current `line_spacing` semantics to avoid a
  silent break; add the points path alongside. Wire the same lines-vs-points
  distinction onto `space_before` / `space_after`. **Guardrail (cheap, do with it):**
  reject an absurd multiple (e.g. `> 5`) unless explicitly forced — a value that large
  is almost always a points-vs-multiple confusion. Library + CLI + MCP + docs table.
  Low COM risk (`LineRule*` are plain bools). **Spike CONFIRMED (2026-06-10,
  `scripts/text_model_spike.py`):** all three `LineRule*` flags set cleanly and read
  back (`msoTrue=-1` multiple / `msoFalse=0` points); `SpaceWithin` stores a bare
  number whose unit `LineRuleWithin` selects. Adopting the PowerPoint-native pair
  (`line_spacing` multiple + `line_spacing_points`). The same spike also pinned the
  **reset primitive** (re-setting `.Text` does *not* clear run overrides → reset
  re-applies layout/master defaults; the `CustomLayout` placeholder is the readable
  source of truth) and the **autofit reads** (`TextFrame2.AutoSize` clean; classic
  `Margin*`/`WordWrap` readable; no `Font.AutoScale` → coarse overflow heuristic).

- [x] **Paragraph-structured writing — `set_paragraphs([...])`.** Even with `\n`
  normalization, the reviewer wants to *not rely on newline inference* for list
  authoring. A paragraph-oriented verb takes a list of `{text, list_type?,
  indent_level?, ...}` and builds each as its own addressable `para:`, resetting list
  state cleanly — the "safe bullet list" path. Pure plumbing over the existing
  `set_text` + `apply_list` + `format_paragraph` verbs; no new COM. MCP op
  `set_paragraphs` (the reviewer's `write_paragraphs`), CLI equivalent.

- [x] **Normalize / reset direct formatting — the recovery verb.** Once a placeholder
  is in a bad state (5 pt font, giant spacing, overflow), reformatting it is
  unreliable. Add a verb that strips *direct* run/paragraph formatting so text falls
  back to the layout/master defaults. COM: a `TextRange`-level reset — clear run font
  overrides (size/bold/color → inherited) and paragraph overrides
  (`SpaceWithin`/`Before`/`After`, `IndentLevel`, bullet). PowerPoint exposes no single
  "clear formatting" call, so **spike the reset primitive** (likely: re-set the text so
  it re-inherits, or null each override). **Companion:** `reset_placeholder_to_layout`
  — restore a placeholder's geometry + formatting from its `CustomLayout` placeholder.
  CLI `text reset-format` / MCP `text_reset_format`. **Spike-first** (the reset
  primitive is the uncertain bit).

- [x] **Text-frame / autofit diagnostics (read).** Surface the state that makes a
  "formatting spiral" visible *before* it bites: `TextFrame.AutoSize`
  (`ppAutoSizeNone` / `ShapeToFitText` / `TextToFitShape`), `WordWrap`, the four
  `MarginLeft/Right/Top/Bottom`, and an overflow signal (compare text extent to the
  frame, or read the shrink-to-fit `Font.AutoScale` / `LineSpacingReduction` when
  present). A read op `text_frame_status`, or extend the `anchor` read with autofit +
  margins + an overflow-risk flag. All reads; low risk.

- [x] **Extend paragraph diagnostics with spacing + run mix.** `paragraph_to_dict`
  already carries bullet/indent/alignment/effective-font; add `space_before` /
  `space_after` / `line_spacing` (with the lines-vs-points mode from the first item)
  and a **mixed-run summary** (the distinct font sizes in a paragraph) so an agent can
  *see* a stray 5 pt run before it renders. Pure read extension.

- [x] **Optional validation / warnings.** Proactively flag suspicious inputs rather
  than silently applying them: a line-spacing multiple `> 5`, a font size `< 8` pt, a
  list applied to a single paragraph full of soft breaks. Return as a non-fatal
  `warnings` array on the edit result (the structured-I/O contract has room), or
  reject-unless-forced for the catastrophic ones. Pairs with the `line_spacing`
  guardrail above.

- **Docs (do alongside the code):** a **"PowerPoint text-model gotchas"** section
  *and* a **formatting-field reference table** (each field → unit, exact COM mapping,
  valid range, per-paragraph-vs-per-run scope, multi-paragraph behavior) — the
  reviewer asked for both explicitly. The single highest-value doc change for agent
  reliability; it's what turns "trial and error" into "read the table once."

---

## v1.7 — automated presentation development: media + narrated video

The **highest-ceiling** capability and the one most native to pptlive's reason for
existing: an agent builds a deck, drops a **spoken narration** on each slide, and
exports a finished **video** — a self-running presentation produced end-to-end from
a prompt. `python-pptx` can author the slides on disk but cannot drive the **video
encoder** (a live-app-only COM service), so this is a capability the file-based
sibling structurally *cannot* have. Promoted out of "opportunistic media" once the
2026-06-11 spike proved the whole chain works on pure COM with **no new dependency**.

> **Why it's a tier, not a one-liner.** The three pieces (embed audio, auto-play +
> pace, encode video) compose into a single agent workflow — "turn these talking
> points into a narrated explainer video" — that nothing else in the object model
> reaches. It also pairs naturally with an LLM **TTS** step upstream (the agent writes
> the script, synthesizes speech to a `.wav`/`.mp3`, then this tier embeds + renders).

- [~] **Insert media — `slide.add_audio(path)` / `slide.add_video(path)` — SPIKE
  CONFIRMED (2026-06-11, `scripts/media_video_spike.py`, net-zero).**
  `Shapes.AddMediaObject2(FileName, LinkToFile=False, SaveWithDocument=True, Left,
  Top, Width, Height)` embeds the clip → `Shape.MediaType`→2 (sound) / 3 (movie);
  **`Shape.MediaFormat.Length` reads the clip duration in ms** (the lever for slide
  pacing), with `Muted`/`Volume`/`StartPoint`/`EndPoint` readable. `LinkToFile=False`
  + `SaveWithDocument=True` embeds (portable deck); expose `link=` for the
  link-don't-embed case. Absolute path (the picture-fill footgun almost certainly
  recurs — resolve it in the wrapper).
- [~] **Auto-play + per-slide pacing — SPIKE CONFIRMED.**
  `Shape.AnimationSettings.PlaySettings.PlayOnEntry = msoTrue` makes the narration
  play on slide entry (reads back -1), `HideWhileNotPlaying` hides the speaker icon;
  pace the slide to the clip with `SlideShowTransition.AdvanceOnTime = msoTrue` +
  `.AdvanceTime = clip_seconds` (reuses the v0.4.0 transition surface). So
  `add_audio(path, *, auto_play=True, pace_slide=True)` reads `MediaFormat.Length`
  and sets the advance time for you — the deck self-times to its narration.
- [~] **Export to video — `deck.export_video(path)` — SPIKE CONFIRMED; the key
  surprise is it *marshals*.** Unlike PDF's `ExportAsFixedFormat` (which won't pass
  the late-bound `_com` dispatch, forcing PDF through `SaveAs`), **`Presentation.
  CreateVideo` has an all-scalar signature and marshals cleanly:**
  `CreateVideo(FileName, UseTimingsAndNarrations=True, DefaultSlideDuration,
  VertResolution, FramesPerSecond, Quality)`. It is **async** — poll
  `Presentation.CreateVideoStatus` (`PpMediaTaskStatus`: None=0, InProgress=1,
  Queued=2, Done=3, Failed=4); the spike saw Queued→InProgress→Done in ~3 s for a
  480p clip and wrote a real 183 KB `.mp4`. It is a **read** (exports the whole deck,
  no mutation, no rebind, dirty flag preserved) — same contract as `export_pdf`. The
  wrapper returns immediately with a pollable status **and** offers a blocking
  `wait=True` convenience that pumps `CreateVideoStatus` to terminal. Alternates noted
  but not adopted: `SaveAs(path, ppSaveAsMP4=39 / ppSaveAsWMV=37)` (no status handle).
- **Constants:** `PpMediaType` (sound=2 / movie=3), `PpMediaTaskStatus`
  (None/InProgress/Queued/Done/Failed), and the `ppSaveAsMP4=39`/`ppSaveAsWMV=37`
  alternates folded into `PpSaveAsFileType` — added as each verb needs them.
- **CLI/MCP:** CLI `media add --audio/--video PATH [--no-autoplay] [--no-pace]` and
  `export-video PATH [--resolution/--fps/--quality/--no-timings] [--wait]`; MCP
  `ppt_edit` `media_add` + `ppt_render` `export_video` (returns the task status, or
  the finished path when `wait`). The encode is async, so the non-wait MCP form
  returns a status the agent can poll on a later call.
- **Open questions for the build spike:** (1) does `CreateVideoStatus` distinguish a
  *failed* encode cleanly enough to map to an exit code (the spike only saw the happy
  path)? (2) video **embed** size — `SaveWithDocument=True` on a large `.mp4` bloats
  the deck; document the link-vs-embed tradeoff. (3) overlap with **recorded
  narration** (`SlideShowSettings`/`.PlayNarration`) — `AddMediaObject2` + per-slide
  audio is the simpler, more controllable path, but note the native record path exists.

---

## Opportunistic — pull in when a workflow needs it

Real features, lower or situational leverage. Build on demand rather than
speculatively.

- [ ] **Deeper tables.** Merge/split (`Cell.Merge`/`.Split`), cell fill &
  borders (`Cell.Shape.Fill`, `Cell.Borders`), column width / row height,
  built-in table styles (`Table.ApplyStyle(styleId)`), header/banding flags
  (`.FirstRow`/`.HorizBanding`). Extends v0.5.
- [~] **Deeper charts / SmartArt — TEXT COLOUR SHIPPED (2026-06-09,
  PPTLIVE-009).** `Chart.recolor_text(color)` recolors every shown chart text
  element (legend, both axis tick labels, title, per-series data labels) +
  `ChartArea` default; `SmartArt.recolor_text(color)` recolors every node label —
  the only colour path for these anchor-less composite shapes (CLI `chart/smartart
  recolor-text`; MCP `chart_recolor_text`/`smartart_recolor_text`). Coarse "all
  text → X" only. **Still open:** title/legend/axis *content & geometry*
  (`Chart.Axes(...)`, `.HasTitle` text), per-element (vs. whole-shape) text
  targeting, and **fill** colour — per-series chart fill and SmartArt node-shape
  fill (no text-anchor; needs its own spike). Extends v0.7b/v0.8 from *content* to
  *appearance*.
- [~] **Shape arrangement — Z-ORDER SHIPPED (2026-06-08, PPTLIVE-008).**
  `Shape.reorder("front"|"back"|"forward"|"backward")` over `Shape.ZOrder` returns
  the new 1-based slot (CLI `shape order --to`; MCP `shape_order`) — so a new
  background panel slides *behind* existing content. **Still open:**
  `ShapeRange.Group`/`Ungroup`, `.Align`/`.Distribute`, and **connectors**
  (`Shapes.AddConnector` + `ConnectorFormat.BeginConnect(shape, site)`) for
  agent-built diagrams. Note: grouping **changes z-order indices** — interacts with
  the `shape:S:N` drift hazard (now mitigable via `shapeid:S:ID`, below); document it.
- **Media + narrated-video export — PROMOTED to its own tier, [v1.7](#v17--automated-presentation-development-media--narrated-video).**
  The spike (2026-06-11) proved the *build-deck → narrate → export MP4* chain works
  end-to-end on pure COM with no new dependency, so it graduated from this bucket.
- [ ] **OLE / other embeds.** `Shapes.AddOLEObject`. Niche; pull in on demand.
- [~] **Durable re-identification handle — `shapeid:S:ID` SHIPPED (2026-06-08,
  PPTLIVE-010); file-persisted Tags still open.** `slide.shapes.by_id(ID)` /
  `anchor_by_id("shapeid:S:ID")` resolves a shape by its stable `Shape.Id` (the
  `id` already in every shape listing) — delete-proof and restack-proof, unlike the
  volatile `shape:S:N` z-order index, and more general than `alt_text` (pictures
  only). It is the recommended "remember-this-shape" handle **within a session**.
  **Still open — `Shape.Tags`** (`.Add(name, value)`/`.Item(name)`): arbitrary
  key/value pairs **persisted in the file**, so they survive *save/reopen* where a
  runtime `Shape.Id` may not — the cross-session complement to `shapeid`. Low cost;
  pull in when a workflow needs cross-session re-find.
- [ ] **Document properties / metadata.** `Presentation.BuiltInDocumentProperties`
  (title, author, subject, keywords) + `CustomDocumentProperties`. Cheap read,
  occasional write.
- [~] **Canonical deck targeting — `--doc` now matches `Name` *or* full path
  (2026-06-10); opaque `doc_id` still open.** The gpt-5.4 review flagged a
  render-vs-edit lookup inconsistency when two decks shared a display name; `--doc`
  now falls back from the (non-unique) `Name` to the `FullName` path, disambiguating
  the common collision. **Still open:** a fully **opaque, stable `doc_id`** returned
  by `status` and accepted everywhere — distinguishing display name / path /
  active-doc token, and covering the never-saved-deck case where `FullName` is just a
  bare name. Pull in if name-or-path proves insufficient in practice.

---

## Deferred (real, but lower leverage)

- [ ] **Event sinks / async.** `Application` events —
  `WindowSelectionChange` (react to what the user just selected),
  `SlideShowNextSlide`, `PresentationCloseFinal`, `SlideSelectionChanged`. The
  reactive surface: an agent that *responds* to the user rather than only acting.
  Real value for a presenter-assistant, but it forces the async/threaded model
  the project has deliberately deferred (STA event pump → callback registry, per
  `spec.md` principle #6). Big lift; revisit when a reactive workflow is
  concretely needed.
- [ ] **SmartArt follow-ups (from v0.8).** orgChart assistant/branch nodes (the
  write-only `type` issue), layout availability across Office versions, widening
  past the 7 core layouts on demand.
- [ ] **Theme/master follow-ups (from v0.9).** Multi-master / per-`Design`
  styling (`Presentation.Designs`), per-layout backgrounds
  (`CustomLayouts(i).Background`), **non-solid master backgrounds**
  (gradient/picture — overlaps v1.2's fill helper, so do them together),
  East-Asian/Complex-Script theme fonts beyond the `--script` opt-in, legacy
  `.ppt` behaviour.
- [ ] **Full layout authoring.** Add/rename `CustomLayouts`, place placeholders
  programmatically. Rare for an agent (templates usually pre-exist); deep COM.

---

## Cross-cutting (carry forward from IMPLEMENTATION.md)

- [ ] **CI** — a Windows+PowerPoint runner for the smoke suite + a cross-OS
  unit-test job (still open from bootstrap).
- [ ] **Smoke fixtures** — a checked-in `.pptx` with known slides / placeholders
  / a table / notes / **a comment / a hyperlink / a styled shape** (extend the
  fixture as each tier lands) so smoke tests have a stable target.
- [ ] **HRESULT coverage** — widen `_BUSY_HRESULTS` as real `com_error`s surface
  (note: the show-running-rejects-edits assumption was *overturned* in v0.6, so
  there's no slideshow busy HRESULT to add from that path — only genuine modal
  dialogs).
- [ ] **`.mcpb` / version sync** — keep `mcpb/manifest.json` +
  `mcpb/pyproject.toml` in lockstep with root `pyproject.toml` on each bump.

---

## How to read the ordering

The tiers above are **leverage × COM-availability**, not difficulty. v1.0/v1.1
are near-free and high-value (parity + export). v1.2 (styling) is the gate to
*good-looking* agent output. v1.3 (comments) is the gate to *collaborative*
decks — and carries the one genuine COM risk worth spiking before committing
(modern threaded comments may simply not be COM-visible). v1.4/v1.5 round out
navigation and motion. v1.7 (media + narrated video) is the **highest ceiling** —
"automated presentation development", an agent that hands back a finished video —
and now de-risked, though it sits after the authoring tiers that feed it (you
narrate a deck you can first *build* and *style*). Everything below the line is real
but waits for a concrete workflow to pull it in — the same demand-driven discipline
that kept the constants module from being pre-populated.

When in doubt, the rule that built v0–v0.9 still holds: **open the equivalent
wordlive module if one exists, spike the COM behaviour on a live deck, write the
one-line finding here, then harden across all four front-ends (wrapper, CLI, MCP,
tests) together.**

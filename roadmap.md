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
> loop is in. **And v1.1's first cut landed (2026-06-09): `deck.snapshot()`** — a
> whole-deck low-resolution render (one PNG per slide, `max_dim` long-edge cap) so a
> vision model can *see* the whole deck at a predictable, uniform per-slide token
> cost — the token-aware "did my styling land across all slides" read, ported from
> wordlive's snapshot but shorter (`Slide.Export` renders a sized PNG directly, no
> PDF/PyMuPDF).
>
> What's still thin is the rest of **appearance and behaviour** (gradients,
> effects, motion, navigation) and a handful of **specced-but-unbuilt** items (the
> CLI `exec` verb, PDF export, save).

**Status legend:** `[ ]` not started · `[~]` in progress · `[x]` shipped.
Spike-first remains the rule: confirm each COM behaviour on a live deck, write a
one-line finding, *then* harden.

---

## Tiering (read this first)

| Tier | Theme | Why now | COM risk |
| ---- | ----- | ------- | -------- |
| **v1.0** | **find / replace + `exec` CLI** | Last wordlive-parity gap; deck-wide search is table-stakes for "change X everywhere" | Low — `TextRange.Find/Replace` exist |
| **v1.1** | **Output: save & PDF/image export** *(started: `deck.snapshot()` low-res whole-deck render shipped; PDF export + save/save_as open)* | Trivial COM, huge practical payoff ("export the deck to PDF"); the one thing every agent eventually wants | Low — `ExportAsFixedFormat`, `SaveAs` |
| **v1.2** | **Shape styling — fill / line / effects** *(started: solid fill/line + z-order shipped; gradients/effects/per-slide bg open)* | Biggest *authoring* gap: agents can place a shape but can't colour it; blocks good-looking decks | Low-med — fills are easy, gradient stops fiddly |
| **v1.3** | **Review loop — comments** *(SHIPPED 2026-06-09 — read + add/reply/delete, threaded; resolve-state not COM-readable)* | "Address the reviewer's comments" is a killer workflow; read is side-effect-free & polite | **Low** — read (incl. threads), add & reply all verified live; comment-less-deck identity solved via legacy-`Add` fallback |
| **v1.4** | **Navigation & structure — hyperlinks, sections, headers/footers** | Makes multi-slide decks navigable and organized | Low |
| **v1.5** | **Motion — transitions & animations** | Polish; transitions are trivial, animations are the long tail | Med — `TimeLine` effect enums are large/fiddly |
| **opportunistic** | deeper tables/charts, arrangement, media, tags, metadata | Pull in on demand when a workflow needs it | varies |
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
- [ ] **`exec --script ops.json` at the CLI.** The batch surface is **already
  built for MCP** (`ppt_batch`, one shared `attach()`, `atomic` fences each
  `edit` into one undo entry) but the **CLI `exec`** specced in `spec.md` /
  `IMPLEMENTATION.md` never landed. Lift the MCP `_*_core` helpers behind a CLI
  `exec` verb so a single process applies a `{"label", "ops":[…]}` script as one
  Ctrl-Z. Op set already proposed in `IMPLEMENTATION.md`. No new COM — pure
  plumbing reuse. (Symbolic `shape:@label` binding stays deferred, Open Q #3.)

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

- [ ] **`deck.export_pdf(path)`** over `Presentation.ExportAsFixedFormat(Path,
  FixedFormatType=ppFixedFormatTypePDF, Intent, …)` (or `SaveAs(path,
  ppSaveAsPDF=32)`). High value, near-zero risk. CLI `deck export-pdf --out
  PATH`; MCP `ppt_render` op `deck_pdf`.
- [x] **Whole-deck image render — SHIPPED as `deck.snapshot()` (2026-06-09).**
  The multi-slide complement to v0.4's single-slide PNG, reframed around LLM token
  cost (ported from wordlive's snapshot): renders slides to PNG with a `max_dim`
  **long-edge cap** so a vision model can *see* the whole deck at a predictable,
  uniform per-slide budget. `deck.snapshot(out=None, *, slides=None, fmt, max_dim)`
  returns one `Snapshot(slide, png, path)` per slide (`slides` = all / one int /
  inclusive span); CLI `snapshot --slide/--slides/--out/--max-dim`; MCP `ppt_render`
  op `deck_snapshot` (one "slide N" label + image block per slide). Built on
  `Slide.export_image` (no PDF/PyMuPDF, no new dependency). The earlier folder-based
  `deck.export_images` (v0.4) stays for bulk-to-disk. **Still open:** a `jpg`-quality
  knob and per-slide `width`/`height` overrides (snapshot's lever is `max_dim`
  only).
- [ ] **`deck.save()` / `save_as(path)`** over `Presentation.Save` /
  `SaveAs(path, format)` — explicit-only, never implicit. Expose
  `Presentation.Saved` (dirty flag) + `.Path` on `status`/`deck` reads so an
  agent can see unsaved state. **Spike:** `Save` on a never-saved deck raises
  (no path) — surface as a clear error, not a COM blob.

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
  (partial alpha), gradient (`.OneColorGradient`/`.TwoColorGradient`/`.PresetGradient`
  + `.GradientStops`), picture (`.UserPicture(path)`), patterned fills, and line
  `.DashStyle` / arrowheads — the fiddly/long-tail cuts below.
- [ ] **Effects (second cut):** `Shape.Shadow` (`MsoShadowFormat`),
  `.Glow`, `.SoftEdge`, `.Reflection`, `.ThreeD`. Start with shadow (the common
  ask); the rest are opportunistic.
- [ ] **Per-slide background:** `Slide.FollowMasterBackground = msoFalse` +
  `Slide.Background.Fill` — the per-slide override of v0.9's master background
  (which is deck-wide). Same `MsoFillFormat` surface as shape fill, so it falls
  out of the same helper.
- **Constants:** `MsoFillType`, `MsoGradientStyle`/`MsoPresetGradientType`,
  `MsoLineDashStyle`, `MsoPatternType`, `MsoShadowType` — added as each verb
  needs them (don't pre-populate, per convention #7).
- **CLI/MCP:** the solid cut shipped as CLI **`shape fill`** + `shape add
  --fill/--line/--line-width` and MCP `format` (`fill_color`/`line_color`/
  `line_width`) — *not* the originally-sketched `format-shape` /
  `shape_set_fill`/`shape_set_line` ops. The remaining `--fill-gradient`/
  `--line-dash`/`--shadow` knobs land with their cuts above. All through
  `deck.edit()`.
- **Spike:** gradient stops are the fiddly bit (`GradientStops.Insert2(color,
  position, transparency, brightness)` ordering/clearing) — ship **solid + line +
  simple two-colour gradient** first, defer multi-stop. Confirm `UserPicture`
  takes an absolute path (the `Export` relative-path footgun likely recurs).

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

- [ ] **Hyperlinks & actions.** `Shape.ActionSettings(ppMouseClick).Hyperlink`
  (`.Address` for URLs/files, `.SubAddress` for "jump to slide N") and the
  text-level `TextRange.ActionSettings`. Lets an agent build clickable
  navigation, agenda links, "back to TOC" buttons. `set_hyperlink(anchor, *,
  url=None, slide=None)`; reads emit any existing link per shape/run.
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

- [ ] **Slide transitions (easy, ship first).** `Slide.SlideShowTransition`:
  `.EntryEffect` (`PpEntryEffect`, a large but flat enum), `.Duration`,
  `.AdvanceOnTime`/`.AdvanceTime` (auto-advance), `.AdvanceOnClick`, sound.
  `slide.set_transition(effect, *, duration, advance_after)`; reads round-trip
  cleanly. Friendly names → `PpEntryEffect` (the `chart_type_for` pattern).
- [ ] **Animations (fiddly, second cut).** Modern path: `Slide.TimeLine.
  MainSequence.AddEffect(Shape, EffectId, Level, Trigger)` →
  `Effect.Timing`/`.EffectType`/`.Exit`/`.EffectParameters`. `MsoAnimEffect` is a
  **huge** enum and the trigger/timing model is intricate. Start with
  **entrance/exit/emphasis on a whole shape** (the 80% ask: "fade this in"),
  defer per-paragraph and motion-path effects. **Spike:** read-back fidelity
  (does an `AddEffect` round-trip its `EffectType`/`Timing`, or is some of it
  write-only like SmartArt assistant nodes?) before promising a `read()`.
- **Caveat:** like SmartArt, expect some properties to be **write-only /
  non-round-tripping** — find them in the spike and scope the first cut to what
  reconstructs.

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
- [ ] **Media.** `Shapes.AddMediaObject2` (video/audio embed/link),
  `Shapes.AddOLEObject`. Niche but occasionally asked.
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
navigation and motion. Everything below the line is real but waits for a concrete
workflow to pull it in — the same demand-driven discipline that kept the
constants module from being pre-populated.

When in doubt, the rule that built v0–v0.9 still holds: **open the equivalent
wordlive module if one exists, spike the COM behaviour on a live deck, write the
one-line finding here, then harden across all four front-ends (wrapper, CLI, MCP,
tests) together.**

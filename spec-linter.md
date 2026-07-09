# Linter + formatting regularizer — design sketch

> Status: **sketch (2026-07-08)** — nothing built yet. This is the PowerPoint
> sibling of wordlive's `spec-linter.md`, written (like `spec.md`) as **the diff
> against wordlive**. Read wordlive's `spec-linter.md` first; this document keeps
> the shared engine one paragraph long and spends its length on the parts
> PowerPoint's 2-D object model forces to be different. Roadmap home: a new
> Priority-1 tier in `roadmap.md`; `IMPLEMENTATION.md` will track staged progress.

> **The one-liner.** Audit a deck for presentation-quality defects
> (`deck.lint()`), then autofix the mechanical ones in one atomic-undo step
> (`deck.regularize()`). Pure composition over shipped write verbs (`format`,
> `shape_move`/`shape_resize`, `shape_align`, `set_headers_footers`, `write`,
> `set_paragraphs`, `table_set_border`) — **no new COM write surface**; the new
> work is a richer *format-probe read* and the rule engine.

---

## 1. Why this, why now

The user's own words for the problem this solves: *"the final cleanup of a
PowerPoint file is extremely time intensive — making sure all the headers are the
same font and size, the shapes are lined up properly, no empty bullets, typos,
checking table formatting especially numerical tables, and making sure the
copyright/confidential marking is present, checking presence/absence of slide
numbers, slide size."*

Every one of those is:

- **objective** (you can write the rule down),
- **mechanical** (the fix is deterministic), and
- **already expressible** in pptlive's verbs (`format`, `shape_move`,
  `shape_align`/`distribute`, `set_headers_footers`, `set_paragraphs`,
  `table_set_fill`/`set_border`, cell `format`).

That combination is exactly what an agent should own — and, more than in Word, it
maps to what a *human* does in the last hour before a deck ships. The linter is
the highest-leverage next feature for pptlive for the same reason it was for
wordlive: it's **composition**, not new COM.

The user-list → cluster map (the spec is organized to answer all eight):

| User's ask | Cluster (§5) | Primitive (§5b) |
|---|---|---|
| "all the headers are the same font and size" | `titles` — `title-font-consistent` | P2 peer-mode |
| "shapes are lined up properly" | `alignment` — `edge-alignment`, `title-position-consistent`, `shape-off-slide` | P3 geometry |
| "no empty bullets" | `typography` — `empty-bullet`, `trailing-empty-paragraph` | P4 text-scan |
| "typos" | `proofing` — deferred, needs a spike (§9) | (spell-check probe) |
| "table formatting, especially numerical tables" | `tables` — `table-numeric-right-align`, `table-decimal-consistent` | P2 + cell-scan |
| "copyright / confidential marking present" | `boilerplate` — `confidentiality-notice`, `copyright-notice` | P5 deck-scan |
| "presence/absence of slide numbers" | `boilerplate` — `slide-number-present` | P5 deck-scan |
| "slide size" | `deck` — `slide-size` | P5 deck-scan |

## 2. The core reframing — where PowerPoint forces the biggest redesign

This is **the** diff, so it leads. Wordlive's whole linter rests on one idea
(§2 there): *consistency = "no direct formatting fighting the applied style,"*
because Word documents are style-driven and the defects are direct overrides that
drifted from a named style. **That idea mostly doesn't port**, because:

> **PowerPoint has no named paragraph styles.** There is no `Presentation.Styles`
> analog; pptlive's v0.3 note records this — "styling" in PowerPoint *is* direct
> font formatting. Placeholders inherit from the **layout → master text-style
> cascade** (`SlideMaster.TextStyles(ppTitleStyle/ppBodyStyle)`), but a free
> textbox has **no style at all** — every property on it is direct.

So Word's two-layer "effective vs style" compare only even *applies* to
placeholder text, and even there the baseline is a cascade, not one named style.
The reframing pptlive needs is that **consistency in a deck is dominated by two
axes Word barely has**, plus a weaker version of Word's cascade compare:

| Axis | What "consistent" means | Detection primitive | Example |
|---|---|---|---|
| **Cross-object uniformity** | peers look alike | **P2 · mode/dominant across peers** — collect a property over a peer set (all title placeholders, all body bullets at one level, all cells in a table column), flag the minority against the dominant value | every slide title is Calibri 40 except slide 7's at 36 |
| **Spatial regularity** | objects are placed on a grid | **P3 · geometry scan** — `geometry_report()` (shipped) + edge-clustering | two boxes whose left edges are 3 pt apart (meant to align); a title that jumps 12 pt between slides; a shape hanging off the slide |
| **Cascade adherence** *(the Word analog, weakened)* | a placeholder's direct formatting matches its master/layout text style | **P1 · placeholder-cascade probe** — effective font vs the resolved master/layout baseline (`format_info()`, §7) | a title placeholder overriding the master title font |

The **mode-across-peers** primitive (P2) is the star, and it's genuinely
PowerPoint-native — wordlive only used it once (`table-style-consistent`, "flag
minority vs dominant"); here it carries the headline rules the user asked for
first. It's also *robust to templates*: it judges the deck against **itself**
(internal consistency), so it works with no configuration on a deck built from any
template. A profile (§6) can instead pin the target to the master cascade or to
explicit house-style values — the brand path.

The three **kinds** and their config story are otherwise unchanged from wordlive:

| Kind | Needs config? | Detection | Example |
|---|---|---|---|
| **consistency** | no | minority ≠ dominant peer value (or ≠ cascade baseline) | mixed title fonts; a stray bullet indent |
| **structural** | no | objective defect | empty bullet; shape off-slide; numeric column left-aligned |
| **policy** | yes (a profile) | value ≠ profile target | slide size must be 16:9; footer must carry the confidentiality notice |

Consistency + structural rules ship with sensible defaults and need no config;
policy rules are opt-in and read their target from a **profile** (§6).

## 3. Surface

```python
findings = deck.lint(rules=None, within=None, profile=None)                 # pure read
report   = deck.regularize(rules=None, within=None, profile=None,
                           dry_run=False, allow_content=False)
```

- **`deck.lint(...)`** → a list of `Finding`s (§4). Read-only: mutates nothing,
  moves no view/Selection, leaves `Saved` untouched. **Cheaper than wordlive's** —
  no repagination step, because slide geometry is always available (no page-flow
  model to settle).
- **`deck.regularize(...)`** → applies the **fixable** subset inside a single
  `deck.edit("Regularize formatting")` (one Ctrl-Z reverts the whole pass via
  `StartNewUndoEntry`, exactly like every other pptlive write), returning
  `{applied, skipped, deferred, findings}`. `dry_run=True` plans without writing.
- **`rules`** selects/deselects by id or tag (`["titles", "alignment"]`,
  `{"exclude": ["edge-alignment"]}`); `None` = the default set (all on-by-default
  consistency + structural; no policy unless a `profile` enables them).
- **`within`** scopes the audit — **and here PowerPoint diverges cleanly from
  wordlive.** Word's `within` is a `(start, end)` character span. pptlive has no
  deck-wide character stream, so `within` is a **container anchor**: `slide:S`
  (audit one slide), a `Slide`, or a shape anchor (`shape:S:N` / `shapeid:S:ID` /
  `ph:S:KIND` — audit one shape and, for peer rules, still compare it against its
  deck-wide peers but only *emit* findings on it). Scope is by **containment**, not
  span overlap.

### Surfaces (all four must agree — the pptlive contract)

- **Python:** `deck.lint` / `deck.regularize`.
- **CLI:** `pptlive lint [--rules …] [--profile …] [--within ANCHOR]` (JSON
  findings) and `pptlive regularize [--dry-run] [--allow-content] …`.
- **exec op:** `regularize` is a **write** op (joins an atomic batch and rides its
  one undo entry, `own_undo=False` inside `run_batch` — same pattern wordlive
  uses); `lint` stays a read (CLI/MCP only, no op — like `selection`/`geometry`).
- **MCP:** `ppt_read op="lint"`, `ppt_edit op="regularize"` (both flow through the
  shared `_read_core`/`_edit_core`, so `ppt_batch` gets them for free).

## 4. The `Finding` shape

Ported near-verbatim from wordlive; the two diffs are the **anchor grammar** and a
new **`slide`** field, because a finding in a deck is located in 2-D (which slide,
which shape) rather than at a character offset.

```jsonc
{
  "rule": "title-font-consistent",       // stable id
  "kind": "consistency",                 // consistency | structural | policy
  "severity": "warning",                 // error | warning | info
  "slide": 7,                            // NEW vs wordlive — the 1-based slide
  "anchor_id": "shapeid:7:5",            // a real, DRIFT-PROOF anchor (see below)
  "message": "Title font 'Calibri 36' differs from the deck's dominant 'Calibri 40'.",
  "fixable": true,
  "fix": {                               // present iff fixable; a real exec op
    "op": "format", "anchor_id": "shapeid:7:5", "font": "Calibri", "size": 40
  },
  "adds_content": false,                 // gate for insert/destroy fixes (§8)
  "observed": "Calibri 36",
  "expected": "Calibri 40"
}
```

`fix.op`/`fix.args` are literally a pptlive **exec op** (or a list of them), so
`regularize` is "lint, then run each finding's `fix` through the existing
`run_batch` loop" — the fix path reuses the audited, warning-emitting op
vocabulary instead of a parallel writer.

**Anchor drift is a real hazard here that wordlive didn't have.** `regularize`
applies many fixes in one batch, and some fixes (`shape_align`, a delete) **shift
z-order**, so a finding anchored by the volatile `shape:S:N` could point at the
wrong shape by the time its fix runs. Therefore **findings anchor by the
drift-proof forms** — `shapeid:S:ID` (stable `Shape.Id`) or `ph:S:KIND` — never
`shape:S:N`. This is the same discipline the CLAUDE.md "shapeid everywhere" rule
already imposes on chained edits, applied to the linter.

## 5. The rule catalogue (v1)

Mapped to the recurring hand-edits, grouped by the user's list. Each row: how it's
**detected** and how it's **fixed** (the exec op). All fixes idempotent unless
noted. Default `on` = runs in the `rules=None` set.

### Titles & text consistency  *(P2 peer-mode — the headline cluster)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `title-font-consistent` | consistency | title placeholders (`ph:S:title`/`ctrtitle`) whose effective font name/size/bold/color ≠ the deck's dominant title value | `format(font=, size=, …)` to the dominant | on |
| `body-font-consistent` | consistency | body placeholders whose per-level font ≠ the dominant body font at that indent level | `format` to the dominant | on |
| `off-theme-color` | consistency | a run/shape color that is a **direct RGB** close to — but not equal to — a theme slot (uses the shipped `color_source`: `direct` vs `theme`) | `format(color=<theme slot>)` | off (`branding` tag) |
| `mixed-runs-in-title` | consistency | a title whose font reads mixed across runs where uniformity is expected (the `text_frame_status`/`run_sizes` mixed tell) | report-only (which run is the outlier needs a run-walk) | on (report) |

The dominant value is the **mode** over the peer set; ties and near-uniform decks
(one clear majority) are handled by a threshold (a rule fires only when a clear
dominant exists — `≥ 60%` of peers, configurable). This is the internal-consistency
default; a `house_style` profile (§6) replaces "dominant" with the pinned target.

### Alignment & geometry  *(P3 — the biggest net-new vs Word)*

Built on the **shipped** `Slide.geometry_report()` (slide size + each shape's
`box` + `off_slide` flag + `overlaps` pairs) and `deck.page_setup()`.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `shape-off-slide` | structural | a shape wholly or mostly outside the slide bounds (`off_slide`) | report-only (nudging onto the slide needs judgment; opt-in `shape_move`) | on (report) |
| `edge-alignment` | consistency | 2+ shapes whose left/right/top/bottom/center edges fall within a tolerance (default 3 pt) of each other but aren't equal — the "meant to line up" defect | `shape_align` to the shared edge (shipped op) | off (`alignment` tag) |
| `title-position-consistent` | consistency | title placeholder `left`/`top`/`width`/`height` that deviates from the deck-dominant title box (the "jumpy title" defect) | `shape_move`/`shape_resize` to the dominant box | on |
| `placeholder-off-layout` | consistency | a placeholder moved/resized off its `CustomLayout` position (uses the shipped `reset_to_layout` baseline) | `shape_reset_layout` | off (`alignment` tag) |
| `overlap-unintended` | structural | non-overlapping-by-design shapes overlapping (from `overlaps`, filtered to text-bearing shapes) | report-only | off (`alignment` tag) |

`edge-alignment` clusters edge coordinates across a slide's shapes; a cluster
whose spread is `0 < spread ≤ tol` is a finding, fixed by snapping the members to
the cluster's dominant (or mean) value via `shape_align`. Rotation is reported,
not accounted for (axis-aligned only — same honest scope as `geometry_report`).

### Whitespace, empty content & typography  *(P4 text-scan — ports from wordlive §A)*

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `empty-bullet` | structural | an empty paragraph between two non-empty bullets in a placeholder/textbox | delete the paragraph via `set_paragraphs` (rewrite without it) | on — **`adds_content`** (destroys content; withheld unless `--allow-content`) |
| `trailing-empty-paragraph` | structural | trailing blank paragraph(s) in a text frame | trim via `set_paragraphs` | on — **`adds_content`** |
| `trailing-whitespace` | structural | paragraph text ends in space/tab | `find_replace` regex trim | on |
| `double-space` | consistency | runs of 2+ spaces in body text | `find_replace` collapse | on |
| `space-before-punctuation` | consistency | ` ,` ` .` ` ;` ` :` | collapse | on |
| `empty-placeholder` | structural | a content/body placeholder with no text on an otherwise-populated slide (heuristic; prompt text is *not* real `.Text` — see caveat) | report-only | off (`finalization` tag) |
| `straight-quotes` / `em-dash-usage` | consistency | (ported from wordlive §A, false-positive-prone) | smart-quote / report | off (`typography` tag) |

**Caveat (pinned as a spike):** PowerPoint placeholder *prompt* text ("Click to
add title") is **not** returned by `TextRange.Text` — it's a rendering artifact. So
a "prompt text left in" rule likely isn't detectable via `.Text`; `empty-placeholder`
detects the *absence* of text on a slide that has other content, which is the
useful signal. Confirm no COM property exposes prompt-vs-real before promising more.

### Tables  *(P2 + a cell-scan; the user's "numerical tables")*

Built on shipped cell anchors (`cell:S:N:R:C`), cell `format`, and
`table_set_fill`/`table_set_border`.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `table-numeric-right-align` | structural/heuristic | a column whose non-empty body cells nearly all parse as numbers (`$`, `%`, `,`, `(neg)`) but aren't right-aligned | per-cell `format(alignment="right")` | on |
| `table-decimal-consistent` | consistency | a numeric column with mixed decimal-place counts (`1.2` next to `1.20` next to `1`) | report-only (rounding is a content decision) | off (`tables` tag) |
| `table-empty-cell` | structural | an empty cell inside an otherwise-full data row | report-only | off (`tables` tag) |
| `table-header-not-emphasized` | consistency | row 1 of a data table not bold / not shaded while data rows are plain | `format(bold=True)` + `table_set_fill` on row 1 | off (`tables` tag) |
| `table-fill-consistent` | consistency | banding/fill that's uneven across data rows (minority vs dominant per the shipped per-cell `fill` read) | `table_set_fill` | off (`tables` tag) |

### Boilerplate, slide numbers & deck-level  *(P5 deck-scan)*

Built on shipped `Slide.headers_footers` / `Master.headers_footers`,
`deck.page_setup()`, `deck.find()`, and `deck.sections`.

| id | kind | detect | fix | default |
|---|---|---|---|---|
| `slide-number-present` | policy | no visible slide-number placeholder across content slides | `set_headers_footers(slide_number=True)` | off (`boilerplate` tag) — **`adds_content`** |
| `confidentiality-notice` | policy | profile-supplied notice text not found in any footer/shape (via `find`) | `set_headers_footers(footer=<text>)` | off (profile; `boilerplate`) — **`adds_content`** |
| `copyright-notice` | policy | profile `©` / text not present | `set_headers_footers(footer=…)` | off (profile; `boilerplate`) — **`adds_content`** |
| `slide-size` | policy | `page_setup()` aspect ≠ profile target (default 16:9) | report-only (a global resize reflows every slide — loud) | off (`deck` tag) |
| `footer-consistent` | consistency | footer text differs across slides that should share one (minority vs dominant) | `set_headers_footers(footer=<dominant>)` | off (`boilerplate` tag) |
| `orphan-slide-transition` | consistency | one slide with a transition the rest of its section lacks (or vice-versa) | report-only | off (`deck` tag) |

## 5b. Catalogue v2 — brainstormed backlog

The v1 set (§5) needs three probes: the peer-mode helper (P2), `geometry_report`
(P3, shipped), and the text-scan (P4). The backlog pushes into the remaining
primitives. **Batch by primitive**, not by category — build the primitive once,
light up its cluster.

### Detection primitives (build order)

| Primitive | COM / shipped surface | Unlocks |
|---|---|---|
| **P1 · Placeholder-cascade probe** | `format_info()` (§7, new): effective font vs resolved master/layout text style | `*-font-consistent` against the *master* baseline; `placeholder-off-layout` font half; `off-theme-color` |
| **P2 · Peer-mode scan** | a `deck` walk collecting a property across a peer set, `mode()` helper | every `*-consistent` rule; the headline cluster |
| **P3 · Geometry scan** | `Slide.geometry_report()` (shipped) + edge-clustering | `edge-alignment`, `title-position-consistent`, `shape-off-slide`, `overlap-unintended` |
| **P4 · Text scan** | per-frame paragraph/run walk + `find` | whitespace, empty-bullet, table numeric parse, prompt-text |
| **P5 · Deck scan** | `page_setup`, `headers_footers`, `sections`, `find`, `deck.theme` | slide size, boilerplate/notices, slide numbers, transition/section hygiene |
| **P6 · Proofing (spiked ✅)** | **borrowed** — a hidden `Word.Application` (`CheckSpelling`/`GetSpellingSuggestions`); PowerPoint's own COM has none | `spelling` (typos) — §9 |

### Tag taxonomy

`titles`, `fonts`, `alignment`, `geometry`, `typography`, `tables`, `boilerplate`
(alias for the slide-number + notice sub-cluster), `branding` (theme/house-style),
`deck`, `finalization`, `proofing`, `accessibility`. `--rules alignment` /
`--rules finalization` become the headline ergonomics; profiles (§6) toggle tags +
supply policy targets. A composite `--rules finalization` (off-slide junk +
empty placeholders + empty bullets + notices + slide numbers) is the natural
"is-this-ready-to-present?" pre-send check.

### Default stance (mirrors wordlive's 2026-06-19 decision)

Unambiguous structural defects (off-slide shape, numeric column not right-aligned,
empty bullet) ship **on**. Opinion-flavored consistency rules (edge-alignment
tolerance, em-dash, decimal rounding) ship **off** behind a tag/profile. Anything
that **adds or destroys content** (delete an empty bullet, insert a footer notice,
turn on slide numbers) is fixable but flagged `adds_content=True` and **withheld**
unless the caller opts in (§8).

## 6. Profiles (policy rules + house style)

A **profile** is a small declarative config that (a) enables policy rules and
supplies their targets and (b) optionally pins consistency targets to an explicit
**house style** instead of the deck's own dominant values — the brand/template
path, which is *stronger* in PowerPoint than in Word because brand decks ship
strict templates.

```jsonc
// pptlive.lint.json  (or passed inline)
{
  "extends": "default",
  "rules": {
    "slide-size":            { "enabled": true, "target": "16:9" },
    "slide-number-present":  { "enabled": true },
    "confidentiality-notice":{ "enabled": true, "text": "CONFIDENTIAL — Do not distribute" },
    "copyright-notice":      { "enabled": true, "text": "© 2026 Acme Corp" },
    "edge-alignment":        { "enabled": true, "tolerance": 3.0 },
    "double-space":          { "enabled": false }
  },
  "house_style": {           // optional: pin consistency targets instead of "dominant"
    "title": { "font": "Calibri", "size": 40, "bold": true, "color": "#1F3864",
               "left": 38, "top": 30 },
    "body":  { "font": "Calibri", "size": 18 },
    "theme_palette": "acme"  // named theme the deck must match (deck.theme)
  }
}
```

Without a profile: consistency rules judge each object against the deck's own
dominant peer value (internal consistency), structural rules run, policy rules stay
off. With a `house_style`, consistency rules judge against the named targets *and*
can fix deck-wide by updating the **master text style** (`master_format_text_style`)
or **theme** (`theme_set_color`/`set_font`) so the whole deck follows — the
brand-enforcement path. `Profile.load` accepts a path / dict / `None`, exactly like
wordlive's `_lint_profile.py`. CLI `--profile PATH`; a discoverable default file
name (`pptlive.lint.json`) so a repo/template can check one in.

## 7. The format-probe read + baseline resolution

Two things land here (mirroring wordlive §7): a **new public read** — the read
mirror of `format`/`format_paragraph` — and the **baseline resolution** the
consistency rules compare against. The pptlive diff is entirely in the baseline.

### 7a. `anchor.format_info()` — the public read (new surface)

Returns an anchor's *effective* paragraph + character formatting, each field
annotated with whether it's a **direct value** and (for placeholders) the
**cascade baseline** it should inherit:

```jsonc
// deck.anchor_by_id("ph:7:title").format_info()
{
  "anchor_id": "ph:7:title",
  "placeholder": "title",
  "cascade": "layout → master:ppTitleStyle",   // where the baseline comes from
  "paragraph": {
    "alignment":   {"value": "center", "baseline": "center", "override": false},
    "space_after": {"value": "6pt",    "baseline": "0pt",    "override": true},
    "indent_level":{"value": 1,        "baseline": 1,        "override": false}
  },
  "font": {
    "name":  {"value": "Calibri", "baseline": "Calibri", "override": false},
    "size":  {"value": 36.0,      "baseline": 40.0,      "override": true},
    "bold":  {"value": true,      "baseline": true,      "override": false},
    "color": {"value": "#1F3864", "baseline": null, "source": "direct"},  // reuses color_source
    "mixed": ["size"]   // fields that vary across runs
  }
}
```

- **Read-only**, same vocabulary as the write verbs, so read and write mirror
  field-for-field.
- `value` = effective; `baseline` = the resolved master/layout text-style value
  for that placeholder kind + indent level; `override` = `value ≠ baseline`. For a
  **free textbox** (no cascade) `baseline` is `null` and `override` is undefined —
  those anchors are judged only by the peer-mode rules, never by cascade adherence.
- Reuses the **shipped** `color_source` (`direct`/`theme`/`mixed`) + `theme_color`
  for the color field, so it's consistent with existing shape reads.
- Surfaces: `anchor.format_info()` Python; CLI `read format --anchor-id ID`; MCP
  `ppt_read op="format_info"`. **No exec op** (pure read).

### 7b. Baseline resolution (the diff from wordlive §7b)

Word resolved the baseline from `Range.ParagraphStyle` (one named style, and Word
resolves the `BaseStyle` chain for you). PowerPoint has to **walk the cascade by
hand**: a placeholder's baseline is its matching placeholder on the slide's
`CustomLayout`, and failing that the master `TextStyles(kind).Levels(indentLevel)`.
pptlive already reaches every rung — `Shape.reset_to_layout()` walks the
`CustomLayout` placeholder, and `deck.master.read()` reads the text styles — so
resolution is composition over shipped reads, not new COM. **Spike to pin
(before hardening):** confirm `CustomLayout` placeholder → master `TextStyles`
fallback resolves a concrete font for each placeholder kind + level across a couple
of templates (the analog of wordlive's live-validated override detection).

**The peer-mode baseline (P2, the more-used path) needs no cascade at all:** it
collects `format_info()["font"]["*"]["value"]` across the peer set and takes the
mode. This is why the peer rules work on any deck with zero config, and why they're
the default over cascade-adherence rules.

### 7c. Targeted fix (idempotent)

Same contract as wordlive: the default fix writes the **target value back as a
direct property** (`format(size=40)`), which is visually correct and **idempotent**
(re-running writes the same value → no-op). `reset_to_layout` / `shape_reset_layout`
is the aggressive strip-to-layout equivalent of Word's `Font.Reset()` (for
`placeholder-off-layout`) — behind a tag, not default, because it nukes intentional
overrides. **Idempotency is a test invariant:** build a messy deck → `regularize` →
`regularize` again → assert the second pass's `applied` is empty (the reason
targeted-write is the default).

## 8. Politeness & safety

- `lint` is a pure read — no view/Selection move, no repagination (cheaper than
  Word's), `Saved` untouched.
- `regularize` runs inside `deck.edit("Regularize formatting")` → snapshots/restores
  the viewed slide + Selection, one atomic Ctrl-Z for the whole pass (pptlive's
  shipped `StartNewUndoEntry` fence). Because a `regularize` pass can `shape_align`
  and delete, it inherits the CLAUDE.md **"follow the work"** consideration — but
  since it's a *pure-formatting* pass by default (no slide adds), it keeps the polite
  view-restore (it does not opt into `allow_view_move`).
- **The `adds_content` gate (ported verbatim).** A fixable finding whose fix
  inserts or destroys content — delete an empty bullet, insert a footer notice, turn
  on slide numbers — sets `adds_content=True`. `regularize` **withholds** those by
  default and reports them in a `deferred` bucket; the caller opts in with
  `allow_content=True` / `--allow-content` / `allow_content: true`. Pure in-place
  formatting/alignment fixes leave `adds_content=False` and apply by default.
- **No Track Changes** — PowerPoint has none, so wordlive's track-changes-aware
  bullet is simply dropped.

## 9. Deferred (v1 boundaries)

- **Typos / proofing (P6) — SPIKED 2026-07-08 (`scripts/proofing_spike.py`), and
  it's a GO via a borrowed engine.** PowerPoint's own COM is a **dead-end**: no
  `Application.CheckSpelling`/`GetSpellingSuggestions`, no
  `TextRange.SpellingErrors` (only `Application.LanguageSettings` and
  `TextRange.LanguageID` → `1033` exist). *But* a **hidden `Word.Application`
  borrowed over COM is a fully working checker** — `CheckSpelling(word)` returns
  the right bool and `GetSpellingSuggestions(word)` returns candidates; Word runs
  `Visible=False` (PowerPoint can't). Both live misspellings on the test deck were
  caught with correct suggestions (`favoritely`→*favorite*, `Potatoe`→*Potatoes*).
  Because pptlive **tokenizes each frame itself** (regex → `(start, length)`), the
  token-level check yields **exact spans** — tier-A anchoring on a tier-B
  primitive, so a `spelling` rule can emit a precise `(shapeid, start, length)`
  finding plus a `set_text`/`find_replace` fix from the top suggestion.
  **Design notes the spike pinned:**
  - The borrow needs a **scratch invisible document** (`word.Documents.Add()`):
    `CheckSpelling` works on bare strings, but `GetSpellingSuggestions` raises "no
    document is open" without one. Spin Word up **once per lint pass**, reuse it for
    every token, `doc.Close(0)` + `Quit()` in a `finally`.
  - **Gate on Word being installed** — `Dispatch("Word.Application")` failing means
    proofing is *unavailable*, not an error: skip the `proofing` cluster gracefully
    (Word is near-ubiquitous on an Office box, but PowerPoint-only installs exist).
  - **False-positive hygiene:** skip ALL-CAPS acronyms, tokens with digits, and
    known brand terms; honor `TextRange.LanguageID` later (v1 = en-US default).
  - This is the natural home for a future pptlive **`proofing()`** surface (the
    missing mirror of wordlive's) — the `spelling` rule is its first consumer, and
    it belongs in a **later batch** (P6), after the formatting clusters, since it
    carries the cross-app dependency. `_com.py` gets a small `word_speller()` helper
    (the only place that may `Dispatch("Word.Application")`), keeping the pywin32
    seam rule intact.
- **Accessibility** — alt-text-missing on pictures/charts (pptlive already reads
  `alt_text` on every shape, so `image-alt-missing` is a *cheap* early structural
  rule worth pulling into v1), reading-order, contrast; the deeper set belongs with
  a future *prepare-for-sharing* product that can call the linter.
- **Cross-object *format* deep-compare** — matching not just font but effects/fills
  across peers; start with font/size/color/position.
- **A custom-rule plugin API** — ship the built-in catalogue first; add
  extensibility only on a concrete need.

## 10. Build order

Mirrors wordlive's primitive-driven staging, resequenced so the user's top asks
land first.

1. **Foundation:** `anchor.format_info()` (§7a) + the peer-mode helper (P2) + the
   cascade baseline resolver (§7b, spiked). Every rule consumes one of these.
2. **P2 headline cluster (the user's #1):** `title-font-consistent`,
   `body-font-consistent`, `title-position-consistent`. Highest signal, pure
   composition over `format`/`shape_move`.
3. **P3 geometry cluster (the user's #2):** `shape-off-slide`, `edge-alignment` (→
   `shape_align`), `placeholder-off-layout`. Built on the shipped `geometry_report`.
4. **`regularize`** (the `run_batch`-over-findings loop) + the idempotency smoke
   test + the `adds_content` gate.
5. **P4 text cluster (the user's #3/#5):** `empty-bullet`, whitespace rules,
   `table-numeric-right-align` + the numeric-column parse.
6. **P5 deck cluster (the user's #6/#7/#8) + profiles (§6):** `slide-number-present`,
   `confidentiality-notice`/`copyright-notice`, `slide-size`, the `Profile` loader.
7. **Wire all four front-ends** (Python / CLI / exec op / MCP) + both SKILL guides +
   docs (a "hand off a clean deck" cookbook entry).
8. **Later:** P1 cascade-adherence rules against the master baseline; the
   `house_style` brand-enforcement fixes (`master_format_text_style`/`theme_set_*`);
   the proofing spike (§9); accessibility.

Steps 1–4 are the foundation slice (the first shippable linter: the two clusters
the user named first, plus the fix loop). The rest continue primitive-driven, each
a batch that builds one probe and lights up its cluster.

## 11. Open questions

1. **Peer-set definition for the mode rules.** Is the peer set for
   `title-font-consistent` *all* title placeholders deck-wide, or per-section? A
   deck with a distinct section-divider style is a legitimate two-mode deck. Lean:
   deck-wide dominant with a per-section override when `deck.sections` exist and a
   section is internally uniform but differs from the deck.
2. **`edge-alignment` tolerance + fix ambiguity.** What default tolerance (3 pt?)
   avoids false positives, and when a cluster of 3 shapes is near-aligned, snap to
   the mode, the mean, or the "most intentional" (e.g. the one matching a layout
   placeholder edge)? Needs a live pass on real decks.
3. **When is `title-position-consistent` fixed vs reported?** Moving a title is more
   visible than a font tweak; do we default it to report-only and only auto-move
   under `--allow-content` or a profile opt-in? (Leaning report-by-default for
   *position*, auto-fix for *font*.)
4. **Proofing dependency (§9).** Resolve the spelling-COM spike before committing to
   a `proofing` cluster at all.
5. **Slide identity in findings.** Emit `slide:S` (index, what users say) — but a
   `regularize` pass that reorders slides could invalidate a later finding's slide
   index. Do findings also carry the stable `SlideID` (Open Q #5 from `spec.md`),
   and does `regularize` re-resolve by it between fixes?

# pptlive — MCP Tool Test Session Feedback

**Date:** 2026-06-08
**Context:** Built an 8-slide deck *about* pptlive, entirely through the MCP tools (`ppt_read` / `ppt_edit` / `ppt_batch` / `ppt_render` / `ppt_show`), as a live shakedown. Goal was full verb coverage + catching friction.
**Environment:** PowerPoint (desktop COM), blank `Presentation1`, stock Office theme — fonts Aptos Display / Aptos, `accent1 #156082`, 960×540 (16:9).

---

## Summary

The core loop is solid and genuinely pleasant to drive: `ppt_batch` atomic grouping with an inline `render` at the tail ("build a slide, then look at it" in one round-trip) is the standout feature and worked every time once the machine was awake. One high-severity content bug (`write` newline handling) caused essentially all the build friction; one medium correctness gap (placeholder ambiguity); the rest are minor/doc-level. Two apparent showstoppers turned out to be an idle-stall artifact, not code.

---

## What worked well

- **`ppt_batch` atomic + embedded render** — one undo entry per batch, inline image back. Excellent. Used it for almost every slide.
- **`ppt_read`** — `status / slides / outline / slide / anchor / theme / layouts` all fast and accurate. `slide` read is the workhorse for anchor discovery.
- **`shape_add` returns the new shape's full descriptor** (incl. `anchor_id`) in the result — re-identifying without a follow-up read is a nice touch.
- **Placeholder anchors (`ph:S:KIND`) resolve correctly on freshly-added slides within the same batch** — no drift issue, as designed.
- Table: `shape_add kind=table` → default theme table style applied (header band + banding); `cell:S:N:R:C` writes and **per-cell `format` (font)** worked cleanly.
- Chart: `shape_add kind=chart` + `chart_set_type` (returned canonical `column_clustered` / `bar_clustered`); auto-inherited theme accents.
- SmartArt: `shape_add kind=smartart smartart_kind=cycle` with inline `nodes` rendered a polished, theme-colored diagram.
- Autoshape + textbox `shape_add`; **per-paragraph `format` with `indent_level`** (level 1 vs 2) worked once paragraphs were real (see #1).
- `master_format_text_style` applied deck-wide without moving the view.

---

## Issues (ranked)

### 1. `write` + `\n` collapses to a single paragraph, not multiple — **HIGH**

**Repro:**
```
ppt_edit op=write anchor_id=ph:2:body text="line A\nline B\nline C\nline D"
ppt_read op=anchor anchor_id=shape:2:2
```
**Expected:** four paragraphs → `para:2:2:1 … para:2:2:4`, individually addressable.
**Actual:** one paragraph (`para:2:2:1`) whose `text` is the four lines joined by `\n`. `para:2:2:2` → `not_found`.

**Why it's sneaky:** in a *bulleted* placeholder PowerPoint renders a bullet per soft-break line, so the slide *looks* like 4 paragraphs. The same write into an `object` placeholder addressed by `shape:` renders with no per-line bullets. Both are structurally **one paragraph** — the visual difference made it look like behavior diverged by anchor kind, but it doesn't. The reader reporting a single `\n`-joined paragraph compounds the confusion: an agent gets `not_found` on `para:…:2` with no way to see why.

**Impact:** per-line formatting and indenting are impossible on `\n`-written bodies. Every multi-level bullet slide had to be rebuilt via `write mode=insert_after` (the only path that creates real, addressable paragraphs). This was ~80% of the session's friction.

**Recommendation (pick one):**
- Preferred: make `write` split `\n` into real paragraph marks. That's the natural expectation for an LLM building a bullet list, and it makes `para:` anchors usable immediately.
- Or: keep soft-break semantics but **document loudly** that `insert_after` is required for separate paragraphs, *and* make the reader surface line-break structure (e.g. report the soft-break count) so the `not_found` is explicable.

---

### 2. Placeholder resolution has no ambiguity guard — **MEDIUM**

**Repro:** On a **Two Content** slide (two `object` placeholders, `shape:5:2` left / `shape:5:3` right):
```
ppt_read op=anchor anchor_id=ph:5:body   →   resolves, returns the LEFT placeholder, no error
```
**Expected:** either an `ambiguous` error (exit 5), consistent with `find_replace`, or an indexed form to disambiguate.
**Actual:** silently resolves to the first match.

**Two sub-problems:**
- `find_replace` errors on multi-match, but **placeholder resolution silently picks the first** — inconsistent ambiguity handling across the API.
- The content-placeholder kind reports as `object`, which **isn't in the documented `ph:` KIND set** (title/ctrtitle/subtitle/body/footer/date/slidenum), yet `body` aliases to it. Net effect: **there is no `ph:` form that addresses the *second* content placeholder** — you must fall back to `shape:S:N`.

**Recommendation:** error (or warn) on multi-match; add an indexed placeholder form (e.g. `ph:5:body#2`) or document that Two Content / Comparison bodies must be reached via `shape:`/`.Name`; add `object` to the documented KIND list with its `body` aliasing noted.

---

### 3. Render + batch "hangs" — **RETRACTED, not a bug**

Two 4-minute timeouts early on (`ppt_render slide_image` standalone, and `ppt_batch` with pure edits / `embed=false`) were caused by the **machine going idle mid-COM-call**, not by the tools. Both work flawlessly when the machine is awake; reproduced clean dozens of times after.

**Optional enhancement:** the failure surfaced as a generic 4-minute MCP timeout whose message *speculated* the server had crashed (it hadn't — reads/edits succeeded instantly before and after). A faster, more specific error on a stalled COM call (e.g. a short keepalive ping with an "app not responding / possibly idle" category) would save the operator a 4-minute wait and a misdiagnosis.

---

## Minor notes

- **Chart series order** didn't follow dict insertion order — passed `{"Edit live": …, "Regenerate": …}` but `Regenerate` rendered as series 1. Worth documenting the ordering rule (or preserving insertion order).
- **`master_format_text_style` requires `level`** even for `style=title` (which conceptually has one level). Error message is clear; consider defaulting `title`→`level 1`.
- **Unsaved-deck `name` is a moving target** — `status.name` flipped `Presentation1`→`pptlive` the instant the title was written (it tracks the window caption). `doc`-by-name targeting is therefore unstable on unsaved decks; consider also accepting a stable id/index.

---

## Open follow-ups for next session

- **Render sweep slides 2–8** to confirm the `master_format_text_style title` cascade took and didn't collide with any direct title formatting (session ended right after applying it — unverified visually).
- Verify `find` / `find_replace` behave correctly **across the soft-line-break single-paragraph bodies** from #1 (does a match span lines cleanly?).
- Exercise **`theme_set_color` / `theme_set_font` / `master_set_background`** — deliberately skipped this round to preserve the cohesive look; untested.
- Test **`shape_add kind=picture`** — no image asset on hand this session.
- Test the **slide-show / view verbs**: `ppt_show start/next/previous/goto/black/white/end` and `ppt_render navigate` — entirely untested.
- Confirm `shape:S:N` **z-order drift** re-resolution behaves as documented after deletes/reorders.

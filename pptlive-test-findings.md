# pptlive — Test Findings & Known Issues

**Component:** `pptlive` MCP extension (live PowerPoint control over COM)
**Tools under test:** `ppt_read`, `ppt_edit`, `ppt_batch`, `ppt_render`, `ppt_show`
**Environment:** Claude Desktop (MCP host) → live, unsaved deck (`Presentation1`), 8-slide self-describing test artifact
**Last updated:** 2026-06-08

**Status legend**
- **Confirmed** — reproduced and verified live in this session.
- **Reported** — carried from prior-session notes; not re-verified this session.

**Severity legend**
- **High** — silently produces wrong/malformed output, or blocks first use.
- **Medium** — surprising behavior or missing guardrail; correct usage requires undocumented knowledge.
- **Low** — cosmetic, or affects only error-message ergonomics.

---

## Open issues

### PPTLIVE-001 — `write` docstring contradicts behavior: `\n` makes soft breaks, not paragraphs
- **Severity:** High
- **Status:** Confirmed
- **Component:** `ppt_edit` op=`write`

**Summary.** The `write` op docstring states *"embed `\n` for multiple paragraphs."* In practice, a `\n`-joined string becomes **one** paragraph containing soft line breaks (`<a:br>`), not multiple paragraphs (`<a:p>`). An agent that trusts the documented contract builds a structurally malformed body every time. This is the root cause of the "paragraph trap" (the single highest friction source in prior testing).

**Evidence.** Slide 2's body (`shape:2:2`) was built from `\n`-joined text and reads back as a single paragraph:
- `ppt_read op=anchor anchor_id="shape:2:2"` → `paragraphs` array has length **1** (`para:2:2:1`), whose `text` contains three literal `\n`.
- `ppt_read op=anchor anchor_id="para:2:2:2"` → `not_found` (no second paragraph exists).
- Contrast slide 7 (`shape:7:2`), built with `insert_after`, which reads back as **6** real paragraphs joined by `\r`, each individually addressable (`para:7:2:1`…`:6`) with independent indent level and run formatting.

So the serializer's `\n` vs `\r` distinction is real and meaningful: **`\r` = paragraph break, `\n` = soft line break.** The defect is purely that the docstring labels `\n` as the paragraph separator.

**Impact.** Bodies built per the docs are a single bullet with soft breaks. Individual lines cannot be addressed (`para:S:N:P`), re-formatted, or re-indented. The failure is invisible in a render (see NON-ISSUE-2), so it goes unnoticed until something tries to target a line.

**Suggested fix (pick one):**
1. **Fix the docs (preferred).** Change the docstring to: *"embed `\n` for soft line breaks within a paragraph; create separate paragraphs with mode=`insert_after` (or embed `\r`)."* Soft-break-within-paragraph is a legitimate capability and shouldn't be lost — it just shouldn't hide behind the word "paragraph."
2. **Fix the impl.** Make `write` split `\n` into real paragraphs to honor the current wording. (Costs the soft-break capability unless a separate escape is added.)

**Follow-up to confirm:** verify whether `write` with embedded `\r` yields real paragraphs. If yes, the corrected doc can simply say "`\r` = paragraph, `\n` = soft break."

---

### PPTLIVE-002 — `tool_search("powerpoint")` surfaces only 2 of 5 tools
- **Severity:** Medium (first-run blocker)
- **Status:** Confirmed
- **Component:** MCP registry / tool descriptions

**Summary.** Searching the deferred-tool registry for the bare product name `powerpoint` returned only `ppt_batch` and `ppt_edit`. The orientation tools `ppt_read`, `ppt_render`, and `ppt_show` were not surfaced, so the documented first step (`ppt_read op=status`) targets an unloaded tool.

**Evidence.** First `tool_search("powerpoint")` → loaded `ppt_batch`, `ppt_edit` only. A second search with verb keywords (`ppt_read render show slides status`) was required to load the remaining three.

**Suggested fix.** Seed each tool's registry description with its action verbs ("read slides / inspect anchors", "render slide to image", "run/drive slideshow") so the product name alone retrieves the whole set. Optionally make the five co-retrieve as a bundle.

---

### PPTLIVE-003 — No read op exposes direct-vs-inherited run formatting
- **Severity:** Medium
- **Status:** Confirmed
- **Component:** `ppt_read` (op=`anchor`, op=`slide`)

**Summary.** Reads return resolved/effective text and paragraph attributes, but there is no way to read back whether a run's formatting (bold/color/font) is **direct** (applied on the run) or **inherited** (from layout/master). This makes it impossible to verify a cascade purely by reading — you must mutate the master and render to disambiguate.

**Evidence.** Verifying the deck-wide title cascade (`master_format_text_style style=title level=1 color=#156082`) could not be settled by reads. It required flipping the master title to a jarring color, rendering, and observing the flip (then restoring). A render proves the *result*, not the *mechanism*; only the destructive-then-restore probe proved inheritance.

**Suggested fix.** Add an inheritance-aware read — e.g. `op=anchor` returns per-run `{value, source: "direct"|"inherited"}` for font attributes, or a dedicated `op=runs`/`op=effective_format`. Lower priority than 001 but it's the difference between "looks right" and "is right."

---

### PPTLIVE-004 — `ph:S:body` silently resolves to first of two `object` placeholders
- **Severity:** Medium
- **Status:** Reported (prior session)
- **Component:** anchor resolution (`ph:S:KIND`)

**Summary.** On Two Content / Comparison layouts, `ph:S:body` silently resolves to the **first** of the two `object` placeholders with **no ambiguity error**. Content intended for the second column lands in the first.

**Workaround.** Address each column explicitly by z-order: `shape:S:N`.

**Suggested fix.** When a `ph:` semantic kind matches more than one placeholder on a slide, raise an `ambiguous` error (consistent with `find_replace`'s ambiguity guard) listing the candidate `shape:S:N` anchors, rather than silently picking the first.

---

### PPTLIVE-005 — Ambiguity error references CLI flags, not MCP param names
- **Severity:** Low
- **Status:** Confirmed
- **Component:** `ppt_edit` op=`find_replace` error message

**Summary.** The `ambiguous` error reads: *"N matches for '…'; pass `--all` or `--occurrence N` to disambiguate."* Those are **CLI flag** names. The MCP tool parameters are `replace_all` (boolean) and `occurrence` (int). An agent acting over MCP gets pointed at identifiers that don't exist in the schema.

**Evidence.** `find_replace find="pptlive"` → `error: ambiguous`, message: *"6 matches for 'pptlive'; pass --all or --occurrence N to disambiguate"*.

**Suggested fix.** Branch the hint by surface, or use neutral wording: *"set `replace_all=true` or `occurrence=N` to disambiguate."*

---

### PPTLIVE-006 — `find` `context` field flattens `\n`/`\r` to spaces
- **Severity:** Low (cosmetic)
- **Status:** Confirmed
- **Component:** `ppt_read` op=`find`

**Summary.** The human-readable `context` string in `find` results replaces paragraph/line separators with a single space, so a hit's surrounding structure isn't visible in the preview. The `start` offsets are unaffected and correct (separators are counted as one char each).

**Evidence.** `find "pptlive"` hit on `shape:2:2` showed context *"…open right now pptlive edits in place…"* where the source has a `\n` between "now" and "pptlive". Offset `start=81` correctly counts that `\n`.

**Suggested fix.** Optional: render separators as a visible glyph (e.g. `⏎` for `\r`, `↵` for `\n`) in `context`, or expose the raw separator. Purely a preview-readability nicety.

---

## Minor quirks (Reported — prior-session notes)

These are documented sharp edges rather than defects; capture them so they don't get re-discovered.

- **`master_format_text_style` requires `level`.** Omitting the outline `level` is an error. (Status: Reported; consistent with this session's usage where `level: 1` was always passed.)
- **Chart series order ≠ dict insertion order.** `series` (a `{name: [values]}` map) does not necessarily render in insertion order. If series order matters, do not rely on dict ordering.
- **Unsaved-deck `name` tracks the window caption.** For an unsaved deck, `name` reflects the live window title (e.g. `Presentation1`) and can change. Target by index, not by `name`, for `doc` until the deck is saved. (Status: Confirmed this session — `ppt_read op=status` reported `path: "Presentation1"` for the unsaved deck.)

---

## Not bugs — record so they aren't re-chased

### NON-ISSUE-1 — ~4-minute `render`/`batch` "hang" on an idle machine
A `render` or `batch` that appears to hang for ~4 minutes is caused by the host machine going idle mid-call, **not** by a code defect. Retry once the machine is awake. (Status: Reported.)

### NON-ISSUE-2 — PowerPoint paints a bullet per soft-break line
PowerPoint renders a bullet glyph on **each** soft-break line of a bulleted paragraph. As a result, the PPTLIVE-001 trap (one paragraph with `\n` soft breaks) is **visually identical** to a correctly built multi-paragraph body. The render gives no signal that anything is wrong; only `ppt_read op=anchor`'s `paragraphs` breakdown reveals the true structure. This is PowerPoint behavior, not a pptlive defect — but it's the reason 001 is so easy to miss, so the diagnostic path (read the paragraph breakdown, don't trust the render) is worth documenting. (Status: Confirmed.)

---

## Validated behaviors (regression baseline)

Confirmed working this session — useful as a checklist for future regression runs.

**Deck-wide title cascade**
- `master_format_text_style style=title level=1` cascades to slide titles **through intervening layout placeholders** (verified across Title and Content, Title Only, and Section Header layouts via a master-color flip-and-render).
- The op is correctly **scoped to the title style**: a same-colored non-title element (slide 8's `github.com/...` line) did **not** change when the title style flipped.

**`find` / `find_replace`**
- `find` locates hits **inside** a soft-break single paragraph and **inside** table cells, with correct per-anchor `start` offsets (separators counted as one char).
- `find_replace` single-match **auto-applies** and rewrites only the matched span — **soft-break structure and surrounding run formatting are preserved** (no paragraph fragmentation, no formatting loss). Verified on the slide-2 soft-break body.
- `find_replace` in a **table cell** rewrites cleanly with no cell-structure damage.
- **Ambiguity guard:** multiple matches with no disambiguator → `ambiguous` error, no mutation.
- **`occurrence=N`** targets the Nth match in **document order** (slide → z-order → in-anchor offset).
- **`scope="slide:S"`** correctly narrows a deck-wide-ambiguous term to a single in-scope hit.
- **`replace_all`** hits every match in document order, **including multiple matches within a single anchor**.
- **Smart-quote tolerance:** a curly-apostrophe query matches straight-apostrophe text.
- **Fuzzy ceiling:** matching is whitespace/smart-quote tolerant but **not** punctuation-eliding (a query without em-dashes does not match em-dash-separated source). This is the documented/expected scope, not a defect — recorded so the boundary is known.

**Workflow**
- `ppt_batch` with a trailing `render slide_image` (`embed=true`) successfully does "build-and-look" in a single round-trip; each batch is one undo entry (`atomic=true`).
- Reads (`status` → `slides` → `slide N` → `anchor`) are side-effect-free and do not move the user's view.

---

## Still untested (next sessions)

- **Slideshow verbs** (`ppt_show start/next/previous/goto/black/white/resume/end`) and `ppt_render op=navigate` — deliberately move the user's screen; run while watching the monitor.
- **`shape_add kind=picture`** — needs an image path.
- **`theme_set_color`, `theme_set_font`, `master_set_background`** — global styling ops skipped to preserve the look.
- **`shape:S:N` z-order re-resolution after a delete** — last open *correctness* question: do shape anchors silently re-point to the wrong shape after a delete shifts z-order? (Recommended next, as it's where a silent-corruption bug would live.)

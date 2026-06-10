Absolutely — here’s a practical review write-up based on this session.

---

# pptlive MCP Server Review
## Session notes and next-release recommendations

Overall impression: **very promising**. The server made it possible to inspect, edit, and verify a live PowerPoint deck in a fairly rich way. The read/edit/render/show split is strong, and the anchor model is powerful. The rough edges we hit were mostly around **text formatting semantics**, **placeholder behavior**, and a bit of **deck/render targeting consistency**.

---

## What worked well

### 1. The read model is excellent
The separation between:

- `status`
- `slides`
- `outline`
- `slide`
- `anchor`
- `selection`

is very usable. In practice:
- `outline` was great for quickly understanding the deck
- `slide` was useful for finding target anchors
- `anchor` was helpful for inspecting text content at a specific location

This is a strong design.

### 2. Anchor-based editing is the right abstraction
Using stable-ish targets like:
- `ph:S:KIND`
- `shape:S:N`
- `cell:S:N:R:C`

worked well conceptually. It feels much more robust than trying to target by vague shape names.

### 3. Batch editing is a good idea
The `ppt_batch` tool is especially useful for slide construction and grouped edits. The single-undo-entry design is smart and makes the agent workflow much nicer.

### 4. Render/read/edit loop is exactly the right capability set
Being able to:
1. inspect
2. edit
3. render
4. inspect again

is the killer feature for an agentic PowerPoint server.

---

# Main issues encountered

## 1. `line_spacing` semantics are dangerous / ambiguous
This was the biggest problem in the session.

### What happened
I attempted to set:

```json
"line_spacing": 24
```

assuming the documented semantics meant **24 points**.

Instead, PowerPoint behaved as though the value was interpreted like a **line multiple** or an incompatible unit, causing text to become absurdly spaced and extend off the slide.

### Why this is bad
This is a serious footgun because:
- the current docs imply one meaning
- PowerPoint’s object model has multiple spacing concepts
- the effect is catastrophic visually
- the bad state is hard to recover from without rewriting content

### Recommendation
Make `line_spacing` unambiguous in the API.

#### Best option
Split into explicit fields like:

- `line_spacing_mode`: `"multiple" | "points" | "at_least"`
- `line_spacing_value`: number

or more PowerPoint-native:

- `line_spacing_multiple`
- `line_spacing_points`

#### Minimum option
Keep `line_spacing`, but document exactly:
- what unit it uses
- how that maps to PowerPoint internals
- examples:
  - single spaced
  - 1.5 spaced
  - double spaced
  - exact 18 pt

#### Strong recommendation
Add guardrails:
- reject huge values unless explicitly forced
- warn on suspicious values like `24` if the mode is “multiple”
- possibly clamp impossible values

---

## 2. Multi-line text writing into placeholders behaves inconsistently
### What happened
When writing content into body placeholders, text that was intended as multiple bullets/paragraphs sometimes ended up as:

- one paragraph with embedded line breaks
instead of
- multiple actual paragraphs / bullet items

This made later formatting unreliable.

### Symptoms
- bullets not behaving like bullets
- formatting operations applying strangely
- visual overflow / collapse
- line spacing and list formatting interacting badly

### Recommendation
When `write` receives `text` with `\n`, define and document exactly what it means:

#### Possible interpretations
1. `\n` = new paragraph
2. `\n` = soft line break
3. heuristic based on target type

For slide authoring, **`"\n"` should almost certainly mean new paragraph**, especially in text placeholders.

#### Even better
Add an explicit paragraph-oriented API:

- `write_paragraphs: ["a", "b", "c"]`
- or `paragraphs: [...]`

so agents don’t have to rely on newline interpretation.

---

## 3. Formatting a placeholder can push it into a bad state
### What happened
Some content placeholders ended up with bizarre formatting states:
- tiny font sizes (e.g. 5 pt)
- giant spacing
- overflow off the slide

Once a placeholder got into that state, trying to “repair” it via formatting alone was unreliable.

### Likely cause
Placeholder text inheritance in PowerPoint is tricky. Direct formatting on placeholders may be interacting with:
- master styles
- layout styles
- existing paragraph runs
- list levels
- legacy formatting already present in the placeholder

### Recommendation
A few ideas:

#### A. Add a “normalize text frame” op
Something like:

- `reset_text_format`
- `clear_direct_formatting`
- `normalize_placeholder_text`

that strips weird local formatting while preserving text.

#### B. Add “replace paragraphs cleanly”
A write mode that:
- deletes all paragraphs
- recreates them from scratch
- resets list state consistently

#### C. Expose more paragraph/run diagnostics in `ppt_read`
For each paragraph:
- bullet on/off
- indent level
- line spacing mode/value
- space before/after
- run font sizes if mixed

This would make debugging much easier.

---

## 4. The docs around paragraph formatting need more precision
The current tool surface is nice, but some fields are underspecified in ways that matter a lot.

### Particularly risky fields
- `line_spacing`
- `space_before`
- `space_after`
- `indent_level`
- `list_type`
- `bullet_char`

### Recommendation
For each formatting field, document:
- units
- exact PowerPoint mapping
- valid ranges
- whether it applies to:
  - selection
  - paragraph
  - all paragraphs in anchor
  - mixed runs
- how it behaves if the anchor contains multiple paragraphs

A short table in the docs would help a lot.

---

## 5. Render/doc targeting seemed a bit inconsistent
### What happened
At one point, edits succeeded but a preview render failed due to what seemed like deck name/path lookup inconsistency.

### Recommendation
Tighten deck targeting semantics.

Helpful improvements:
- make `status` return a canonical document identifier
- allow all later commands to target that exact identifier
- distinguish:
  - display name
  - path
  - active-doc token / opaque handle

If possible, return and accept a stable `doc_id` rather than only a name.

---

## 6. Better recovery-oriented operations would help agents a lot
When things go visually wrong, the recovery path is currently awkward.

### Useful additions
- `clear_formatting` on text anchors
- `reset_placeholder_to_layout`
- `rewrite_as_bullets` / `set_paragraphs`
- `fit_text` or `autofit_status`
- `text_frame_status` read op showing:
  - margins
  - auto-size / shrink-to-fit behavior
  - overflow risk

These would reduce the chance of “formatting spiral” bugs.

---

# Suggested concrete API improvements

## A. Replace ambiguous text formatting fields
Instead of:

```json
{
  "line_spacing": 24
}
```

prefer something like:

```json
{
  "line_spacing_mode": "multiple",
  "line_spacing_value": 1.2
}
```

or

```json
{
  "line_spacing_mode": "exact_points",
  "line_spacing_value": 18
}
```

---

## B. Add paragraph-structured writing
Example:

```json
{
  "op": "write_paragraphs",
  "anchor_id": "ph:2:body",
  "paragraphs": [
    {"text": "Launch banana responsibly", "list_type": "bulleted"},
    {"text": "Measure impact radius", "list_type": "bulleted"},
    {"text": "Never brief legal first", "list_type": "bulleted"}
  ]
}
```

This would be much safer than newline-based inference.

---

## C. Add a “normalize” or “reset local formatting” op
Example:

```json
{
  "op": "text_reset_format",
  "anchor_id": "ph:3:body",
  "scope": "all_paragraphs"
}
```

Could strip direct formatting and re-apply inherited layout/master defaults.

---

## D. Add richer read diagnostics for text
A verbose text inspection mode would be extremely helpful, especially for debugging.

For example, `ppt_read(op="anchor")` could optionally include:
- paragraph count
- whether each paragraph is bulleted/numbered
- list level
- spacing before/after
- line spacing mode/value
- run formatting summary
- effective font size(s)

---

## E. Safer defaults / validation
The server could proactively reject suspicious inputs.

Examples:
- `line_spacing > 5` if interpreted as multiple
- font size below, say, 8 pt unless forced
- list formatting on a single paragraph with embedded line breaks could trigger a warning

Even just returning warnings would help.

---

# UX / agent-facing documentation suggestions

## Add a “PowerPoint gotchas” section
A short section in the docs explaining:

- placeholders can inherit complex formatting
- `\n` behavior in write ops
- paragraph formatting applies per paragraph, not per run
- line spacing semantics
- when to rewrite text vs reformat existing text

This would save a lot of trial and error.

## Add examples for common safe patterns
For example:

### Safe bullet list creation
1. write paragraphs individually or via explicit paragraph API
2. apply list formatting
3. apply paragraph spacing
4. avoid ambiguous line spacing unless mode is explicit

### Safe placeholder repair
1. read anchor
2. reset local formatting
3. rewrite paragraphs cleanly
4. render slide to verify

---

# Suggested release notes / bug list

## Bugs observed
1. **`line_spacing` semantics/documentation mismatch**
   - Setting `24` produced catastrophic spacing instead of 24 pt exact spacing.

2. **Multi-line writes may collapse into one paragraph with line breaks**
   - Especially problematic in content placeholders.

3. **Formatting placeholders can produce unstable visual states**
   - Tiny font, giant spacing, overflow.

4. **Render/doc targeting inconsistency**
   - Edit succeeded while render lookup failed in at least one case.

---

# Priority recommendations

## High priority
- Fix / clarify `line_spacing`
- Add explicit paragraph-writing support
- Improve text diagnostics
- Add a formatting reset / normalize operation

## Medium priority
- Stabilize doc targeting with canonical IDs
- Add validation/warnings for dangerous formatting values

## Nice to have
- text-frame overflow/autofit introspection
- higher-level bullet/list authoring helpers
- placeholder repair helpers

---

# Bottom line

The server is already **very capable**, and the architecture feels right. The main pain points weren’t about missing core functionality — they were about **PowerPoint text model sharp edges leaking through the abstraction**.

If the next release improves:
- formatting semantics,
- paragraph handling,
- and recovery/debuggability,

it’ll become dramatically more reliable for agentic slide authoring.

If you want, I can turn this into either:
1. a polished **GitHub issue / release-feedback draft**, or  
2. a more technical **API change proposal** with suggested parameter schemas.
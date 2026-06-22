# pptlive — code-review issue log

Findings from the 2026-06-22 fan-out review (7 agents, full source sweep). Use
this as the triage backlog; work items across sessions and tick `Status` as you go.

**Severity:** Critical (none found) · High (correctness bugs with user-visible
impact) · Medium (fragility / inconsistency that can bite) · Low (consistency,
robustness, docs).

**Status legend:** `todo` · `wip` · `done` · `wontfix` · `needs-repro`.

Two High items (P-01, P-02) were **verified against the code** during the review;
the rest are agent findings not yet independently reproduced.

---

## Triage table

| ID | Sev | File:line | Title | Status |
|----|-----|-----------|-------|--------|
| P-01 | High | `_findreplace.py:136` | Right-boundary sentinel fuses paragraphs on replace | **done** ✅ |
| P-02 | High | `_batch.py:1512` | `run_batch` per-op catch misses `ValueError`/`FileNotFoundError` | **done** ✅ |
| P-03 | High | `_anchors.py:265` | `set_paragraphs` on a `Paragraph` anchor silently corrupts/drops | **done** ✅ |
| P-04 | High | `_slides.py:658` | Placeholder-geometry apply isn't atomic; bad KIND leaves half-built slide | **done** ✅ |
| P-05 | High | `_slides.py:163` | `_find_placeholder` unwrapped COM + swallows transient busy | **done** ✅ |
| P-06 | Med | `_batch.py:540` | Inconsistent `ValueError→BatchOpError` wrapping across handlers | todo |
| P-07 | Med | `_shapes.py:835` | `effect_to_dict` unguarded reads break whole `animations()` listing | todo |
| P-08 | Med | `_shapes.py:1696` | `ShapeById`/`reorder` use `ZOrderPosition`, can emit wrong `shape:S:N` | todo |
| P-09 | Med | `_presentation.py:741` | `active`/`list` remap `PowerPointBusyError` to not-found (wrong exit) | todo |
| P-10 | Med | `_charts.py:67` | Non-numeric chart value raises unhelpful bare `float()` error | todo |
| P-11 | Med | `_anchors.py:337` | `format_text` double-parses color; validation lives in two places | todo |
| P-12 | Med | `_shapes.py:1529` | `set_hyperlink` re-resolves shape COM twice (TOCTOU) | todo |
| P-13 | Med | `_presentation.py:690` | `go_to(select=True)` swallows all non-busy select failures | todo |
| P-14 | Med | `_presentation.py:122` | `save_as("pdf")` may not redirect to `export_pdf` (verify) | todo |
| P-15 | Low | exceptions / cli / batch | No shared `classify()`; exit-code ladder duplicated 2× (wordlive-drift) | todo |
| P-16 | Low | `cli/main.py` + `__init__.py` | No `__version__` / `--version` / `--about` (wordlive-drift) | todo |
| P-17 | Low | `_findreplace` / `_presentation.py:663` | `find_replace` drops wordlive's `normalized_equal` re-verify | todo |
| P-18 | Low | `_tables.py:413` | `set_fill` validates color per-cell, not up-front (vs `set_border`) | todo |
| P-19 | Low | `_smartart.py:169` | No runtime cross-check that depth-first walk == `AllNodes.Count` | todo |
| P-20 | Low | `_theme.py:96` | Theme palette read uses `color_hex` not `color_hex_or_none` | todo |
| P-21 | Low | `_snapshot.py:190` | Multi-slide filenames can get wrong extension when `fmt`≠`out` suffix | todo |
| P-22 | Low | `_batch.py:1293` | `deck_snapshot` temp dirs never cleaned up (long-lived MCP leak) | todo |
| P-23 | Low | `_shapes.py:524` | Swallowed `LockAspectRatio` failure silently reintroduces aspect-snap | todo |
| P-24 | Low | `_shapes.py:1742` | Shape name lookup returns first of duplicates, no `AmbiguousMatchError` | todo |
| P-25 | Low | `_selection.py:91` | Selection restore keys on non-unique `Shape.Name` | todo |
| P-26 | Low | `_sections.py:99` | `add()` validates inside `translate_com_errors` (vs siblings) | todo |
| P-27 | Low | `_tables.py:385` | `_resolve_axis` accepts `bool` selector (inconsistent w/ other validators) | todo |
| P-28 | Low | `_charts.py:264,346` | `_reflects_data`/`_attempt_axis` `except` masks busy / non-COM errors | todo |
| P-29 | Low | `_com.py:133` | `assert last is not None` stripped under `python -O` | todo |
| P-30 | Low | `cli/commands.py:3362` | `set-paragraphs --json` option shadows global `--json/--text` | todo |
| P-31 | Low | `cli/commands.py:496` | `snapshot_cmd` hand-rolls `sys.exit(1)` instead of `_run` boundary | todo |
| P-32 | Low | `cli/commands.py:1316` | `shape add --kind chart` with no data passes `None,None` to `add_chart` | todo |
| P-33 | Low | `constants.py:1696` | `theme_color_for` `or`-fallback breaks if a slot were `0`; rebuilds dict | todo |
| P-34 | Low | `_smartart.py:315` / `_tables.py:457` | double color-parse; `set_border` skips `weight` validation | todo |
| P-35 | Low | `mcp/server.py:338` | `shape_type` default `"rectangle"` diverges from all-`None` convention | todo |
| P-36 | Low | `mcp/server.py:188` | Image-return relies on `-> Any` passthrough w/o `structured_output=False` note | todo |

---

## High

### P-01 — `_findreplace.py:136` — Right-boundary sentinel fuses paragraphs on replace  *(verified)*
**Category:** bug / wordlive-drift
The end-of-segment sentinel is `out_offsets.append(len(s))`. wordlive uses
`out_offsets.append(out_offsets[-1] + 1 if out_offsets else len(s))`. Because pptlive
folds the paragraph mark `\r`→`\n` and soft break `\v`→space and then strips trailing
whitespace, a match that ends just before a stripped/collapsed trailing char gets a
span that runs to `len(s)` — i.e. it **includes the trailing `\r`**. `find_replace`
then runs `tr.Characters(start+1, end-start).Text = replace`, overwriting the
paragraph mark and **fusing the paragraph into the next one** — exactly the footgun
wordlive's sentinel fix prevents. Verified: e.g. haystack `"Hello\r"`, needle `"Hello"`
→ normalized `"Hello"`, offsets `[0,1,2,3,4]`, sentinel `6` → match span
`original[0:6]` = `"Hello\r"`.
**Fix:** Port wordlive's sentinel: `out_offsets.append(out_offsets[-1] + 1 if out_offsets else len(s))`.

### P-02 — `_batch.py:1512` — `run_batch` per-op catch misses `ValueError`/`FileNotFoundError`  *(verified)*
**Category:** error-handling
The per-command `try` catches only `PptliveError`. Several handlers call straight into
the library, which raises **bare** `ValueError` (`format_paragraph` line_spacing>5 /
indent range; `set_paragraphs` empty list) and `FileNotFoundError` (`shape_add`
kind=picture, `shape_set_picture`). Only `_edit_slide_add` (:540) and `_render_save_as`
(:1345) pre-wrap into `BatchOpError`. An unwrapped error escapes the per-op loop, is
never recorded as a `{ok:false,error,message}` entry, defeats `stop_on_error=False`, and
unwinds already-appended results. MCP then mis-reports it as a whole-batch `ToolError`
(contradicting the ppt_batch "failures reported in place" contract); CLI `exec` exits 1
with the batch lost.
**Fix:** Broaden the catch to `except (PptliveError, ValueError, FileNotFoundError)`,
mapping the latter two to `error="invalid_args"` (mirroring `_mcp_errors`).

### P-03 — `_anchors.py:265` — `set_paragraphs` on a `Paragraph` anchor silently corrupts/drops
**Category:** correctness
`set_paragraphs` writes all items as one joined `Text` block onto `self._text_range()`
then formats per-item via `self.paragraphs`. Correct only on a **whole-shape** anchor.
Inherited on `Paragraph` (defined on base `Anchor`): on a `Paragraph` the joined
multi-paragraph block is written into one paragraph's range (corrupting/splitting), and
`getattr(self, "paragraphs", None)` is `None` so per-item formatting is silently dropped
and the returned id list is empty — looks like success.
**Fix:** Gate `set_paragraphs` to whole-shape anchors (raise a clear error on
`Paragraph`), or at least surface the silent format-drop instead of returning `[]`.

### P-04 — `_slides.py:658` — Placeholder-geometry apply isn't atomic; bad KIND leaves half-built slide
**Category:** bug
`add()` creates+inserts the slide, then `_apply_placeholder_geometry` resolves each
`ph:S:KIND` and moves it. A typo'd/ambiguous KIND raises `AmbiguousMatchError`/
`AnchorNotFoundError` **post-creation, mid-loop** → the new slide exists with some
placeholders moved and some not. The pre-COM `_validate_placeholders_arg` establishes
"fail before mutating" intent but doesn't complete it for KIND resolution.
**Fix:** Resolve (and disambiguate) every requested KIND first; apply geometry in a
second pass so a bad KIND fails without leaving a partially-positioned slide.

### P-05 — `_slides.py:163` — `_find_placeholder` unwrapped COM + swallows transient busy
**Category:** error-handling
`_find_placeholder` reads `Shapes`, `PlaceholderFormat.Type`, `Name`, `Id` directly with
no local `translate_com_errors`, and its per-shape `except Exception: continue` (line
~185) swallows a transient `PowerPointBusyError`, silently dropping a placeholder that
momentarily failed to read — unlike `title`/`layout_name`/`has_notes` which re-raise busy.
**Fix:** Wrap the iteration in `translate_com_errors`; narrow the `except` to re-raise
`PowerPointBusyError`.

---

## Medium

### P-06 — `_batch.py:540` — Inconsistent `ValueError→BatchOpError` wrapping across handlers
**Category:** inconsistency
Only `_edit_slide_add` and `_render_save_as` wrap library `ValueError`/`FileExistsError`.
Peers (`_edit_format`, `_edit_set_paragraphs`, `_edit_shape_add` picture,
`_edit_shape_set_picture`, `_edit_shape_picture_fill`) don't — the root of P-02.
**Fix:** Handle `(ValueError, FileNotFoundError)` centrally in the `_<tool>_core`
dispatchers / `run_batch` so the taxonomy is uniform.

### P-07 — `_shapes.py:835` — `effect_to_dict` unguarded reads break whole `animations()` listing
**Category:** error-handling
`effect_to_dict` reads `Effect.Timing`/`Duration`/`TriggerDelayTime`/`Shape.Id`/
`EffectType`/`Exit`/`TriggerType` with no `_safe` guard, unlike every other read-to-dict
helper in the file. One unreadable property fails the entire `slide.animations()` read
(and `Slide.read()`) instead of degrading.
**Fix:** Wrap the individual reads in `_safe(...)` with sensible defaults.

### P-08 — `_shapes.py:1696` — `ShapeById`/`reorder` use `ZOrderPosition`, can emit wrong `shape:S:N`
**Category:** inconsistency
`ShapeById.index`/`to_dict` and `reorder` return `int(sh.ZOrderPosition)`, but addressing
elsewhere is by `Shapes(idx)` collection index. These usually match but aren't guaranteed
equal (groups, certain placeholder orderings). If they diverge, the emitted `shape:S:N`
resolves to a *different* shape than the `shapeid` it was read from. `find_shape_by_id`
already computes the true collection index — use it.
**Fix:** Use the `idx` from `find_shape_by_id` for the z-order index; standardize one
notion of "1-based position" across `reorder`/`ShapeById.index`/`to_dict`.

### P-09 — `_presentation.py:741` — `active`/`list` remap `PowerPointBusyError` to not-found
**Category:** error-handling
`active` turns *any* exception into `PresentationNotFoundError` (exit 2); `list()` maps a
busy `ActivePresentation` read to `active_name=None`. Both lose the busy distinction
(exit 3) the rest of the codebase carefully preserves.
**Fix:** Re-raise `PowerPointBusyError` before the generic mapping.

### P-10 — `_charts.py:67` — Non-numeric chart value raises unhelpful bare `float()` error
**Category:** error-handling
`_normalize_series` eagerly `float(v)`s every value; a non-numeric value raises bare
`ValueError("could not convert string to float: 'abc'")` rather than the friendly,
series-named diagnostics the method otherwise emits (the docstring promises clean
pre-COM errors).
**Fix:** Coerce per-value with a message naming the offending series/value.

### P-11 — `_anchors.py:337` — `format_text` double-parses color; validation in two places
**Category:** fragile-pattern
`format_text` calls `parse_color(color)` to validate then passes raw `color` to
`apply_font`, which parses again. The master-text-style path shares `apply_font` but may
not pre-validate — validation lives inconsistently.
**Fix:** Parse once; pass the parsed RGB int into `apply_font` (single validation point).

### P-12 — `_shapes.py:1529` — `set_hyperlink` re-resolves shape COM twice (TOCTOU)
**Category:** fragile-pattern
`self._com_shape()` is called once to mutate and again inside
`_hyperlink_to_dict(self._com_shape())` for the readback. For `ShapeById`/
`PlaceholderShape` that's two live scans; a shift between them could describe a different
shape than was mutated.
**Fix:** Capture `sh = self._com_shape()` once and reuse for both.

### P-13 — `_presentation.py:690` — `go_to(select=True)` swallows all non-busy select failures
**Category:** error-handling
`shape.com.Select()` re-raises busy but `except Exception: pass` hides every other genuine
select failure with no signal, even though `select=True` was explicitly requested.
**Fix:** Narrow the swallow or log; ideally only swallow the specific "selection not
possible" HRESULT.

### P-14 — `_presentation.py:122` — `save_as("pdf")` may not redirect to `export_pdf` (verify)
**Category:** error-handling
pptlive delegates pdf rejection to `save_format_for` instead of wordlive's explicit
redirect message. Confirm `save_format_for("pdf")` raises a message that names
`export_pdf`; if not, add an explicit pre-check mirroring wordlive.
**Fix:** Verify the message; add redirect if missing.

---

## Low

### P-15 — exceptions/cli/batch — No shared `classify()`; exit-code ladder duplicated 2×
**Category:** wordlive-drift
wordlive centralizes failure labelling in `exceptions.classify()`. pptlive duplicates the
isinstance ladder in `cli/main.py:_exit_for` and `_batch.py:_error_code` (kept in lockstep
by hand, comments admit it), and drops wordlive's `retryable` signal.
**Fix:** Port `classify(exc) -> (code, retryable)` into `exceptions.py`; have both
front-ends consume it.

### P-16 — `cli/main.py` + `__init__.py` — No `__version__` / `--version` / `--about`
**Category:** wordlive-drift
wordlive's CLI provides `--version`/`--about` and exports `__version__`; pptlive has
neither (no `__version__` anywhere in `src/pptlive`). `pptlive --version` is a real
contract gap for an agent-driven tool.
**Fix:** Add `__version__` (via `importlib.metadata.version`) and `@click.version_option`.

### P-17 — `_findreplace` / `_presentation.py:663` — `find_replace` drops `normalized_equal` re-verify
**Category:** fragile-pattern / wordlive-drift
wordlive re-reads the located range and confirms it still normalizes-equal before
overwriting; pptlive overwrites the COM range unconditionally. Compounds P-01.
**Fix:** Re-read `tr.Characters(start+1, end-start).Text`, confirm normalized-equal to the
captured `text` before assigning; skip/raise otherwise.

### P-18 — `_tables.py:413` — `set_fill` validates color per-cell, not up-front
**Category:** inconsistency
Unlike `set_border` (validates `parse_color` up front), `set_fill` defers to
`apply_shape_fill` inside the `for r/for c` loop — re-parsed per cell. Correct (no partial
mutation) but inconsistent with its sibling and the docstring's "validated before any COM"
wording.
**Fix:** Validate color once before the loop, mirroring `set_border`.

### P-19 — `_smartart.py:169` — No cross-check that depth-first walk == `AllNodes.Count`
**Category:** correctness
`read()`'s `node_index` counter assumes `AllNodes` enumerates in exactly depth-first
order (the load-bearing spike assumption). No runtime check; a layout that violates it
(assistant/hidden nodes) would make `format_node` address the wrong node silently.
**Fix:** Assert/log if the recursive walk count != `AllNodes.Count`.

### P-20 — `_theme.py:96` — Theme palette read uses `color_hex` not `color_hex_or_none`
**Category:** inconsistency
Every other color readback uses the sentinel-guarded `color_hex_or_none` (→ `None`, never
a wrong `#000000`). The palette read could surface the `0x80000000` automatic sentinel as
`#000000`.
**Fix:** Use `color_hex_or_none` for palette slots.

### P-21 — `_snapshot.py:190` — Multi-slide filenames can get wrong extension
**Category:** correctness
`build_snapshots` reuses the caller's `out` suffix; `out="deck.png"` with `fmt="jpg"`
writes JPEG bytes into `deck-s1.png`. Single-slide path has the same mismatch.
**Fix:** Derive the suffix from `fmt`, or validate `out` suffix matches `fmt`.

### P-22 — `_batch.py:1293` — `deck_snapshot` temp dirs never cleaned up
**Category:** fragile-pattern
`tempfile.mkdtemp(prefix="pptlive_snap_")` per call, never removed; accumulates in a
long-lived MCP server. Path must outlive the call (bytes read back in `_render_reply`).
**Fix:** Reap old snapshot dirs per call, or clean a per-process temp root on shutdown.

### P-23 — `_shapes.py:524` — Swallowed `LockAspectRatio` failure reintroduces aspect-snap
**Category:** error-handling
In `replace_picture`, `new_com.LockAspectRatio = msoFalse` is wrapped in
`except Exception: pass`; if it silently fails, the subsequent width/height assignments
snap to the new image's ratio — the exact pitfall this code exists to prevent — invisibly.
**Fix:** Narrow the except / log; verify the box stuck after setting it.

### P-24 — `_shapes.py:1742` — Shape name lookup returns first of duplicates
**Category:** inconsistency
`ShapeCollection.__getitem__` by name returns the first match; PowerPoint allows duplicate
names. Contradicts the stated `.Name` uniqueness assumption and the `AmbiguousMatchError`
convention used for placeholders.
**Fix:** Document first-match, or raise `AmbiguousMatchError` listing `shape:S:N` candidates.

### P-25 — `_selection.py:91` — Selection restore keys on non-unique `Shape.Name`
**Category:** fragile-pattern
`snapshot`/`restore` round-trips selection via `Shape.Name` + `Shapes.Range(names)`; names
aren't unique, so colliding names can restore the wrong shape (or silently drop the
selection). COM's `Range` doesn't take ids, so it's a known limitation — document it.
**Fix:** Document the duplicate-name caveat; optionally capture `Id`s and re-resolve when
names collide.

### P-26 — `_sections.py:99` — `add()` validates inside `translate_com_errors`
**Category:** fragile-pattern
The `before_slide` type/range checks run inside the `with translate_com_errors` block,
diverging from `rename`/`delete`/`move` which validate before the `with`.
**Fix:** Move validation above the `with`.

### P-27 — `_tables.py:385` — `_resolve_axis` accepts `bool` selector
**Category:** fragile-pattern
`rows=True`→index 1 silently; `parse_color`/`dash_style_for`/`border_edges_for` all reject
`bool`. Inconsistent.
**Fix:** Reject `bool` selectors explicitly.

### P-28 — `_charts.py:264,346` — `_reflects_data`/`_attempt_axis` `except` masks errors
**Category:** fragile-pattern
`_attempt_axis` catches all `PptliveError` (a busy → mis-reported as "axis absent");
`_reflects_data`'s `except PptliveError` doesn't cover a non-COM `TypeError` from
`list(XValues)` on a degenerate chart (escapes the retry loop).
**Fix:** Re-raise `PowerPointBusyError` in `_attempt_axis`; broaden `_reflects_data`'s
except to `(PptliveError, TypeError, ValueError)`.

### P-29 — `_com.py:133` — `assert last is not None` stripped under `python -O`
**Category:** fragile-pattern
`retry_on_busy` relies on `assert` before `raise last`; under `-O` with `attempts<=0`,
`raise None` → `TypeError`. Unreachable with the default `attempts=4` but latent.
**Fix:** Replace the assert with an explicit guard / validate `attempts >= 1`.

### P-30 — `cli/commands.py:3362` — `set-paragraphs --json` shadows global `--json/--text`
**Category:** fragile-pattern
The subcommand's value-taking `--json` (dest `paragraphs_json`) reuses the spelling of the
global format flag — a documented footgun; no other command reuses a global flag name.
**Fix:** Rename to `--paragraphs`/`--data` (keep dest).

### P-31 — `cli/commands.py:496` — `snapshot_cmd` hand-rolls `sys.exit(1)`
**Category:** inconsistency
The lone deck command that catches `ValueError` locally + `sys.exit(1)` instead of letting
`_run` classify it; error-message format will diverge from siblings over time.
**Fix:** Drop the local try/except; rely on `_run`.

### P-32 — `cli/commands.py:1316` — `shape add --kind chart` with no data passes `None,None`
**Category:** error-handling
The guard only enforces both-or-neither for `--categories`/`--series`; supplying neither
passes `None,None` into `add_chart`, yielding a library/COM error rather than a crisp
UsageError like the picture/table kinds. Confirm `add_chart(None,None)` is intended.
**Fix:** Document the empty-chart default, or add a UsageError for symmetry.

### P-33 — `constants.py:1696` — `theme_color_for` `or`-fallback + rebuilds dict
**Category:** fragile-pattern
`_THEME_COLORS.get(key) or {...}.get(key)` treats a falsy hit as miss (safe only because
indices are 1-12, never 0) and rebuilds the underscore-stripped dict every call.
**Fix:** Use explicit `is None` membership; precompute the alias dict at module load.

### P-34 — `_smartart.py:315` / `_tables.py:457` — double color-parse; `set_border` skips `weight` validation
**Category:** inconsistency / error-handling
`format_node` validates color then re-parses in `_apply_node_font`. `set_border` validates
color/dash/edges but passes `weight` through to COM unchecked (`weight=-5` reaches COM).
**Fix:** Pass the parsed int through / add a `weight >= 0` precheck (or document pass-through).

### P-35 — `mcp/server.py:338` — `shape_type` default `"rectangle"` diverges from all-`None`
**Category:** inconsistency
The lone shape/geometry param with a non-`None` default; the handler already coalesces
`p.get("shape_type") or "rectangle"`, so the schema default is redundant and only applies
to `kind="shape"`.
**Fix:** Default `shape_type: str | None = None` for parity.

### P-36 — `mcp/server.py:188` — Image-return relies on `-> Any` passthrough w/o explicit guard
**Category:** wordlive-drift
wordlive guards its image tool with `@mcp.tool(structured_output=False)` + a comment;
pptlive relies on `-> Any` + `CallToolResult` passthrough (correct per the memory note) but
the safety is implicit and would silently regress to double-encoding if the return
annotation were narrowed to a concrete type.
**Fix:** Add a load-bearing comment at the registration / on `ppt_render`/`ppt_batch`, and
pin it with a test asserting the image rides in `content` exactly once.

---

## Notes / non-issues confirmed during review

- Theme-sentinel color guard (`color_hex_or_none`) is applied consistently across shape
  fill/line/background/font/gradient/effects — **no** path returns a wrong `#000000`
  (the one exception is the theme *palette* read, P-20).
- `replace_picture` z-order restore (send-to-back then `BringForward` `z-1` times) is
  **correct**; only the swallowed `LockAspectRatio` failure (P-23) is a weakness.
- line_spacing / space_before / space_after mode handling + the documented `ValueError`
  guards (both-forms-of-a-pair, multiple > 5 without `force`, indent 1-5) are correct.
- Table row/col intersection, 1-based bounds, `AnchorNotFoundError` kinds, refuse-last-
  row/column guards, and border-edge index mapping are all correct.
- `_batch.py` enum/handler registries are sound (import-time drift guard asserts every op
  has a handler; no missing/duplicate/mis-wired handlers).
- CLI `__all__` matches imported public names; exit-code mapping centralized in
  `_exit_for`/`_BATCH_EXIT_FOR`; the `--json/--text` emit contract is honored uniformly.
- Comment threading recursion guard, headers/footers visible-guard + auto-show, snapshot
  `max_dim` vs `width`/`height` mutual exclusion, and sections 1-based math are all correct.

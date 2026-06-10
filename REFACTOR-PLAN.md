# pptlive — deferred refactor plan

Structural cleanups surfaced by the 2026-06-09 comprehensive review.

## Shipped (branch `chore/py311-and-dispatch-refactors`)

The two large refactors this file originally tracked are **done**, behind a
Python-floor bump that unblocked the cleaner typing:

1. **Python floor → 3.11+.** Dropped 3.10 (library + `mcpb` bundle now
   `requires-python >= 3.11`; `ruff`/`mypy` target `py311`). Coordinated with a
   matching bump in sibling `wordlive`, so parity holds. *(commit: "drop Python
   3.10, target 3.11–3.15")*
2. **CLI `_deck_command` decorator.** The `def go(): with attach() as ppt: deck =
   _pick_deck(...); …; _run(ctx, go)` envelope (repeated ~58×) is now one
   scaffold-only decorator; each mutator keeps its own `with deck.edit(label):`
   fence. `commands.py` shed ~260 lines. `status` and the offline `llm-help` /
   `install-*` stay un-decorated by design. *(commit: "collapse per-command
   scaffold into a `_deck_command` decorator")*
3. **MCP op → `StrEnum` + handler registry.** The triplicated op list (the
   `Literal`, the `if op == …` chain, the docstring) became a per-tool `StrEnum`
   (the single source of truth, typing the `op:` param) plus a per-op handler
   registry. An import-time `assert set(*_OPS) == set(*Op)` and a docstring-mention
   test make drift a hard failure. *(commit: "replace triplicated op Literals with
   per-tool Enum + handler registry")*

## Minor backlog — cleared (2026-06-10)

The low-priority items the review flagged but the dispatch refactors didn't touch
have now all been resolved (one of them as a misdiagnosis):

1. **Busy-swallow in defensive reads — FIXED (decision: re-raise busy only).**
   `_com.safe_read` now lets a genuine `PowerPointBusyError` propagate (the
   taxonomy's retryable exit 3) while still degrading every *other* per-property
   failure to its default — so a busy app surfaces honestly instead of masquerading
   as a missing field. Scoped deliberately: the broad `except Exception` in
   `_comments._identity_from_comment` stays (a failed identity lift *should* fall
   back to the legacy `Add`, so swallowing is correct there). Covered by
   `test_safe_read_propagates_busy` / `test_safe_read_degrades_a_failing_property`.
2. **Z-order-by-Id scan deduplicated — DONE.** The "scan `Shapes` for a matching
   `.Id`" loop is now one helper, `_shapes.find_shape_by_id(slide_com, id) ->
   (index, shape) | None`, shared by `ShapeById._com_shape` (raises if `None`) and
   `_selection._zorder_index` (returns the index; lazy-imports to avoid a cycle).
3. **`--doc` selector now matches `Name` *or* full path — DONE.**
   `PresentationCollection.__getitem__` matches the display `Name` first (common
   case), then falls back to a `FullName` pass, so two same-named decks in different
   folders are disambiguable by path. Covered by
   `test_doc_selector_matches_by_name_then_full_path`.
4. **Paragraph-break-set "inconsistency" — NOT A BUG (no change).** On inspection
   this was a misdiagnosis. `para:S:N:P` resolves through COM
   `TextRange.Paragraphs(P, 1)`, which breaks on `\r` **only** — `\v` is an explicit
   *soft* break that stays within a paragraph (`_anchors.SOFT_BREAK`). So
   `read_selection` counting only `\r` is *correct* for matching `para:` addressing;
   "aligning" it to `\r\v\n` would have made it over-count soft breaks. `_slides.
   _paragraphs` splits on all three **by design** — it's a tidy-display helper, not
   an addressing one, so the two diverging is intended. *Residual open question (a
   live smoke spike, not a code edit):* whether `\n` ever appears in a real
   `TextRange.Text` read and, if so, whether `.Paragraphs()` treats it as a break —
   if it does, the `\r`-only count would undercount and should add `\n` (but never
   `\v`). Left for the next smoke session.
5. **`_charts.recolor_text` non-atomic — FIXED.** Its core (chart-area + legend /
   title / data-label sets) now runs under `_com.retry_on_busy`, the same idempotent
   retry `set_data` uses, so a transient busy mid-recolor retries the whole block
   instead of leaving a half-recolored chart. Axes stay best-effort (already
   probe-and-skip) outside the fence.
6. **`Snapshot.png` misnomer — FIXED (renamed to `.image`).** The field holds the
   chosen `fmt`'s bytes (JPEG when `fmt="jpg"`), so `png` was misleading; renamed to
   `Snapshot.image` across the dataclass, the CLI consumer, tests, and the docs /
   SKILL guides. (MCP only reads `.slide`/`.path`, so it was untouched.)

---

*Source: the 2026-06-09 comprehensive review (5 inspection agents over core
modules, feature modules, CLI/MCP, docs, tests). The bug fixes, small cleanups,
docs, and the two structural refactors landed first; the minor backlog above was
then cleared on 2026-06-10. Nothing from this review remains open except the one
`\n`-in-live-reads smoke spike noted in item 4. New product direction (the
gpt-5.4 LLM-feedback round) lives in `roadmap.md`, not here.*

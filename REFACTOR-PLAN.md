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

---

## Minor backlog (still open)

Low-priority items the review flagged but the dispatch refactors didn't touch.
Fold into a future cleanup session if cheap, otherwise leave noted:

- **Busy-swallow in defensive reads.** `_com.safe_read` (and the `except
  Exception` in `_comments._identity_from_comment`) swallow *every* exception,
  including a real `PowerPointBusyError` that the taxonomy says should propagate as
  exit 3. Letting `PptliveError` through from `safe_read` is more correct but is a
  behavior change across many read paths — make it a deliberate decision, and note
  wordlive's helpers swallow broadly too.
- **Z-order-by-Id scan duplicated** in `ShapeById._com_shape` (`_shapes.py`) and
  `_selection._zorder_index` — same `for sh in slide.Shapes: if int(sh.Id) == …`
  loop, two places. Share one helper.
- **`PresentationCollection` matches decks by `Name`** (`_presentation.py`), which
  is non-unique across folders; `FullName` would disambiguate the `--doc` global.
  Real but rare correctness edge.
- **`_selection.read_selection` paragraph index counts only `\r`** while
  `_slides._paragraphs` splits on `\r\v\n` — a caret after a `\v`/`\n` can compute a
  paragraph number inconsistent with `para:S:N:P` addressing. Align the break set.
- **`_charts.recolor_text` is non-atomic** — many discrete COM sets in one
  `translate_com_errors` block with no `retry_on_busy`; a busy mid-recolor leaves a
  partial. Wrap in `retry_on_busy` like `set_data`, or document best-effort.
- **`_snapshot.Snapshot.png` field is named `png` but holds whatever `fmt` produced**
  (jpg bytes when `fmt="jpg"`). Rename to `image`/`data` or constrain `fmt` at that
  layer.

---

*Source: the 2026-06-09 comprehensive review (5 inspection agents over core
modules, feature modules, CLI/MCP, docs, tests). The bug fixes, small cleanups,
docs, and the two structural refactors above have all landed; this file now tracks
only the minor backlog.*

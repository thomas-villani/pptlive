# pptlive — deferred refactor plan

Structural cleanups surfaced by the 2026-06-09 comprehensive review and
deliberately deferred (the review shipped the bug fixes + small cleanups + docs;
these two refactors touch ~4000 lines and were held for a focused session).

Both are **behavior-preserving** — the fake-COM unit suite (`uv run pytest`,
currently 602 passing) is the safety net. Land each as its own commit and keep
`ruff` / `mypy` / `pytest` green at every step.

---

## 1. CLI: a `_with_deck` helper to collapse the per-command scaffold

**Problem.** Every command in `src/pptlive/cli/commands.py` repeats the same
scaffold. As of this review: **59** `def go():` closures, **59**
`with attach() as ppt:`, **58** `_pick_deck(ppt, ctx.obj["doc_name"])`, and **34**
`with deck.edit(label):` fences. This boilerplate is the dominant reason
`commands.py` is ~2840 lines, and it's 58 chances for the deck-selection / edit /
error-boundary contract to drift between commands.

The shape repeated everywhere (e.g. `slides_cmd`, `commands.py:641`):

```python
def slides_cmd(ctx: click.Context) -> None:
    def go() -> None:
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            rows = deck.slides.list()
            emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_slides(rows))
    _run(ctx, go)
```

Mutating commands add one more layer — `with deck.edit(label):` around the body.

**Target.** A context-manager helper that owns `attach` + `_pick_deck` + the
optional `edit()` fence + the `_run` error boundary, so a command body is just
"deck in, emit out":

```python
@contextmanager
def _with_deck(ctx, *, edit: str | None = None):
    """Yield the selected deck inside the standard scaffold.

    Opens PowerPoint, picks the --doc deck, optionally fences a one-Ctrl-Z
    edit() block when `edit` is a label, and routes PptliveError through the
    _run boundary (stderr + exit code). Read commands pass edit=None.
    """
    def go():
        with attach() as ppt:
            deck = _pick_deck(ppt, ctx.obj["doc_name"])
            if edit is not None:
                with deck.edit(edit):
                    yield deck            # see note on generator+contextmanager below
            else:
                yield deck
    _run(ctx, go)
```

Note: a `@contextmanager` that also needs `_run` to wrap the `yield` can't be a
naive generator — `_run` must call a function, not re-enter the generator. Two
clean ways to do it:

- **(a) Decorator form (recommended).** `_deck_command(edit=None)` decorates the
  command body `fn(ctx, deck, **kwargs)`; the decorator builds `go()`, runs it
  through `_run`, and passes the resolved `deck` (and, for mutators, an open
  edit scope) into `fn`. Command bodies lose the `def go()` / `with attach()` /
  `_pick_deck` / `_run` lines entirely.
- **(b) Callback form.** `_run_with_deck(ctx, edit, lambda deck: ...)` — same
  effect, less magic, slightly noisier call sites.

Recommended target for a read command:

```python
@click.command(name="slides")
@click.pass_context
@_deck_command()                                 # read: no edit fence
def slides_cmd(ctx, deck) -> None:
    rows = deck.slides.list()
    emit(rows, as_text=not ctx.obj["as_json"], text=_fmt_slides(rows))
```

…and a mutator:

```python
@_deck_command(edit="CLI: add {kind} on slide {slide_index}")   # label may need lazy formatting
def shape_add(ctx, deck, *, slide_index, kind, ...) -> None:
    ...
```

**Watch-outs (per-command quirks the helper must preserve):**
- **Dynamic edit labels.** Many fences use an f-string built from the command's
  args (`f"CLI: replace {anchor_id}"`). The decorator can't see those at decoration
  time — accept `edit` as either a static string or a `Callable[..., str]` over the
  bound kwargs, or open the fence inside the body for those commands.
- **Commands that emit on a caught error then re-raise.** `status` (commands.py:~620)
  catches, `emit`s, then re-raises so `_run` still sets the exit code. Don't let the
  helper swallow that pattern — keep the body able to `raise`.
- **Commands that don't pick a deck** (`status` itself reads all decks; `llm-help` /
  `install-skill` / `install-mcp` never `attach`). Leave these un-decorated.
- **`shape add` text fix (this review).** The autoshape `set_text` lives *inside* the
  edit fence — make sure the migration keeps it there.
- **The `--doc` global** flows through `ctx.obj["doc_name"]`; the helper must read it
  the same way (don't capture it at decoration time).

**Payoff.** Command bodies drop ~4 lines each across ~58 commands (~230 lines),
and the deck-selection + edit-fence + error-boundary contract lives in exactly
one place. Expect `commands.py` to shrink by 8–10%.

**Verification.** `uv run pytest tests/test_cli.py` after each batch; the CLI tests
already cover read/mutate/exit-code paths for most commands. Migrate in groups
(reads first — lowest risk — then mutators), running the suite between groups.

---

## 2. MCP: an op→handler registry to kill the triplicated op list

**Problem.** In `src/pptlive/mcp/server.py` each op is named in **three** places
that must be kept in sync by hand:

1. The `op: Literal[...]` in the tool signature (`ppt_read` ~`:686`, `ppt_edit`
   ~`:752`, `ppt_render` ~`:988`/`:1073` — note `ppt_render`'s Literal now also
   lists `deck_snapshot`/`deck_pdf`/`save`/`save_as` after this review).
2. The giant `if op == "...": ... return` chain in the matching `_*_core`
   function (`_read_core:217`, `_edit_core:269` — a ~290-line chain,
   `_render_core:602`, `_show_core:654`).
3. The prose op list in each tool's docstring (the agent-facing guide).

Adding or renaming an op means editing all three, and they silently drift (the
docs review already found the docstring/table lagging the dispatch chain).

**Target.** A handler registry per tool — each op is one self-registering
function, and the `Literal` is derived from the registry keys:

```python
# read_ops: dict[str, Callable[[Any, dict], dict]]
READ_OPS: dict[str, ReadHandler] = {}

def read_op(name):                     # decorator
    def reg(fn): READ_OPS[name] = fn; return fn
    return reg

@read_op("slides")
def _read_slides(ppt, p):
    return {"slides": _pick_deck(ppt, p.get("doc")).slides.list()}

@read_op("find")
def _read_find(ppt, p):
    _require(p.get("text") is not None, "read op='find' requires `text`")
    deck = _pick_deck(ppt, p.get("doc"))
    matches = deck.find(p["text"], scope=p.get("scope"))
    return {"count": len(matches), "matches": matches}

def _read_core(ppt, op, p):
    handler = READ_OPS.get(op)
    if handler is None:
        raise ToolError(f"invalid_args: unknown read op {op!r}")
    return handler(ppt, p)
```

Then the tool's `Literal` is generated once from the keys so it can't drift:

```python
ReadOp = Literal[tuple(READ_OPS)]      # or build the Literal/enum at import time
```

(If a literal-from-runtime-dict fights the type checker, keep a single
`_READ_OP_NAMES: tuple[str, ...]` constant that both the `Literal` and a startup
assertion `set(READ_OPS) == set(_READ_OP_NAMES)` reference — the assertion turns
silent drift into an import-time failure.)

**Watch-outs:**
- **`status` reads all decks** and must run *before* `_pick_deck` (it doesn't pick
  one). Keep that as a handler that simply doesn't call `_pick_deck`.
- **`_edit_core` takes a `Presentation`, not the app** (`deck`, already picked +
  fenced by the caller). Keep the read vs edit handler signatures distinct
  (`(ppt, p)` for read/render, `(deck, p)` for edit/show) — don't force one shape.
- **`ppt_batch`** dispatches the same ops; point it at the same registries so batch
  and the dedicated tools can't diverge (today they share `_*_core`, so this keeps
  that property).
- **`ppt_render` embedding.** Some render ops return an image reply
  (`_render_reply`) and some a plain dict (`deck_pdf`/`save`/`save_as`); keep that
  decision in `ppt_render` around the handler return, not inside each handler.
- **Docstrings.** The per-op prose still lives in the tool docstring (agents read
  it). Consider a short `help=` string on each handler and assembling the docstring
  from them, or at least add an import-time check that every registered op is
  mentioned in the docstring so #3 stops drifting.

**Payoff.** One op = one function = one registry key. The `Literal` is derived, the
dispatch is a dict lookup, and a startup assertion makes any missing piece a hard
error instead of silent drift. `_edit_core`'s ~290-line chain becomes ~25 small
handlers.

**Verification.** `uv run pytest tests/test_mcp.py` covers every op across
read/edit/render/show/batch; run it after each tool's migration. Migrate one tool
at a time (`ppt_read` → `ppt_show` → `ppt_render` → `ppt_edit`, easiest to hardest).

---

## Minor backlog (small items also deferred by the review)

Low-priority; fold into the above session if cheap, otherwise leave noted:

- **Busy-swallow in defensive reads.** `_com.safe_read` (and the `except Exception`
  in `_comments._identity_from_comment`) swallow *every* exception, including a real
  `PowerPointBusyError` that the taxonomy says should propagate as exit 3. Letting
  `PptliveError` through from `safe_read` is more correct but is a behavior change
  across many read paths — make it a deliberate decision, and note wordlive's
  helpers swallow broadly too. (Flagged HIGH-ish by the feature-module review.)
- **Z-order-by-Id scan duplicated** in `ShapeById._com_shape` (`_shapes.py`) and
  `_selection._zorder_index` — same `for sh in slide.Shapes: if int(sh.Id) == …`
  loop, two places. Share one helper.
- **`PresentationCollection` matches decks by `Name`** (`_presentation.py:~730`),
  which is non-unique across folders; `FullName` would disambiguate the `--doc`
  global. Real but rare correctness edge.
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
and doc updates from that review already landed; this file is the deferred
remainder.*

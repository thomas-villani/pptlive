"""Linter + formatting regularizer — `deck.lint()` / `deck.regularize()`.

Audit a deck for presentation-quality defects (`lint`, a pure read), then autofix
the mechanical ones in one atomic-undo step (`regularize`, a write). Pure
composition over shipped verbs — `format`, `shape_move` / `shape_resize`,
`shape_align`, `shape_reset_layout` — so there is **no new COM write surface**
here; the new work is the `format_info()` probe (`_format_info.py`) and this rule
engine, ported from wordlive's `_linting.py`.

The PowerPoint reframing (`spec-linter.md` §2): a deck has ~no named styles, so
"consistency" is not Word's *direct override vs. style*. It is (P2) **mode across
peers** — every title alike, every level-1 bullet alike — which needs no
configuration because the deck is judged against itself; and (P3) **spatial
regularity** over the shipped `geometry_report()`. Each rule is `consistency`,
`structural`, or `policy` (profile-driven). A rule emits `Finding`s; a *fixable*
finding carries an op-shaped `fix` (literally an `exec` op, or a list of them),
so `regularize` is "lint, then run each finding's `fix` through the batch op loop".

**Anchor drift.** A regularize pass applies many fixes in one batch, so findings
anchor by the drift-proof `shapeid:S:ID` (every finding also carries `shapeid`
and `slide`); a paragraph-level finding anchors by `para:S:N:P` — stable across
the pure-formatting fixes in this catalogue, which never delete or restack.

Rule modules (`_linting_consistency`, `_linting_geometry`) are imported at the
bottom for their side effect of registering rules.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Callable, Iterator, Sequence
from dataclasses import asdict, dataclass, replace
from typing import TYPE_CHECKING, Any

from ._lint_profile import Profile
from .exceptions import BatchOpError

if TYPE_CHECKING:
    from ._presentation import Presentation

# A finding's fix is one exec op, or a list of ops applied in order.
FixOps = dict[str, Any] | list[dict[str, Any]]

_SEVERITY_RANK = {"error": 0, "warning": 1, "info": 2}


@dataclass(frozen=True)
class Finding:
    """One linter result. `fix` is present iff `fixable` — an op-shaped dict (or
    list of them) `regularize` runs verbatim through the batch op loop.

    `slide` is the 1-based slide; `anchor_id` the anchor the finding is *about*
    (drift-proof `shapeid:S:ID`, or `para:S:N:P` for one paragraph); `shapeid`
    always names the host shape's stable handle. `adds_content` marks a fix that
    adds or destroys content (withheld by `regularize` unless `allow_content`)."""

    rule: str
    kind: str  # consistency | structural | policy
    severity: str  # error | warning | info
    slide: int | None
    anchor_id: str
    message: str
    shapeid: str | None = None
    fixable: bool = False
    fix: FixOps | None = None
    adds_content: bool = False
    observed: str | None = None
    expected: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class Scope:
    """The `within=` audit scope, by **containment**: `slide` limits findings to
    one slide, `shape_id` (with `slide`) to one shape. `None`/`None` = the deck.
    Peer rules still compute their dominant value deck-wide; scope only limits
    which findings are *emitted*."""

    slide: int | None = None
    shape_id: int | None = None

    def covers(self, slide: int | None, shape_id: int | None = None) -> bool:
        """Whether a finding on `slide` (and, if given, shape `shape_id`) is in
        scope. Called with the slide alone it answers "could anything on this slide
        be in scope?" — the per-slide short-circuit a rule takes before walking."""
        if self.slide is not None and slide != self.slide:
            return False
        if shape_id is None or self.shape_id is None:
            return True
        return shape_id == self.shape_id


@dataclass(frozen=True)
class Rule:
    """A registered rule: identity/metadata plus a `check` that yields findings.

    `check(walk, scope, profile)` reads the deck through the per-pass `DeckWalk`
    cache and yields `Finding`s. Every check receives the resolved `Profile`;
    consistency / structural rules mostly ignore it, `policy` rules read their
    target from it, and opinionated rules read a `tolerance`. `tags` lets a
    caller select a family (`["alignment"]`) as well as by id. `default_on`
    controls the `rules=None` set: unambiguous defects ship on, opinionated rules
    (`edge-alignment`'s tolerance call) ship off until named / tagged / enabled by
    a profile; policy rules are off by `kind` regardless."""

    id: str
    kind: str
    severity: str
    tags: tuple[str, ...]
    check: Callable[[DeckWalk, Scope, Profile], Iterator[Finding]]
    default_on: bool = True


# ---------------------------------------------------------------------------
# the per-pass read cache
# ---------------------------------------------------------------------------


class DeckWalk:
    """Memoised deck reads for the duration of one lint pass.

    Every rule wants the same walks — each slide's `shapes.list()`, each text
    shape's paragraphs, each placeholder's `format_info()` — and a default run
    has several rules. A lint pass is a pure read, so one walk apiece is enough;
    a fresh `DeckWalk` per pass means any edit (a `regularize` fix, a user
    keystroke) invalidates it by construction."""

    def __init__(self, deck: Presentation) -> None:
        self.deck = deck
        self._slides: list[dict[str, Any]] | None = None
        self._geometry: dict[int, dict[str, Any]] = {}
        self._paragraphs: dict[str, list[dict[str, Any]]] = {}
        self._format_info: dict[str, dict[str, Any]] = {}
        self._page_setup: dict[str, float] | None = None

    def slides(self) -> list[dict[str, Any]]:
        """`[{index, id, layout, shapes: [shape dicts]}]` for the whole deck."""
        if self._slides is None:
            rows: list[dict[str, Any]] = []
            for s in self.deck.slides:
                rows.append(
                    {
                        "index": s.index,
                        "id": s.id,
                        "layout": s.layout_name,
                        "shapes": s.shapes.list(),
                    }
                )
            self._slides = rows
        return self._slides

    def page_setup(self) -> dict[str, float]:
        if self._page_setup is None:
            self._page_setup = self.deck.page_setup()
        return self._page_setup

    def geometry(self, slide: int) -> dict[str, Any]:
        if slide not in self._geometry:
            self._geometry[slide] = self.deck.slides[slide].geometry_report()
        return self._geometry[slide]

    def paragraphs(self, shapeid: str) -> list[dict[str, Any]]:
        """`paragraphs.list()` of the shape at `shapeid:S:ID` (rows carry the
        live `para:S:N:P` anchor ids)."""
        if shapeid not in self._paragraphs:
            anchor: Any = self.deck.anchor_by_id(shapeid)
            self._paragraphs[shapeid] = anchor.paragraphs.list()
        return self._paragraphs[shapeid]

    def format_info(self, anchor_id: str) -> dict[str, Any]:
        if anchor_id not in self._format_info:
            self._format_info[anchor_id] = self.deck.anchor_by_id(anchor_id).format_info()
        return self._format_info[anchor_id]


# ---------------------------------------------------------------------------
# the peer-mode primitive (P2)
# ---------------------------------------------------------------------------

#: A peer value must be shared by at least this fraction of peers to count as
#: the group's *dominant* value (spec §5: "a rule fires only when a clear
#: dominant exists").
DEFAULT_DOMINANCE = 0.6


def dominant(
    values: Sequence[Any], *, threshold: float = DEFAULT_DOMINANCE
) -> tuple[Any, int, int] | None:
    """The mode of `values` if it is a clear majority — `(value, count, total)` —
    else `None`. `None` entries (mixed / unreadable peers) are ignored; a group of
    fewer than two readable peers has no dominant value."""
    readable = [v for v in values if v is not None]
    total = len(readable)
    if total < 2:
        return None
    value, count = Counter(readable).most_common(1)[0]
    if count / total < threshold:
        return None
    return value, count, total


def shape_text(shape: dict[str, Any]) -> str:
    return str(shape.get("text") or "")


def is_title_placeholder(shape: dict[str, Any]) -> bool:
    return shape.get("placeholder") in ("title", "ctrtitle")


def is_body_placeholder(shape: dict[str, Any]) -> bool:
    return shape.get("placeholder") in ("body", "object")


# ---------------------------------------------------------------------------
# registry
# ---------------------------------------------------------------------------

_RULES: list[Rule] = []


def _register_rule(rule: Rule) -> None:
    _RULES.append(rule)


def rules_catalog() -> list[dict[str, Any]]:
    """Every registered rule as `{id, kind, severity, tags, default_on}`."""
    return [
        {
            "id": r.id,
            "kind": r.kind,
            "severity": r.severity,
            "tags": list(r.tags),
            "default_on": r.default_on and r.kind != "policy",
        }
        for r in _RULES
    ]


def _select_rules(rules: Any, profile: Profile) -> list[Rule]:
    """Resolve the `rules=` selector (composed with `profile`) into the rule set.

    - `None` — the default set: every on-by-default consistency + structural rule.
    - a list of ids/tags — only rules matching an id or carrying a tag (ignores
      `default_on`, so `["alignment"]` lights up the whole opt-in cluster).
    - `{"exclude": [...]}` — the default set minus the listed ids/tags.

    A profile then composes: rules it enables are unioned in (on the default-based
    paths), rules it disables are removed (on every path). An unknown id/tag in a
    list is a `ValueError`, so a typo can't silently run nothing.
    """

    def default_set() -> list[Rule]:
        return [r for r in _RULES if r.kind in ("consistency", "structural") and r.default_on]

    known = {r.id for r in _RULES} | {t for r in _RULES for t in r.tags}
    if rules is None:
        selected = default_set()
    elif isinstance(rules, dict):
        excluded = set(rules.get("exclude", []))
        _check_known(excluded, known)
        selected = [r for r in default_set() if not (r.id in excluded or excluded & set(r.tags))]
    else:
        wanted = set(rules)
        _check_known(wanted, known)
        selected = [r for r in _RULES if r.id in wanted or wanted & set(r.tags)]

    chosen = {r.id: r for r in selected}
    default_based = rules is None or isinstance(rules, dict)
    for r in _RULES:
        state = profile.is_enabled(r.id)
        if state is True and default_based:
            chosen[r.id] = r
        elif state is False:
            chosen.pop(r.id, None)
    return [r for r in _RULES if r.id in chosen]


def _check_known(names: set[Any], known: set[str]) -> None:
    unknown = sorted(str(n) for n in names if n not in known)
    if unknown:
        raise ValueError(
            f"unknown lint rule id/tag: {', '.join(unknown)}; known: {', '.join(sorted(known))}"
        )


# ---------------------------------------------------------------------------
# public entry points (Presentation.lint / Presentation.regularize delegate here)
# ---------------------------------------------------------------------------


def _resolve_scope(deck: Presentation, within: Any) -> Scope:
    if within is None:
        return Scope()
    from ._anchors import Paragraph  # noqa: PLC0415
    from ._slides import Slide  # noqa: PLC0415

    if isinstance(within, Slide):
        return Scope(slide=within.index)
    if isinstance(within, str):
        kind, _, rest = within.partition(":")
        if kind == "slide":
            try:
                return Scope(slide=deck.slides[int(rest)].index)
            except ValueError as exc:
                raise ValueError(f"within: bad slide anchor {within!r}") from exc
        anchor: Any = deck.anchor_by_id(within)
    else:
        anchor = within
    shape = anchor.shape if isinstance(anchor, Paragraph) else anchor
    shape_id = getattr(shape, "shape_id", None)
    return Scope(slide=anchor.slide.index, shape_id=int(shape_id) if shape_id else None)


def run_lint(
    deck: Presentation, *, rules: Any = None, within: Any = None, profile: Any = None
) -> list[Finding]:
    """Run the selected rules and return findings, ranked by severity then slide.

    Pure read: nothing is mutated, no view/Selection move, `Saved` untouched.
    `profile` (a path / dict / `Profile` / `None`) opts rules in, supplies their
    targets, and can override a rule's severity — resolved once and threaded to
    every rule. See `Presentation.lint`.
    """
    resolved = Profile.load(profile)
    scope = _resolve_scope(deck, within)
    selected = _select_rules(rules, resolved)
    walk = DeckWalk(deck)
    findings: list[Finding] = []
    for rule in selected:
        for finding in rule.check(walk, scope, resolved):
            sev = resolved.severity_for(finding.rule)
            if sev is not None and sev != finding.severity:
                finding = replace(finding, severity=sev)
            findings.append(finding)
    findings.sort(key=lambda f: (_SEVERITY_RANK.get(f.severity, 9), f.slide or 0, f.anchor_id))
    return findings


def _fix_ops(fix: FixOps) -> list[dict[str, Any]]:
    return list(fix) if isinstance(fix, list) else [fix]


def regularize(
    deck: Presentation,
    *,
    rules: Any = None,
    within: Any = None,
    profile: Any = None,
    dry_run: bool = False,
    allow_content: bool = False,
    own_undo: bool = True,
) -> dict[str, Any]:
    """Apply the fixable findings; return `{applied, skipped, deferred, findings}`.

    Runs `run_lint`, then applies every fixable finding's `fix` op(s). With
    `own_undo=True` (the `Presentation.regularize` path) the fixes run through
    `run_batch` — one `deck.edit("Regularize formatting")`, one Ctrl-Z, the
    polite view restore kept (a pure-formatting pass never opts into "follow the
    work"). With `own_undo=False` (the `regularize` *exec op*, already inside a
    batch's `deck.edit`) they apply via `_edit_core` directly, so the surrounding
    batch stays one undo entry. `dry_run=True` plans without writing.

    Fixes flagged `adds_content` are **withheld** unless `allow_content=True`;
    withheld fixes land in the `deferred` bucket so the caller sees what an opt-in
    would apply. A failing fix op raises `BatchOpError` naming the op and the
    finding (earlier fixes in the pass stay applied — one Ctrl-Z reverts them).
    """
    findings = run_lint(deck, rules=rules, within=within, profile=profile)
    fixable = [f for f in findings if f.fixable and f.fix is not None]
    to_apply = [f for f in fixable if allow_content or not f.adds_content]
    deferred = [f for f in fixable if not allow_content and f.adds_content]
    skipped = [f.to_dict() for f in findings if not (f.fixable and f.fix is not None)]
    report: dict[str, Any] = {
        "applied": [],
        "skipped": skipped,
        "deferred": [f.to_dict() for f in deferred],
        "findings": [f.to_dict() for f in findings],
        "dry_run": dry_run,
    }
    if dry_run or not to_apply:
        return report

    ops: list[dict[str, Any]] = []
    owners: list[Finding] = []
    for f in to_apply:
        for op in _fix_ops(f.fix):  # type: ignore[arg-type]
            ops.append(dict(op))
            owners.append(f)

    from ._batch import _edit_core, run_batch  # noqa: PLC0415 — _batch imports us

    if own_undo:
        results = run_batch(
            deck._ppt,
            deck,
            ops,
            atomic=True,
            stop_on_error=True,
            follow_view=False,
            label="Regularize formatting",
        )
        for entry in results:
            if not entry.get("ok", False):
                owner = owners[int(entry["index"])]
                raise BatchOpError(
                    f"regularize: fix for {owner.rule} ({owner.anchor_id}) failed — "
                    f"{entry.get('error')}: {entry.get('message')}"
                )
        report["ops_run"] = len(results)
    else:
        for i, op in enumerate(ops):
            p = {k: v for k, v in op.items() if k != "op"}
            try:
                _edit_core(deck, str(op["op"]), p)
            except Exception as exc:
                owner = owners[i]
                raise BatchOpError(
                    f"regularize: fix for {owner.rule} ({owner.anchor_id}) failed — {exc}"
                ) from exc
        report["ops_run"] = len(ops)
    report["applied"] = [f.to_dict() for f in to_apply]
    return report


# Rule modules register themselves on import; they need `Rule` / `Finding` /
# `_register_rule` defined first, hence the bottom-of-module import.
from . import (  # noqa: E402,F401
    _linting_consistency,
    _linting_geometry,
)

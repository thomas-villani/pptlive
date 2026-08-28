"""P2 — the peer-mode consistency rules (`spec-linter.md` §5 "Titles & text").

The headline cluster, and the user's first ask: *"all the headers are the same
font and size"*. Each rule collects one property across a **peer set**, takes the
mode via `dominant()`, and flags the minority against it. The peer set is keyed by
**placeholder kind + layout name**: a section-header title is legitimately styled
unlike a content-slide title, and a Comparison body is legitimately smaller than a
Title-and-Content body, so peers are "the same placeholder on the same layout"
(the spec's Open Q #1, answered by the layout rather than by section). A deck is
judged against **itself**, so the rules need no configuration and work on any
template; a `house_style` profile can later pin the target instead.

The targeted fix writes the dominant value back as direct formatting (`format` /
`shape_move` / `shape_resize`) — visually correct and **idempotent** (re-running
writes the same value → no-op), the same contract as wordlive §7c.

- `title-font-consistent` (on) — title / ctrtitle font name, size, bold, colour vs
  the dominant peer value. Fix: `format`.
- `mixed-runs-in-title` (on, report-only) — a title whose runs disagree on a font
  field (the `format_info()["mixed"]` tell); which run is the outlier needs a
  run-walk, so it only reports. Such a title is left out of the peer compare.
- `body-font-consistent` (on) — each body/content paragraph's font name + size vs
  the dominant value at its **indent level** on that layout. Fix: `format` on the
  `para:S:N:P` anchor.
- `title-position-consistent` (on) — the title box (`left/top/width/height`) vs
  the dominant title box on that layout (the "jumpy title"). Fix: `shape_move` +
  `shape_resize`. Deviations under 1 pt are ignored.

Imported by `_linting` for its side effect of registering the rules.
"""

from __future__ import annotations

from collections import defaultdict
from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._linting import (
    DEFAULT_DOMINANCE,
    DeckWalk,
    Finding,
    Rule,
    Scope,
    _register_rule,
    dominant,
    is_body_placeholder,
    is_title_placeholder,
    shape_text,
)

if TYPE_CHECKING:
    from ._lint_profile import Profile

#: `format_info()` font field → the `format` op keyword that writes it.
_FONT_FIX_KEY = {"name": "font", "size": "size", "bold": "bold", "color": "color"}
_TITLE_FIELDS = ("name", "size", "bold", "color")
_BODY_FIELDS = ("name", "size")
_POSITION_TOLERANCE = 1.0


def _threshold(profile: Profile, rule_id: str) -> float:
    cfg = profile.config_for(rule_id).get("dominance")
    return float(cfg) if isinstance(cfg, int | float) else DEFAULT_DOMINANCE


def _font_label(values: dict[str, Any], fields: tuple[str, ...]) -> str:
    parts: list[str] = []
    if "name" in fields and values.get("name") is not None:
        parts.append(str(values["name"]))
    if "size" in fields and values.get("size") is not None:
        parts.append(f"{values['size']:g} pt")
    if "bold" in fields and values.get("bold") is True:
        parts.append("bold")
    if "color" in fields and values.get("color") is not None:
        parts.append(str(values["color"]))
    return " ".join(parts) or "(unreadable)"


def _norm(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 2)
    if isinstance(value, str):
        return value.casefold()
    return value


def _title_peers(walk: DeckWalk) -> dict[tuple[Any, Any], list[dict[str, Any]]]:
    """Title placeholders grouped by `(layout, placeholder kind)`, each with its
    `format_info()` font values."""
    groups: dict[tuple[Any, Any], list[dict[str, Any]]] = defaultdict(list)
    for slide in walk.slides():
        for shape in slide["shapes"]:
            if not is_title_placeholder(shape) or not shape_text(shape).strip():
                continue
            info = walk.format_info(shape["shapeid"])
            values = {f: info["font"][f]["value"] for f in _TITLE_FIELDS}
            groups[(slide["layout"], shape["placeholder"])].append(
                {
                    "slide": slide["index"],
                    "shapeid": shape["shapeid"],
                    "id": shape["id"],
                    "name": shape["name"],
                    "values": values,
                    "mixed": [f for f in info["mixed"] if f in _TITLE_FIELDS],
                    "geometry": shape["geometry"],
                }
            )
    return groups


def _check_title_font_consistent(
    walk: DeckWalk, scope: Scope, profile: Profile
) -> Iterator[Finding]:
    threshold = _threshold(profile, "title-font-consistent")
    for (_layout, kind), peers in _title_peers(walk).items():
        keyed = [
            None if p["mixed"] else tuple(_norm(p["values"][f]) for f in _TITLE_FIELDS)
            for p in peers
        ]
        dom = dominant(keyed, threshold=threshold)
        if dom is None:
            continue
        dom_key, count, total = dom
        dom_values = next(p["values"] for p, k in zip(peers, keyed, strict=True) if k == dom_key)
        for peer, key in zip(peers, keyed, strict=True):
            if key is None or key == dom_key or not scope.covers(peer["slide"], peer["id"]):
                continue
            fix: dict[str, Any] = {"op": "format", "anchor_id": peer["shapeid"]}
            diffs = [
                f
                for f in _TITLE_FIELDS
                if _norm(peer["values"][f]) != _norm(dom_values[f]) and dom_values[f] is not None
            ]
            for f in diffs:
                fix[_FONT_FIX_KEY[f]] = dom_values[f]
            yield Finding(
                rule="title-font-consistent",
                kind="consistency",
                severity="warning",
                slide=peer["slide"],
                anchor_id=peer["shapeid"],
                shapeid=peer["shapeid"],
                message=(
                    f"Slide {peer['slide']} {kind} '{_font_label(peer['values'], _TITLE_FIELDS)}' "
                    f"differs from the deck's dominant {kind} "
                    f"'{_font_label(dom_values, _TITLE_FIELDS)}' ({count} of {total} peers)."
                ),
                fixable=bool(diffs),
                fix=fix if diffs else None,
                observed=_font_label(peer["values"], _TITLE_FIELDS),
                expected=_font_label(dom_values, _TITLE_FIELDS),
            )


def _check_mixed_runs_in_title(walk: DeckWalk, scope: Scope, profile: Profile) -> Iterator[Finding]:
    for (_layout, kind), peers in _title_peers(walk).items():
        for peer in peers:
            if not peer["mixed"] or not scope.covers(peer["slide"], peer["id"]):
                continue
            yield Finding(
                rule="mixed-runs-in-title",
                kind="consistency",
                severity="info",
                slide=peer["slide"],
                anchor_id=peer["shapeid"],
                shapeid=peer["shapeid"],
                message=(
                    f"Slide {peer['slide']} {kind} mixes {', '.join(peer['mixed'])} across its "
                    "runs; a title is expected to be uniform (read the paragraphs to find the "
                    "outlier run)."
                ),
                fixable=False,
                observed=f"mixed: {', '.join(peer['mixed'])}",
                expected="one font across the title",
            )


def _check_body_font_consistent(
    walk: DeckWalk, scope: Scope, profile: Profile
) -> Iterator[Finding]:
    threshold = _threshold(profile, "body-font-consistent")
    groups: dict[tuple[Any, int], list[dict[str, Any]]] = defaultdict(list)
    for slide in walk.slides():
        for shape in slide["shapes"]:
            if not is_body_placeholder(shape) or not shape_text(shape).strip():
                continue
            for row in walk.paragraphs(shape["shapeid"]):
                if not str(row.get("text") or "").strip():
                    continue
                font = row["font"]
                values = {"name": font.get("font"), "size": font.get("size")}
                groups[(slide["layout"], int(row.get("indent_level") or 1))].append(
                    {
                        "slide": slide["index"],
                        "shapeid": shape["shapeid"],
                        "id": shape["id"],
                        "anchor_id": row["anchor_id"],
                        "values": values,
                    }
                )
    for (_layout, level), peers in groups.items():
        keyed = [
            None
            if any(p["values"][f] is None for f in _BODY_FIELDS)
            else tuple(_norm(p["values"][f]) for f in _BODY_FIELDS)
            for p in peers
        ]
        dom = dominant(keyed, threshold=threshold)
        if dom is None:
            continue
        dom_key, count, total = dom
        dom_values = next(p["values"] for p, k in zip(peers, keyed, strict=True) if k == dom_key)
        for peer, key in zip(peers, keyed, strict=True):
            if key is None or key == dom_key or not scope.covers(peer["slide"], peer["id"]):
                continue
            fix: dict[str, Any] = {"op": "format", "anchor_id": peer["anchor_id"]}
            for f in _BODY_FIELDS:
                if _norm(peer["values"][f]) != _norm(dom_values[f]):
                    fix[_FONT_FIX_KEY[f]] = dom_values[f]
            yield Finding(
                rule="body-font-consistent",
                kind="consistency",
                severity="info",
                slide=peer["slide"],
                anchor_id=peer["anchor_id"],
                shapeid=peer["shapeid"],
                message=(
                    f"Slide {peer['slide']} level-{level} body text "
                    f"'{_font_label(peer['values'], _BODY_FIELDS)}' differs from the dominant "
                    f"'{_font_label(dom_values, _BODY_FIELDS)}' ({count} of {total} peers)."
                ),
                fixable=True,
                fix=fix,
                observed=_font_label(peer["values"], _BODY_FIELDS),
                expected=_font_label(dom_values, _BODY_FIELDS),
            )


_BOX_KEYS = ("left", "top", "width", "height")


def _box_key(geo: dict[str, Any]) -> tuple[float, ...]:
    return tuple(round(float(geo[k]) * 2) / 2 for k in _BOX_KEYS)  # snap to 0.5 pt


def _check_title_position_consistent(
    walk: DeckWalk, scope: Scope, profile: Profile
) -> Iterator[Finding]:
    threshold = _threshold(profile, "title-position-consistent")
    for (_layout, kind), peers in _title_peers(walk).items():
        keyed: list[Any] = [_box_key(p["geometry"]) for p in peers]
        dom = dominant(keyed, threshold=threshold)
        if dom is None:
            continue
        dom_key, count, total = dom
        dom_box = dict(zip(_BOX_KEYS, dom_key, strict=True))
        for peer in peers:
            if not scope.covers(peer["slide"], peer["id"]):
                continue
            geo = peer["geometry"]
            diffs = {
                k: dom_box[k]
                for k in _BOX_KEYS
                if abs(float(geo[k]) - dom_box[k]) > _POSITION_TOLERANCE
            }
            if not diffs:
                continue
            ops: list[dict[str, Any]] = []
            move = {k: diffs[k] for k in ("left", "top") if k in diffs}
            size = {k: diffs[k] for k in ("width", "height") if k in diffs}
            if move:
                ops.append({"op": "shape_move", "anchor_id": peer["shapeid"], **move})
            if size:
                ops.append({"op": "shape_resize", "anchor_id": peer["shapeid"], **size})
            observed = ", ".join(f"{k}={float(geo[k]):g}" for k in diffs)
            expected = ", ".join(f"{k}={v:g}" for k, v in diffs.items())
            yield Finding(
                rule="title-position-consistent",
                kind="consistency",
                severity="warning",
                slide=peer["slide"],
                anchor_id=peer["shapeid"],
                shapeid=peer["shapeid"],
                message=(
                    f"Slide {peer['slide']} {kind} box ({observed}) deviates from the dominant "
                    f"{kind} box ({expected}; {count} of {total} peers) — the 'jumpy title'."
                ),
                fixable=True,
                fix=ops,
                observed=observed,
                expected=expected,
            )


for _rule in (
    Rule(
        id="title-font-consistent",
        kind="consistency",
        severity="warning",
        tags=("titles", "fonts"),
        check=_check_title_font_consistent,
    ),
    Rule(
        id="mixed-runs-in-title",
        kind="consistency",
        severity="info",
        tags=("titles", "fonts"),
        check=_check_mixed_runs_in_title,
    ),
    Rule(
        id="body-font-consistent",
        kind="consistency",
        severity="info",
        tags=("fonts",),
        check=_check_body_font_consistent,
    ),
    Rule(
        id="title-position-consistent",
        kind="consistency",
        severity="warning",
        tags=("titles", "alignment", "geometry"),
        check=_check_title_position_consistent,
    ),
):
    _register_rule(_rule)

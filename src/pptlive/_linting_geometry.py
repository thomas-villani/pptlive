"""P3 — the geometry / alignment rules (`spec-linter.md` §5 "Alignment & geometry").

The user's second ask — *"the shapes are lined up properly"* — and the biggest
net-new cluster vs wordlive, since Word has no 2-D canvas. Built on the shipped
`Slide.geometry_report()` (slide size, per-shape `box` + `off_slide`, `overlaps`)
and `shapes.list()` geometry; axis-aligned only (rotation is reported, not
accounted for — rotated shapes are skipped by `edge-alignment`).

- `shape-off-slide` (on, report-only) — a shape wholly or partly outside the slide
  bounds. Nudging it back needs judgment (which edge? shrink or move?), so it only
  reports. The other reason it matters: the oversize-and-bleed picture habit
  poisons this flag forever — `crop_to_fit` is the fix.
- `edge-alignment` (off — `alignment` tag) — 2+ shapes on a slide whose left /
  right / top / bottom edges fall within a tolerance (default **3 pt**, profile
  `tolerance`) of each other but aren't equal: the "meant to line up" defect.
  Fix: `shape_align` to that edge relative to the selection (snaps the members to
  their extreme member — deterministic, idempotent). Off by default because the
  tolerance is an opinion; `--rules alignment` lights it up.
- `placeholder-off-layout` (off — `alignment` tag) — a placeholder moved / resized
  off its `CustomLayout` position (> 1 pt). Fix: `shape_reset_layout`, which also
  restores the layout's default font size (the verb's contract) — aggressive, so
  opt-in.
- `overlap-unintended` (off — `alignment` tag, report-only) — two text-bearing
  shapes whose boxes intersect (from `overlaps`, filtered to shapes with text, so a
  picture behind a caption doesn't fire).

Imported by `_linting` for its side effect of registering the rules.
"""

from __future__ import annotations

from collections.abc import Iterator
from typing import TYPE_CHECKING, Any

from ._format_info import layout_placeholder_for
from ._linting import DeckWalk, Finding, Rule, Scope, _register_rule, shape_text

if TYPE_CHECKING:
    from ._lint_profile import Profile

DEFAULT_EDGE_TOLERANCE = 3.0
_LAYOUT_TOLERANCE = 1.0
_EDGES = ("left", "top", "right", "bottom")


def _check_shape_off_slide(walk: DeckWalk, scope: Scope, profile: Profile) -> Iterator[Finding]:
    for slide in walk.slides():
        if not scope.covers(slide["index"]):
            continue
        geo = walk.geometry(slide["index"])
        size = geo["slide_size"]
        for entry in geo["shapes"]:
            if not entry["off_slide"] or not scope.covers(slide["index"], entry["id"]):
                continue
            box = entry["box"]
            observed = (
                f"left={box['left']:g}, top={box['top']:g}, right={box['right']:g}, "
                f"bottom={box['bottom']:g}"
            )
            yield Finding(
                rule="shape-off-slide",
                kind="structural",
                severity="warning",
                slide=slide["index"],
                anchor_id=entry["shapeid"],
                shapeid=entry["shapeid"],
                message=(
                    f"Slide {slide['index']} shape '{entry['name']}' extends beyond the "
                    f"{size['width']:g}×{size['height']:g} pt slide ({observed}). Move or "
                    "resize it (a bleeding picture wants shape_crop_to_fit, not a bigger box)."
                ),
                fixable=False,
                observed=observed,
                expected=f"within 0..{size['width']:g} × 0..{size['height']:g}",
            )


def _edge_value(geo: dict[str, Any], edge: str) -> float:
    left, top = float(geo["left"]), float(geo["top"])
    if edge == "left":
        return left
    if edge == "top":
        return top
    if edge == "right":
        return left + float(geo["width"])
    return top + float(geo["height"])


def _clusters(
    items: list[tuple[float, dict[str, Any]]], tol: float
) -> list[list[tuple[float, Any]]]:
    """Group sorted `(value, shape)` pairs into runs where each value sits within
    `tol` of the run's first value."""
    out: list[list[tuple[float, Any]]] = []
    for value, shape in sorted(items, key=lambda t: t[0]):
        if out and value - out[-1][0][0] <= tol:
            out[-1].append((value, shape))
        else:
            out.append([(value, shape)])
    return out


def _check_edge_alignment(walk: DeckWalk, scope: Scope, profile: Profile) -> Iterator[Finding]:
    cfg = profile.config_for("edge-alignment").get("tolerance")
    tol = float(cfg) if isinstance(cfg, int | float) else DEFAULT_EDGE_TOLERANCE
    for slide in walk.slides():
        if not scope.covers(slide["index"]):
            continue
        candidates = [
            s
            for s in slide["shapes"]
            if s.get("geometry")
            and not s.get("connector")
            and abs(float(s["geometry"].get("rotation", 0.0))) < 0.01
        ]
        for edge in _EDGES:
            items = [(_edge_value(s["geometry"], edge), s) for s in candidates]
            for cluster in _clusters(items, tol):
                if len(cluster) < 2:
                    continue
                spread = cluster[-1][0] - cluster[0][0]
                if spread <= 0.01:
                    continue
                members = [s for _v, s in cluster]
                if not any(scope.covers(slide["index"], m["id"]) for m in members):
                    continue
                names = ", ".join(f"'{m['name']}'" for m in members)
                observed = ", ".join(f"{v:g}" for v, _s in cluster)
                target = cluster[0][0] if edge in ("left", "top") else cluster[-1][0]
                yield Finding(
                    rule="edge-alignment",
                    kind="consistency",
                    severity="info",
                    slide=slide["index"],
                    anchor_id=members[0]["shapeid"],
                    shapeid=members[0]["shapeid"],
                    message=(
                        f"Slide {slide['index']}: the {edge} edges of {names} are within "
                        f"{spread:.1f} pt ({observed}) but not equal — align them to "
                        f"{target:g}."
                    ),
                    fixable=True,
                    fix={
                        "op": "shape_align",
                        "anchors": [m["shapeid"] for m in members],
                        "how": edge,
                        "relative_to": "selection",
                    },
                    observed=f"{edge}: {observed}",
                    expected=f"{edge}: {target:g}",
                )


def _check_placeholder_off_layout(
    walk: DeckWalk, scope: Scope, profile: Profile
) -> Iterator[Finding]:
    deck = walk.deck
    for slide in walk.slides():
        if not scope.covers(slide["index"]):
            continue
        slide_obj = deck.slides[slide["index"]]
        for shape in slide["shapes"]:
            if not shape.get("placeholder") or not scope.covers(slide["index"], shape["id"]):
                continue
            layout_ph = layout_placeholder_for(
                slide_obj.shapes.by_id(shape["id"]).com, slide_obj.com
            )
            if layout_ph is None:
                continue
            geo = shape["geometry"]
            want = {
                "left": float(layout_ph.Left),
                "top": float(layout_ph.Top),
                "width": float(layout_ph.Width),
                "height": float(layout_ph.Height),
            }
            diffs = [k for k in want if abs(float(geo[k]) - want[k]) > _LAYOUT_TOLERANCE]
            if not diffs:
                continue
            observed = ", ".join(f"{k}={float(geo[k]):g}" for k in diffs)
            expected = ", ".join(f"{k}={want[k]:g}" for k in diffs)
            yield Finding(
                rule="placeholder-off-layout",
                kind="consistency",
                severity="info",
                slide=slide["index"],
                anchor_id=shape["shapeid"],
                shapeid=shape["shapeid"],
                message=(
                    f"Slide {slide['index']} {shape['placeholder']} placeholder sits off its "
                    f"layout position ({observed}; layout: {expected}). shape_reset_layout "
                    "restores the box and the layout's default font size."
                ),
                fixable=True,
                fix={"op": "shape_reset_layout", "anchor_id": shape["shapeid"]},
                observed=observed,
                expected=expected,
            )


def _check_overlap_unintended(walk: DeckWalk, scope: Scope, profile: Profile) -> Iterator[Finding]:
    for slide in walk.slides():
        if not scope.covers(slide["index"]):
            continue
        by_anchor = {s["anchor_id"]: s for s in slide["shapes"]}
        geo = walk.geometry(slide["index"])
        for pair in geo["overlaps"]:
            a, b = by_anchor.get(pair["a"]), by_anchor.get(pair["b"])
            if a is None or b is None:
                continue
            if not (shape_text(a).strip() and shape_text(b).strip()):
                continue
            if not (scope.covers(slide["index"], a["id"]) or scope.covers(slide["index"], b["id"])):
                continue
            yield Finding(
                rule="overlap-unintended",
                kind="structural",
                severity="info",
                slide=slide["index"],
                anchor_id=a["shapeid"],
                shapeid=a["shapeid"],
                message=(
                    f"Slide {slide['index']}: text shapes '{a['name']}' and '{b['name']}' "
                    f"overlap by {pair['area']:.0f} pt² ({b['shapeid']})."
                ),
                fixable=False,
                observed=f"overlap {pair['area']:.0f} pt² with {b['shapeid']}",
                expected="no overlap between text-bearing shapes",
            )


for _rule in (
    Rule(
        id="shape-off-slide",
        kind="structural",
        severity="warning",
        tags=("alignment", "geometry", "finalization"),
        check=_check_shape_off_slide,
    ),
    Rule(
        id="edge-alignment",
        kind="consistency",
        severity="info",
        tags=("alignment", "geometry"),
        check=_check_edge_alignment,
        default_on=False,
    ),
    Rule(
        id="placeholder-off-layout",
        kind="consistency",
        severity="info",
        tags=("alignment", "geometry"),
        check=_check_placeholder_off_layout,
        default_on=False,
    ),
    Rule(
        id="overlap-unintended",
        kind="structural",
        severity="info",
        tags=("alignment", "geometry"),
        check=_check_overlap_unintended,
        default_on=False,
    ),
):
    _register_rule(_rule)

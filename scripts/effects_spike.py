"""Spike harness — pin shape *effects* (shadow / glow / soft-edge / reflection / 3-D)
against real PowerPoint.

The roadmap's v1.2 "Effects (second cut)" is deferred and unspiked. This pokes the
raw `MsoShapeFormat` effect sub-objects directly (no wrappers exist yet) to learn
which set+read-back cleanly (so a `read()` can reconstruct them) and which are
write-only / non-round-tripping (the SmartArt-assistant-node hazard):

- **Shadow** — `Shape.Shadow` (`MsoShadowFormat`): `.Type` preset, `.Visible`,
  `.ForeColor.RGB`, `.Transparency`, `.Blur`, `.Size`, `.OffsetX`/`.OffsetY`,
  `.Style`. Both the preset path (`.Type = N`) and the individual-property path.
- **Glow** — `Shape.Glow` (`GlowFormat`): `.Color.RGB`, `.Radius`, `.Transparency`.
- **SoftEdge** — `Shape.SoftEdge`: `.Type` (0..6 presets), `.Radius`.
- **Reflection** — `Shape.Reflection`: `.Type` (0..9 presets).
- **ThreeD** — `Shape.ThreeD` (`ThreeDFormat`): `.Visible`, `.Depth`, bevels,
  `.SetThreeDFormat(preset)`, rotation.

Run:  uv run python scripts/effects_spike.py

Net-zero / polite exactly like `scripts/style_spike.py`.
"""

from __future__ import annotations

import json
from typing import Any

import pptlive as pl
from pptlive import _selection

_RED = 0x0000FF
_BLUE = 0xFF0000


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:200]


def _read(obj: Any, props: tuple[str, ...]) -> dict[str, Any]:
    """Read a tuple of properties, capturing per-prop errors (some are write-only)."""
    out: dict[str, Any] = {}
    for name in props:
        try:
            val = getattr(obj, name)
            # ColorFormat sub-objects -> read .RGB
            out[name] = val if isinstance(val, (int, float, str)) else _err(
                TypeError(f"non-scalar {type(val).__name__}")
            )
        except Exception as exc:
            out[name] = _err(exc)
    return out


def probe_shadow(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=20.0, top=20.0, width=120.0, height=90.0)
        sd = sh.com.Shadow
        out["default"] = _read(sd, ("Type", "Visible", "Transparency", "Blur", "Size", "OffsetX", "OffsetY", "Style"))
        # Preset path
        try:
            sd.Type = 25  # an outer-shadow preset (msoShadow25)
            out["preset_set_ok"] = True
        except Exception as exc:
            out["preset_set_error"] = _err(exc)
        # Individual-property path (the modern, flexible one)
        try:
            sd.Visible = -1
            sd.ForeColor.RGB = _RED
            sd.Transparency = 0.4
            sd.Blur = 8.0
            sd.Size = 100.0
            sd.OffsetX = 4.0
            sd.OffsetY = 4.0
            out["props_set_ok"] = True
        except Exception as exc:
            out["props_set_error"] = _err(exc)
        out["after"] = _read(sd, ("Type", "Visible", "Transparency", "Blur", "Size", "OffsetX", "OffsetY", "Style"))
        try:
            out["after"]["ForeColor.RGB"] = int(sd.ForeColor.RGB)
        except Exception as exc:
            out["after"]["ForeColor.RGB"] = _err(exc)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_glow(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=160.0, top=20.0, width=120.0, height=90.0)
        g = sh.com.Glow
        out["default"] = _read(g, ("Radius", "Transparency"))
        try:
            g.Color.RGB = _BLUE
            g.Radius = 10.0
            g.Transparency = 0.2
            out["set_ok"] = True
        except Exception as exc:
            out["set_error"] = _err(exc)
        out["after"] = _read(g, ("Radius", "Transparency"))
        try:
            out["after"]["Color.RGB"] = int(g.Color.RGB)
        except Exception as exc:
            out["after"]["Color.RGB"] = _err(exc)
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_soft_edge(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=300.0, top=20.0, width=120.0, height=90.0)
        se = sh.com.SoftEdge
        out["default"] = _read(se, ("Type", "Radius"))
        try:
            se.Type = 4  # a preset soft-edge radius
            out["set_ok"] = True
        except Exception as exc:
            out["set_error"] = _err(exc)
        out["after"] = _read(se, ("Type", "Radius"))
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_reflection(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=20.0, top=130.0, width=120.0, height=90.0)
        r = sh.com.Reflection
        out["default"] = _read(r, ("Type",))
        try:
            r.Type = 5  # a preset reflection
            out["set_ok"] = True
        except Exception as exc:
            out["set_error"] = _err(exc)
        out["after"] = _read(r, ("Type",))
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def probe_threed(shapes: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    try:
        sh = shapes.add_shape("rectangle", left=160.0, top=130.0, width=120.0, height=90.0)
        td = sh.com.ThreeD
        out["default"] = _read(td, ("Visible", "Depth", "RotationX", "RotationY", "PresetMaterial", "BevelTopType", "BevelTopInset", "BevelTopDepth"))
        try:
            td.SetThreeDFormat(1)  # a preset 3-D format
            out["preset_set_ok"] = True
        except Exception as exc:
            out["preset_set_error"] = _err(exc)
        try:
            td.Depth = 20.0
            td.BevelTopType = 1
            td.BevelTopInset = 6.0
            td.BevelTopDepth = 6.0
            out["props_set_ok"] = True
        except Exception as exc:
            out["props_set_error"] = _err(exc)
        out["after"] = _read(td, ("Visible", "Depth", "RotationX", "RotationY", "PresetMaterial", "BevelTopType", "BevelTopInset", "BevelTopDepth"))
    except Exception as exc:
        out["error"] = _err(exc)
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        findings["slide_count_before"] = count_before

        snap = _selection.snapshot(ppt)
        findings["viewed_slide"] = snap.slide_index

        temp_ids: list[int] = []
        try:
            with deck.edit("effects spike: shadow/glow/softedge/reflection/3d"):
                temp = deck.slides.add(layout="blank")
                temp_ids.append(temp.id)
                shapes = deck.slides[temp.index].shapes
                findings["shadow"] = probe_shadow(shapes)
                findings["glow"] = probe_glow(shapes)
                findings["soft_edge"] = probe_soft_edge(shapes)
                findings["reflection"] = probe_reflection(shapes)
                findings["threed"] = probe_threed(shapes)
                findings["final_shape_count"] = len(shapes)
        finally:
            deleted: list[int] = []
            try:
                with deck.edit("effects spike: cleanup"):
                    for _ in range(len(temp_ids) + 2):
                        victim = None
                        for idx in range(len(deck.slides), 0, -1):
                            try:
                                sid = deck.slides[idx].id
                            except Exception:
                                continue
                            if sid in temp_ids and sid not in deleted:
                                victim = (idx, sid)
                                break
                        if victim is None:
                            break
                        deck.slides[victim[0]].delete()
                        deleted.append(victim[1])
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            findings["cleaned_up_ids"] = deleted
            _selection.restore(ppt, snap)

        count_after = len(deck.slides)
        findings["slide_count_after"] = count_after
        findings["net_zero_ok"] = count_after == count_before

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

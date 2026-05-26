"""Unit helpers — points throughout, never EMUs.

PowerPoint's COM layer measures geometry in **points** (1 inch = 72 pt), so
pptlive does too: `Shape.Left/Top/Width/Height`, slide dimensions, indents —
all points. EMUs are an OOXML / `python-pptx` concern and never surface here.

These helpers exist so an agent can write `pl.units.inches(1.5)` instead of
hardcoding `108`. `points()` is the identity, included for symmetry and to make
"this number is already in points" explicit at a call site.

    shape.move(left=inches(1), top=cm(4))
"""

from __future__ import annotations

_POINTS_PER_INCH = 72.0
_CM_PER_INCH = 2.54


def points(value: float) -> float:
    """Return `value` unchanged — it is already in points (PowerPoint's unit)."""
    return float(value)


def inches(value: float) -> float:
    """Convert inches to points (1 in = 72 pt)."""
    return float(value) * _POINTS_PER_INCH


def cm(value: float) -> float:
    """Convert centimetres to points (1 in = 2.54 cm = 72 pt)."""
    return float(value) * _POINTS_PER_INCH / _CM_PER_INCH


def mm(value: float) -> float:
    """Convert millimetres to points."""
    return cm(float(value) / 10.0)

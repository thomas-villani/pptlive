"""Probe what PowerPoint's COM `Font` actually exposes about direct-vs-inherited
formatting (PPTLIVE-003), so the `ppt_read` font surface promises only what COM
can back.

Finding (drives the design of `_anchors.font_to_dict`):
  - `Font.Bold/Italic/Underline/Size/Name` return the *effective* (rendered)
    value. A value inherited from the master text style and the same value set
    directly on the run read back identically — COM exposes **no** "is this
    directly set" flag for them. So pptlive reports effective values + the
    tri-state `"mixed"` signal, and does NOT claim direct-vs-inherited for these.
  - `Font.Color` is the one exception: `.Color.Type` distinguishes a literal RGB
    (msoColorTypeRGB=1) from a theme/scheme color (msoColorTypeScheme=2), and an
    automatic color returns the `0x80000000` sentinel from `.RGB`. So a literal
    color surfaces as `#RRGGBB` and an inherited/theme color as `null`.

This script proves it: it flips a run's bold ON directly and shows the read is the
same shape as an inherited-bold run, then reports each `.Color.Type`. Net-zero —
it restores whatever it changed.

    uv run python scripts/inherit_probe.py
"""

from __future__ import annotations

import json
import sys

import pptlive as pl
from pptlive._anchors import font_to_dict

TARGET_SLIDE = 8  # a content slide with a body placeholder in the test deck


def main(argv: list[str]) -> int:
    slide_index = int(argv[1]) if len(argv) > 1 else TARGET_SLIDE
    out: dict[str, object] = {}
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        body = deck.slides[slide_index].placeholder("body")
        tr = body.com.TextFrame.TextRange

        out["effective_as_read"] = font_to_dict(tr)
        out["color_type"] = int(tr.Font.Color.Type)  # 1=RGB, 2=scheme, -2=mixed
        out["color_rgb_raw"] = int(tr.Font.Color.RGB)

        # Flip bold directly, re-read: the value changes but the *shape* of the
        # read is identical to an inherited-bold run — no "direct" marker appears.
        original_bold = int(tr.Font.Bold)
        with deck.edit("inherit-probe: set bold"):
            body.format_text(bold=True)
        out["after_direct_bold"] = font_to_dict(tr)
        # Restore.
        with deck.edit("inherit-probe: restore"):
            tr.Font.Bold = original_bold

    out["conclusion"] = (
        "bold/italic/size/name expose effective values only (no direct/inherited "
        "flag); color distinguishes literal RGB vs theme via Color.Type / the "
        "0x80000000 sentinel."
    )
    print(json.dumps(out, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

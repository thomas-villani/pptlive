"""pptlive — drive a running Microsoft PowerPoint instance from Python.

xlwings, but for PowerPoint, and built for LLM agents. The live-app sibling of
`python-pptx` (which works the file on disk) and the PowerPoint sibling of
`wordlive`.

Quick start:

    import pptlive as pl

    with pl.attach() as ppt:
        deck = ppt.presentations.active
        with deck.edit("Set the agenda"):          # preserves the viewed slide
            deck.anchor_by_id("ph:2:title").set_text("Agenda")
            deck.anchor_by_id("ph:2:body").set_text("Intro\\nDemo\\nQ&A")

Note: `edit()` preserves the user's view and selection *and* is an atomic-undo
scope — PowerPoint groups a block's edits into a single Ctrl-Z (fenced with
`StartNewUndoEntry`). See `_edit.EditScope` for the mechanism and caveats.
"""

from __future__ import annotations

from . import constants, units
from ._anchors import Anchor, Notes
from ._app import PowerPoint, attach, connect
from ._edit import EditScope
from ._presentation import Presentation, PresentationCollection
from ._selection import SelectionSnapshot
from ._shapes import PlaceholderShape, Shape, ShapeCollection
from ._slides import Slide, SlideCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    LayoutNotFoundError,
    NoTextFrameError,
    PowerPointBusyError,
    PowerPointNotRunningError,
    PptliveError,
    PresentationNotFoundError,
    SlideNotFoundError,
)

__all__ = [
    "AmbiguousMatchError",
    "Anchor",
    "AnchorNotFoundError",
    "ComError",
    "EditScope",
    "LayoutNotFoundError",
    "NoTextFrameError",
    "Notes",
    "PlaceholderShape",
    "PowerPoint",
    "PowerPointBusyError",
    "PowerPointNotRunningError",
    "Presentation",
    "PresentationCollection",
    "PresentationNotFoundError",
    "PptliveError",
    "SelectionSnapshot",
    "Shape",
    "ShapeCollection",
    "Slide",
    "SlideCollection",
    "SlideNotFoundError",
    "attach",
    "connect",
    "constants",
    "units",
]

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

from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _pkg_version

try:
    __version__ = _pkg_version("pptlive")
except PackageNotFoundError:  # not installed (e.g. running from a source tree)
    __version__ = "0.0.0+unknown"

from . import constants, units
from ._anchors import Anchor, Notes, Paragraph, ParagraphCollection
from ._app import PowerPoint, attach, connect
from ._charts import Chart
from ._comments import Comment, CommentCollection
from ._edit import EditScope
from ._headersfooters import HeadersFooters
from ._presentation import Presentation, PresentationCollection, VideoExportResult
from ._sections import SectionCollection
from ._selection import SelectionInfo, SelectionSnapshot
from ._shapes import PlaceholderShape, Shape, ShapeById, ShapeCollection, TextFrameStatus
from ._show import SlideShow
from ._slides import Slide, SlideCollection
from ._smartart import SmartArt
from ._snapshot import Snapshot
from ._tables import Cell, Table
from ._theme import Master, Theme
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
    ReplaceVerificationError,
    SlideNotFoundError,
    SlideShowNotRunningError,
    UnsavedPresentationError,
    VideoExportError,
)

__all__ = [
    "AmbiguousMatchError",
    "Anchor",
    "AnchorNotFoundError",
    "Cell",
    "Chart",
    "ComError",
    "Comment",
    "CommentCollection",
    "EditScope",
    "HeadersFooters",
    "LayoutNotFoundError",
    "Master",
    "NoTextFrameError",
    "Notes",
    "Paragraph",
    "ParagraphCollection",
    "PlaceholderShape",
    "PowerPoint",
    "PowerPointBusyError",
    "PowerPointNotRunningError",
    "Presentation",
    "PresentationCollection",
    "PresentationNotFoundError",
    "PptliveError",
    "ReplaceVerificationError",
    "SectionCollection",
    "SelectionInfo",
    "SelectionSnapshot",
    "Shape",
    "ShapeById",
    "ShapeCollection",
    "Slide",
    "SlideCollection",
    "SlideNotFoundError",
    "SlideShow",
    "SlideShowNotRunningError",
    "SmartArt",
    "Snapshot",
    "Table",
    "TextFrameStatus",
    "Theme",
    "UnsavedPresentationError",
    "VideoExportError",
    "VideoExportResult",
    "__version__",
    "attach",
    "connect",
    "constants",
    "units",
]

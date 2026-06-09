"""Diagnose the view-reset regression: after an MCP action the user's viewed
slide jumps back to slide 1, despite `EditScope` snapshotting + restoring it.

This probes both the **library** path (so we can peek at the viewed slide
*inside* the `edit()` scope, before restore) and the **MCP** path (the surface
the user actually drives from Claude Desktop — before/after only, since each tool
call opens + closes its own `attach()`).

For each op we record the viewed slide index (`PowerPoint.viewed_slide_index`)
at three points:
  - **before**: just after parking the view on TARGET_SLIDE, before `edit()`;
  - **inside**: inside the scope, after the op, before restore (library path only);
  - **after**: once the scope/tool call has returned and restore has run.
Plus the current thread name, to check the apartment/threading assumption.

Reading `before != after` (with `after == 1`) reproduces the bug. `inside`
localises it: if `inside` already moved, the *op* moves the view and we need a
restore that catches it; if only `after` is wrong, the snapshot/restore itself is
at fault. Each phase cleans up the structural change it makes.

Usage (PowerPoint open, with a deck of >=3 slides):

    uv run python scripts/view_repro.py            # run every phase
    uv run python scripts/view_repro.py lib        # library-path phases only
    uv run python scripts/view_repro.py mcp        # MCP-path phases only
"""

from __future__ import annotations

import json
import sys
import threading
from collections.abc import Callable
from typing import Any

import pptlive as pl

TARGET_SLIDE = 3


def _goto(ppt: pl.PowerPoint, index: int) -> None:
    try:
        ppt.com.ActiveWindow.View.GotoSlide(index)
    except Exception:
        pass


def _viewed() -> int | None:
    """Read the live viewed-slide index through a fresh, short attach."""
    with pl.attach() as ppt:
        return ppt.viewed_slide_index()


# ---------------------------------------------------------------------------
# Library-path probes — peek at the viewed slide *inside* the edit() scope.
# `op` mutates the deck and returns a zero-arg cleanup to undo any structural
# change; text edits are restored by the cleanup too.
# ---------------------------------------------------------------------------


def _lib_probe(name: str, op: Callable[[pl.Presentation], Callable[[], None]]) -> dict[str, Any]:
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        _goto(ppt, TARGET_SLIDE)
        before = ppt.viewed_slide_index()
        thread = threading.current_thread().name
        inside: int | None = None

        def cleanup() -> None:
            return None

        with deck.edit(f"view_repro: {name}") as _scope:
            cleanup = op(deck)
            inside = ppt.viewed_slide_index()
        after = ppt.viewed_slide_index()
        try:
            cleanup()
        except Exception as exc:  # noqa: BLE001
            return {
                "phase": f"lib:{name}",
                "thread": thread,
                "before": before,
                "inside": inside,
                "after": after,
                "cleanup_error": repr(exc),
            }
        return {
            "phase": f"lib:{name}",
            "thread": thread,
            "before": before,
            "inside": inside,
            "after": after,
            "moved": before != after,
        }


def _op_write(deck: pl.Presentation) -> Callable[[], None]:
    title = deck.slides[TARGET_SLIDE].placeholder("title")
    original = title.text
    title.set_text("view_repro write probe")
    return lambda: title.set_text(original)


def _op_shape_add(deck: pl.Presentation) -> Callable[[], None]:
    shape = deck.slides[TARGET_SLIDE].shapes.add_textbox("view_repro probe box")
    return shape.delete


def _op_slide_add(deck: pl.Presentation) -> Callable[[], None]:
    new = deck.slides.add()
    return new.delete


def _op_slide_duplicate(deck: pl.Presentation) -> Callable[[], None]:
    copy = deck.slides[TARGET_SLIDE].duplicate()
    return copy.delete


def _op_find_replace(deck: pl.Presentation) -> Callable[[], None]:
    # Self-contained: stamp a unique token, replace it, then clean the box.
    token = "ViewReproTokenAlpha"
    shape = deck.slides[TARGET_SLIDE].shapes.add_textbox(token)
    deck.find_replace(token, "ViewReproTokenOmega", all=True)
    return shape.delete


def _op_batch_like(deck: pl.Presentation) -> Callable[[], None]:
    # Mimic an atomic batch: several mutations under one edit() scope.
    new = deck.slides.add()
    box = deck.slides[TARGET_SLIDE].shapes.add_textbox("view_repro batch box")

    def _cleanup() -> None:
        box.delete()
        new.delete()

    return _cleanup


# ---------------------------------------------------------------------------
# MCP-path probes — the real surface. Park the view, call the tool, re-read.
# ---------------------------------------------------------------------------


def _mcp_probe(name: str, call: Callable[[], Any], cleanup: Callable[[], None]) -> dict[str, Any]:
    with pl.attach() as ppt:
        _goto(ppt, TARGET_SLIDE)
    before = _viewed()
    thread_seen: dict[str, str] = {}

    orig_thread = threading.current_thread().name
    try:
        call()
    finally:
        thread_seen["after_call"] = threading.current_thread().name
    after = _viewed()
    try:
        cleanup()
    except Exception:  # noqa: BLE001
        pass
    return {
        "phase": f"mcp:{name}",
        "thread": orig_thread,
        "before": before,
        "after": after,
        "moved": before != after,
    }


def _run_mcp_phases() -> list[dict[str, Any]]:
    import importlib

    server = importlib.import_module("pptlive.mcp.server")

    out: list[dict[str, Any]] = []

    # write via ppt_edit
    with pl.attach() as ppt:
        title = ppt.presentations.active.slides[TARGET_SLIDE].placeholder("title")
        original = title.text
    out.append(
        _mcp_probe(
            "edit_write",
            lambda: server.ppt_edit(
                op="write", anchor_id=f"ph:{TARGET_SLIDE}:title", text="view_repro mcp write"
            ),
            lambda: server.ppt_edit(
                op="write", anchor_id=f"ph:{TARGET_SLIDE}:title", text=original
            ),
        )
    )

    # atomic batch: slide_add + write
    out.append(
        _mcp_probe(
            "batch_atomic",
            lambda: server.ppt_batch(
                commands=[
                    {"tool": "edit", "op": "slide_add"},
                    {
                        "tool": "edit",
                        "op": "write",
                        "anchor_id": f"ph:{TARGET_SLIDE}:title",
                        "text": "view_repro mcp batch",
                    },
                ],
                atomic=True,
                embed=False,
            ),
            _cleanup_last_slide,
        )
    )
    return out


def _cleanup_last_slide() -> None:
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        last = len(deck.slides)
        deck.slides[last].delete()


def main(argv: list[str]) -> int:
    which = argv[1] if len(argv) > 1 else "all"
    results: list[dict[str, Any]] = []

    if which in ("all", "lib"):
        results.append(_lib_probe("write", _op_write))
        results.append(_lib_probe("shape_add", _op_shape_add))
        results.append(_lib_probe("slide_add", _op_slide_add))
        results.append(_lib_probe("slide_duplicate", _op_slide_duplicate))
        results.append(_lib_probe("find_replace", _op_find_replace))
        results.append(_lib_probe("batch_like", _op_batch_like))
    if which in ("all", "mcp"):
        results.extend(_run_mcp_phases())

    print(json.dumps(results, indent=2, default=str))
    moved = [r["phase"] for r in results if r.get("moved")]
    print(f"\n=== phases that moved the view: {moved or 'none'} ===", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

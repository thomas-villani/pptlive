"""Click entry point for `pptlive`. JSON in, JSON out, deterministic exit codes."""

from __future__ import annotations

import json
import sys
from typing import Any

import click

from ..exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    NoTextFrameError,
    PowerPointBusyError,
    PowerPointNotRunningError,
    PptliveError,
    PresentationNotFoundError,
)

# Exit codes per spec.md §"Error taxonomy → exit codes":
EXIT_OK = 0
EXIT_OTHER = 1
EXIT_ANCHOR_NOT_FOUND = 2  # anchor / slide / shape / presentation not found; zero find matches
EXIT_POWERPOINT_BUSY = 3  # PowerPoint busy / modal dialog (a running show does not block edits)
EXIT_POWERPOINT_NOT_RUNNING = 4
EXIT_AMBIGUOUS_MATCH = 5
EXIT_NO_TEXT_FRAME = 6  # the one genuinely new code


def emit(payload: Any, *, as_text: bool = False, text: str | None = None) -> None:
    """One JSON object on stdout per invocation.

    With `--json` (default), `payload` is dumped as JSON. With `--text`, `text`
    (if given) is echoed verbatim; otherwise we fall back to pretty-printed JSON
    of `payload` so machines and humans see the same data.
    """
    if as_text:
        if text is not None:
            click.echo(text)
        else:
            click.echo(payload if isinstance(payload, str) else json.dumps(payload, indent=2))
    else:
        click.echo(json.dumps(payload, ensure_ascii=False))


def _exit_for(exc: PptliveError) -> int:
    # NoTextFrameError before the generic fallthrough; AnchorNotFoundError covers
    # SlideNotFoundError (subclass) and zero find matches.
    if isinstance(exc, NoTextFrameError):
        return EXIT_NO_TEXT_FRAME
    if isinstance(exc, AnchorNotFoundError):
        return EXIT_ANCHOR_NOT_FOUND
    if isinstance(exc, AmbiguousMatchError):
        return EXIT_AMBIGUOUS_MATCH
    if isinstance(exc, PowerPointBusyError):
        return EXIT_POWERPOINT_BUSY
    if isinstance(exc, PowerPointNotRunningError):
        return EXIT_POWERPOINT_NOT_RUNNING
    if isinstance(exc, PresentationNotFoundError):
        return EXIT_ANCHOR_NOT_FOUND
    return EXIT_OTHER


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.option("--json/--text", "as_json", default=True, help="Output format (default JSON).")
@click.option(
    "--doc", "doc_name", default=None, help="Target presentation by name (default: active)."
)
@click.pass_context
def main(ctx: click.Context, as_json: bool, doc_name: str | None) -> None:
    """pptlive — drive a running Microsoft PowerPoint instance."""
    ctx.ensure_object(dict)
    ctx.obj["as_json"] = as_json
    ctx.obj["doc_name"] = doc_name


def _run(ctx: click.Context, fn: Any) -> None:
    """Top-level error boundary: classify PptliveError into exit codes."""
    try:
        fn()
    except PptliveError as exc:
        click.echo(f"error: {exc}", err=True)
        sys.exit(_exit_for(exc))


# Register subcommands. Import here to avoid a circular dependency at module load.
from . import commands  # noqa: E402

commands.register(main)

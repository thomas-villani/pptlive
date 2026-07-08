"""Click entry point for `pptlive`. JSON in, JSON out, deterministic exit codes."""

from __future__ import annotations

import json
import sys
from typing import Any

import click

from .. import __version__
from ..exceptions import (
    EXIT_CODE_FOR,
    PptliveError,
    classify,
)

# Exit codes per spec.md §"Error taxonomy → exit codes":
EXIT_OK = 0
EXIT_OTHER = 1


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
    # The taxonomy lives in exceptions.classify() (shared with _batch._error_code);
    # here we just map its code token to the spec's exit int.
    return EXIT_CODE_FOR.get(classify(exc), EXIT_OTHER)


@click.group(context_settings={"help_option_names": ["-h", "--help"]})
@click.version_option(__version__, "-V", "--version", message="pptlive %(version)s")
@click.option("--json/--text", "as_json", default=True, help="Output format (default JSON).")
@click.option(
    "--doc", "doc_name", default=None, help="Target presentation by name (default: active)."
)
@click.pass_context
def main(ctx: click.Context, as_json: bool, doc_name: str | None) -> None:
    """pptlive — drive a running Microsoft PowerPoint instance.

    LLM agent? Run `pptlive llm-help` for the full agent guide in one shot: the
    anchor model, every command, and the exit-code taxonomy (add `--python` for
    the Python-API guide). `pptlive install-skill` drops those guides into
    `.agents/skills/`, and `pptlive install-mcp` registers the MCP server.
    """
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
    except (ValueError, OSError) as exc:
        # ValueError: library-level input validation (e.g. a line_spacing multiple
        # > 5, an out-of-range indent level). OSError: file-touching verbs —
        # FileNotFoundError from `shape set-picture` / `picture-fill` / `add --kind
        # picture` / `media add`, FileExistsError from `save-as`. Both a clean exit
        # 1, not a traceback — mirroring the MCP boundary, which already maps
        # ValueError / FileNotFoundError / FileExistsError to `invalid_args`.
        click.echo(f"error: {exc}", err=True)
        sys.exit(EXIT_OTHER)


# Register subcommands. Import here to avoid a circular dependency at module load.
from . import commands  # noqa: E402

commands.register(main)

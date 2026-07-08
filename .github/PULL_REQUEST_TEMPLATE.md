<!-- Keep one logical change per PR. -->

## What & why

<!-- What does this change, and why? Link any issue (e.g. Fixes #123). -->

## The surfaces

pptlive exposes most capabilities across surfaces that must stay in sync. Tick
what this PR touches — and confirm the rest genuinely don't need it:

- [ ] Python API
- [ ] CLI verb (`cli/commands.py`)
- [ ] `exec` op / MCP op (`_batch.py`)
- [ ] MCP tool wrapper (`mcp/server.py`)
- [ ] Both SKILL guides (`_skill/pptlive-cli` + `_skill/pptlive-python`)
- [ ] N/A — not a user-facing capability change

## Checklist

- [ ] `uv run ruff format .` and `uv run ruff check .` pass
- [ ] `uv run mypy` passes
- [ ] `uv run pytest` passes; added/updated tests for the change
- [ ] Spiked / smoke-tested against real PowerPoint on Windows (or N/A — no COM behaviour changed)
- [ ] Updated the docs that track this change (`cli.md` / `mcp.md` / `python-api.md` / `README.md` / `SKILL.md`)
- [ ] Added a `CHANGELOG.md` entry under `[Unreleased]`

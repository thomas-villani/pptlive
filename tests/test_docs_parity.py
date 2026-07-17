"""Every dispatchable op must appear in the docs an agent or user actually reads.

The op tables drifted for two straight releases before anyone noticed: v0.7.0 shipped
`table_set_fill` / `shape_set_picture` / `smartart_format_node` / … and v0.8.0's
arrangement + run-link + `media_set` ops, none of which reached `README.md` or
`docs/mcp.md`. The cause is visible in the history — the media commit updated `docs/`,
the arrangement commit didn't — and the CHANGELOG's own definition of done ("Library +
CLI + MCP + both SKILL guides + tests") never mentions `docs/` at all.

`_batch.py` already turns an enum/registry mismatch into an import-time failure, so
code-side drift is impossible; this closes the same loop on the *docs* side.

Matching is on the **backticked** form (`` `slide_add` ``) rather than a bare
substring: ops like `write`, `format`, `find`, and `links` are ordinary English words
that appear in prose throughout both files, so a substring check would pass while
documenting nothing.

Needs no PowerPoint and no Office — it reads two markdown files.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from pptlive._batch import EditOp, ReadOp, RenderOp, ShowOp

REPO = Path(__file__).resolve().parent.parent
README = REPO / "README.md"
MCP_DOC = REPO / "docs" / "mcp.md"

#: Every op the four dispatchers accept, as `(tool, op)`.
ALL_OPS: list[tuple[str, str]] = [
    *(("ppt_read", str(o)) for o in ReadOp),
    *(("ppt_edit", str(o)) for o in EditOp),
    *(("ppt_render", str(o)) for o in RenderOp),
    *(("ppt_show", str(o)) for o in ShowOp),
]

#: Ops that are deliberately undocumented, with a reason. Keep this empty if you can:
#: an entry here is a promise that the op is genuinely internal, not a shortcut around
#: a failing test.
UNDOCUMENTED_OK: dict[str, str] = {}


def _doc_text(path: Path) -> str:
    if not path.exists():  # pragma: no cover - the repo always ships these
        pytest.fail(f"documentation file is missing: {path}")
    return path.read_text(encoding="utf-8")


@pytest.mark.parametrize(
    ("doc_path", "label"),
    [(README, "README.md"), (MCP_DOC, "docs/mcp.md")],
    ids=["readme", "mcp-doc"],
)
def test_every_op_is_documented(doc_path: Path, label: str) -> None:
    """Each op appears, backticked, in the op tables users and agents read."""
    text = _doc_text(doc_path)
    missing = [
        f"{tool} op={op!r}"
        for tool, op in ALL_OPS
        if op not in UNDOCUMENTED_OK and f"`{op}`" not in text
    ]
    assert not missing, (
        f"{len(missing)} op(s) dispatchable by _batch.py but absent from {label}.\n"
        "Add them to the op table (backticked), or record a reason in "
        "UNDOCUMENTED_OK:\n  " + "\n  ".join(missing)
    )


def test_undocumented_ok_has_no_stale_entries() -> None:
    """An excuse must name a real op — otherwise it's silently excusing nothing."""
    known = {op for _, op in ALL_OPS}
    stale = set(UNDOCUMENTED_OK) - known
    assert not stale, f"UNDOCUMENTED_OK names ops that no longer exist: {sorted(stale)}"

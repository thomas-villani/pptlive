"""The linter through its other three front-ends — CLI (`lint` / `regularize` /
`read format`), the `exec` batch op (`regularize`), and MCP (`ppt_read` `lint` /
`format_info`, `ppt_edit` `regularize`). The rule logic itself is covered by
`test_linting.py`; these pin the contracts (exit codes, payload shapes, the
`invalid_args` mapping) on the default fixture deck, whose slide-1 title is
unique on its layout — so the deck lints clean apart from nothing at all.
"""

from __future__ import annotations

import json
from typing import Any

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main


def _json(result: Any) -> Any:
    return json.loads(result.output)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def test_cli_lint_emits_count_and_findings(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["lint"])
    assert result.exit_code == 0, result.output
    out = _json(result)
    assert out["count"] == len(out["findings"])
    assert all(
        {"rule", "kind", "severity", "slide", "anchor_id", "shapeid", "fixable", "fix"} <= set(f)
        for f in out["findings"]
    )


def test_cli_lint_rule_and_exclude_are_exclusive(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["lint", "--rule", "titles", "--exclude", "fonts"])
    assert result.exit_code == 2  # click UsageError


def test_cli_lint_unknown_rule_is_exit_1(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["lint", "--rule", "nope"])
    assert result.exit_code == 1
    assert "unknown lint rule" in result.output


def test_cli_lint_bad_profile_is_exit_1(fake_powerpoint: Any, tmp_path: Any) -> None:
    bad = tmp_path / "p.json"
    bad.write_text("{not json", encoding="utf-8")
    result = CliRunner().invoke(main, ["lint", "--profile", str(bad)])
    assert result.exit_code == 1
    assert "not valid JSON" in result.output


def test_cli_lint_text_mode(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["--text", "lint", "--rule", "alignment"])
    assert result.exit_code == 0
    assert result.output.strip()  # "(no findings)" or one line per finding


def test_cli_regularize_dry_run(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["regularize", "--dry-run"])
    assert result.exit_code == 0, result.output
    out = _json(result)
    assert out["dry_run"] is True
    assert set(out) >= {"applied", "skipped", "deferred", "findings"}
    assert fake_powerpoint._undo_entries == 0


def test_cli_regularize_runs_one_edit(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["--text", "regularize", "--allow-content"])
    assert result.exit_code == 0, result.output
    assert result.output.startswith("fixed ")


def test_cli_read_format(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["read", "format", "--anchor-id", "ph:2:body"])
    assert result.exit_code == 0, result.output
    out = _json(result)
    assert out["anchor_id"] == "ph:2:body"
    assert out["placeholder"] == "body"
    assert out["cascade"] == "layout"  # the fixture's slide 2 carries a layout body
    assert out["font"]["size"] == {"value": 18.0, "baseline": 28.0, "override": True}
    text = CliRunner().invoke(main, ["--text", "read", "format", "--anchor-id", "ph:2:body"])
    assert "font.size: 18.0 (baseline 28.0) *override*" in text.output


def test_cli_read_format_missing_anchor_exit_2(fake_powerpoint: Any) -> None:
    result = CliRunner().invoke(main, ["read", "format", "--anchor-id", "ph:9:title"])
    assert result.exit_code == 2


# ---------------------------------------------------------------------------
# exec (run_batch) — regularize is a write op inside the batch's one undo entry
# ---------------------------------------------------------------------------


def test_exec_regularize_op_rides_the_batch_undo(fake_powerpoint: Any, tmp_path: Any) -> None:
    script = {
        "label": "tidy",
        "ops": [
            {"op": "format", "anchor_id": "ph:2:title", "size": 40},
            {"op": "regularize", "rules": ["alignment"], "dry_run": True},
        ],
    }
    path = tmp_path / "ops.json"
    path.write_text(json.dumps(script), encoding="utf-8")
    result = CliRunner().invoke(main, ["exec", "--script", str(path)])
    assert result.exit_code == 0, result.output
    out = _json(result)
    assert out["ok"] is True and out["count"] == 2
    report = out["results"][1]["result"]
    assert report["dry_run"] is True and "findings" in report
    assert fake_powerpoint._undo_entries == 1  # the batch's fence, not a nested one


# ---------------------------------------------------------------------------
# MCP
# ---------------------------------------------------------------------------

mcp = pytest.importorskip("mcp")


def test_mcp_read_lint_and_format_info(fake_powerpoint: Any) -> None:
    from pptlive.mcp.server import ppt_read

    out = ppt_read("lint", rules=["alignment"], within="slide:2")
    assert set(out) == {"count", "findings"}
    assert all(f["slide"] == 2 for f in out["findings"])

    info = ppt_read("format_info", anchor_id="ph:2:body")
    assert info["cascade"] == "layout"
    assert info["font"]["size"]["override"] is True


def test_mcp_read_lint_unknown_rule_is_invalid_args(fake_powerpoint: Any) -> None:
    from mcp.server.fastmcp.exceptions import ToolError

    from pptlive.mcp.server import ppt_read

    with pytest.raises(ToolError, match="invalid_args"):
        ppt_read("lint", rules=["nope"])


def test_mcp_edit_regularize(fake_powerpoint: Any) -> None:
    from pptlive.mcp.server import ppt_edit

    out = ppt_edit("regularize", rules={"exclude": ["fonts"]}, dry_run=True)
    assert out["dry_run"] is True
    assert set(out) >= {"applied", "skipped", "deferred", "findings"}
    live = ppt_edit("regularize", profile={"rules": {"edge-alignment": {"enabled": True}}})
    assert live["dry_run"] is False

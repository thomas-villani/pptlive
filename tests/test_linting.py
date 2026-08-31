"""The linter / regularizer (`_linting`, `_format_info`, `_lint_profile`) against
the fake COM — the P2 peer-mode and P3 geometry rules, the fix loop, idempotency,
scoping, profiles, and the `format_info()` cascade baseline.

The deck is built here rather than in `conftest` because the rules need *peers*:
several Title-and-Content slides whose layout placeholders carry the cascade
values the live spike pinned (title 44 pt; body 28/24/20/18/18 by indent level),
plus one deliberately drifted slide.
"""

from __future__ import annotations

from typing import Any

import conftest as cf
import pytest

import pptlive
from pptlive import _linting
from pptlive._lint_profile import Profile
from pptlive._linting import dominant
from pptlive.exceptions import BatchOpError

TITLE_BOX = {"left": 66.0, "top": 30.0, "width": 828.0, "height": 100.0}
BODY_BOX = {"left": 66.0, "top": 143.75, "width": 828.0, "height": 342.625}


def _layout_placeholders() -> list[cf._FakeShape]:
    title = cf._FakeShape(
        name="Title 1",
        shape_id=1,
        shape_type=cf._MSO_PLACEHOLDER,
        text="Click to edit Master title style",
        placeholder_type=cf._PH_TITLE,
        **TITLE_BOX,
    )
    title.TextFrame.TextRange.Font.Size = 44.0
    title.TextFrame.TextRange.Font.Name = "Aptos Display"
    body = cf._FakeShape(
        name="Content Placeholder 2",
        shape_id=2,
        shape_type=cf._MSO_PLACEHOLDER,
        text="L1\rL2\rL3\rL4\rL5",
        placeholder_type=cf._PH_OBJECT,
        **BODY_BOX,
    )
    for i, size in enumerate((28.0, 24.0, 20.0, 18.0, 18.0)):
        para = body.TextFrame._paras[i]
        para.IndentLevel = i + 1
        para.Font.Size = size
        para.Font.Name = "Aptos"
    return [title, body]


def _content_slide(
    slide_id: int,
    title: str,
    body: str,
    *,
    title_size: float = 44.0,
    title_box: dict[str, float] | None = None,
    extra: list[cf._FakeShape] | None = None,
) -> cf._FakeSlide:
    t = cf._FakeShape(
        name="Title 1",
        shape_id=2,
        shape_type=cf._MSO_PLACEHOLDER,
        text=title,
        placeholder_type=cf._PH_TITLE,
        **(title_box or TITLE_BOX),
    )
    t.TextFrame.TextRange.Font.Size = title_size
    t.TextFrame.TextRange.Font.Name = "Aptos Display"
    b = cf._FakeShape(
        name="Content Placeholder 2",
        shape_id=3,
        shape_type=cf._MSO_PLACEHOLDER,
        text=body,
        placeholder_type=cf._PH_OBJECT,
        **BODY_BOX,
    )
    b.TextFrame.TextRange.Font.Size = 28.0
    b.TextFrame.TextRange.Font.Name = "Aptos"
    return cf._FakeSlide(
        slide_id=slide_id,
        layout_name="Title and Content",
        shapes=[t, b, *(extra or [])],
        layout_placeholders=_layout_placeholders(),
    )


def _lint_deck() -> cf._FakePresentation:
    s1 = _content_slide(256, "Agenda", "Intro\rDemo")
    # A level-2 bullet on slide 1 (for the per-level baseline).
    s1.Shapes._shapes[1].TextFrame._paras[1].IndentLevel = 2
    s1.Shapes._shapes[1].TextFrame._paras[1].Font.Size = 24.0
    s2 = _content_slide(257, "Results", "Revenue\rChurn")
    off = cf._FakeShape(
        name="Runaway 9",
        shape_id=9,
        shape_type=cf._MSO_TEXT_BOX,
        text="bleeds",
        left=900.0,
        top=100.0,
        width=200.0,
        height=40.0,
    )
    # The drifted slide: shrunk title, moved title box, one small body bullet.
    s3 = _content_slide(
        258,
        "Outlook",
        "Plan\rRisks",
        title_size=36.0,
        title_box={**TITLE_BOX, "top": 45.0},
        extra=[off],
    )
    s3.Shapes._shapes[1].TextFrame._paras[1].Font.Size = 20.0
    s4 = _content_slide(259, "Q&A", "Questions")
    # A mixed-run title (bold reads the msoTriStateMixed sentinel).
    s5 = _content_slide(260, "Mixed", "x")
    s5.Shapes._shapes[0].TextFrame._paras[0].Font.Bold = -2
    # A blank slide with two nearly-aligned text boxes that also overlap.
    a = cf._FakeShape(
        name="Card A",
        shape_id=2,
        shape_type=cf._MSO_TEXT_BOX,
        text="A",
        left=100.0,
        top=100.0,
        width=200.0,
        height=80.0,
    )
    b = cf._FakeShape(
        name="Card B",
        shape_id=3,
        shape_type=cf._MSO_TEXT_BOX,
        text="B",
        left=102.0,
        top=150.0,
        width=200.0,
        height=80.0,
    )
    s6 = cf._FakeSlide(slide_id=261, layout_name="Blank", shapes=[a, b])
    return cf._FakePresentation(
        name="Lint.pptx",
        full_name=r"C:\decks\Lint.pptx",
        slides=[s1, s2, s3, s4, s5, s6],
    )


@pytest.fixture
def lint_app(monkeypatch: pytest.MonkeyPatch) -> cf._FakeApplication:
    app = cf._FakeApplication([_lint_deck()])
    from pptlive import _com

    monkeypatch.setattr(_com, "get_active_powerpoint", lambda: app)
    monkeypatch.setattr(_com, "launch_powerpoint", lambda: app)
    return app


@pytest.fixture
def lint_deck(lint_app: cf._FakeApplication):  # type: ignore[no-untyped-def]
    with pptlive.attach() as ppt:
        yield ppt.presentations.active


def _by_rule(findings: list[dict[str, Any]], rule: str) -> list[dict[str, Any]]:
    return [f for f in findings if f["rule"] == rule]


# ---------------------------------------------------------------------------
# dominant() — the peer-mode primitive
# ---------------------------------------------------------------------------


def test_dominant_needs_a_clear_majority() -> None:
    assert dominant([44, 44, 36]) == (44, 2, 3)
    assert dominant([44, 36]) is None  # 50/50: no dominant
    assert dominant([44, None, 44, 36, None]) == (44, 2, 3)  # None peers ignored
    assert dominant([44]) is None  # a lone peer has nothing to be consistent with
    assert dominant([1, 1, 2, 2, 3], threshold=0.4) == (1, 2, 5)


# ---------------------------------------------------------------------------
# format_info() — the probe + cascade baseline
# ---------------------------------------------------------------------------


def test_format_info_reports_layout_baseline_and_override(lint_deck: Any) -> None:
    clean = lint_deck.anchor_by_id("ph:1:title").format_info()
    assert clean["placeholder"] == "title"
    assert clean["cascade"] == "layout"
    assert clean["font"]["size"] == {"value": 44.0, "baseline": 44.0, "override": False}
    assert clean["font"]["name"]["baseline"] == "Aptos Display"
    assert clean["mixed"] == []

    drifted = lint_deck.anchor_by_id("ph:3:title").format_info()
    assert drifted["font"]["size"] == {"value": 36.0, "baseline": 44.0, "override": True}
    assert drifted["paragraph"]["indent_level"]["value"] == 1


def test_format_info_body_baseline_is_per_indent_level(lint_deck: Any) -> None:
    level2 = lint_deck.anchor_by_id("para:1:2:2").format_info()
    assert level2["paragraph"]["indent_level"]["value"] == 2
    assert level2["font"]["size"] == {"value": 24.0, "baseline": 24.0, "override": False}
    level1 = lint_deck.anchor_by_id("para:1:2:1").format_info()
    assert level1["font"]["size"]["baseline"] == 28.0


def test_format_info_free_textbox_has_no_cascade(lint_deck: Any) -> None:
    info = lint_deck.anchor_by_id("shapeid:6:2").format_info()
    assert info["placeholder"] is None
    assert info["cascade"] is None
    assert info["font"]["size"] == {"value": 18.0, "baseline": None, "override": None}


def test_format_info_flags_mixed_runs(lint_deck: Any) -> None:
    info = lint_deck.anchor_by_id("ph:5:title").format_info()
    assert info["mixed"] == ["bold"]
    assert info["font"]["bold"]["override"] is None


def test_format_info_falls_back_to_master_without_layout_placeholder(deck: Any) -> None:
    # The default fixture's slide 1 (Title Slide) has no layout placeholders, so
    # the cascade falls through to the master text style.
    info = deck.anchor_by_id("ph:1:title").format_info()
    assert info["cascade"] == "master:title"
    assert info["font"]["size"]["baseline"] == 18.0  # the fake master's default


# ---------------------------------------------------------------------------
# lint — the default set
# ---------------------------------------------------------------------------


def test_lint_default_set_finds_the_drifted_slide(lint_deck: Any) -> None:
    findings = lint_deck.lint()
    rules = {f["rule"] for f in findings}
    assert rules == {
        "title-font-consistent",
        "body-font-consistent",
        "title-position-consistent",
        "shape-off-slide",
        "mixed-runs-in-title",
    }
    # Every finding names its slide + a drift-proof host handle.
    assert all(f["slide"] and str(f["shapeid"]).startswith("shapeid:") for f in findings)

    (title,) = _by_rule(findings, "title-font-consistent")
    assert title["slide"] == 3
    assert title["anchor_id"] == "shapeid:3:2"
    assert title["fixable"] is True
    assert title["fix"] == {"op": "format", "anchor_id": "shapeid:3:2", "size": 44.0}
    assert title["observed"].startswith("Aptos Display 36 pt")
    assert title["expected"].startswith("Aptos Display 44 pt")
    assert "3 of 3 peers" in title["message"] or "3 of 4 peers" in title["message"]

    (body,) = _by_rule(findings, "body-font-consistent")
    assert body["anchor_id"] == "para:3:2:2"
    assert body["shapeid"] == "shapeid:3:3"
    assert body["fix"] == {"op": "format", "anchor_id": "para:3:2:2", "size": 28.0}

    (pos,) = _by_rule(findings, "title-position-consistent")
    assert pos["anchor_id"] == "shapeid:3:2"
    assert pos["fix"] == [{"op": "shape_move", "anchor_id": "shapeid:3:2", "top": 30.0}]

    (off,) = _by_rule(findings, "shape-off-slide")
    assert off["anchor_id"] == "shapeid:3:9"
    assert off["fixable"] is False
    assert "960×540" in off["message"]

    (mixed,) = _by_rule(findings, "mixed-runs-in-title")
    assert mixed["slide"] == 5 and mixed["fixable"] is False


def test_lint_is_severity_ranked(lint_deck: Any) -> None:
    sev = [f["severity"] for f in lint_deck.lint()]
    assert sev == sorted(sev, key=lambda s: {"error": 0, "warning": 1, "info": 2}[s])


def test_lint_is_a_pure_read(lint_app: cf._FakeApplication, lint_deck: Any) -> None:
    before = lint_app._undo_entries
    lint_deck.lint()
    assert lint_app._undo_entries == before
    assert lint_deck.slides[3].placeholder("title").geometry()["top"] == 45.0


def test_lint_opt_in_alignment_cluster(lint_deck: Any) -> None:
    findings = lint_deck.lint(rules=["alignment"])
    rules = {f["rule"] for f in findings}
    assert {"edge-alignment", "overlap-unintended", "placeholder-off-layout"} <= rules
    assert "body-font-consistent" not in rules  # an explicit tag list picks only its set

    edges = [f for f in _by_rule(findings, "edge-alignment") if f["slide"] == 6]
    assert sorted(f["fix"]["how"] for f in edges) == ["left", "right"]  # both edges drift
    (edge,) = [f for f in edges if f["fix"]["how"] == "left"]
    assert edge["fix"] == {
        "op": "shape_align",
        "anchors": ["shapeid:6:2", "shapeid:6:3"],
        "how": "left",
        "relative_to": "selection",
    }
    assert edge["expected"] == "left: 100"

    overlaps = _by_rule(findings, "overlap-unintended")
    # Slide 6's cards, plus slide 3 where the moved-down title now clips its body.
    assert sorted(f["slide"] for f in overlaps) == [3, 6]
    assert all(f["fixable"] is False for f in overlaps)

    (layout,) = _by_rule(findings, "placeholder-off-layout")
    assert layout["anchor_id"] == "shapeid:3:2"
    assert layout["fix"] == {"op": "shape_reset_layout", "anchor_id": "shapeid:3:2"}


def test_lint_edge_alignment_tolerance_from_profile(lint_deck: Any) -> None:
    tight = lint_deck.lint(
        rules=["edge-alignment"], profile={"rules": {"edge-alignment": {"tolerance": 1.0}}}
    )
    assert not [f for f in tight if f["slide"] == 6]  # 2 pt apart > 1 pt tolerance


def test_lint_exclude_and_unknown_selector(lint_deck: Any) -> None:
    findings = lint_deck.lint(rules={"exclude": ["titles"]})
    assert {f["rule"] for f in findings} == {"body-font-consistent", "shape-off-slide"}
    with pytest.raises(ValueError, match="unknown lint rule"):
        lint_deck.lint(rules=["no-such-rule"])


def test_lint_within_scopes_emission_not_peers(lint_deck: Any) -> None:
    only3 = lint_deck.lint(within="slide:3")
    assert only3 and all(f["slide"] == 3 for f in only3)
    # The peer compare still ran deck-wide: slide 3's title is judged against 1/2/4.
    assert _by_rule(only3, "title-font-consistent")
    clean = lint_deck.lint(within=lint_deck.slides[1])
    assert clean == []
    one_shape = lint_deck.lint(within="shapeid:3:9")
    assert [f["rule"] for f in one_shape] == ["shape-off-slide"]


# ---------------------------------------------------------------------------
# profiles
# ---------------------------------------------------------------------------


def test_profile_disables_and_reranks(lint_deck: Any) -> None:
    profile = {
        "rules": {
            "title-font-consistent": {"enabled": False},
            "shape-off-slide": {"severity": "error"},
            "edge-alignment": {"enabled": True},
        }
    }
    findings = lint_deck.lint(profile=profile)
    rules = {f["rule"] for f in findings}
    assert "title-font-consistent" not in rules
    assert "edge-alignment" in rules  # a profile enable unions an off-by-default rule in
    assert findings[0]["rule"] == "shape-off-slide" and findings[0]["severity"] == "error"


def test_profile_load_variants(tmp_path: Any) -> None:
    assert Profile.load(None) == Profile()
    p = tmp_path / "pptlive.lint.json"
    p.write_text('{"rules": {"edge-alignment": {"tolerance": 2}}, "house_style": {"title": {}}}')
    loaded = Profile.load(p)
    assert loaded.config_for("edge-alignment") == {"tolerance": 2}
    assert loaded.is_enabled("edge-alignment") is True
    assert loaded.is_enabled("other") is None
    assert loaded.house_style == {"title": {}}
    bad = tmp_path / "bad.json"
    bad.write_text("[1, 2]")
    with pytest.raises(ValueError, match="JSON object"):
        Profile.load(bad)
    with pytest.raises(ValueError, match="'rules' must be an object"):
        Profile.load({"rules": 3})


# ---------------------------------------------------------------------------
# regularize — the fix loop
# ---------------------------------------------------------------------------


def test_regularize_applies_fixes_in_one_undo_entry_and_is_idempotent(
    lint_app: cf._FakeApplication, lint_deck: Any
) -> None:
    before = lint_app._undo_entries
    report = lint_deck.regularize()
    assert lint_app._undo_entries == before + 1  # one deck.edit fence for the whole pass
    assert {f["rule"] for f in report["applied"]} == {
        "title-font-consistent",
        "body-font-consistent",
        "title-position-consistent",
    }
    assert {f["rule"] for f in report["skipped"]} == {"shape-off-slide", "mixed-runs-in-title"}
    assert report["deferred"] == []
    assert report["ops_run"] == 3
    assert report["dry_run"] is False

    # The drifted slide now matches its peers.
    title = lint_deck.slides[3].placeholder("title")
    assert title.format_info()["font"]["size"]["value"] == 44.0
    assert title.geometry()["top"] == 30.0
    assert lint_deck.anchor_by_id("para:3:2:2").format_info()["font"]["size"]["value"] == 28.0

    # Idempotent: a second pass applies nothing.
    again = lint_deck.regularize()
    assert again["applied"] == [] and "ops_run" not in again
    assert {f["rule"] for f in again["findings"]} == {"shape-off-slide", "mixed-runs-in-title"}


def test_regularize_dry_run_writes_nothing(lint_app: cf._FakeApplication, lint_deck: Any) -> None:
    before = lint_app._undo_entries
    report = lint_deck.regularize(dry_run=True)
    assert report["dry_run"] is True and report["applied"] == []
    assert [f["rule"] for f in report["findings"] if f["fixable"]]
    assert lint_app._undo_entries == before
    assert lint_deck.slides[3].placeholder("title").format_info()["font"]["size"]["value"] == 36.0


def test_regularize_alignment_cluster_snaps_edges(lint_deck: Any) -> None:
    report = lint_deck.regularize(rules=["edge-alignment"], within="slide:6")
    assert [f["rule"] for f in report["applied"]] == ["edge-alignment"] * 2
    lefts = [s["geometry"]["left"] for s in lint_deck.slides[6].shapes.list()]
    assert lefts == [100.0, 100.0]
    assert lint_deck.lint(rules=["edge-alignment"], within="slide:6") == []


def test_regularize_surfaces_a_failing_fix(lint_deck: Any, monkeypatch: pytest.MonkeyPatch) -> None:
    broken = _linting.Finding(
        rule="title-font-consistent",
        kind="consistency",
        severity="warning",
        slide=3,
        anchor_id="shapeid:3:999",
        message="x",
        fixable=True,
        fix={"op": "format", "anchor_id": "shapeid:3:999", "size": 44},
    )
    monkeypatch.setattr(_linting, "run_lint", lambda *a, **k: [broken])
    with pytest.raises(BatchOpError, match="shapeid:3:999"):
        lint_deck.regularize()


def test_regularize_content_gate_defers(lint_deck: Any, monkeypatch: pytest.MonkeyPatch) -> None:
    content = _linting.Finding(
        rule="fake-content-rule",
        kind="structural",
        severity="info",
        slide=1,
        anchor_id="shapeid:1:2",
        message="x",
        fixable=True,
        fix={"op": "format", "anchor_id": "shapeid:1:2", "size": 40},
        adds_content=True,
    )
    monkeypatch.setattr(_linting, "run_lint", lambda *a, **k: [content])
    held = lint_deck.regularize()
    assert held["applied"] == [] and [f["rule"] for f in held["deferred"]] == ["fake-content-rule"]
    assert lint_deck.slides[1].placeholder("title").format_info()["font"]["size"]["value"] == 44.0
    opted = lint_deck.regularize(allow_content=True)
    assert [f["rule"] for f in opted["applied"]] == ["fake-content-rule"]
    assert lint_deck.slides[1].placeholder("title").format_info()["font"]["size"]["value"] == 40.0


def test_rules_catalog_marks_default_membership() -> None:
    catalog = {r["id"]: r for r in _linting.rules_catalog()}
    assert catalog["title-font-consistent"]["default_on"] is True
    assert catalog["edge-alignment"]["default_on"] is False
    assert "alignment" in catalog["edge-alignment"]["tags"]

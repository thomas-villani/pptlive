"""Save / save-as / PDF-export (v1.1) — the explicit file-output verbs.

Against the fake, `_FakePresentation` mirrors the 2026-06-09 COM spike: `Path` is
empty for a never-saved deck (bare `FullName`), `SaveAs(path, 24)` writes a pptx
stub and *rebinds* the working file, `SaveAs(path, 32)` writes a PDF stub but
leaves `FullName`/`Path`/`Saved` untouched (a pure export), and `Save()` clears
the dirty flag. So the rebind-vs-export distinction, the never-saved guard, and
the overwrite refusal are all provable without a real PowerPoint. The default
fake deck is saved at `C:\\decks\\Pitch.pptx`.
"""

from __future__ import annotations

import json
from typing import Any

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.exceptions import UnsavedPresentationError

# -- library: saved flag ----------------------------------------------------


def test_saved_reads_the_dirty_flag(deck: Any) -> None:
    assert deck.saved is True  # a freshly-opened deck is clean
    deck.com.Saved = 0  # msoFalse — an edit dirtied it
    assert deck.saved is False


def test_list_includes_saved_and_path(ppt: Any) -> None:
    row = ppt.presentations.list()[0]
    assert row["saved"] is True
    assert row["path"].endswith("Pitch.pptx")


# -- library: save() --------------------------------------------------------


def test_save_persists_and_clears_dirty(deck: Any) -> None:
    deck.com.Saved = 0  # dirty
    path = deck.save()
    assert path.endswith("Pitch.pptx")
    assert deck.saved is True


def test_save_on_never_saved_deck_raises(deck: Any) -> None:
    # A never-saved deck has a bare FullName -> empty Path; save() must refuse
    # rather than let PowerPoint silently route it to the default cloud folder.
    deck.com.FullName = "Presentation1"
    assert deck.com.Path == ""
    with pytest.raises(UnsavedPresentationError):
        deck.save()


# -- library: save_as() -----------------------------------------------------


def test_save_as_writes_pptx_and_rebinds(deck: Any, tmp_path: Any) -> None:
    target = tmp_path / "copy.pptx"
    written = deck.save_as(target)
    assert written == str(target.resolve())
    assert target.read_bytes().startswith(b"PK\x03\x04")  # zip == pptx
    # SaveAs rebinds: the open deck IS the new file now.
    assert deck.path == str(target.resolve())
    assert deck.saved is True


def test_save_as_refuses_to_clobber_then_allows_with_overwrite(deck: Any, tmp_path: Any) -> None:
    target = tmp_path / "existing.pptx"
    target.write_bytes(b"old")
    with pytest.raises(FileExistsError):
        deck.save_as(target)
    assert target.read_bytes() == b"old"  # untouched
    written = deck.save_as(target, overwrite=True)
    assert written == str(target.resolve())
    assert target.read_bytes().startswith(b"PK\x03\x04")


def test_save_as_rejects_pdf_format(deck: Any, tmp_path: Any) -> None:
    with pytest.raises(ValueError, match="export_pdf"):
        deck.save_as(tmp_path / "x.pdf", fmt="pdf")


def test_save_as_rejects_unknown_format(deck: Any, tmp_path: Any) -> None:
    with pytest.raises(ValueError, match="unsupported save format"):
        deck.save_as(tmp_path / "x.odp", fmt="odp")


# -- library: export_pdf() --------------------------------------------------


def test_export_pdf_writes_pdf_without_touching_the_working_file(deck: Any, tmp_path: Any) -> None:
    deck.com.Saved = 0  # dirty working file
    original_path = deck.path
    out = tmp_path / "deck.pdf"
    written = deck.export_pdf(out)
    assert written == str(out.resolve())
    assert out.read_bytes().startswith(b"%PDF-")
    # A read: no rebind, dirty flag preserved (the .pptx is untouched).
    assert deck.path == original_path
    assert deck.saved is False


def test_export_pdf_absolutizes_a_relative_path(deck: Any, tmp_path: Any, monkeypatch: Any) -> None:
    monkeypatch.chdir(tmp_path)
    written = deck.export_pdf("rel.pdf")
    assert written == str((tmp_path / "rel.pdf").resolve())
    assert (tmp_path / "rel.pdf").read_bytes().startswith(b"%PDF-")


# -- CLI --------------------------------------------------------------------


def test_cli_save(fake_powerpoint: Any) -> None:
    fake_powerpoint.ActivePresentation.Saved = 0
    res = CliRunner().invoke(main, ["--json", "save"])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["saved"] is True
    assert payload["path"].endswith("Pitch.pptx")
    assert fake_powerpoint.ActivePresentation.Saved == -1  # cleared


def test_cli_save_never_saved_is_exit_1(fake_powerpoint: Any) -> None:
    fake_powerpoint.ActivePresentation.FullName = "Presentation1"
    res = CliRunner().invoke(main, ["save"])
    assert res.exit_code == 1
    assert "never been saved" in res.output + str(res.exception or "")


def test_cli_save_as_writes_and_reports_path(fake_powerpoint: Any, tmp_path: Any) -> None:
    out = tmp_path / "out.pptx"
    res = CliRunner().invoke(main, ["--json", "save-as", str(out)])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["format"] == "pptx"
    assert payload["path"] == str(out.resolve())
    assert out.read_bytes().startswith(b"PK\x03\x04")


def test_cli_save_as_refuses_overwrite_then_allows(fake_powerpoint: Any, tmp_path: Any) -> None:
    out = tmp_path / "out.pptx"
    out.write_bytes(b"old")
    res = CliRunner().invoke(main, ["save-as", str(out)])
    assert res.exit_code == 1
    assert "overwrite" in res.output + str(res.exception or "")
    assert out.read_bytes() == b"old"
    res2 = CliRunner().invoke(main, ["--json", "save-as", str(out), "--overwrite"])
    assert res2.exit_code == 0
    assert out.read_bytes().startswith(b"PK\x03\x04")


def test_cli_export_pdf(fake_powerpoint: Any, tmp_path: Any) -> None:
    out = tmp_path / "deck.pdf"
    res = CliRunner().invoke(main, ["--json", "export-pdf", str(out)])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["path"] == str(out.resolve())
    assert out.read_bytes().startswith(b"%PDF-")


def test_cli_status_flags_unsaved(fake_powerpoint: Any) -> None:
    fake_powerpoint.ActivePresentation.Saved = 0
    res = CliRunner().invoke(main, ["--text", "status"])
    assert res.exit_code == 0
    assert "(unsaved)" in res.output

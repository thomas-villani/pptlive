"""Deck snapshot (v1.1): `Presentation.snapshot` — the token-aware vision read.

Against the fake, `Slide.Export` writes a 24-byte stub PNG that encodes its pixel
size (so dims round-trip) and records the requested size. So we can prove the
`max_dim` long-edge cap math, slide selection, the per-slide byte return, file
placement (single vs `<stem>-sN<suffix>`), and the CLI's path-vs-base64 emit —
all without a real PowerPoint. The default fake deck is three 960x540 pt slides
(native 1280x720 px at 96 DPI).
"""

from __future__ import annotations

import base64
import json

import pytest
from click.testing import CliRunner

from pptlive import _snapshot
from pptlive.cli.main import main
from pptlive.exceptions import SlideNotFoundError

_PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _png_dims(data: bytes) -> tuple[int, int]:
    """Recover (width, height) from the stub PNG's IHDR the fake writes."""
    return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")


def _last_export(fake, slide_index):  # type: ignore[no-untyped-def]
    return fake.ActivePresentation.Slides(slide_index).last_export


# -- _capped_dims (pure math, no COM) ---------------------------------------


def test_capped_dims_caps_long_edge_and_keeps_aspect() -> None:
    # 960x540 (16:9) capped to 1000 on the long edge -> width 1000, height scaled.
    assert _snapshot._capped_dims(960.0, 540.0, 1000) == (1000, 562)


def test_capped_dims_none_means_native() -> None:
    assert _snapshot._capped_dims(960.0, 540.0, None) is None


def test_capped_dims_never_upscales_past_native() -> None:
    # A max_dim far above native (1280 px long edge) clamps to native, not up.
    assert _snapshot._capped_dims(960.0, 540.0, 99999) == (1280, 720)


def test_capped_dims_portrait_caps_the_taller_edge() -> None:
    # Long edge is the height here, so it's the one pinned to 1000.
    assert _snapshot._capped_dims(540.0, 960.0, 1000) == (562, 1000)


# -- snapshot() selection ---------------------------------------------------


def test_snapshot_all_slides_returns_one_per_slide(deck) -> None:  # type: ignore[no-untyped-def]
    snaps = deck.snapshot()
    assert [s.slide for s in snaps] == [1, 2, 3]
    assert all(s.png.startswith(_PNG_SIG) for s in snaps)
    assert all(s.path is None for s in snaps)  # no out -> bytes only


def test_snapshot_single_slide(deck) -> None:  # type: ignore[no-untyped-def]
    snaps = deck.snapshot(slides=2)
    assert [s.slide for s in snaps] == [2]


def test_snapshot_range_inclusive(deck) -> None:  # type: ignore[no-untyped-def]
    snaps = deck.snapshot(slides=(1, 2))
    assert [s.slide for s in snaps] == [1, 2]


def test_snapshot_out_of_range_raises(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SlideNotFoundError):
        deck.snapshot(slides=9)
    with pytest.raises(SlideNotFoundError):
        deck.snapshot(slides=(2, 9))


# -- max_dim cap ------------------------------------------------------------


def test_snapshot_max_dim_caps_the_export(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    snaps = deck.snapshot(slides=1, max_dim=1000)
    # the cap reaches Slide.Export as the requested pixel size...
    rec = _last_export(fake_powerpoint, 1)
    assert (rec["Width"], rec["Height"]) == (1000, 562)
    # ...and the returned bytes carry those dims.
    assert _png_dims(snaps[0].png) == (1000, 562)


def test_snapshot_no_max_dim_is_native(deck, fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    deck.snapshot(slides=1)
    rec = _last_export(fake_powerpoint, 1)
    assert (rec["Width"], rec["Height"]) == (1280, 720)


# -- file placement ---------------------------------------------------------


def test_snapshot_out_single_writes_to_that_path(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "one.png"
    snaps = deck.snapshot(out, slides=2)
    assert snaps[0].path == out
    assert out.is_file()
    assert out.read_bytes() == snaps[0].png


def test_snapshot_out_multiple_writes_stem_sN(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "deck.png"
    snaps = deck.snapshot(out)
    names = [s.path.name for s in snaps]
    assert names == ["deck-s1.png", "deck-s2.png", "deck-s3.png"]
    assert all(s.path.is_file() for s in snaps)


def test_snapshot_jpg_format(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    snaps = deck.snapshot(tmp_path / "d.jpg", slides=1, fmt="jpg")
    assert snaps[0].path.suffix == ".jpg"
    assert _last_export(fake_powerpoint, 1)["FilterName"] == "JPG"


def test_snapshot_unknown_format_raises_before_com(deck) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown image format"):
        deck.snapshot(slides=1, fmt="webp")


# -- CLI --------------------------------------------------------------------


def test_cli_snapshot_base64_without_out(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "snapshot"])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["count"] == 3
    assert payload["selector"] == "all slides"
    assert [img["slide"] for img in payload["images"]] == [1, 2, 3]
    # no --out -> base64 inline, no path
    first = payload["images"][0]
    assert "path" not in first
    assert base64.b64decode(first["base64"]).startswith(_PNG_SIG)


def test_cli_snapshot_writes_files_with_out(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "deck.png"
    res = CliRunner().invoke(main, ["--json", "snapshot", "--out", str(out)])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert [img["path"] for img in payload["images"]] == [
        str(tmp_path / "deck-s1.png"),
        str(tmp_path / "deck-s2.png"),
        str(tmp_path / "deck-s3.png"),
    ]
    assert all("base64" not in img for img in payload["images"])


def test_cli_snapshot_single_slide(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "snapshot", "--slide", "2"])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["count"] == 1
    assert payload["selector"] == "slide 2"
    assert payload["images"][0]["slide"] == 2


def test_cli_snapshot_range_and_max_dim(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["--json", "snapshot", "--slides", "1-2", "--max-dim", "1000"])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["count"] == 2
    assert payload["max_dim"] == 1000
    assert _png_dims(base64.b64decode(payload["images"][0]["base64"])) == (1000, 562)


def test_cli_snapshot_rejects_both_selectors(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["snapshot", "--slide", "1", "--slides", "1-2"])
    assert res.exit_code == 1
    assert "at most one" in res.output + str(res.exception or "")


def test_cli_snapshot_rejects_malformed_range(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["snapshot", "--slides", "oops"])
    assert res.exit_code == 1


def test_cli_snapshot_out_of_range_is_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    res = CliRunner().invoke(main, ["snapshot", "--slide", "9"])
    assert res.exit_code == 2  # SlideNotFoundError -> anchor-not-found

"""Pictures (v0.7): alt text as a re-identification handle + per-shape export.

Against the fake, `Shape.Export` writes a stub PNG and records the
PpShapeFormat int + pixel size, and `AlternativeText` is a plain settable
attribute — enough to prove alt-text round-trips (wrapper + listings + CLI +
add_picture), the export filter/dimension logic, and that both stay polite.
"""

from __future__ import annotations

import json
import os

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main

_PNG_SIG = b"\x89PNG\r\n\x1a\n"
_SHAPE_PNG = 2  # PpShapeFormat.PNG
_SHAPE_JPG = 1  # PpShapeFormat.JPG


def _json(result):  # type: ignore[no-untyped-def]
    return json.loads(result.output)


def _picture(deck):  # type: ignore[no-untyped-def]
    # slide 2, z-order 3 is "Picture 3" (300x200 pt, no text frame).
    return deck.slides[2].shapes[3]


def _last_export(fake, slide_index, z):  # type: ignore[no-untyped-def]
    return fake.ActivePresentation.Slides(slide_index).Shapes(z).last_export


# -- alt text (wrapper) -----------------------------------------------------


def test_alt_text_defaults_empty(deck) -> None:  # type: ignore[no-untyped-def]
    assert _picture(deck).alt_text == ""


def test_set_alt_text_round_trips(deck) -> None:  # type: ignore[no-untyped-def]
    pic = _picture(deck)
    pic.set_alt_text("Quarterly revenue chart")
    assert pic.alt_text == "Quarterly revenue chart"


def test_alt_text_appears_in_listing(deck) -> None:  # type: ignore[no-untyped-def]
    _picture(deck).set_alt_text("Logo")
    rows = deck.slides[2].shapes.list()
    assert rows[2]["alt_text"] == "Logo"
    assert rows[0]["alt_text"] == ""  # untagged shapes carry an empty string


def test_add_picture_with_alt_text(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    img = tmp_path / "logo.png"
    img.write_bytes(_PNG_SIG)
    pic = deck.slides[3].shapes.add_picture(img, alt_text="Company logo")
    assert pic.alt_text == "Company logo"


# -- per-shape export (wrapper) ---------------------------------------------


def test_export_returns_path_and_writes_png(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "pic.png"
    path = _picture(deck).export_image(out)
    assert path == out
    assert path.read_bytes().startswith(_PNG_SIG)


def test_export_default_dims_are_native_pixels(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    # 300x200 pt shape -> native 400x267 px (96 DPI), filter = PpShapeFormat.PNG.
    _picture(deck).export_image(tmp_path / "p.png")
    rec = _last_export(fake_powerpoint, 2, 3)
    assert (rec["Width"], rec["Height"]) == (400, 267)
    assert rec["Filter"] == _SHAPE_PNG


def test_export_jpg_filter(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    path = _picture(deck).export_image(tmp_path / "p.jpg", fmt="jpg")
    assert path.suffix == ".jpg"
    assert _last_export(fake_powerpoint, 2, 3)["Filter"] == _SHAPE_JPG


def test_export_unknown_format_raises(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown image format"):
        _picture(deck).export_image(tmp_path / "p.tiff", fmt="tiff")  # no TIFF for shapes


def test_export_relative_path_resolved_absolute(deck, monkeypatch, tmp_path) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    path = _picture(deck).export_image("rel.png")
    assert path.is_absolute()
    assert path == tmp_path / "rel.png"


def test_export_temp_default_path(deck) -> None:  # type: ignore[no-untyped-def]
    path = _picture(deck).export_image()  # no out: a temp file
    try:
        assert path.is_file()
        assert path.suffix == ".png"
    finally:
        os.remove(path)


def test_export_is_polite(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Title 1")
    before_type, before_view = fake_powerpoint._selection_type, fake_powerpoint._viewed
    _picture(deck).export_image(tmp_path / "p.png")
    assert fake_powerpoint._selection_type == before_type
    assert fake_powerpoint._viewed == before_view


# -- CLI --------------------------------------------------------------------


def test_cli_shape_set_alt(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(
        main, ["shape", "set-alt", "--anchor-id", "shape:2:3", "--alt-text", "Hero image"]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["ok"] is True
    assert payload["alt_text"] == "Hero image"


def test_cli_shape_add_picture_with_alt(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    img = tmp_path / "p.png"
    img.write_bytes(_PNG_SIG)
    result = CliRunner().invoke(
        main,
        [
            "shape",
            "add",
            "--slide",
            "3",
            "--kind",
            "picture",
            "--path",
            str(img),
            "--alt-text",
            "Diagram",
        ],
    )
    assert result.exit_code == 0
    assert _json(result)["alt_text"] == "Diagram"


def test_cli_shape_export(fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "shape.png"
    result = CliRunner().invoke(
        main, ["shape", "export", "--anchor-id", "shape:2:3", "--out", str(out)]
    )
    assert result.exit_code == 0
    payload = _json(result)
    assert payload["anchor_id"] == "shape:2:3"
    assert payload["path"] == str(out)
    assert out.read_bytes().startswith(_PNG_SIG)


def test_cli_shape_export_unknown_anchor_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    result = CliRunner().invoke(main, ["shape", "export", "--anchor-id", "shape:2:99"])
    assert result.exit_code == 2


def test_cli_set_alt_non_shape_anchor_exit_2(fake_powerpoint) -> None:  # type: ignore[no-untyped-def]
    # notes:S is a valid anchor but not a Shape -> exit 2.
    result = CliRunner().invoke(
        main, ["shape", "set-alt", "--anchor-id", "notes:2", "--alt-text", "x"]
    )
    assert result.exit_code == 2

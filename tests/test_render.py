"""Slide render (v0.4): `Slide.export_image` / `Presentation.export_images`.

Against the fake, `Slide.Export` writes a 24-byte stub PNG and records the
filter + pixel size, so we can prove path resolution, format selection, the
default/native and aspect-fill dimension logic, and that export is polite (it
goes through no `edit()` and never touches the Selection).
"""

from __future__ import annotations

import os

import pytest

_PNG_SIG = b"\x89PNG\r\n\x1a\n"


def _last_export(fake, slide_index):  # type: ignore[no-untyped-def]
    return fake.ActivePresentation.Slides(slide_index).last_export


def test_export_returns_path_and_writes_file(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    out = tmp_path / "slide2.png"
    path = deck.slides[2].export_image(out)
    assert path == out
    assert path.is_file()
    assert path.read_bytes().startswith(_PNG_SIG)


def test_export_default_dims_are_native_pixels(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    # A 960x540 pt (16:9) slide exports at its native 1280x720 px (96 DPI).
    deck.slides[2].export_image(tmp_path / "s.png")
    rec = _last_export(fake_powerpoint, 2)
    assert (rec["Width"], rec["Height"]) == (1280, 720)
    assert rec["FilterName"] == "PNG"


def test_export_requested_dims_honored(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    deck.slides[2].export_image(tmp_path / "s.png", width=800, height=600)
    rec = _last_export(fake_powerpoint, 2)
    assert (rec["Width"], rec["Height"]) == (800, 600)


def test_export_one_dim_fills_from_aspect_ratio(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    # width only on a 16:9 slide -> height follows (1920 -> 1080).
    deck.slides[2].export_image(tmp_path / "s.png", width=1920)
    rec = _last_export(fake_powerpoint, 2)
    assert (rec["Width"], rec["Height"]) == (1920, 1080)


def test_export_jpg_format(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    path = deck.slides[2].export_image(tmp_path / "s.jpg", fmt="jpg")
    assert path.suffix == ".jpg"
    assert _last_export(fake_powerpoint, 2)["FilterName"] == "JPG"


def test_export_unknown_format_raises(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(ValueError, match="unknown image format"):
        deck.slides[2].export_image(tmp_path / "s.webp", fmt="webp")


def test_export_relative_path_resolved_absolute(deck, monkeypatch, tmp_path) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    path = deck.slides[2].export_image("relative.png")
    assert path.is_absolute()
    assert path == tmp_path / "relative.png"


def test_export_temp_default_path(deck) -> None:  # type: ignore[no-untyped-def]
    path = deck.slides[2].export_image()  # no --out: a temp file
    try:
        assert path.is_file()
        assert path.suffix == ".png"
    finally:
        os.remove(path)


def test_export_is_polite(deck, fake_powerpoint, tmp_path) -> None:  # type: ignore[no-untyped-def]
    fake_powerpoint._viewed = 2
    fake_powerpoint._select_shapes("Title 1")
    before_type, before_view = fake_powerpoint._selection_type, fake_powerpoint._viewed
    deck.slides[2].export_image(tmp_path / "s.png")
    assert fake_powerpoint._selection_type == before_type
    assert fake_powerpoint._viewed == before_view


def test_export_images_whole_deck(deck, tmp_path) -> None:  # type: ignore[no-untyped-def]
    paths = deck.export_images(tmp_path)
    assert [p.name for p in paths] == ["slide-001.png", "slide-002.png", "slide-003.png"]
    assert all(p.is_file() for p in paths)

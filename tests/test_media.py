"""Media insertion + narrated-video export (v1.7).

Against the fake, `_FakeShapes.AddMediaObject2` makes a `msoMedia` shape with a
1.2 s nominal clip (movie for a video extension, else sound) and `PlaySettings`
the wrapper writes auto-play onto; `_FakePresentation.CreateVideo` /
`CreateVideoStatus` "finish" the encode by the first poll and write a stub MP4
(a path containing "fail" reports Failed). So auto-play, slide pacing, the media
read, and the async export loop are all provable without a real PowerPoint.
"""

from __future__ import annotations

import json
from typing import Any

import pytest
from click.testing import CliRunner

from pptlive.cli.main import main
from pptlive.exceptions import AnchorNotFoundError, VideoExportError


def _audio(tmp_path: Any) -> str:
    p = tmp_path / "narration.wav"
    p.write_bytes(b"RIFFfake wav")
    return str(p)


def _video(tmp_path: Any) -> str:
    p = tmp_path / "clip.mp4"
    p.write_bytes(b"\x00\x00\x00\x18ftypmp42")
    return str(p)


# -- library: insertion -----------------------------------------------------


def test_add_audio_inserts_media_with_autoplay_and_pacing(deck: Any, tmp_path: Any) -> None:
    slide = deck.slides[1]
    with deck.edit("add audio"):
        shape = slide.add_audio(_audio(tmp_path))
    assert shape.has_media is True
    media = shape.media
    assert media["type"] == "sound"
    assert media["length_s"] == 1.2
    assert media["autoplay"] is True
    # pace_slide auto-advances the slide to the clip length (1.2 s).
    tr = slide.transition()
    assert tr["advance_on_time"] is True
    assert tr["advance_time"] == pytest.approx(1.2)


def test_add_video_reads_back_as_movie(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add video"):
        shape = deck.slides[1].add_video(_video(tmp_path))
    assert shape.has_media is True
    assert shape.media["type"] == "movie"


def test_add_audio_opt_out_of_autoplay_and_pacing(deck: Any, tmp_path: Any) -> None:
    slide = deck.slides[1]
    with deck.edit("add audio quiet"):
        shape = slide.add_audio(_audio(tmp_path), autoplay=False, pace_slide=False)
    assert shape.media["autoplay"] is False
    assert slide.transition()["advance_on_time"] is False  # untouched


def test_add_media_missing_file_raises_before_com(deck: Any, tmp_path: Any) -> None:
    with pytest.raises(FileNotFoundError), deck.edit("add audio missing"):
        deck.slides[1].add_audio(str(tmp_path / "nope.wav"))


def test_shape_read_includes_media(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        deck.slides[1].add_audio(_audio(tmp_path))
    rows = deck.slides[1].shapes.list()
    media_rows = [r for r in rows if r.get("has_media")]
    assert len(media_rows) == 1
    assert media_rows[0]["media"]["type"] == "sound"


def test_media_property_on_non_media_shape_raises(deck: Any) -> None:
    # The default deck's first shape is a placeholder, not media.
    shape = deck.slides[1].shapes[1]
    assert shape.has_media is False
    with pytest.raises(AnchorNotFoundError):
        _ = shape.media


# -- library: video export --------------------------------------------------


def test_export_video_blocks_and_writes_file(deck: Any, tmp_path: Any) -> None:
    out = tmp_path / "deck.mp4"
    result = deck.export_video(out, resolution=480)
    assert result.ok is True
    assert result.status == "done"
    assert result.path == str(out.resolve())
    assert out.read_bytes().startswith(b"\x00\x00\x00\x18ftyp")


def test_export_video_no_wait_returns_immediately(deck: Any, tmp_path: Any) -> None:
    out = tmp_path / "deck.mp4"
    result = deck.export_video(out, wait=False)
    assert result.ok is False  # in-flight: caller polls video_status()
    assert result.path == str(out.resolve())


def test_export_video_failed_raises(deck: Any, tmp_path: Any) -> None:
    out = tmp_path / "will-fail.mp4"
    with pytest.raises(VideoExportError):
        deck.export_video(out)


def test_video_status_none_before_any_export(deck: Any) -> None:
    result = deck.video_status()
    assert result.status == "none"
    assert result.ok is False


def test_video_status_done_after_export(deck: Any, tmp_path: Any) -> None:
    deck.export_video(tmp_path / "deck.mp4")
    result = deck.video_status()
    assert result.status == "done"
    assert result.ok is True


# -- CLI --------------------------------------------------------------------


def test_cli_media_add_audio(fake_powerpoint: Any, tmp_path: Any) -> None:
    res = CliRunner().invoke(
        main,
        ["--json", "media", "add", "--slide", "1", "--kind", "audio", "--path", _audio(tmp_path)],
    )
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["has_media"] is True
    assert payload["media"]["type"] == "sound"


def test_cli_media_add_video(fake_powerpoint: Any, tmp_path: Any) -> None:
    res = CliRunner().invoke(
        main,
        ["--json", "media", "add", "--slide", "1", "--kind", "video", "--path", _video(tmp_path)],
    )
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["media"]["type"] == "movie"


def test_cli_export_video(fake_powerpoint: Any, tmp_path: Any) -> None:
    out = tmp_path / "out.mp4"
    res = CliRunner().invoke(main, ["--json", "export-video", str(out), "--resolution", "480"])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["status"] == "done"
    assert payload["path"] == str(out.resolve())
    assert out.read_bytes().startswith(b"\x00\x00\x00\x18ftyp")


def test_cli_video_status(fake_powerpoint: Any) -> None:
    res = CliRunner().invoke(main, ["--json", "video-status"])
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["status"] == "none"  # no export requested yet


def test_add_audio_surfaces_playsettings_failure(
    deck: Any, tmp_path: Any, monkeypatch: Any
) -> None:
    # The autoplay/hide_icon write is no longer swallowed: if PlaySettings can't be
    # reached, the failure surfaces instead of silently dropping the caller's
    # request (which would break the narrate → auto-advance → export_video flow).
    fake_shape_cls = type(deck.slides[1].shapes[1].com)

    def _boom(self: Any) -> Any:
        raise RuntimeError("PlaySettings unavailable")

    monkeypatch.setattr(fake_shape_cls, "AnimationSettings", property(_boom))
    with pytest.raises(RuntimeError), deck.edit("add audio"):
        deck.slides[1].add_audio(_audio(tmp_path))


def test_export_video_rejects_out_of_range_params(deck: Any, tmp_path: Any) -> None:
    # Validate before any COM: a bad param is a clean ValueError, not a confusing
    # raw CreateVideo COM error.
    out = tmp_path / "deck.mp4"
    with pytest.raises(ValueError, match="quality"):
        deck.export_video(out, quality=101)
    with pytest.raises(ValueError, match="resolution"):
        deck.export_video(out, resolution=0)
    with pytest.raises(ValueError, match="fps"):
        deck.export_video(out, fps=0)

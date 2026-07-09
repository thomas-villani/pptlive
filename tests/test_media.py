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


# -- library: playback long-tail (mute / volume / trim) ---------------------


def test_media_read_includes_trim_window(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    media = shape.media
    # Untrimmed: start 0 → end = clip length (1.2 s).
    assert media["start_s"] == 0.0
    assert media["end_s"] == 1.2


def test_set_media_playback_mute_and_volume_round_trip(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    with deck.edit("set playback"):
        out = shape.set_media_playback(muted=True, volume=0.25)
    assert out["muted"] is True
    assert out["volume"] == 0.25
    # Re-read confirms it persisted.
    assert shape.media["muted"] is True
    assert shape.media["volume"] == 0.25
    with deck.edit("unmute"):
        shape.set_media_playback(muted=False)
    assert shape.media["muted"] is False


def test_set_media_playback_trim_seconds_round_trip(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    with deck.edit("trim"):
        out = shape.set_media_playback(start=0.2, end=1.0)
    assert out["start_s"] == 0.2
    assert out["end_s"] == 1.0


def test_set_media_playback_partial_trim_keeps_other_edge(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    with deck.edit("trim start only"):
        out = shape.set_media_playback(start=0.3)
    assert out["start_s"] == 0.3
    assert out["end_s"] == 1.2  # unchanged (clip length)


def test_set_media_playback_volume_out_of_range_raises(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    for bad in (-0.1, 1.5):
        with pytest.raises(ValueError, match="volume"), deck.edit("bad volume"):
            shape.set_media_playback(volume=bad)


def test_set_media_playback_bad_trim_raises_before_com(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    # end <= start, end > length, start < 0 all rejected with a clean ValueError.
    with pytest.raises(ValueError, match="greater than"), deck.edit("e"):
        shape.set_media_playback(start=0.8, end=0.4)
    with pytest.raises(ValueError, match="exceeds"), deck.edit("e"):
        shape.set_media_playback(end=5.0)
    with pytest.raises(ValueError, match="start"), deck.edit("e"):
        shape.set_media_playback(start=-0.5, end=1.0)


def test_set_media_playback_requires_an_option(deck: Any, tmp_path: Any) -> None:
    with deck.edit("add audio"):
        shape = deck.slides[1].add_audio(_audio(tmp_path))
    with pytest.raises(ValueError, match="at least one"), deck.edit("noop"):
        shape.set_media_playback()


def test_set_media_playback_on_non_media_raises(deck: Any) -> None:
    shape = deck.slides[1].shapes[1]  # a placeholder
    with pytest.raises(AnchorNotFoundError), deck.edit("nope"):
        shape.set_media_playback(muted=True)


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


def test_cli_media_set(fake_powerpoint: Any, tmp_path: Any) -> None:
    runner = CliRunner()
    add = runner.invoke(
        main,
        ["--json", "media", "add", "--slide", "1", "--kind", "audio", "--path", _audio(tmp_path)],
    )
    assert add.exit_code == 0
    anchor = json.loads(add.output)["anchor_id"]
    res = runner.invoke(
        main,
        [
            "--json",
            "media",
            "set",
            "--anchor-id",
            anchor,
            "--muted",
            "--volume",
            "0.3",
            "--start",
            "0.2",
            "--end",
            "1.0",
        ],
    )
    assert res.exit_code == 0
    payload = json.loads(res.output)
    assert payload["ok"] is True
    assert payload["media"]["muted"] is True
    assert payload["media"]["volume"] == 0.3
    assert payload["media"]["start_s"] == 0.2
    assert payload["media"]["end_s"] == 1.0


def test_cli_media_set_bad_volume_exit_1(fake_powerpoint: Any, tmp_path: Any) -> None:
    runner = CliRunner()
    add = runner.invoke(
        main,
        ["--json", "media", "add", "--slide", "1", "--kind", "audio", "--path", _audio(tmp_path)],
    )
    anchor = json.loads(add.output)["anchor_id"]
    res = runner.invoke(main, ["--json", "media", "set", "--anchor-id", anchor, "--volume", "9"])
    assert res.exit_code == 1


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

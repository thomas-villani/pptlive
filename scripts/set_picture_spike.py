"""Spike — pin the COM mechanism for *re-sourcing a picture in place*.

`add_picture` exists; what's missing from the edit surface is swapping the image
of an **existing** picture shape without delete-and-recreate at the agent level
(which would force the LLM to re-derive geometry / name / alt text / z-order).

The genuine unknown: does `Shape.Fill.UserPicture(path)` actually *replace the
displayed image* of a real `msoPicture` shape, or does it merely set a fill that
sits behind the (unchanged) picture? PowerPoint's UI "Change Picture" has no
obvious single COM verb. So this probe builds a red picture, tries each candidate
mechanism to turn it green, then **exports the shape and samples the centre
pixel** — pixels are ground truth, not COM properties.

Mechanisms tried:
1. `pic.Fill.UserPicture(green)`              — the fill trick.
2. `pic.PictureFormat` replace               — probe for any in-place verb / error.
3. delete + AddPicture(green) at copied       — the fallback that always works;
   geometry/rotation/name/alt text/z-order      confirm it lands green at the same box.

Run against a *running* PowerPoint with a deck open:

    uv run python scripts/set_picture_spike.py

Net-zero and polite: everything is built on a single temporary blank slide
appended then deleted in a `finally`; the viewed slide + Selection are restored;
`net_zero_ok` confirms the deck slide count is unchanged.
"""

from __future__ import annotations

import json
import os
import struct
import tempfile
import zlib
from typing import Any

import pptlive as pl
from pptlive import _selection


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _solid_png(rgb: tuple[int, int, int], size: int = 64) -> bytes:
    """A solid-colour PNG of `size`x`size` (pure stdlib, no PIL)."""
    r, g, b = rgb
    row = b"\x00" + bytes((r, g, b)) * size  # filter byte 0 + RGB triples
    raw = row * size
    comp = zlib.compress(raw, 9)

    def _chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + tag
            + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)  # 8-bit RGB
    return (
        b"\x89PNG\r\n\x1a\n" + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", comp) + _chunk(b"IEND", b"")
    )


def _write_png(rgb: tuple[int, int, int]) -> str:
    fd, path = tempfile.mkstemp(prefix="pptlive_setpic_", suffix=".png")
    os.write(fd, _solid_png(rgb))
    os.close(fd)
    return path


def _png_info(png_path: str) -> tuple[int, int, int, int, bytes, bytes] | None:
    """`(width, height, bit_depth, color_type, idat, plte)` for a PNG, or None."""
    try:
        with open(png_path, "rb") as fh:
            data = fh.read()
    except OSError:
        return None
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        return None
    pos = 8
    width = height = bit_depth = color_type = 0
    idat = b""
    plte = b""
    while pos + 8 <= len(data):
        (length,) = struct.unpack(">I", data[pos : pos + 4])
        tag = data[pos + 4 : pos + 8]
        chunk = data[pos + 8 : pos + 8 + length]
        if tag == b"IHDR":
            width, height, bit_depth, color_type = struct.unpack(">IIBB", chunk[:10])
        elif tag == b"PLTE":
            plte = chunk
        elif tag == b"IDAT":
            idat += chunk
        elif tag == b"IEND":
            break
        pos += 12 + length
    return width, height, bit_depth, color_type, idat, plte


def _center_pixel(png_path: str) -> tuple[int, int, int] | None:
    """Decode an RGB / RGBA / palette PNG, return its centre pixel, or None.

    Handles PowerPoint's `Shape.Export` output, which comes back as a **1-bit
    palette** PNG (`bit_depth 1, color_type 3`) for a solid image — the reason the
    first cuts read `null`. Supports palette depths 1/2/4/8 and 8-bit truecolor.
    """
    info = _png_info(png_path)
    if info is None:
        return None
    width, height, bit_depth, color_type, idat, plte = info
    if color_type == 3:
        if bit_depth not in (1, 2, 4, 8):
            return None
    elif color_type in (2, 6):
        if bit_depth != 8:
            return None
    else:
        return None
    channels = {2: 3, 6: 4, 3: 1}[color_type]
    # Bytes per scanline (sub-byte palette packs `bit_depth` bits per pixel).
    if color_type == 3 and bit_depth < 8:
        stride = (width * bit_depth + 7) // 8
        bpp = 1  # filter byte-distance for <8-bit images is 1
    else:
        stride = width * channels
        bpp = channels
    try:
        raw = zlib.decompress(idat)
    except zlib.error:
        return None
    out = bytearray()
    prev = bytearray(stride)
    p = 0
    for _y in range(height):
        if p >= len(raw):
            break
        ft = raw[p]
        p += 1
        line = bytearray(raw[p : p + stride])
        p += stride
        for i in range(stride):
            a = line[i - bpp] if i >= bpp else 0
            bb = prev[i]
            c = prev[i - bpp] if i >= bpp else 0
            if ft == 1:
                line[i] = (line[i] + a) & 0xFF
            elif ft == 2:
                line[i] = (line[i] + bb) & 0xFF
            elif ft == 3:
                line[i] = (line[i] + ((a + bb) >> 1)) & 0xFF
            elif ft == 4:
                pp = a + bb - c
                pa, pbb, pc = abs(pp - a), abs(pp - bb), abs(pp - c)
                pr = a if (pa <= pbb and pa <= pc) else (bb if pbb <= pc else c)
                line[i] = (line[i] + pr) & 0xFF
        out.extend(line)
        prev = line
    cx, cy = width // 2, height // 2
    if color_type == 3:
        if bit_depth == 8:
            idx = out[cy * stride + cx]
        else:
            byte = out[cy * stride + (cx * bit_depth) // 8]
            shift = 8 - bit_depth - (cx * bit_depth) % 8
            idx = (byte >> shift) & ((1 << bit_depth) - 1)
        po = idx * 3
        if po + 2 >= len(plte):
            return None
        return (plte[po], plte[po + 1], plte[po + 2])
    base = cy * stride + cx * channels
    return (out[base], out[base + 1], out[base + 2])


def _close(px: tuple[int, int, int] | None, want: tuple[int, int, int], tol: int = 40) -> bool:
    if px is None:
        return False
    return all(abs(a - b) <= tol for a, b in zip(px, want, strict=False))


_RED = (220, 30, 30)
_GREEN = (30, 200, 60)


def probe_fill_userpicture(shapes: Any, green: str) -> dict[str, Any]:
    """Does Fill.UserPicture replace a *picture* shape's displayed image?"""
    out: dict[str, Any] = {}
    tmp: str | None = None
    try:
        red = _write_png(_RED)
        try:
            pic = shapes.add_picture(red, left=40.0, top=40.0, width=120.0, height=120.0)
            before = pic.export_image(fmt="png")
            info = _png_info(str(before))
            if info is not None:
                out["export_ihdr"] = {
                    "w": info[0],
                    "h": info[1],
                    "bit_depth": info[2],
                    "color_type": info[3],
                }
            out["center_before"] = _center_pixel(str(before))
            os.remove(before)
            pic.com.Fill.UserPicture(os.path.abspath(green))
            out["fill_type_after"] = int(pic.com.Fill.Type)
            after = pic.export_image(fmt="png")
            tmp = str(after)
            out["center_after"] = _center_pixel(tmp)
            out["image_became_green"] = _close(out["center_after"], _GREEN)
            out["still_red"] = _close(out["center_after"], _RED)
            pic.delete()
        finally:
            os.remove(red)
    except Exception as exc:
        out["error"] = _err(exc)
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)
    return out


def _resolve_by_id(shapes_com: Any, shape_id: int) -> Any | None:
    for i in range(1, int(shapes_com.Count) + 1):
        sh = shapes_com(i)
        if int(sh.Id) == int(shape_id):
            return sh
    return None


def probe_delete_readd(shapes: Any, green: str) -> dict[str, Any]:
    """The fallback that actually works: AddPicture(green) under a stack of other
    shapes, copy geometry/name/alt, delete old, re-resolve the new shape by its
    stable `Shape.Id` (z-order drifts on delete), and restack to the old slot.

    Pins the two pitfalls the real wrapper must handle: picture **aspect-lock**
    (so the copied box must be applied with `LockAspectRatio` off) and **z-order
    drift** (so the new shape is held by `Shape.Id`, never a z-order index).
    """
    out: dict[str, Any] = {}
    tmp: str | None = None
    try:
        red = _write_png(_RED)
        try:
            shapes_com = shapes._com_collection
            # A decoy on top so the picture sits mid-stack (z-order restore matters).
            shapes.add_shape("rectangle", left=600.0, top=300.0, width=40.0, height=40.0)
            old = shapes.add_picture(
                red, left=200.0, top=40.0, width=120.0, height=90.0, alt_text="logo-v1"
            )
            shapes.add_shape("oval", left=600.0, top=360.0, width=40.0, height=40.0)
            oc = old.com
            geom = {
                "Left": float(oc.Left),
                "Top": float(oc.Top),
                "Width": float(oc.Width),
                "Height": float(oc.Height),
                "Rotation": float(oc.Rotation),
            }
            name = str(oc.Name)
            alt = str(oc.AlternativeText or "")
            zpos = int(oc.ZOrderPosition)
            out["old_box"] = geom
            out["old_z"] = zpos

            # AddPicture lands on top; capture its stable Id immediately.
            new_com = shapes_com.AddPicture(
                os.path.abspath(green), 0, -1, geom["Left"], geom["Top"], -1.0, -1.0
            )
            new_id = int(new_com.Id)
            try:
                new_com.LockAspectRatio = 0  # msoFalse — let the box override aspect
            except Exception as exc:
                out["lock_set_error"] = _err(exc)
            new_com.Left, new_com.Top = geom["Left"], geom["Top"]
            new_com.Width, new_com.Height = geom["Width"], geom["Height"]
            new_com.Rotation = geom["Rotation"]
            new_com.AlternativeText = alt

            oc.Delete()
            # Re-resolve by Id — z-order indices shifted when `old` went away.
            new_com = _resolve_by_id(shapes_com, new_id)
            out["reresolved_by_id"] = new_com is not None
            new_com.Name = name  # set after delete so no name clash

            # Restore z-order: send to back, then step forward to the old slot.
            new_com.ZOrder(1)  # msoSendToBack
            for _ in range(zpos - 1):
                new_com.ZOrder(2)  # msoBringForward (3 is *send backward*)
            out["restored_z"] = int(new_com.ZOrderPosition)
            out["z_preserved"] = out["restored_z"] == zpos

            out["new_box"] = {
                "Left": float(new_com.Left),
                "Top": float(new_com.Top),
                "Width": float(new_com.Width),
                "Height": float(new_com.Height),
                "Rotation": float(new_com.Rotation),
            }
            out["new_name"] = str(new_com.Name)
            out["new_alt"] = str(new_com.AlternativeText or "")
            out["box_preserved"] = out["new_box"] == geom
            out["name_preserved"] = out["new_name"] == name
            out["alt_preserved"] = out["new_alt"] == alt

            sfd, shot = tempfile.mkstemp(prefix="pptlive_setpic_shot_", suffix=".png")
            os.close(sfd)  # close the fd or Windows locks the file from PowerPoint
            os.remove(shot)
            shot = os.path.abspath(shot)
            new_com.Export(shot, 2)  # PpShapeFormat.PNG
            tmp = shot
            out["center_after"] = _center_pixel(tmp)
            out["image_is_green"] = _close(out["center_after"], _GREEN)
        finally:
            os.remove(red)
    except Exception as exc:
        out["error"] = _err(exc)
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)
    return out


def main() -> int:
    findings: dict[str, Any] = {}
    green = _write_png(_GREEN)
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        findings["active_deck"] = deck.name
        count_before = len(deck.slides)
        snap = _selection.snapshot(ppt)

        temp_id: int | None = None
        try:
            with deck.edit("set_picture spike"):
                temp = deck.slides.add(layout="blank")
                temp_id = temp.id
                shapes = deck.slides[temp.index].shapes
                findings["fill_userpicture"] = probe_fill_userpicture(shapes, green)
                findings["delete_readd"] = probe_delete_readd(shapes, green)
        except Exception as exc:
            findings["fatal"] = _err(exc)
        finally:
            try:
                with deck.edit("set_picture spike: cleanup"):
                    for idx in range(len(deck.slides), 0, -1):
                        try:
                            if deck.slides[idx].id == temp_id:
                                deck.slides[idx].delete()
                                break
                        except Exception:
                            continue
            except Exception as exc:
                findings["cleanup_error"] = _err(exc)
            _selection.restore(ppt, snap)

        findings["net_zero_ok"] = len(deck.slides) == count_before

    if os.path.exists(green):
        os.remove(green)
    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

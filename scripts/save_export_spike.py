"""Spike — verify v1.1 save / save-as / PDF-export COM behavior on real PowerPoint.

Probes the surface for `deck.save()` / `deck.save_as(path)` / `deck.export_pdf(path)`:

  * `Presentation.Saved` (dirty flag) + `Presentation.Path` on a never-saved deck
    (Path is expected to be "" — the signal that `Save()` has no target yet).
  * `Presentation.Save()` on a never-saved deck — does it raise (no path), or
    silently no-op / pop a dialog? (the roadmap "Spike" question to resolve).
  * `Presentation.SaveAs(FileName, FileFormat)` to .pptx
    (ppSaveAsOpenXMLPresentation = 24) and the effect on `.Path` / `.Saved`.
  * `Presentation.SaveAs(FileName, ppSaveAsPDF=32)` vs.
    `Presentation.ExportAsFixedFormat(Path, FixedFormatType=ppFixedFormatTypePDF=2)`
    — which produces a PDF, what each does to the in-memory deck's `.Path`
    (SaveAs *rebinds* the working file; ExportAsFixedFormat should not), and
    whether either wants an absolute path.

Run against a *running* PowerPoint:

    uv run python scripts/save_export_spike.py

Net-zero and polite: all work happens on a throwaway presentation created
windowless via `Presentations.Add(WithWindow=msoFalse)`, written only into a
TemporaryDirectory, and closed without saving in a `finally`. The user's open
deck and view are never touched. Prints one JSON findings object.
"""

from __future__ import annotations

import json
import os
import tempfile
from typing import Any

import pptlive as pl

_PDF_SIG = b"%PDF-"
_PPTX_SIG = b"PK\x03\x04"  # .pptx is a zip

# PowerPoint magic constants (probed here, promoted to constants.py once verified).
PP_SAVE_AS_OPEN_XML = 24  # ppSaveAsOpenXMLPresentation
PP_SAVE_AS_PDF = 32  # ppSaveAsPDF
PP_FIXED_FORMAT_PDF = 2  # ppFixedFormatTypePDF
MSO_FALSE = 0  # msoFalse (WithWindow)


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:400]


def _head(path: str, n: int = 8) -> bytes:
    try:
        with open(path, "rb") as fh:
            return fh.read(n)
    except Exception:
        return b""


def _file_info(path: str) -> dict[str, Any]:
    exists = os.path.isfile(path)
    head = _head(path) if exists else b""
    return {
        "exists": exists,
        "size": os.path.getsize(path) if exists else 0,
        "is_pdf": head.startswith(_PDF_SIG),
        "is_pptx": head.startswith(_PPTX_SIG),
    }


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        app = ppt.com
        active_before: str | None
        try:
            active_before = str(app.ActivePresentation.Name)
        except Exception:
            active_before = None
        findings["active_deck_before"] = active_before
        count_before = int(app.Presentations.Count)

        pres = None
        with tempfile.TemporaryDirectory(prefix="pptlive_save_") as tmp:
            try:
                # Throwaway, windowless deck with one slide so SaveAs/Export have content.
                pres = app.Presentations.Add(MSO_FALSE)
                pres.Slides.Add(1, 11)  # ppLayoutBlank-ish; index 1, any layout int

                # --- never-saved state -------------------------------------------------
                findings["never_saved"] = {
                    "Path": repr(str(pres.Path)),
                    "FullName": repr(str(pres.FullName)),
                    "Name": str(pres.Name),
                    "Saved": int(pres.Saved),  # MsoTriState; -1 True / 0 False
                }

                # --- Save() on a never-saved deck: does it raise? ---------------------
                try:
                    pres.Save()
                    findings["save_never_saved"] = {
                        "raised": False,
                        "Path_after": repr(str(pres.Path)),
                    }
                except Exception as exc:
                    findings["save_never_saved"] = {"raised": True, "error": _err(exc)}

                # --- SaveAs .pptx (relative vs absolute) ------------------------------
                pptx_abs = os.path.join(tmp, "deck.pptx")
                try:
                    pres.SaveAs(pptx_abs, PP_SAVE_AS_OPEN_XML)
                    findings["save_as_pptx"] = {
                        "file": _file_info(pptx_abs),
                        "Path_after": repr(str(pres.Path)),  # SaveAs rebinds working file
                        "FullName_after": repr(str(pres.FullName)),
                        "Saved_after": int(pres.Saved),
                    }
                except Exception as exc:
                    findings["save_as_pptx"] = {"error": _err(exc)}

                # --- Save() now that it has a path ------------------------------------
                try:
                    pres.Slides.Add(2, 11)  # dirty it
                    saved_dirty = int(pres.Saved)
                    pres.Save()
                    findings["save_with_path"] = {
                        "dirty_before": saved_dirty,
                        "Saved_after": int(pres.Saved),
                    }
                except Exception as exc:
                    findings["save_with_path"] = {"error": _err(exc)}

                # --- ExportAsFixedFormat -> PDF (should NOT rebind .Path) -------------
                path_before_export = str(pres.Path)
                pdf_export = os.path.join(tmp, "export.pdf")
                try:
                    pres.ExportAsFixedFormat(pdf_export, PP_FIXED_FORMAT_PDF)
                    findings["export_as_fixed_format_pdf"] = {
                        "file": _file_info(pdf_export),
                        "Path_unchanged": str(pres.Path) == path_before_export,
                        "Path_after": repr(str(pres.Path)),
                    }
                except Exception as exc:
                    findings["export_as_fixed_format_pdf"] = {"error": _err(exc)}

                # --- SaveAs ppSaveAsPDF=32 (the alternative PDF path) -----------------
                pdf_saveas = os.path.join(tmp, "saveas.pdf")
                path_before_saveas = str(pres.Path)
                try:
                    pres.SaveAs(pdf_saveas, PP_SAVE_AS_PDF)
                    findings["save_as_pdf"] = {
                        "file": _file_info(pdf_saveas),
                        "Path_after": repr(str(pres.Path)),
                        "rebound_working_file": str(pres.Path) != path_before_saveas,
                    }
                except Exception as exc:
                    findings["save_as_pdf"] = {"error": _err(exc)}

                # --- relative-path probe for ExportAsFixedFormat ----------------------
                rel = "pptlive_rel_probe.pdf"
                cwd = os.getcwd()
                try:
                    pres.ExportAsFixedFormat(rel, PP_FIXED_FORMAT_PDF)
                    landed = os.path.join(cwd, rel)
                    findings["export_relative_path"] = {
                        "landed_in_cwd": os.path.isfile(landed),
                        "cwd": cwd,
                    }
                    if os.path.isfile(landed):
                        os.remove(landed)
                except Exception as exc:
                    findings["export_relative_path"] = {"error": _err(exc)}

            finally:
                if pres is not None:
                    try:
                        pres.Saved = -1  # mark clean so Close() never prompts
                    except Exception:
                        pass
                    try:
                        pres.Close()
                    except Exception as exc:
                        findings["close_error"] = _err(exc)

        findings["net_zero_ok"] = int(app.Presentations.Count) == count_before
        try:
            findings["active_deck_after"] = (
                str(app.ActivePresentation.Name) if int(app.Presentations.Count) else None
            )
        except Exception:
            findings["active_deck_after"] = None

    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

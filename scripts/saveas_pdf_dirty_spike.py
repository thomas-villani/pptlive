"""Spike — does `SaveAs(path, ppSaveAsPDF=32)` corrupt the deck's dirty flag?

`export_pdf` will use `SaveAs(path, 32)` (the verified PDF path; ExportAsFixedFormat
won't marshal under late binding). A PDF export must be a *read* — it must NOT mark
the working .pptx as saved when it still has unsaved edits, or a later `deck.save()`
guard / the user's "unsaved changes" prompt would be wrong. This dirties a throwaway
deck, exports PDF, and checks whether `.Saved` flipped. Net-zero (windowless, temp,
closed in finally).

    uv run python scripts/saveas_pdf_dirty_spike.py
"""

from __future__ import annotations

import json
import os
import tempfile
from typing import Any

import pptlive as pl

PP_SAVE_AS_PDF = 32
MSO_FALSE = 0


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        app = ppt.com
        count_before = int(app.Presentations.Count)
        pres = None
        with tempfile.TemporaryDirectory(prefix="pptlive_dirty_") as tmp:
            try:
                pres = app.Presentations.Add(MSO_FALSE)
                pres.Slides.Add(1, 11)
                # Bind a real working file so .Path is non-empty and dirty state is meaningful.
                pres.SaveAs(os.path.join(tmp, "work.pptx"), 24)  # ppSaveAsOpenXMLPresentation
                # Now dirty it with an edit.
                pres.Slides.Add(2, 11)
                findings["dirty_before_export"] = int(pres.Saved)  # expect 0 (msoFalse = dirty)
                name_before = str(pres.Name)
                path_before = str(pres.Path)

                pdf = os.path.join(tmp, "out.pdf")
                pres.SaveAs(pdf, PP_SAVE_AS_PDF)

                findings["after_pdf_export"] = {
                    "Saved": int(pres.Saved),  # if -1, the export wrongly marked it clean
                    "Name_unchanged": str(pres.Name) == name_before,
                    "Path_unchanged": str(pres.Path) == path_before,
                    "Name_now": str(pres.Name),
                    "pdf_ok": os.path.isfile(pdf) and open(pdf, "rb").read(5) == b"%PDF-",  # noqa: SIM115
                }
                # If the export flipped Saved to clean, confirm we can restore it.
                if int(pres.Saved) != 0:
                    pres.Saved = MSO_FALSE
                    findings["restore_dirty_works"] = int(pres.Saved) == 0
            except Exception as exc:
                findings["error"] = _err(exc)
            finally:
                if pres is not None:
                    try:
                        pres.Saved = -1
                    except Exception:
                        pass
                    try:
                        pres.Close()
                    except Exception as exc:
                        findings["close_error"] = _err(exc)
        findings["net_zero_ok"] = int(app.Presentations.Count) == count_before
    print(json.dumps(findings, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

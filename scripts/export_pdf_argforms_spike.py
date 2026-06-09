"""Spike — pin down the working `ExportAsFixedFormat` arg form for PDF export.

The first save/export spike found positional `ExportAsFixedFormat(path, 2)` raises
`TypeError: The Python instance can not be converted to a COM object`. This tries
several call forms (named kwargs, +Intent, the `...2` variant) to find the one that
binds, and whether it rebinds `.Path` or wants an absolute path. Net-zero: throwaway
windowless deck, temp dir, closed in `finally`.

    uv run python scripts/export_pdf_argforms_spike.py
"""

from __future__ import annotations

import json
import os
import tempfile
from typing import Any

import pptlive as pl

PP_FIXED_FORMAT_PDF = 2  # ppFixedFormatTypePDF
PP_FFI_SCREEN = 1  # ppFixedFormatIntentScreen
PP_FFI_PRINT = 2  # ppFixedFormatIntentPrint
MSO_FALSE = 0


def _err(exc: Exception) -> str:
    return f"{type(exc).__name__}: {exc}"[:300]


def _pdf_info(path: str) -> dict[str, Any]:
    if not os.path.isfile(path):
        return {"exists": False}
    with open(path, "rb") as fh:
        head = fh.read(5)
    return {"exists": True, "size": os.path.getsize(path), "is_pdf": head == b"%PDF-"}


def main() -> int:
    findings: dict[str, Any] = {}
    with pl.attach() as ppt:
        app = ppt.com
        count_before = int(app.Presentations.Count)
        pres = None
        with tempfile.TemporaryDirectory(prefix="pptlive_pdf_") as tmp:
            try:
                pres = app.Presentations.Add(MSO_FALSE)
                pres.Slides.Add(1, 11)
                path0 = str(pres.Path)

                forms: list[tuple[str, Any]] = []

                def attempt(label: str, fn: Any) -> None:
                    out = os.path.join(tmp, f"{label}.pdf")
                    try:
                        fn(out)
                        forms.append(
                            (
                                label,
                                {
                                    "ok": True,
                                    "file": _pdf_info(out),
                                    "path_unchanged": str(pres.Path) == path0,
                                },
                            )
                        )
                    except Exception as exc:
                        forms.append((label, {"ok": False, "error": _err(exc)}))

                attempt(
                    "kw_path_type",
                    lambda o: pres.ExportAsFixedFormat(Path=o, FixedFormatType=PP_FIXED_FORMAT_PDF),
                )
                attempt(
                    "kw_with_intent",
                    lambda o: pres.ExportAsFixedFormat(
                        Path=o, FixedFormatType=PP_FIXED_FORMAT_PDF, Intent=PP_FFI_SCREEN
                    ),
                )
                attempt(
                    "pos_with_intent",
                    lambda o: pres.ExportAsFixedFormat(o, PP_FIXED_FORMAT_PDF, PP_FFI_SCREEN),
                )
                # The `...2` variant exists on newer builds (adds a few trailing params).
                attempt(
                    "fixedformat2",
                    lambda o: pres.ExportAsFixedFormat2(
                        Path=o, FixedFormatType=PP_FIXED_FORMAT_PDF, Intent=PP_FFI_SCREEN
                    ),
                )

                findings["forms"] = dict(forms)
                findings["path0"] = repr(path0)

                # Relative-path behavior for the winning form (kw_with_intent if it worked).
                rel = "pptlive_export_rel.pdf"
                cwd = os.getcwd()
                try:
                    pres.ExportAsFixedFormat(
                        Path=rel, FixedFormatType=PP_FIXED_FORMAT_PDF, Intent=PP_FFI_SCREEN
                    )
                    landed = os.path.join(cwd, rel)
                    findings["relative"] = {"landed_in_cwd": os.path.isfile(landed), "cwd": cwd}
                    if os.path.isfile(landed):
                        os.remove(landed)
                except Exception as exc:
                    findings["relative"] = {"error": _err(exc)}
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

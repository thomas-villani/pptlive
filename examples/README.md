# pptlive examples

Runnable demos you can point at a **live** PowerPoint and watch happen. They are
not part of the test suite (CI never runs them) — they exist to show the library
end-to-end on a real instance.

## Prerequisites

- Windows with Microsoft PowerPoint installed.
- pptlive installed in the environment you run them from:

  ```powershell
  uv sync                      # from the repo, into .venv
  # or, as an installed package:
  uv pip install pptlive
  ```

The Python scripts (`examples/python/`) **attach to a running PowerPoint and
create a fresh presentation** to draw into (`pl.connect()` launches PowerPoint
if none is open). The `pptlive` CLI itself only ever *attaches* — it never
launches or creates a presentation — so the PowerShell scripts
(`examples/powershell/`) start PowerPoint themselves via raw COM
(`New-Object -ComObject PowerPoint.Application`) and add a blank deck before
handing off to `pptlive`. Either way, PowerPoint stays visible the whole time —
that's the point; you watch the edits land — and no script touches a deck you
already have open. Nothing is saved to disk and nothing is closed; the new deck
is yours to keep or discard.

Every mutation goes through `deck.edit(...)`, so each demo step is **one Ctrl-Z**
and your viewed slide / selection is preserved.

## Python (`examples/python/`)

| Script | Shows |
| ------ | ----- |
| `01_quickstart.py`   | attach, add a title + content slide, set text, format a run |
| `02_build_a_deck.py` | a table, a column chart (with data), and a SmartArt diagram |
| `03_restyle_deck.py` | deck-wide styling: theme palette + fonts, master text styles, background |
| `04_finish_and_export.py` | the finishing pass: `find_replace` a typo, `snapshot()`, `save_as` + `export_pdf` |

```powershell
uv run python examples/python/01_quickstart.py
```

## PowerShell / CLI (`examples/powershell/`)

The same ideas driven entirely through the `pptlive` CLI — useful for shell
automation and for seeing the one-JSON-object-per-call contract.

| Script | Shows |
| ------ | ----- |
| `quickstart.ps1`    | `status`, add a slide, write placeholders, read it back |
| `build_report.ps1`  | build a small report deck: slides, a table, and a chart |

```powershell
# you may need to allow the script to run in this session first:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
./examples/powershell/quickstart.ps1
```

> The PowerShell demos assume `pptlive` is on `PATH` (true after `uv pip install
> pptlive`). From the repo, prefixing the individual `pptlive` calls *inside* the
> script with `uv run` doesn't work — those calls happen in the child `pptlive`
> process the script invokes, which never sees `uv run`'s venv. Instead run the
> whole script under `uv run` so the venv's `Scripts` dir is on `PATH` for the
> entire child process: `uv run pwsh -File examples/powershell/quickstart.ps1`.
> Alternatively, activate the venv first (`.venv\Scripts\Activate.ps1`) and then
> run the `.ps1` script normally. Each call prints one JSON object to stdout; the
> scripts parse it with `ConvertFrom-Json`.

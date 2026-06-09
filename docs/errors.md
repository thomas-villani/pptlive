# Errors & exit codes

pptlive translates pywin32's `pywintypes.com_error` into a small, typed
exception hierarchy. The CLI maps those exceptions to deterministic exit codes
so LLM tool-use loops can branch on the failure mode without parsing error
text.

## Exception hierarchy

```
Exception
└── PptliveError
    ├── PowerPointNotRunningError
    ├── PresentationNotFoundError
    ├── AnchorNotFoundError
    │   ├── SlideNotFoundError
    │   └── LayoutNotFoundError
    ├── NoTextFrameError
    ├── SlideShowNotRunningError
    ├── UnsavedPresentationError
    ├── AmbiguousMatchError
    ├── PowerPointBusyError
    └── ComError
```

`PptliveError` is the catch-all base — `except pptlive.PptliveError` catches
every typed error pptlive raises. Anything that wasn't a COM error in the first
place (e.g. a `ValueError` from your own code) bubbles up unchanged.

## Reference

### `PptliveError`
Base class. Catch this if you want one `try` for every pptlive failure.

### `PowerPointNotRunningError`
No PowerPoint instance is running. Raised by
[`attach()`](python-api.md#pptlive.attach) and by
[`connect(launch_if_missing=False)`](python-api.md#pptlive.connect).
**Not retryable** within a session — PowerPoint has to actually be running.
Maps to exit **4**.

### `PresentationNotFoundError`
The requested presentation isn't open. Raised by
`ppt.presentations[name]` and by `ppt.presentations.active` when no deck is
active. The missing name is on `.name`. Maps to exit **2**.

### `AnchorNotFoundError`
A shape, placeholder, paragraph, table cell, notes anchor — or an entire slide
or layout — you asked for doesn't exist. `.kind` names the thing that was
missing (`"shape"`, `"placeholder"`, `"paragraph"`, `"table"`, `"cell"`,
`"slide"`, `"layout"`, …) and `.name` is what you asked for. **Retryable after
re-reading** the slide / shape listing — the deck may have changed (and z-order
drifts). Maps to exit **2**.

### `SlideNotFoundError`
A slide index is out of range. Subclass of
[`AnchorNotFoundError`](#anchornotfounderror), so it shares exit code 2 and
`except AnchorNotFoundError` catches it too. The bad index is on `.index`.
**Retryable after re-reading `deck.slides.list()`.**

### `LayoutNotFoundError`
A layout name (or index) you passed to `slides.add` / `set_layout` doesn't
exist in the deck. Subclass of [`AnchorNotFoundError`](#anchornotfounderror)
(exit 2). The error lists the deck's available layout names. **Retryable after
reading `deck.layouts()`** to see what's actually defined.

### `NoTextFrameError`
A text operation (`set_text`, `format_text`, paragraph verbs) hit a shape with
no text frame — a picture, a line, a connector. This is the one genuinely new
code relative to wordlive: it's common enough (an LLM tries to set text on a
decorative shape) to deserve a deterministic exit code instead of a bare COM
failure. **Not retryable** against that shape — pick a text-bearing one. Maps
to exit **6**.

### `SlideShowNotRunningError`
A `deck.show` control verb (`next`, `goto`, `black`, …) was called when no
slide show is running. It's a precondition failure, not a missing anchor — so
it maps to the generic exit code (**1**), not 2. Start a show first with
`deck.show.start()`; `deck.show.state()` is the one verb that never raises when
nothing is running (it just reports `running: false`).

### `UnsavedPresentationError`
`deck.save()` was called on a deck that has never been saved (no file path yet),
so there's nothing to save *to*. PowerPoint's own `Save()` doesn't raise here —
on a OneDrive/SharePoint build it silently uploads to the default cloud folder —
so pptlive guards on the empty path and raises instead. Use `deck.save_as(path)`
to give it a file first. A precondition failure, so it maps to exit **1**.

### `AmbiguousMatchError`
A fuzzy match resolved to more than one target without disambiguation. The
exception carries the candidates so an agent can pick one and retry.
**Retryable** by narrowing the request. Maps to exit **5**.

### `PowerPointBusyError`
PowerPoint rejected the COM RPC — usually a modal dialog is open (Save As,
Insert, a Format pane prompt) or the app is mid-operation. **Retryable** with
exponential back-off. The HRESULT is on `.hresult`; `.retryable` is always
`True`. Maps to exit **3**.

!!! note "A running slide show does *not* block edits"
    The spec originally assumed editing during a live show would reject as
    busy. A 2026 spike overturned that: a `set_text` mid-show *succeeds* and
    raises nothing. So a running show is not, by itself, a `PowerPointBusyError`
    — this exception stays the home for genuine modal-dialog `RPC_E_*`
    rejections.

### `ComError`
Catch-all for any other classified COM error. Carries `.hresult` and
`.description` (when pywin32 surfaces one). Not retryable in general; treat as
a bug in your code or a PowerPoint-side problem. Maps to exit **1**.

## HRESULT mapping

Only one HRESULT family is special-cased: the "PowerPoint is momentarily
unavailable" codes that map to [`PowerPointBusyError`](#powerpointbusyerror).
Everything else becomes a generic [`ComError`](#comerror) with the HRESULT
preserved.

| HRESULT       | Mnemonic                         | pptlive exception     |
| ------------- | -------------------------------- | --------------------- |
| `0x80010001`  | `RPC_E_CALL_REJECTED`            | `PowerPointBusyError` |
| `0x80010005`  | `RPC_E_SERVERCALL_REJECTED`      | `PowerPointBusyError` |
| `0x8001010A`  | `RPC_E_SERVERCALL_RETRYLATER`    | `PowerPointBusyError` |
| any other     | —                                | `ComError`            |

The classification logic lives in the `_BUSY_HRESULTS` set in
[`src/pptlive/exceptions.py`](https://github.com/thomas-villani/pptlive/blob/main/src/pptlive/exceptions.py).
If you find a code that should be treated as busy/retryable, it goes there.

## CLI exit codes

The CLI maps the exception hierarchy onto seven exit codes, defined in
[`src/pptlive/cli/main.py`](https://github.com/thomas-villani/pptlive/blob/main/src/pptlive/cli/main.py):

| Exit | Exception(s)                                          | Meaning                                       | Retry?                          |
| ---- | ----------------------------------------------------- | --------------------------------------------- | ------------------------------- |
| `0`  | —                                                     | success                                       | —                               |
| `1`  | `PptliveError` (default), `SlideShowNotRunningError`, `UnsavedPresentationError`, `ComError` | other / unclassified / no show running / never-saved deck | depends on cause                |
| `2`  | `AnchorNotFoundError`, `SlideNotFoundError`, `LayoutNotFoundError`, `PresentationNotFoundError` | anchor / slide / shape / layout / deck missing | yes, after re-reading content   |
| `3`  | `PowerPointBusyError`                                 | modal dialog or busy RPC                      | **yes**, with back-off          |
| `4`  | `PowerPointNotRunningError`                           | no PowerPoint instance                        | only if the user launches PowerPoint |
| `5`  | `AmbiguousMatchError`                                 | fuzzy match hit more than one target          | **yes**, after disambiguating   |
| `6`  | `NoTextFrameError`                                    | text op on a shape with no text frame         | no — pick a text-bearing shape  |

## Retry guidance

The only exception explicitly designed to be retryable is
[`PowerPointBusyError`](#powerpointbusyerror). A typical retry loop:

```python
import time
import pptlive as pl

def with_retry(fn, *, attempts=4, base=0.5):
    for i in range(attempts):
        try:
            return fn()
        except pl.PowerPointBusyError:
            if i == attempts - 1:
                raise
            time.sleep(base * (2 ** i))   # 0.5, 1, 2, 4 seconds


def update_title():
    with pl.attach() as ppt:
        deck = ppt.presentations.active
        with deck.edit("Update title"):
            deck.anchor_by_id("ph:1:title").set_text("Q3 Results")


with_retry(update_title)
```

For the CLI (PowerShell):

```powershell
foreach ($i in 1..4) {
    pptlive write --anchor-id ph:1:title --text "Q3 Results"
    if ($LASTEXITCODE -eq 0) { break }
    if ($LASTEXITCODE -ne 3) { exit $LASTEXITCODE }   # only retry exit code 3
    Start-Sleep -Seconds ($i * $i)                    # quadratic-ish back-off
}
```

[`AnchorNotFoundError`](#anchornotfounderror) is *also* effectively retryable —
but only after you've re-read the slide / shape listing, since z-order drifts
and the deck may have changed since your last call.

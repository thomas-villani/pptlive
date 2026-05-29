# Getting started

This page takes you from zero to a first polite edit, both from Python and
from the CLI.

## Prerequisites

- **Windows.** pptlive talks to PowerPoint over COM (`pywin32`); there is no
  cross-platform path.
- **Microsoft PowerPoint**, installed and running, with a deck open. pptlive
  *attaches* to the running app — it never closes it, and it can't run
  PowerPoint hidden (`Application.Visible = False` is refused by the app).
- **Python 3.10+**.

## Install

```bash
pip install pptlive

# add to a project
uv add pptlive
```

`pywin32` is pulled in automatically on Windows. Click is the only other
runtime dependency. The MCP server is an optional extra — see
[MCP server](mcp.md).

## Hello, deck

Open a presentation in PowerPoint, then run:

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active
    print(deck.name)

    for entry in deck.outline():
        print(f"{entry['slide']:>3}. {entry['title']}")
        for bullet in entry["bullets"]:
            print(f"      • {bullet}")
```

`attach()` connects to the *already-running* PowerPoint instance — it won't
launch one. If PowerPoint isn't running you get a
[`PowerPointNotRunningError`](errors.md). Use
[`pl.connect(launch_if_missing=True)`](python-api.md#pptlive.connect) when
you'd rather launch PowerPoint if it isn't already up.

Every shape listing carries an `anchor_id` like `ph:3:title` or `shape:3:2` —
those strings are how the CLI and LLM tool-use loops address text. See
[Anchor IDs](concepts.md#anchor-ids) for the scheme.

## Your first polite edit

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active

    with deck.edit("Update the agenda slide"):
        deck.anchor_by_id("ph:2:title").set_text("Agenda")
        deck.anchor_by_id("ph:2:body").set_text("Intro\nDemo\nQ&A")
```

A few things are happening that aren't obvious from the code:

- `deck.edit("…")` calls `Application.StartNewUndoEntry()` on entry, fencing
  the block. PowerPoint groups the in-session edits that follow, so both
  mutations collapse into a single Ctrl-Z.
- Before the block runs, the user's **viewed slide** and shape/text
  `Selection` are snapshotted. On exit they're restored — your script does not
  jump the user to a different slide. See
  [Politeness](concepts.md#politeness-model).
- If a placeholder is missing, you get a typed
  [`AnchorNotFoundError`](errors.md), not a raw COM error. Setting text on a
  shape with no text frame (a picture, a line) raises
  [`NoTextFrameError`](errors.md).

## Same task from the CLI

The CLI is intentionally thin over the Python API. Same atomic-undo, same
politeness, JSON on stdout:

```bash
# What's open?
pptlive status

# What's in the active deck?
pptlive slides            # one row per slide (index, layout, title, …)
pptlive outline           # titles + body bullets
pptlive slide read 2      # every shape on slide 2 (anchor_id, name, type, geometry)

# Mutate (each is one Ctrl-Z).
pptlive write --anchor-id ph:2:title --text "Agenda"
pptlive write --anchor-id ph:2:body  --text "Intro\nDemo\nQ&A"
```

Every command emits one JSON object on stdout (`--text` if you'd rather read
it) and uses deterministic exit codes:

| Exit | Meaning                          |
| ---- | -------------------------------- |
| `0`  | ok                               |
| `2`  | anchor / slide / shape / deck not found |
| `3`  | PowerPoint busy / modal dialog   |
| `4`  | PowerPoint not running           |
| `5`  | ambiguous match                  |
| `6`  | shape has no text frame          |
| `1`  | other error                      |

Full reference: [CLI](cli.md).

## Two more everyday flows

### Let a vision model *see* the slide it just built

PowerPoint renders the live, unsaved state — so you can build, look, and
iterate:

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    png = deck.slides[4].export_image(width=1280)   # temp PNG (or pass a path)
    # ...hand `png` to your image tool, look, then revise.
```

`export_image` is polite — it doesn't move the user's view. See
[Cookbook §5](cookbook.md#5-build-look-iterate-the-vision-loop).

### Read what the user is looking at

Most useful when you want to feed the user's current focus to an LLM, or act
on a hotkey-driven selection:

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    sel = deck.selection()             # SelectionInfo: type, slide, anchor_id, …

    if sel.anchor_id:
        print(f"User is on slide {sel.slide}, selection -> {sel.anchor_id}")
    else:
        print("Nothing selected.")
```

`selection()` never moves the user. To *act* on the selection, target the
opt-in [`here:`](concepts.md#anchor-ids) anchor. See
[Cookbook §6](cookbook.md#6-act-on-whatever-the-user-is-pointing-at).

## Where to next

- [Concepts](concepts.md) — the ideas that shape every pptlive API:
  politeness, semantic anchors, the hierarchical anchor scheme, and
  `EditScope`.
- [Cookbook](cookbook.md) — end-to-end recipes including an LLM tool-use loop.
- [Python API](python-api.md) — auto-generated reference for every public
  symbol.
- **Driving an LLM agent?** The [MCP server](mcp.md) exposes the same live
  control to Claude Desktop and other MCP clients.

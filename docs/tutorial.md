# Tutorial: build a narrated deck

This is a guided, end-to-end walkthrough. By the end you'll have built a short
presentation from scratch in a *live* PowerPoint window, rendered it to check
your work, narrated each slide, and exported a finished MP4 — touching every
layer of pptlive along the way.

It assumes you've done the one-minute [Getting started](getting-started.md)
(install + a first polite edit). Everything here is Python; the matching CLI
commands are shown where they're handy.

!!! note "Follow along live"
    Open PowerPoint with a **blank presentation** before you start. pptlive
    attaches to the running app, so you'll watch each step happen on screen.

## What you'll build

A three-slide micro-deck — a title, a content slide, and a closing — each with
its own voiceover, exported as a self-playing video. We'll go in the order you'd
actually work: look, build, *look again*, narrate, export.

## Step 1 — Attach and look around

Always start with a side-effect-free read, so you know what you're working with.

```python
import pptlive as pl

with pl.attach() as ppt:
    deck = ppt.presentations.active
    print(deck.name, "—", len(deck.slides), "slide(s)")
    for row in deck.slides.list():
        print(row["index"], row["layout"], "|", row["title"])
```

`attach()` connects to the already-running app (it never launches or closes it).
A blank deck reports one slide on the default layout. From the CLI that's just
`pptlive status` and `pptlive slides`.

## Step 2 — Build the slides

We'll add our slides, then drop the blank starter. Slides are created from a
**layout** (by friendly name); each new slide exposes its placeholders by
semantic kind (`ctrtitle`, `subtitle`, `title`, `body`).

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    starter = [s.id for s in deck.slides]          # remember the blank slide

    with deck.edit("Build the deck"):
        # Title slide
        s1 = deck.slides.add(layout="title_slide")
        deck.slides[s1.index].placeholder("ctrtitle").set_text("Project Nimbus")
        deck.slides[s1.index].placeholder("subtitle").set_text("A weather balloon, but ambitious")

        # Content slide — newlines become separate, addressable paragraphs
        s2 = deck.slides.add(layout="title_and_content")
        deck.slides[s2.index].placeholder("title").set_text("Why now?")
        deck.slides[s2.index].placeholder("body").set_text(
            "Helium is cheaper than ever\nThe sky is right there\nVibes"
        )

        # Closing slide
        s3 = deck.slides.add(layout="title_and_content")
        deck.slides[s3.index].placeholder("title").set_text("Ad astra (ish)")
        deck.slides[s3.index].placeholder("body").set_text("Questions → up there ↑")

    # Drop the original blank slide so the deck is exactly our three
    with deck.edit("Remove starter slide"):
        for sid in starter:
            for i in range(len(deck.slides), 0, -1):
                if deck.slides[i].id == sid:
                    deck.slides[i].delete()
                    break
```

Two things worth noticing:

- **`deck.edit(label)`** fences each block into a single Ctrl-Z *and* restores
  the user's view and selection when it exits. Wrap every mutation in it.
- **Anchors are hierarchical.** `placeholder("title")` resolves the title
  placeholder on that slide; you could equally address it by `ph:S:title`. See
  [Concepts → Anchor IDs](concepts.md#anchor-ids).

!!! tip "Layout names are template-dependent"
    `add(layout="title_slide")` resolves friendly aliases against the deck's real
    layout names. If a name isn't found you get a `LayoutNotFoundError` listing
    the available ones — run `deck.layouts()` (or `pptlive slide layouts`) to see
    them.

## Step 3 — Look at what you built

You don't have to guess whether it looks right — PowerPoint renders the live,
unsaved state, so you can *see* it and iterate.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    shots = deck.snapshot(max_dim=1000)            # one PNG per slide
    for shot in shots:
        print(f"slide {shot.slide} -> {shot.path}")
    # Hand the PNG paths to a vision model, or just open them.
```

`snapshot` is a **read** — it never moves the user's view. `max_dim` caps each
slide's long edge in pixels, a predictable per-slide budget for a vision model.
From the CLI: `pptlive snapshot --out check.png --max-dim 1000`.

This is the loop that changes how you work: **build → look → revise.** If a title
overflowed or a color came out wrong, you'd catch it here before narrating.

## Step 4 — Narrate each slide

Now the fun part. Generate a voiceover for each slide (any TTS that writes an
audio file works — below uses the [`llm`](https://llm.datasette.io/) CLI's
`tts`), then attach it. `add_audio` embeds the clip and, by default, **paces the
slide to the clip's length** so the exported video tracks the narration.

```bash
llm tts "Introducing Project Nimbus: a weather balloon, but ambitious." -o s1.mp3 --no-play
llm tts "Why now? Helium is cheap, the sky is right there, and honestly, the vibes." -o s2.mp3 --no-play
llm tts "Ad astra, ish. Questions may be directed upward. Thank you." -o s3.mp3 --no-play
```

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    with deck.edit("Narrate"):
        for i, clip in enumerate(["s1.mp3", "s2.mp3", "s3.mp3"], start=1):
            deck.slides[i].add_audio(clip)         # autoplay + pace the slide (defaults)

    # Confirm it took: every media shape reads back its clip info.
    media = deck.slides[1].read()["shapes"][-1]["media"]
    print(media)   # {'type': 'sound', 'length_s': 4.1, 'autoplay': True, ...}
```

The defaults (`autoplay=True`, `hide_icon=True`, `pace_slide=True`) are what you
want for narration; pass `pace_slide=False` to keep a slide's own timing, or
`add_video(...)` for a video clip (which stays visible). CLI equivalent:
`pptlive media add --slide 1 --kind audio --path s1.mp3`.

## Step 5 — Export the video

`export_video` drives PowerPoint's encoder. It's a **read** (it won't touch your
working file) and, by default, **blocks until the encode finishes** — handy for a
script that wants the file in hand.

```python
with pl.attach() as ppt:
    deck = ppt.presentations.active
    result = deck.export_video("nimbus.mp4", resolution=1080)
    print(result.status, "->", result.path)   # done -> C:\...\nimbus.mp4
    assert result.ok
```

A long encode? Pass `wait=False` to return immediately, then poll
`deck.video_status()` until it reports `done`. From the CLI:
`pptlive export-video nimbus.mp4 --resolution 1080` (add `--no-wait` +
`pptlive video-status` for the non-blocking flow).

## Step 6 — Save the editable deck

pptlive **never auto-saves** — persisting is always explicit. `save_as` writes
the `.pptx` and rebinds the working file to it.

```python
with pl.attach() as ppt:
    ppt.presentations.active.save_as("nimbus.pptx", overwrite=True)
```

## Recap

In one short session you went the whole distance:

1. **Read** the deck without disturbing it (`status` / `slides.list`).
2. **Built** slides from layouts and wrote their placeholders, each block one
   atomic, view-preserving edit.
3. **Rendered** to *see* your work (`snapshot`) — the build-look-revise loop.
4. **Narrated** each slide (`add_audio`, auto-paced to the clip).
5. **Exported** a finished MP4 (`export_video`) and **saved** the source deck.

That's the spine of pptlive: read freely, mutate politely, look, and produce.

## Where to next

- [Concepts](concepts.md) — politeness, the anchor scheme, and `EditScope` in
  depth.
- [Cookbook](cookbook.md) — focused recipes (tables, charts, SmartArt, comments,
  theming, an LLM tool-use loop, and [the narrated-video recipe](cookbook.md#22-narrate-a-deck-and-export-a-video)).
- [CLI](cli.md) and [Python API](python-api.md) — the full reference.
- Driving an agent? The [MCP server](mcp.md) exposes all of this to Claude
  Desktop and other MCP clients.

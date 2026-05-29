---
title: pptlive
hide:
  - navigation
---

# pptlive

**Drive a running Microsoft PowerPoint instance from Python — `xlwings`, but for PowerPoint.**

Built for both human scripting and LLM agents. Windows-only.

<div class="grid cards" markdown>

-   :material-rocket-launch:{ .lg .middle } **New here?**

    ---

    Install, attach to PowerPoint, and run your first polite edit in five minutes.

    [:octicons-arrow-right-24: Getting started](getting-started.md)

-   :material-lightbulb-on:{ .lg .middle } **How it thinks**

    ---

    Anchors, atomic undo, politeness — the ideas that drive the 2-D API.

    [:octicons-arrow-right-24: Concepts](concepts.md)

-   :material-console:{ .lg .middle } **CLI**

    ---

    JSON-in / JSON-out commands designed to drop into an LLM tool-use loop.

    [:octicons-arrow-right-24: CLI reference](cli.md)

-   :material-code-braces:{ .lg .middle } **Python API**

    ---

    Every public class and function, generated from source docstrings.

    [:octicons-arrow-right-24: Python API](python-api.md)

</div>

---

{%
   include-markdown "../README.md"
   start="## Install"
   end="## Development"
%}

## Design principles

- **Politeness first** — operations preserve the user's **viewed slide**,
  shape/text `Selection`, and focus. The user keeps presenting and editing
  alongside you; only verbs that *must* move the screen say so in their name
  (`go_to`, `show.*`, `allow_view_move()`).
- **Semantic anchors over `Selection`** — operations target slides, shapes,
  placeholders, paragraphs, table cells, or notes — never the live selection
  unless you opt in with `here:`.
- **Atomic undo** — every `deck.edit()` block fences a single undo entry with
  `StartNewUndoEntry`, so one Ctrl-Z reverts the whole intent.
- **Escape hatch** — every wrapper exposes `.com` for the raw COM object;
  you're never blocked by missing coverage.

See the [Design](design.md) page for the full rationale.

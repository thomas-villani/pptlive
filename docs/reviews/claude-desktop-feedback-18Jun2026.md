**What's working really well now**

`ppt_batch` with a trailing `render` is the killer pattern. Build a whole slide and see it in one round-trip — that's the workflow that makes iterative design actually possible. I never felt like I was flying blind.

`set_paragraphs` is a huge ergonomic win over the old `write` + `\n` approach. Being able to specify per-paragraph formatting inline — bold, color, size, list_type, space_before — means I can express complex typographic intent in a single op without guessing at inheritance behavior. That's the right primitive for an LLM author.

`shape_add` + `shape_move`/`shape_resize` as a composition pattern works cleanly. Building a card as a background rectangle + foreground textbox is intuitive and the anchor IDs come back immediately, so I can chain format calls on them in the same batch.

`deck_snapshot` at the end for a full-deck visual QA pass is a great quality-of-life feature. One call, see everything.

---

**Friction points still worth addressing**

**The master cascade into custom textboxes** is the biggest behavioral surprise I hit. When I set `master_format_text_style` for body to white, the text inside independently-added textboxes on slide 6 still rendered green. I hadn't targeted those shapes with that call — it cascaded in unexpectedly. For an LLM author this is particularly tricky because I can't easily distinguish "this text is green because I set it green" from "this text is green because the master pulled it there." A `text_frame_status` that surfaced whether a given run's color is directly set vs. theme-inherited would help a lot.

**Body placeholder width defaults to full-width on Title and Content layout.** I had to manually resize `ph:N:body` on slides 4 and 5 after adding a code box on the right — the placeholder was covering my custom shape and truncating text. If I could specify the placeholder geometry on `slide_add` (or if there were a layout variant with a narrower content area by default), that would save a fix-up pass.

**The orphaned arrow on slide 3** — when I placed a `→` textbox between the three phase boxes and the right edge of the slide, it had nothing to point at. That's a design error on my part, but it points to a deeper issue: I have no way to preview spatial layout *before* committing it. Even a lightweight "bounding box report" for shapes I've just added — showing their screen positions relative to each other — would let me catch geometry mistakes before rendering.

**Shape index vs. z-order aliasing still requires care.** I had one moment where `shape_order: back` on a shape returned a different anchor_id because the restack shifted indices. I knew to expect this from prior sessions, but it's a persistent gotcha. A `shapeid:S:ID` addressing mode that survives restacks would be the clean fix — I see the stable ID in the `shape_add` return already, just need that addressing form to work everywhere.

---

**One bigger-picture observation**

The tool is now past the "can it do the thing" threshold and into the "does it do the thing *gracefully*" territory. The core ops are solid. The remaining friction is almost entirely in the feedback loop — when something doesn't land visually, I need more signal about *why* (inherited vs. direct formatting, geometry before commit, stable identity through restacks). That's the layer I'd focus on next.

The delta from the first session to this one is genuinely significant. The `set_paragraphs` addition alone saved probably 6-8 round-trips on this deck. You're clearly using LLM friction as a real design signal, which is exactly the right methodology.

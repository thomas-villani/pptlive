On pptlive

Genuinely good. The thing that made this work is that slide export / snapshot closes the loop — I could build 24 slides, look at all of them for a few thousand tokens, spot that the agenda numerals were too faint and that a card title was colliding with its badge, fix both, re-render just those slides, and confirm. Without that I'd have been writing XML blind and hoping. The --max-dim knob mattering more than any quality setting is the right call — it makes "check every slide" a predictable, affordable operation rather than a gamble.

Two design decisions I'd call out as correct:

- geometry_report() as a pre-render check. Scanning all 24 slides for off_slide took one call and no pixels. It correctly flagged only the title image, which I'd deliberately bled off the right edge. That's a cheap linter for the most common LLM failure mode.
- llm-help returning the whole guide in one shot, including the exit-code taxonomy and the line-spacing footgun. I read it once and never guessed at an API afterward. The --python variant was the right call for this job — I'd have burned 150+ CLI invocations doing it verb-by-verb.

Rough edges I actually hit

1. No text-frame control in the public API. This was the only real gap. add_textbox creates a box with autosize on, so a height you pass is advisory — text grows the box and your layout drifts. I ended up writing this against the COM escape hatch on the very first helper:

tf = sh.com.TextFrame2
tf.AutoSize = 0
tf.WordWrap = -1
tf.VerticalAnchor = _ANCHOR[anchor]
tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0

Every one of those four is needed for precise layout, and the default 0.1″ internal margins in particular silently eat your padding math. read text-frame-status will tell me overflow_risk: possible, but there's no supported way to act on it — the diagnostic exists without the setter. An autosize=/wrap=/anchor=/margins= set of kwargs on add_textbox (and a shape set-text-frame verb) would remove the single biggest reason to reach for .com.

2. set_paragraphs accepts a {...} dict but the documented key list trails off in .... Both help pages show {text, list_type, indent_level, alignment, line_spacing, size, bold, ...} — I couldn't tell whether color and font were in there, so I defensively set structure via set_paragraphs and then looped paragraph(i).format_text(...) for every paragraph. That works, but it's 2–3× the COM round-trips. An exhaustive key table would let a one-pass build be the obvious path.

3. Minor: shape add --kind shape takes --text, but Shapes.add_shape() doesn't. The CLI and Python surfaces diverge here. set_text() right after is fine, but it's a small asymmetry worth closing.

4. Nothing to say about the docs, but worth knowing: the only thing that actually bit me during the run was my own Python syntax errors, twice. Given that scripts are the intended power path, a --dry-run that parses and validates the script without touching PowerPoint would catch that class of error before it half-builds a deck.

One thing I'd add
                                                                                                                                        A crop verb. I had a 2.08:1 tissue panorama and a 1.007: get a full-bleed title panel was to oversize the squareand let 74pt hang off the slide edge. It works, but geometry_report then flags it as a defect forever, so I lose the signal. shape crop --anchor-id ... --left/--top/--right/--bottom (PowerPoin) would turn "cover-fit this image into this box" from atrick into a one-liner — and image placement is a large fraction of what makes a deck look designed rather than assembled.                     
Net: the read → render → look → fix loop is the right abstraction, and it's the part most tooling in this space gets wrong. The gaps are all in the "last 10% of layout precision" band, which is exactlprogress to be.

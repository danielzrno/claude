# Editorial — Mivada design system 02

**Bold magazine.** Black-dominant, oversized headlines, coral punctuation, big pull-quotes.
For thought leadership, marketing pieces, and reports that need to feel confident and editorial.

## Use when
You want impact and voice — a point-of-view document or a keynote-style deck. Pairs a strong
typographic system with high black/coral contrast. (For restrained/corporate work use **Clarity**;
for the website's warm, card-based energy use **Momentum**.)

## Brand (shared — see `../guides/mivada-brand.md`)
Identical to all concepts: coral `#EA493F`, black `#000000`, ink `#111111`, off-white `#FAFAFA`,
warm-neutral `#F2F3EE`, graphite `#323232`, hairline `#E4E4E0`. **Coral is the sharp 10% accent.**

| Token | Hex | Use here |
|-------|-----|----------|
| Coral | `#EA493F` | one word per headline, pull-quotes, index numerals, the "M" chip |
| Black | `#000000` | full-bleed section slides/bands (the dominant surface) |
| Off-white | `#FAFAFA` | the light slides in the black/light sandwich |
| Ink | `#111111` | body on light |

## Type
**Inter** (matched to the site; swap the `--font` variable / `THEME["font"]` / docx font if the licensed
family differs). Headlines run large — Inter **800**, sentence case, tight tracking; uppercase tracked
eyebrows. Hierarchy is carried by **size and black/coral contrast**.

## Motifs used
Black/off-white **sandwich** (alternating slides), oversized headlines, a deliberate **black callout
band**, a large **coral pull-quote**, big coral index numerals, the coral **"M" chip** (white Inter-800 "M").
No accent line under titles; left-aligned body; no emoji; no purple.

## What makes it distinct
The type does the work. Where Clarity whispers and Momentum shows cards, Editorial **shouts in print** —
huge headings, a hard black/coral rhythm, and a pull-quote as a hero moment.

## Files
`editorial.css` · `presentation.html` (8-slide 16:9 deck) · `document.html` (A4) · `build_docx.js`
+ `editorial.docx` (native Word) · `build_pptx.py` + `editorial.pptx` (PowerPoint).

## Make a new artifact
```html
<!-- an on-brand slide: link the theme, write content only -->
<link rel="stylesheet" href="editorial.css">
<section class="slide slide--dark"><p class="eyebrow">02 — Who we are</p>
  <h2 class="h-xl">We turn platforms into <em>human</em> outcomes.</h2></section>
```
```bash
python build_pptx.py                                            # -> editorial.pptx
NODE_PATH=/opt/node22/lib/node_modules node build_docx.js       # -> editorial.docx
# HTML: open presentation.html / document.html, then Print → Save as PDF
```

# Momentum — Mivada design system 03

**The site's own energy.** Coral hero fields, the red "M" chip on cards, pill tags, rounded
photography and KPI cards. The most brand-forward, dynamic concept — closest to mivada.com.

## Use when
Sales pitches, brand-forward decks, dynamic one-pagers — anything that should feel like it came
straight off the Mivada website. (For restrained/corporate work use **Clarity**; for a typographic,
editorial voice use **Editorial**.)

## Brand (shared — see `../guides/mivada-brand.md`)
Identical to all concepts: coral `#EA493F`, coral-deep `#C9362B`, black `#000000`, ink `#111111`,
off-white `#FAFAFA`, warm-neutral `#F2F3EE`, hairline `#E4E4E0`.

| Token | Hex | Use here |
|-------|-----|----------|
| Coral | `#EA493F` | hero fields, the "M" chip, KPI highlight cards, pills |
| Coral-deep | `#C9362B` | the faint diagonal wave inside coral hero fields |
| Off-white | `#FAFAFA` | content slides behind white cards |
| Ink | `#111111` | body and headings on light |

## Type
**Inter** (matched to the site; swap the single `--font` / `THEME["font"]` / docx font if the licensed
family differs). Sentence-case headlines, tight tracking; uppercase tracked eyebrows.

## Motifs used
**Coral hero field** (with a `#C9362B` diagonal wave), the **"M" chip** (coral rounded square + white
Inter-800 "M"), **pill** buttons/tags, **rounded photography** blocks, **KPI cards**, ~12px radii.
No accent line under titles; left-aligned body; no emoji; no purple.

## Generated assets
`make_assets.py` builds the embedded PNGs (`mchip*.png`, `hero_*.png`, `photo_*.png`) used by the PPTX
and DOCX builders. They're committed so `build_pptx.py` runs without a separate step; re-run
`python make_assets.py` to regenerate them (e.g. after a palette change).

## What makes it distinct
It's the live brand in motion — coral fields, M-chip cards and pills give it the warmth and momentum
of the website, where Clarity is airy and Editorial is high-contrast print.

## Files
`momentum.css` · `presentation.html` (8-slide 16:9 deck) · `document.html` (A4) · `build_docx.js`
+ `momentum.docx` (native Word) · `build_pptx.py` + `momentum.pptx` (PowerPoint) · `make_assets.py`
+ asset PNGs.

## Make a new artifact
```html
<link rel="stylesheet" href="momentum.css">
<section class="slide cover cover--coral"><span class="mchip">M</span>
  <h1 class="cover__h">Technology, human first.</h1></section>
```
```bash
python make_assets.py        # (once) regenerate embedded PNGs if the palette changed
python build_pptx.py                                            # -> momentum.pptx
NODE_PATH=/opt/node22/lib/node_modules node build_docx.js       # -> momentum.docx
# HTML: open presentation.html / document.html, then Print → Save as PDF
```

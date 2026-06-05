# Mivada Design Systems — for fast, on-brand artifacts in Claude

Three production-ready design systems for generating **HTML pages, HTML presentations,
A4 documents and PowerPoint decks** that reflect Mivada's look and feel — built to be
produced *fast* and with *managed tokens* in Claude, and engineered so the output
**does not read as AI-generated**.

> Mivada is an Australian technology consultancy — Workday & ERP, payroll, data & AI,
> and intelligent automation — positioned "human first." These systems interpret that
> brand. Drop official brand hex values and fonts in where noted once you have them.

## The three concepts

| | Concept | Feel | Type | Palette | Use when |
|--|---------|------|------|---------|----------|
| **01** | **[Meridian](01-meridian/)** | Editorial consulting — calm, authoritative | Fraunces + Archivo | Ink-navy + teal on warm white | Proposals, capability statements, board decks |
| **02** | **[Signal](02-signal/)** | Data/AI — precise, modern, technical | Hanken Grotesk + JetBrains Mono | Near-black + signal-mint, dot-grid | Data & AI pitches, product/eng, conference talks |
| **03** | **[Village](03-village/)** | Human-first — warm, approachable, premium | Gabarito + Source Serif 4 | Oat + sage + amber, espresso ink | People/change, culture, recruiting, client care |

Each folder contains: a reusable CSS theme, an HTML presentation, an HTML A4 document,
a python-pptx builder + a generated `.pptx`, and a README spec. They all present the
**same Mivada content** (see [`guides/mivada-content.md`](guides/mivada-content.md)) so
you can compare them on design alone.

## Why this is fast + token-light

Styling is written **once** per concept (`<concept>.css` + a PPTX `THEME` + helpers).
New artifacts emit **content only** and inherit the look:

```html
<!-- a whole on-brand slide -->
<link rel="stylesheet" href="meridian.css">
<section class="slide slide--stat">
  <p class="kicker">02 — Who we are</p>
  <h2>We turn enterprise platforms into <em>human outcomes</em>.</h2>
  <ul class="stats"> … </ul>
</section>
```

```python
# a whole on-brand PPTX deck
from build_pptx import Deck
d = Deck()
d.cover("Mivada", "Technology, human first.", "Capability overview · 2026")
d.kpis("Who we are", [("2014","Founded"), ("120+","Specialists"), ("AU + India","Delivery")])
d.save("mivada.pptx")
```

Claude writes a few dozen lines of content, not hundreds of lines of styling.

## The anti-"AI look" discipline

Every artifact is checked against [`guides/anti-ai-tells.md`](guides/anti-ai-tells.md) —
a concrete checklist of the patterns that scream "a model made this" (purple gradients,
Inter/Poppins+Lora, emoji bullets, accent-line-under-title slides, beige defaults,
centered hero + two pills, hype copy) and the fix for each.

## How to use a concept

1. Pick a concept from the table.
2. **HTML page / presentation / document:** copy the relevant `.html`, keep the
   `<link>` to the concept CSS and the Google Fonts `<link>`, replace the content.
   Open in a browser; "Print → Save as PDF" exports decks (one slide/page) and A4 docs.
3. **PowerPoint:** edit the content calls in `build_pptx.py` and run it. Install the two
   named Google fonts for pixel-perfect rendering (otherwise PowerPoint uses the listed
   Office-safe fallbacks).

## Guides

- [`guides/mivada-content.md`](guides/mivada-content.md) — shared positioning, voice, and the example copy.
- [`guides/anti-ai-tells.md`](guides/anti-ai-tells.md) — the "doesn't look AI-generated" checklist.
- [`guides/build-conventions.md`](guides/build-conventions.md) — structure, fonts, HTML/PPTX rules.

## Swapping in official Mivada assets

When you have the real brand kit: (1) replace the palette hex values at the top of each
`<concept>.css` and the `THEME` dict in `build_pptx.py`; (2) swap the Google Fonts link +
`--font-*` variables for the licensed brand fonts; (3) drop the logo into the masthead/cover
slots. Nothing else needs to change.

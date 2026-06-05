# Mivada design systems — fast, on-brand artifacts that don't look AI-generated

Three design systems for generating **HTML pages & presentations, Word documents, and PowerPoint
decks** that are **true to Mivada's real brand** (colours and fonts taken from mivada.com), are fast
and token-light to produce in Claude, and are engineered so the output **doesn't read as AI-generated**.

All three are **one brand, three layout systems** — identical palette and typeface, distinct
composition. The brand reference is [`guides/mivada-brand.md`](guides/mivada-brand.md).

## The real Mivada brand (extracted from the live site)

- **Coral `#EA493F`** (primary accent) · **Black `#000000`** · **Off-white `#FAFAFA`** · warm-neutral
  `#F2F3EE` · ink `#111111` · graphite `#323232`. Coral is the decisive 10% accent; the brand lives
  comfortably on black and off-white.
- **Type:** a clean neo-grotesque — matched with **Inter** (the licensed family isn't exposed in the
  site CSS; swap one variable if it differs).
- **Motifs:** the coral **"M" chip**, **pill buttons**, **coral hero fields**, **rounded photography**,
  full-**black sections**, uppercase eyebrow labels.

## The three concepts

| | Concept | Expression | Use when |
|--|---------|-----------|----------|
| **01** | **[Clarity](01-clarity/)** | Corporate / clean — light, airy, coral as a disciplined accent | Proposals, capability statements, board packs |
| **02** | **[Editorial](02-editorial/)** | Bold magazine — black-dominant + coral, oversized headlines, pull-quotes | Thought leadership, marketing, reports |
| **03** | **[Momentum](03-momentum/)** | The site's own energy — coral hero fields, M-chip cards, pills, rounded photography | Sales pitches, brand-forward decks, dynamic one-pagers |

Each concept ships every format:

| File | Format |
|------|--------|
| `presentation.html` | HTML presentation (16:9, keyboard nav, print-to-PDF) |
| `<concept>.pptx` + `build_pptx.py` | PowerPoint |
| `<concept>.docx` + `build_docx.js` | Word — native `.docx` |
| `document.html` | Word — HTML A4 (print-to-PDF) alternative |
| `<concept>.css` | the reusable theme (token-light core) |
| `README.md` | the concept spec |

All present the **same Mivada content** ([`guides/mivada-content.md`](guides/mivada-content.md)) so you
compare them on design alone. Figures in the examples are illustrative samples.

## Why it's fast + token-light

Styling is written **once** per concept (the CSS theme, the PPTX `THEME`/helpers, the docx style block).
New artifacts emit **content only** and inherit the look — Claude writes a few dozen lines, not hundreds.

## Not looking AI-generated — with a real brand font

[`guides/anti-ai-tells.md`](guides/anti-ai-tells.md) lists the patterns that scream "a model made this."
Because we're being faithful to Mivada's actual (grotesque) typeface, the "don't look generic" work is
carried by the **bold coral/black colour system, the distinct per-concept layouts, and the brand motifs**
(M-chip, pills, coral fields, rounded photography) — not by exotic fonts. No purple, no emoji, no
accent-line-under-title slides, no centered-hero-plus-two-pills filler.

## How to use a concept

1. Pick a concept. 2. **HTML / presentation / A4 doc:** copy the `.html`, keep the `<link>`s, replace the
content; open in a browser; "Print → Save as PDF" exports decks (one slide/page) and A4 docs. 3.
**PowerPoint:** edit the content calls in `build_pptx.py` and run it. 4. **Word:** edit `build_docx.js`
and run `NODE_PATH=/opt/node22/lib/node_modules node build_docx.js`.

For pixel-perfect PowerPoint/Word, install **Inter** (free, Google Fonts); otherwise Office uses Arial.

## Guides

- [`guides/mivada-brand.md`](guides/mivada-brand.md) — the real brand: exact palette, fonts, motifs, voice.
- [`guides/mivada-content.md`](guides/mivada-content.md) — positioning, voice, example copy.
- [`guides/anti-ai-tells.md`](guides/anti-ai-tells.md) — the "doesn't look AI-generated" checklist.
- [`guides/build-conventions.md`](guides/build-conventions.md) — structure, fonts, HTML/PPTX/Word rules.

## Swapping in official assets

When you have the real brand kit: (1) the palette is already exact; (2) if the licensed font differs from
Inter, change the single `--font` variable in each `.css`, the `THEME["font"]` in `build_pptx.py`, and the
font in `build_docx.js`; (3) drop the real logo into the masthead/cover and replace the photo-placeholder
blocks with Mivada photography. Nothing else changes.

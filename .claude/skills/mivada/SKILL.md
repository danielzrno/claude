---
name: mivada
description: >-
  Generate on-brand Mivada documents in the Editorial look & feel — presentations
  (PowerPoint .pptx and self-contained HTML 16:9 decks), capability/one-pager
  documents (native Word .docx and print-ready HTML A4), and translating an existing
  PowerPoint or document into the Mivada style without losing information. Use whenever
  the user asks to create, build, brand, or restyle a Mivada deck, slide, proposal,
  capability statement, report, or one-pager, or to "make this on-brand / Mivada-styled".
---

# Mivada document generator (Editorial)

Produces artifacts that are **true to Mivada's real brand** (extracted from mivada.com),
**fast** to generate (write content only — the look is inherited), and **don't read as
AI-generated**. The house style is **Editorial** (concept 02): black-dominant + coral,
oversized sentence-case headlines, a black callout band, a coral pull-quote.

Repo root for all paths below: the repository top level (where `design-system/` lives).

## Brand quick-reference (authoritative source: `design-system/guides/mivada-brand.md`)

- **Colours (exact):** coral `#EA493F` (the decisive 10% accent), coral-deep `#C9362B`,
  black `#000000`, ink `#111111`, off-white `#FAFAFA`, white `#FFFFFF`, warm-neutral
  `#F2F3EE`, graphite `#323232`, hairline `#E4E4E0`.
- **Type:** **Inter** (headlines 700–800, sentence case, tight `-0.02em`; body 400; eyebrows
  600 uppercase tracked `0.14em`). Office falls back to Arial. One `--font` to swap.
- **Motifs:** white/black **sandwich** of slides, oversized headlines, a deliberate **black
  callout band**, a large **coral pull-quote**, big coral index numerals, **logos** (see below).
- **Voice:** plain, confident, human, Australian English. Pair a human claim with a concrete
  system outcome. Real lines: "Where human understanding turns into system advantage."; "We
  Listen. We Deliver. We're Locals."; "Experience the Mivada Difference."
- **Never:** purple, emoji, an accent line directly under a title, centred hero + two pills,
  hype words, justified/centused body. Keep figures illustrative unless the user gives real ones.

## Logos — always use the real art, right variant per background

Logo files live in `design-system/assets/logos/` (see its README for filenames + the full
placement matrix). **Rule:** dark/coral background → **white** logo; light background → **coral**
(or **black** beside other coral); **wordmark** for covers/mastheads, **"M" mark** for slugs/
footers/chips. If a logo file is absent the templates fall back to the CSS/Pillow "M" chip —
prefer the real files. Never place coral-on-coral or white-on-white.

## Workflows

Pick the matching workflow; read its reference file before building.

1. **New presentation** → PowerPoint `.pptx` and/or HTML deck.
   Read `editorial-system.md`, then compose from the **standard slide library**
   (`design-system/02-editorial/slide_library.py`, `import slide_library as L`) + the `Deck` methods
   in `build_pptx.py` — write content-only calls and run. The catalogue of every standard slide is
   `samples/slide_library_demo.py` → `Editorial_Slide_Library.pptx`; `samples/ams_slides.py` and
   `samples/mivada_full_deck.py` are worked examples. For HTML, copy `presentation.html` + link `editorial.css`
   (the HTML catalogue of every standard slide is `02-editorial/slide-library.html`).
   **Fastest start:** `python design-system/make_deck.py list` shows every standard slide;
   `python design-system/make_deck.py new <name>` scaffolds an editable deck script.

2. **New document** → native Word `.docx` and/or HTML A4.
   Read `editorial-system.md` (docx helpers). Start from
   `design-system/02-editorial/build_docx.js` / `document.html`.

3. **Translate an existing deck/document into Editorial** (no information lost).
   Read `translate-deck.md` and follow it exactly — extract everything first, map slide-by-slide,
   then rebuild in the Editorial system and reconcile against the source.

4. **Logos / brand assets** → read `design-system/assets/logos/README.md`.

## Build & verify (always QA before declaring done)

```bash
# PowerPoint
cd design-system/02-editorial && python build_pptx.py            # -> editorial.pptx
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf editorial.pptx
pdftoppm -jpeg -r 110 editorial.pdf s   # then READ s-*.jpg and fix overflow/contrast

# Word
NODE_PATH=/opt/node22/lib/node_modules node build_docx.js        # -> editorial.docx
python /mnt/skills/public/docx/scripts/office/validate.py editorial.docx
python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf editorial.docx

# HTML: open the .html and "Print → Save as PDF" (decks print one slide/page)
```

Follow `design-system/guides/build-conventions.md` (structure, fonts, PPTX/Word rules) and
sanity-check against `design-system/guides/anti-ai-tells.md` before finishing. Clean temp
PDFs/JPGs; commit only the source + generated `.pptx`/`.docx`.

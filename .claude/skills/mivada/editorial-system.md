# Editorial system — API reference

The reusable theme lives in `design-system/02-editorial/`. Generate **content only**; the look
is inherited. The canonical, always-accurate examples are the builders themselves — read the
`build()` function at the bottom of `build_pptx.py` and `build()` in `build_docx.js` for exact
argument shapes before you call anything.

## PowerPoint — `build_pptx.py`

A `THEME` dict + a `Deck` class. Each method adds one slide; chain them, then `save()`.

```python
from build_pptx import Deck          # or copy build_pptx.py and edit build()
d = Deck()
d.cover(brand, line_pre, line_accent, eyebrow, footL, footR)   # black cover, oversized headline (accent = coral word)
d.kpis(eyebrow, headline_pre, headline_accent, standfirst, stats)   # off-white, headline + big stats row
d.pillars(eyebrow, headline_pre, headline_accent, items)       # black, coral index numerals + items
d.steps(eyebrow, headline_pre, headline_accent, steps)         # numbered sequence (Listen→Design→Deliver→Care)
d.split(eyebrow, headline_pre, headline_accent, standfirst, caps)   # black, two-column capability list
d.pullquote(eyebrow, quote, attr, stats)                       # large coral pull-quote + stats
d.reasons(eyebrow, headline_pre, headline_accent, items)       # black, 3 reasons
d.contact(brand, eyebrow, line_pre, line_accent, footL, footR) # black closing CTA
d.save("out.pptx")
```

- `headline_pre` + `headline_accent`: the headline splits into ink/white text + the **coral** word(s).
- `stats`: list of `(value, label)` (some accept a third "coral"/variant flag — check the signature).
- Slides alternate **black ↔ off-white** (the Editorial sandwich); keep that rhythm.
- Internals you may reuse: `_logo(...)` / `_m_chip(...)` (brand mark), `_eyebrow`, `_slug`,
  `_rect`, `_hline`, `_run`. Backgrounds via `_slide(bg_hex)`; colours via `C("coral")`.
- 16:9, `THEME["font"] = "Inter"`. **No drop shadows** in Editorial. **No rule under titles.**

## Word — `build_docx.js` (docx-js, Node)

Tokens at top (`CORAL="EA493F"`, `INK`, `BLACK`, …, `FONT="Inter"`). Compose paragraphs with the
helpers, push into `build()`'s `children`, then it writes the file.

```js
run(text, { size, bold, color, italic, caps, spacing })   // a styled TextRun
eyebrow(text, { color, before, after })                   // uppercase tracked label
headline(parts, { before, after, size })                  // oversized headline (parts = runs; coral word inside)
rule(color, size, before, after)                          // a paragraph bottom-border rule (NOT a table)
body(text, opts)                                          // body paragraph
```

Editorial doc motifs: oversized lede, a two-column intro/pillar table, a **black callout band**
(a full-width black `TableCell` with white runs), a **coral pull-quote** (left coral border +
italic), a KPI table with a coral header row. A4 (`CONTENT_W = 9304` DXA). Run with
`NODE_PATH=/opt/node22/lib/node_modules node build_docx.js`. Validate with the docx skill's
`validate.py`. Never use unicode bullets — use docx `numbering`.

## HTML (deck + A4 doc) — link `editorial.css`

Self-contained; `presentation.html` is the 16:9 deck, `document.html` the A4. Key classes:

- **Slides:** `section.slide` + `.on-black` or `.on-offwhite` (the sandwich). `.stagger` children
  animate in; honors `prefers-reduced-motion`. `.slug` = corner slide label; `.cover-top` = cover header row.
- **Type:** `.eyebrow`, `.headline`, `.display` (cover), `.standfirst`; `.accent` = coral span; `.index` = big coral numeral.
- **Components:** `.stat`/`.stat__num`/`.stat__label` in `.stats-row`; `.pillar`(`__title`/`__tag`/`__desc`);
  `.step`(`__n`/`__name`/`__desc`) in `.seq`; `.cap`(`__k`/`__v`); `.reason`(`__n`/`__h`/`__b`);
  the pull-quote section; `.col-rule`, `.line`.
- **Brand mark:** `.wordmark` / `.wordmark--rev` (reversed for dark), `.m-chip` / `.m-chip--lg`.
  When real logos exist, swap these for `<img>` of the correct variant (see logos README).
- Deck nav: ← / → / Space; scroll-snap; slide counter; prints one slide per page.

## Adding logos (when files are in `design-system/assets/logos/`)

- **HTML:** replace the `.wordmark`/`.m-chip` element with `<img class="logo" src="…assets/logos/mivada-wordmark-white.svg" alt="Mivada">`, choosing the variant per the placement matrix; size by height (e.g. `height:28px` cover, `18px` slug).
- **PPTX:** in `_logo()`, `slide.shapes.add_picture(path, x, y, height=Inches(h))` (height only → preserves aspect); pick `*-white.png` on dark/coral, `*-coral.png`/`*-black.png` on light; fall back to `_m_chip` if the file is missing.
- **Word:** `new ImageRun({ data: fs.readFileSync(path), transformation: { width, height } })` in the masthead/footer; pick the variant by band colour.
- Always preserve aspect ratio; never place coral-on-coral or white-on-white.

## General-purpose slides (for translating arbitrary decks)

Beyond the bespoke Mivada slides, `Deck` has general builders that take free content — use these
when restyling someone else's deck so no text is dropped:

- `cover_plain(eyebrow, headline_pre, headline_accent, footL, footR)` — black cover, any headline.
- `contact_plain(eyebrow, headline_pre, headline_accent, footL, footR)` — black closing, any headline.
- `content(eyebrow, headline_pre, headline_accent, items, dark=False, slug_txt=None)` — title + a
  coral-bulleted body that **auto-shrinks** to hold an arbitrary number of lines (the lossless catch-all).
- `quote_plain(eyebrow, quote, attr, dark=False)` — a pull-quote with no stats required.
- `chart_coral(eyebrow, headline_pre, headline_accent, categories, series, slug_txt=None)` — a column
  chart re-coloured to the brand (coral primary), all data preserved.

These power the automated translator (`design-system/samples/auto_translate.py`).

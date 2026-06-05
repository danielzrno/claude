# Mivada — Concept 01 · Clarity

Corporate, clean, restrained. Off-white pages, white cards, ink text, and coral
held back to a disciplined accent. Structured grids, generous white space,
hairline dividers, and one black slide as the closing bookend. Calm and
board-ready.

> One of three Mivada concepts. All three share the **exact same brand tokens**
> (palette + Inter); they differ only in layout, density and colour-blocking.

## Use when

Proposals, capability statements and board packs — anything that needs to read
as measured and trustworthy rather than loud. When the audience is a buying
committee or a board, reach for Clarity.

## Palette (shared across all concepts — exact, sampled from mivada.com)

| Token | Hex | Role |
|-------|-----|------|
| Coral | `#EA493F` | Primary — the decisive 10%. Eyebrows, one key word, the M-chip, a short stat rule. |
| Coral deep | `#C9362B` | Gradient / hover partner. |
| Black | `#000000` | Closing bookend slide. |
| Ink | `#111111` | Body + headings on light (soft black, not pure). |
| Off-white | `#FAFAFA` | Primary page background. |
| White | `#FFFFFF` | Cards / raised surfaces. |
| Warm neutral | `#F2F3EE` | Photo placeholder panels. |
| Graphite | `#323232` | Strong secondary text. |
| Mid grey | `#AEAEAE` | Muted captions, slide chrome. |
| Hairline | `#E4E4E0` | Borders / rules on light. |

**60/30/10:** off-white + ink dominate; warm-neutral + grey support; coral is
the sharp 10% — it punctuates, it never floods.

## Type

**Inter** across the whole system (one family — faithful to the site, which uses
a single neo-grotesque for headings and body).

- Headlines: 700–800, **sentence case**, tight tracking (`-0.02em`).
- Body: 400; weight contrast carries the hierarchy.
- Eyebrows: 600 **uppercase**, tracked `0.14em`.

HTML loads Inter from Google Fonts (`wght@400;500;600;700;800`). Office / docx
fall back to **Arial / Helvetica Neue** — PowerPoint and Word substitute
automatically if Inter isn't installed locally.

**Swap note:** if Mivada's real licensed face differs, change it in one place —
`--font` (`clarity.css`), `THEME["font"]` (`build_pptx.py`), or the `FONT`
constant (`build_docx.js`). Nothing else moves.

## Motifs used

- **The M-chip** — coral rounded-square (~20% radius) with a white Inter-800 "M".
  On the cover and the closing bookend, never decoratively repeated.
- **Eyebrow labels** — uppercase, tracked, coral; above every section heading.
- **Hairline dividers** (`#E4E4E0`) — separate stats and list rows.
- **Short coral rule** — a 34px bar beside a credential, the one place a coral
  line appears (and never under a title).
- **White cards** — hairline border, no shadow.
- **Black closing slide** — white text, coral accent.

## What makes Clarity distinct

Where Editorial goes black-dominant and oversized, and Momentum leans into coral
fields and KPI tiles, **Clarity stays light and quiet**. The discipline is the
design: coral is rationed, the grid is visible, and the whitespace does the
talking. Crucially — **no accent line under any title** (the AI-slide tell), no
drop shadows anywhere, no centred hero, body always left-aligned, headlines in
sentence case. The single black bookend is the only dark moment, so it lands.

## Files

| File | What it is |
|------|-----------|
| `clarity.css` | Brand tokens + reusable components (eyebrow, M-chip, hairline, stat, card, pill). The token-light core. |
| `presentation.html` | Self-contained 16:9 deck (8 slides). Links `clarity.css` + Inter. Keyboard nav (←/→/Space), scroll-snap, slide counter, one staggered CSS entrance per slide; prints one slide per page. |
| `document.html` | A4 capability statement. Masthead, two-column body, hairline-bordered "By the numbers" block. `@page A4`, exact colour. |
| `build_pptx.py` | python-pptx `THEME` dict + `Deck` helpers (`cover`, `kpis`, `pillars`, `steps`, `split_caplist`, `stats3`, `reasons`, `contact`). Builds `clarity.pptx`. |
| `clarity.pptx` | Generated PowerPoint (8 slides). |
| `build_docx.js` | docx-js native Word capability statement (A4, Inter, coral KPI table, paragraph-border rules, header/footer). Builds `clarity.docx`. |
| `clarity.docx` | Generated Word document. |
| `README.md` | This file. |

## Build

```bash
python build_pptx.py                                    # → clarity.pptx
NODE_PATH=/opt/node22/lib/node_modules node build_docx.js   # → clarity.docx
# open presentation.html / document.html in a browser; Print → Save as PDF
```

## Make a new artifact (reuse the theme, write only content)

```python
from build_pptx import Deck
d = Deck()
d.cover("Quarterly review · 2026",
        [[("Smarter, faster ", 70, "ink", True, -0.02),
          ("ways", 70, "coral", True, -0.02)], [("to work.", 70, "ink", True, -0.02)]],
        "How we cut a payroll cycle from days to hours.")
d.reasons("Highlights", "Three that mattered.",
          [("01", "Faster close", "Month-end from 9 days to 3."),
           ("02", "Cleaner data", "One governed KPI store."),
           ("03", "Higher adoption", "94% self-service in week one.")], "02")
d.save("review.pptx")
```

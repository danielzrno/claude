# Mivada — Concept 03 · Momentum

**The site's own energy.** Coral hero fields, M-chip tiles, pill buttons, rounded
photography blocks and KPI cards — the brand-forward, dynamic expression and the
concept that sits **closest to the live mivada.com**. Energetic but premium.

> One of three concepts that share Mivada's exact palette and typeface and differ
> only in layout, density and colour-blocking. See `../guides/` for the brand,
> content and anti-AI references.

## Use when

Sales pitches, brand-forward decks, dynamic one-pagers, capability overviews —
anywhere the room should *feel* the brand. (For calm board packs use **01 Clarity**;
for bold thought-leadership use **02 Editorial**.)

## Palette (shared across all three concepts — identical)

| Hex | Token | Role |
|-----|-------|------|
| `#EA493F` | `--coral` | **Primary.** Hero fields, the M-chip, KPI/accents — the decisive 10%. |
| `#C9362B` | `--coral-deep` | Hero diagonal wave, hover/gradient partner. |
| `#000000` | `--black` | Black sections (premium moments). |
| `#111111` | `--ink` | Body + headings on light (soft black, not pure). |
| `#FAFAFA` | `--off-white` | Primary page background. |
| `#FFFFFF` | `--white` | Cards / raised surfaces. |
| `#F2F3EE` | `--warm-neutral` | Warm off-white panels, tags. |
| `#323232` | `--graphite` | Strong secondary text. |
| `#AEAEAE` | `--mid-grey` | Muted captions, slide numbers. |
| `#E4E4E0` | `--hairline` | Borders / rules on light. |

**60/30/10:** off-white + ink dominate; warm-neutral + grey support; **coral is the
sharp 10%** — used decisively as a field, a chip or one word, never spread thin.

## Type — Inter (+ swap note)

One family across the system: **Inter** (headlines 700–800, sentence case, tight
`-0.02em`; body 400; eyebrows 600 uppercase, tracked `0.14em`).

- **HTML** loads Inter from Google Fonts (`wght@400;500;600;700;800`).
- **Office / docx** fall back to **Arial** when Inter isn't installed.
- **Swap:** if Mivada's real licensed face differs, change the single `--font`
  variable (`momentum.css`) / `THEME["font"]` (`build_pptx.py`) / the docx `FONT`
  constant (`build_docx.js`). Nothing else changes.

## Signature motifs

- **The "M" chip** — coral rounded-square (`~20%` radius) + white Inter-800 "M".
  The mark on covers, mastheads and every card. (`mchip.png` / `mchip_white.png`
  are Pillow-rendered for crisp Office embedding; CSS draws it natively for HTML.)
- **Coral hero field** — a full coral panel with a faint darker `#C9362B` diagonal
  wave (CSS gradients in HTML; `hero_coral.png` in Office). Covers + closing.
- **Cards with M-chip tiles** — the site's "We Listen / We Deliver / We're Locals"
  pattern: white surfaces, one tuned soft shadow, an M-chip per tile.
- **Pill buttons & tags** — fully-rounded; coral = active, plus dark and ghost-on-field.
- **Rounded photography blocks** (`~12px`) — tasteful neutral or **coral-duotone**
  placeholder blocks; **no network images** (`photo_*.png`, Pillow-generated).
- **KPI cards** — big coral number, quiet label; white / coral / ink variants.
- **Eyebrow labels** — small uppercase tracked labels above headings. **No accent
  line under titles** (that's an AI tell — whitespace and fields do the work).

### What makes Momentum distinct (and closest to the site)

Where Clarity is disciplined and Editorial is black-and-bold, Momentum leads with
the brand's own warmth: a coral field opens and closes the story, the M-chip recurs
as a tile, content lives in rounded cards and pills, and figures land as KPI cards.
Radius is held to one curve language; the only shadow is one tuned soft card shadow.
Coral is used decisively but never as a thin rainbow. It reads like the studio that
built mivada.com made it — energetic, human, premium.

## Files

| File | What it is |
|------|-----------|
| `momentum.css` | Brand tokens + components (coral field, M-chip, card, pill/tag, photo block, KPI). The token-light core both HTML files link. |
| `presentation.html` | Self-contained 16:9 deck, 8 slides. Keyboard nav (← / → / Space), scroll-snap, slide counter, one CSS staggered entrance per slide; prints one slide/page. |
| `document.html` | A4 capability statement — coral masthead, pillar cards, duotone photo, KPI cards, pill tags. `@page A4`, print-color exact. |
| `build_pptx.py` | python-pptx `THEME` + `Deck` helpers (`cover_coral`, `kpis_cards`, `pillars_chips`, `steps`, `split_photo`, `stats3_quote`, `reasons`, `contact`) → `momentum.pptx`. |
| `build_docx.js` | docx-js native Word builder (coral masthead band, card-bordered sections, KPI table with coral header) → `momentum.docx`. |
| `make_assets.py` | Pillow generator for the embedded PNGs (M-chips, coral hero field, photo blocks). Re-run if you re-tint the brand. |
| `momentum.pptx` / `momentum.docx` | Generated, committed. |
| `*.png` | Embedded assets used by the Office files. |

### Build / regenerate

```bash
python make_assets.py                                       # (re)build embedded PNGs
python build_pptx.py                                        # -> momentum.pptx
NODE_PATH=/opt/node22/lib/node_modules node build_docx.js   # -> momentum.docx
# HTML: open presentation.html / document.html; "Print → Save as PDF".
```

## Make a new artifact (content only — the look is inherited)

```python
from build_pptx import Deck
d = Deck()
d.cover_coral("Mivada", "Technology,\nhuman first.", "Capability overview · 2026")
d.kpis_cards("Who we are",
    [[d.run("We turn platforms into ", 33, "111111", True, -0.8),
      d.run("human outcomes", 33, "EA493F", True, -0.8)]],
    [("Built around how teams really work.", False)],
    [("2014","Founded"), ("120+","Specialists"), ("AU + India","Delivery","coral")])
d.save("new-deck.pptx")
```

> Figures shown across these artifacts are **illustrative samples** for the design
> examples — replace with verified numbers before any real use.

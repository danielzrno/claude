# Signal — Mivada design system 02

Data / AI. Dark by default, precise, technical. An engineering-led brand: the look of a
great developer-tool landing page, not a clichéd "techy gradient." Hairlines and a
dot-grid do the work; mint is the one sharp accent.

## Use when
Data & AI pitches, platform/engineering capability decks, technical proposals,
conference talks, product one-pagers. Reach for **Meridian** instead for board/editorial
work, **Village** for people/change/culture.

## Files
| File | What |
|------|------|
| `signal.css` | The reusable theme — tokens, type scale, dot-grid utility, deck + document components. Write once; HTML emits content only. |
| `presentation.html` | Self-contained 16:9 **dark** deck, 8 slides. Keyboard nav (← / → / Space), scroll-snap, slide counter, one CSS entrance per slide. Prints one slide/page. |
| `document.html` | A4 Capability Statement — **light** (ink on paper). Print-ready. |
| `build_pptx.py` | python-pptx builder: `THEME` dict + `Deck` helpers. Generates a dot-grid PNG and builds the 8 slides. |
| `signal.pptx` | Generated deck (run the script to rebuild). |

## Palette
60 / 30 / 10 — ink dominates, surface + slate support, mint is the sharp accent. Amber is a rare second accent (one emphasis only). Never pure `#000` / `#fff`. No purple, no blue→purple gradients.

| Token | Hex | Role |
|-------|-----|------|
| `--ink` | `#0E1416` | near-black, faint cool-green undertone — primary (dark) background |
| `--surface` | `#161E21` | raised panel on dark |
| `--line` | `#28343A` | hairline / border on dark |
| `--paper` | `#ECEFEC` | off-white text on dark; also the **light** document background |
| `--mint` | `#34E2A8` | signal mint — **primary accent**: key word, KPI numerals, mono markers |
| `--slate` | `#7E97A1` | muted secondary text / secondary accent |
| `--amber` | `#E9B949` | warm highlight — rare 2nd accent, one emphasis only |

Light-document derivatives (in CSS): `--mint-d #128A63`, `--slate-d #50656E`, `--line-lt #C9D3CE` — mint/slate/hairline tuned dark enough to read on paper.

## Type
- **Display & body — Hanken Grotesk.** 800/700 tight headlines (sentence case), 400 body. Office fallback: **Arial**. Stack: `"Hanken Grotesk", Arial, system-ui, sans-serif`.
- **Mono — JetBrains Mono** (500). Kicker labels, KPI captions, figure/source tags, contact line. Office fallback: **Consolas**. Stack: `"JetBrains Mono", Consolas, monospace`.
- Weight contrast carries hierarchy; color the one word that matters, never the whole line.

**Install for pixel-perfect decks** (both free on Google Fonts): **Hanken Grotesk** and **JetBrains Mono**. PowerPoint substitutes Arial / Consolas automatically if they're missing. HTML loads them via a Google Fonts `<link>`.

## Motifs (carry across every artifact)
- Mono kicker labels with markers: `// 04 — data & ai`, arrows `→`, index tags `[03]`.
- **Dot-grid** background texture on dark — *the* signature (radial-gradient in CSS; a generated PNG for PPTX). Subtle; never competes with text.
- KPI tiles: hairline-bordered cells, big Hanken numerals in mint, mono caption under a hairline.
- Pipelines: hairline node cells joined by mint `→` flow arrows on a hairline track.
- Mono figure / source / attribution captions; a single mint "signal dot" (the one glow) on cover and contact.

## Dark-UI / light-print rule
**Screen is dark; print is light.** The deck (`presentation.html`, `signal.pptx`) is dark by default — it's premium and it's where Signal lives. The A4 document (`document.html`) inverts to **paper background, ink text, mint-dark accents, hairline rules**, with the dot-grid kept only as a faint header motif. Reason: dark print wastes toner and bands badly on office printers; a capability statement is meant to be printed and handed over. Same brand, two surfaces — chosen per medium, not by default.

## Radius / shadow / rule language
1px hairline borders (`--line`) define everything. Small radius (**4px**). **No soft drop shadows** — depth is layered surfaces + hairlines. The only glow is one subtle mint halo on a single hero element (the signal dot), used once. No accent line under any title; no full-width colored bars (dot-grid and tick-ruler motifs only).

## Make a new deck
```python
from build_pptx import Deck
d = Deck()
d.cover("mivada.com  //  capability overview",
        [[("Technology,", False)], [("human", True), (" first.", False)]],
        "An Australian technology consultancy.", "33.87°S  151.21°E", "2026")
d.pillars("02", "what we do", [("Four practices, one model.", False)], [...])
d.save("signal.pptx")
```
Helpers: `cover`, `who` + `kpis`, `pillars`, `track`, `pipeline`, `stats3` + `quote`, `reasons`, `contact`. A new deck is a list of helper calls — write the content, inherit the look.

> Figures throughout are illustrative samples for design, not Mivada's published metrics.

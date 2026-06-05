# Mivada — real brand reference (extracted from mivada.com)

Source: the live site (WordPress/Salient theme), captured as a rendered screenshot on
2026-06-05; **colours sampled directly from the page** (exact). The font is a clean
neo-grotesque matched visually — the licensed family name isn't exposed in the site's CSS,
so we use **Inter** as a faithful, swappable stand-in (change one variable if the real font differs).

This file is the source of truth for all three concepts. They share these tokens exactly and
differ only in **layout, composition and treatment**.

## Palette (exact, sampled)

| Hex | Role | Notes |
|-----|------|-------|
| `#EA493F` | **Primary — Mivada coral** | The signature. Hero fields, the "M" chip, key accents, active states. This is the 10% that must feel decisive. |
| `#C9362B` | Coral deep | Gradient/!hover partner; the hero is a coral field with a subtly darker diagonal wave. |
| `#000000` | Black | Dark sections — the brand's "OnBlack" home; logo reverses to white here. |
| `#111111` | Ink | Body text / headings on light. Soft black, not pure. |
| `#FAFAFA` | Off-white | Primary page background. |
| `#FFFFFF` | White | Cards / raised surfaces. |
| `#F2F3EE` | Warm neutral | Warm off-white section panels (sampled `#F2F3EE`/`#EEEFEA`). |
| `#323232` | Graphite | Strong secondary text, dark UI. |
| `#AEAEAE` | Mid grey | Muted text/captions. |
| `#E4E4E0` | Hairline | Borders / rules on light. |

**60/30/10:** off-white + black/ink dominate; warm-neutral + grey support; **coral is the sharp 10%**.
Use coral decisively (a field, a chip, one word) — never timidly spread thin.

## Typography

- **Primary face: Inter** (matched to the site's neo-grotesque). One family across the system —
  faithful to the site, which uses a single grotesque for headings and body.
  - Headlines: 700–800, **sentence case**, tight tracking (`-0.02em`).
  - Body: 400; Labels/eyebrows: 600 **uppercase**, tracked `0.12em`.
  - HTML: load Inter from Google Fonts. Office fallback: **Arial / Helvetica Neue**.
- **Swap note:** if Mivada's real licensed font differs, change the single `--font` variable
  (CSS) / `THEME["font"]` (PPTX) / the docx style font — nothing else.

## Signature motifs (reproduce these — they make it unmistakably Mivada)

1. **The "M" chip** — a **coral rounded-square tile with a white stylised "M"**. Appears as the logo
   mark on cards and covers. Recreate as a shape/SVG (coral `#EA493F`, ~20% radius, white M).
2. **Pill buttons** — fully-rounded buttons; the nav shows a dark pill group with a **coral active pill**.
3. **Coral hero field** — a full coral panel with a faint darker diagonal wave/gradient, white headline.
   Use for covers and section breaks.
4. **Rounded photography** — image blocks with ~10–14px rounded corners; warm, real-people, Australian
   workplace photography. Where no photo exists, use a tasteful neutral/duotone placeholder block.
5. **Black sections** — full-black panels, white text, coral accent — for premium moments (cover/closing).
6. **Eyebrow labels** — small uppercase tracked labels above headings.

## Brand voice (real lines observed on the site — echo this register)

- "Where human understanding turns into system advantage."
- "Smarter, faster ways to work."
- "We Listen. We Deliver. We're Locals."  *(their three-pillar framing — people-first, delivery, local)*
- "Experience the Mivada Difference."

Plain, confident, human, Australian English. Pair a human claim with a concrete system outcome.
(Positioning, services and the illustrative figures remain as in `mivada-content.md`.)

## The three concepts (one brand, three layout systems)

| Concept | Expression | Use when |
|---------|-----------|----------|
| **01 Clarity** | Corporate / clean — light, airy, coral as a disciplined accent, structured grids | Proposals, capability statements, board packs |
| **02 Editorial** | Bold magazine — black-dominant + coral, oversized headlines, asymmetric, pull-quotes | Thought leadership, marketing, reports |
| **03 Momentum** | The site's own energy — coral hero fields, M-chip tiles, pills, rounded photography, KPI cards | Sales pitches, brand-forward decks, dynamic one-pagers |

All three: identical palette + Inter; distinct layout, density and colour-blocking.

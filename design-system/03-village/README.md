# Village — Mivada design system 03

Human-first, warm, premium, approachable. The "my village / We Care" expression of
Mivada: warm editorial, people at the centre, still corporate-credible. **Sage +
amber + espresso on a named oat surface** — deliberately *not* the Anthropic
bone + coral palette.

## Use when

People & change work, culture and recruiting, client-care and account decks,
anything where warmth and the "human first" story carry the message. Pairs with
**Meridian** (formal proposals) and **Signal** (data/AI, technical).

## Why this isn't the Anthropic look

The single most important brand guardrail. Two deliberate moves keep Village clearly
distinct from the `#faf9f5` bone + `#d97757` coral signature:

1. **The primary accent is sage green (`#6E7E54`), not orange.** The 10% colour that
   carries the brand is olive/sage. Amber and clay are warm highlights used sparingly —
   never a coral-orange field.
2. **Oat (`#FAF5EC`) is an *intentional, named brand surface*** — a warm, slightly
   yellower paper chosen on purpose — not the accidental beige/cream default the
   anti-AI checklist warns about. It is warmer and more golden than bone, and it is
   always paired with espresso ink tuned for strong contrast.

## Palette

| Token | Hex | Role |
|-------|-----|------|
| `--ink` | `#241F1B` | Warm espresso brown-black — body text (never pure `#000`) |
| `--oat` | `#FAF5EC` | Soft warm oat — **primary brand surface** (named, intentional) |
| `--card` | `#FFFDF8` | Raised card on oat |
| `--sage` | `#6E7E54` | Sage/olive — **PRIMARY accent**, the 10% (clearly not orange) |
| `--sage-d` | `#586647` | Deeper sage — hover, fine text, pill labels |
| `--amber` | `#E0A33E` | Warm amber/gold — the hand-drawn underline motif + unit accents |
| `--clay` | `#B7593C` | Dusty clay — tertiary, one or two marks only |
| `--ink-60` | `#6E665C` | Secondary text |
| `--line` | `#E4DBC9` | Warm hairline |

**60 / 30 / 10:** oat + espresso dominate; sage supports as the brand accent; amber is
the sharp highlight; clay is rare. Covers/bookends may invert to espresso or sage.

## Type

| Role | Font | Office fallback | Stack |
|------|------|-----------------|-------|
| Display | **Gabarito** (700/800 headlines, 500 sub) | Trebuchet MS | `"Gabarito", "Trebuchet MS", system-ui, sans-serif` |
| Body | **Source Serif 4** (400 body, italic for warmth) | Georgia | `"Source Serif 4", Georgia, serif` |

**Install these two free Google fonts** for pixel-perfect output:
[Gabarito](https://fonts.google.com/specimen/Gabarito) +
[Source Serif 4](https://fonts.google.com/specimen/Source+Serif+4). The HTML files load
them via a Google Fonts `<link>`; the PPTX references the brand names and PowerPoint
substitutes the fallbacks automatically if a font is missing. Weight contrast (800
headline vs 400 serif body) carries the hierarchy — humanist serif sets the human tone.

## Signature motifs

- **One hand-drawn amber underline.** An inline SVG brush-stroke (`.mark`) under exactly
  **one** key word per artifact — "human" on the cover, "people" on contact, "human
  outcomes" in the document intro. Never a per-heading rule (that is the AI-slide tell).
- **Pill tags** — sage outline, fully rounded, tracked caps; for labels/categories.
- **Circular sage tokens** — for step numbers, stat numbers and avatar slots. The one
  place perfect circles live.
- **Custom sage check-mark** — a pale-sage disc + CSS/shape tick. Never an emoji.
- **Quiet `• • •` dot/leaf micro-divider** used sparingly beside eyebrows.
- **Warm humanist serif body, generous leading.**

## Radius / shadow language

Soft and warm but restrained. **Cards carry a 12–16px radius** (`--r: 14px`) — held
consistently; curves live on cards and circular tokens only. **One tuned warm soft
shadow**: `0 14px 34px rgba(60,40,20,.08)` (`--shadow`), low and soft, on raised cards
(and, subtly, circular tokens). No hard offsets, no shadow on flat elements like pills.

## Hard rules (carried from the shared anti-AI checklist)

No emoji. No accent line under any title (the amber word-underline is the deliberate
once-per-page motif). No full-width coloured header/footer bars. Left-align all body and
headings. Sentence-case headlines, tracked caps only on short labels. Tuned ink/paper,
never pure `#000`/`#fff`. No purple. No hype words — numbers over adjectives.

## Files

| File | What |
|------|------|
| `village.css` | The reusable theme: tokens, type scale, `.mark` underline helper, `.pill`, `.token`, `.card`, `.checks`, `.dots`/`.connector`, `.stat`. Shared by both HTML files. |
| `presentation.html` | Self-contained 16:9 deck, 8 slides. Keyboard nav (←/→/Space, Home/End), scroll-snap, slide counter, one CSS choreographed entrance per slide (incl. the underline drawing itself in). Prints one slide/page. |
| `document.html` | A4 print-ready Capability Statement. `@page A4`, `print-color-adjust: exact`. |
| `build_pptx.py` | python-pptx `THEME` dict + `Deck` class. Builds `village.pptx`. |
| `village.pptx` | Generated 8-slide deck. |

## Make a new deck (6 lines)

```python
from build_pptx import Deck
d = Deck()
d.cover("Mivada", "Technology, human first.", "Capability overview · 2026")
d.kpis("Who we are", [("Since 2014","Founded as LJM Infotech"),
                       ("120+ specialists","Certified, not generalists"),
                       ("AU + India","Onshore led, offshore scaled")])
d.save("mivada.pptx")
```

Other helpers: `pillars(title, items)`, `steps(title, sub, steps)`,
`split(title, headline, standfirst, cap_title, caps)`,
`stats3(title, headline, stats, quote, attrib, attrib_sub, initials)`,
`reasons(title, headline, items)`, `contact(headline_a, mark_word, headline_b, lines)`.
A new artifact is content + helper calls; the look is inherited.

> **Note on figures:** all statistics here are illustrative samples for the design
> examples, not Mivada's published metrics. Replace with verified numbers before use.

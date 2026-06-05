# Meridian — Mivada editorial-consulting design system

**Use when:** the artifact should read like a premium consulting report — proposals,
capability statements, board decks. Calm, structured, authoritative, lots of air.

A light theme with ink-navy cover/closing bookends. Depth comes from hairline rules and
warm-stone panels — never drop shadows. The accent (deep teal) is a sharp 10%; brass is rarer still.

## Palette

| Token | Hex | Role |
|-------|-----|------|
| `--ink` | `#14283A` | Deep slate-navy — primary text; background for cover & closing slides |
| `--paper` | `#FBFAF7` | Barely-warm white — primary background (a hair off white, not beige) |
| `--bone` | `#ECE7DB` | Warm stone — secondary panel / fill |
| `--teal` | `#18786A` | **Primary accent** — the sharp 10% (one italic key word, list terms, small marks) |
| `--brass` | `#A87C3D` | Secondary accent — very sparing (pillar indices, rare hairlines) |
| `--stone` | `#C8C1B2` | Hairline / rule colour |
| `--ink-60` | `#5A6675` | Secondary text |

Discipline is 60 / 30 / 10: paper + ink dominate, bone supports, teal is the accent, brass is rare. No purple, never pure `#000`/`#fff`.

## Type

- **Display — Fraunces** (variable; optical sizing). Weight 300 for big standfirsts and hero
  lines, 600/700 for headlines, *italic* for the single emphasis word. Office fallback: **Georgia**.
- **Body / label — Archivo.** 400 body, 500/600 labels; tracked caps for kickers
  (`letter-spacing: .14em`). Office fallback: **Calibri**.
- Stacks: `"Fraunces", Georgia, serif` · `"Archivo", Calibri, system-ui, sans-serif`.

**Install these two Google Fonts** for pixel-perfect decks (PowerPoint falls back to Georgia/Calibri otherwise):

- Fraunces — https://fonts.google.com/specimen/Fraunces (load axes `opsz, ital, wght`)
- Archivo — https://fonts.google.com/specimen/Archivo

## Signature motifs

- Numbered sections: large light Fraunces numerals (`01 / 02 …`) as kickers, with a tracked
  Archivo caps label beside — never a rule under the kicker.
- A large light-weight Fraunces standfirst paragraph.
- Editorial grid with a wide left margin / marginal notes (especially in the document).
- Stats = big Fraunces serif numerals + small Archivo caps labels.
- One italic Fraunces emphasis word inside a headline, coloured teal — the single key word only.
- Hairline rules live in the **margins and between columns**, never directly under a title.

## Radius / shadow / rule language

- **Sharp:** radius `0–2px`.
- **No drop shadows at all.** Depth = hairline rules (`--stone`) + bone panels.
- Signature = thin rules in margins and between columns. **Never a rule directly under a title**
  (the #1 slide tell). A rule *under the masthead* of the document is fine — that's a masthead rule.

## Files

- `meridian.css` — the reusable theme (tokens + type scale + components for deck *and* document).
- `presentation.html` — self-contained 16:9 deck (8 slides), keyboard nav, scroll-snap, prints 1/page.
- `document.html` — A4 print-ready Capability Statement.
- `build_pptx.py` — `THEME` dict + `Deck` class; builds `meridian.pptx`.
- `meridian.pptx` — generated output.

## Make a new deck or document

```python
from build_pptx import Deck                 # PPTX: a deck is a short list of calls
d = Deck()
d.cover("Mivada", [[("Technology,", False)], [("human ", False), ("first.", True)]], "Capability overview · 2026")
d.kpis("01", "Who we are", [("We turn platforms into ", False), ("human", True), (" outcomes.", False)],
       "One-line standfirst here.", [("2014","Founded"), ("120+","Specialists"), ("AU + India","Delivery")])
d.save("meridian.pptx")                      # then: python build_pptx.py
```

```html
<!-- HTML: link the theme + fonts, emit content only -->
<link rel="stylesheet" href="meridian.css">
<section class="slide split">
  <div class="split__main">
    <div class="section-no"><span class="num">01</span><span class="label">Who we are</span></div>
    <h2 class="h-display">We turn platforms into <em>human</em> outcomes.</h2>
  </div>
</section>
```

# Mivada Standard — visual spec

One brand, two backgrounds (black & white), one accent (coral). Inter throughout. The deck
**alternates** black and white slides; `assemble.py` enforces it, `kit.Deck()` does it automatically.

## Tokens (exact)

| Role | Light slide | Dark slide |
|------|-------------|-----------|
| Background | off-white `#FAFAFA` | black `#000000` |
| Card fill | white `#FFFFFF` | `#141414` |
| Card border (1pt) | `#E4E4E0` | `#3A3A3A` |
| Title / heading text | ink `#111111` | white `#FFFFFF` |
| Body text | graphite `#323232` | `#CFCFCF` |
| **Accent (always)** | coral `#EA493F` | coral `#EA493F` |
| Connector / hairline | `#AEAEAE` | `#3A3A3A` |
| Emphasis band fill | ink `#111111` (white text) | `#1C1C1C` (white text) |

Font **Inter** (Office falls back to Arial): headings 700–800, tight `-0.02em`, sentence case;
body 400; eyebrow 700 uppercase, tracked. Numerals/labels can be coral.

## Header furniture (every content slide)

- **Logo** — coral "M" icon top-left at `0.96, 0.5` (`0.44×0.26"`); white "M" on dark slides.
- **Eyebrow** — coral, bold, 11pt, UPPERCASE, **right-aligned**, top-right (`6.5, 0.46`, w 6.0).
- **Title** — `0.92, 0.98`, 33pt bold, **two-tone**: ink/white first part + **coral** accent part,
  sentence case, ends in a period (e.g. *"Across the whole **Workday lifecycle.**"*).
- **Lead** — one line under the title, 12.5pt, graphite/`#CFCFCF`.

## The box idiom — the standard way to do boxes

A box = **fill + 1pt border + a coral label inside.** That's it. Reference: "Four ways we support
you" (dark `#141414`/`#3A3A3A`) and the light equivalent (white/`#E4E4E0`).

- ❌ **No left-edge accent stripe.** ❌ No drop shadow. ❌ No per-card coloured header band.
- The coral comes from the **label/numeral text** (and, sparingly, an outline or a filled chip).
- Variants in the vocabulary: **coral-outline** box (fill none, coral 1.5pt border — slide 7 bands);
  **coral-filled** node/chip (white text — org top-node, table header, milestone tag);
  full-width **black emphasis band** (coral lead + white body); **coral section header** bar.

## Components (see `kit.py`)

- `d.cards(slide, top, [(label, head, body), …])` — a row of standard boxes (stat/feature cards).
- `d.flow(slide, top, [(name, detail), …], emph=(…))` — boxes joined by coral `→` arrows; `emph`
  fills a step coral (use for the pivotal step).
- `d.grid(slide, top, cols, rows)` — a **boxed table** (coral header row, bordered cells). Build
  tables this way, not as native PPTX tables, so the black/white recolour flows through.
- `d.band(slide, top, lead, rest)` — full-width black emphasis band.
- Org chart = coral top-node + coral/grey-outline child chips joined by `#AEAEAE` connectors.
- Gantt = labelled rows; coral phase bars, black milestone, grey aftercare.

## Slide catalogue (keys for `assemble.py --slides`)

Natural theme in brackets — but alternation overrides it; any slide works on either background.

| # | key | what it is |
|---|-----|-----------|
| 1 | `logo-open` | logo open (dark) |
| 2 | `cover` | title cover — *Technology, human first.* (dark) |
| 3 | `who-we-are` | positioning + 3 KPIs (light) |
| 4 | `capability` | proof stats + Workday partner badges (light) |
| 5 | `trusted-by` | client logo wall (light) |
| 6 | `what-we-do` | numbered service pillars (dark) |
| 7 | `engagement-model` | the Workday lifecycle — advise→manage flow + bands (dark) |
| 8 | `engagement-models` | four ways we support you — 2×2 cards (dark) |
| 9 | `velocity` | Workday GO — 90/10 + five-phase flow (light) |
| 10 | `timeline` | 20-week implementation Gantt (light) |
| 11 | `data-handshake` | data certainty — handshake + load cycles (light) |
| 12 | `delivery-team` | two org charts, Mivada + customer (light) |
| 13 | `payroll-testing` | payroll rigour — highlights + parallel cycle (light) |
| 14 | `coverage` | full-stack Workday FIN & HCM — module grid (light) |
| 15 | `ams-itil` | AMS built on ITIL — three columns (dark) |
| 16 | `ams-framework` | AMS operating framework — capability grid (light) |
| 17 | `coverage-model` | onshore/offshore coverage + 24×7 band (light) |
| 18 | `governance-model` | Strategic/Tactical/Operational matrix (light) |
| 19 | `transition` | four-week transition with milestones (light) |
| 20 | `slas` | P1–P4 SLA grid (light) |
| 21 | `commercials` | pricing cards + blended-rate band (light) |
| 22 | `augmentation` | FIN/HCM bench availability (light) |
| 23 | `governance-pyramid` | governance tiers diagram (light) |
| 24 | `team-structure` | AMS org, onshore + offshore (light) |
| 25 | `outcomes` | proven Workday outcomes / case studies (light) |
| 26 | `why-mivada` | a partner, not a vendor — statement (dark) |
| 27 | `contact` | closing — let's talk (dark) |

## Voice & anti-tells

Plain, confident, human, Australian English; pair a human claim with a concrete system outcome.
Keep figures illustrative unless given real ones; keep customer names out unless asked.
**Avoid the tells:** left-edge accent stripes, drop shadows, purple, emoji, an accent line under a
title, centred hero + two pills, hype words, justified/centred body, all-light monotony (alternate!).

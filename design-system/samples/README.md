# Translation test case — legacy deck → Editorial

A worked example of the `/mivada` **translate-deck** workflow: take an existing, off-brand
PowerPoint and re-express **every** slide in the Editorial look without losing information.

> **Stand-in note:** the real example deck wasn't in the repo when this was built, so
> `legacy-overview.pptx` is a representative generic "before" deck (Office default look, Calibri,
> blue accents, bullets, a table, a bar chart, a quote). Re-run the translator on the real file
> the moment it's added — the workflow is identical.

## Files

| File | What it is |
|------|-----------|
| `make_legacy.py` | builds `legacy-overview.pptx` — the generic "before" deck (8 slides + speaker notes) |
| `legacy-overview.pptx` | the **before** (off-brand source) |
| `auto_translate.py` | **general** translator — classifies any deck, rebuilds in Editorial, reconciles (first pass) |
| `legacy-overview_editorial_auto.pptx` | the automated first-pass output (99% tokens; only the "Metric" header label dropped) |
| `translate_to_editorial.py` | **hand-tuned** worked example (richer per-slide mapping) |
| `legacy-overview_editorial.pptx` | the hand-tuned **after** (44/44 critical tokens) |

## Run

```bash
python make_legacy.py                          # (re)build the sample source
python auto_translate.py SOURCE.pptx           # GENERAL first pass -> SOURCE_editorial_auto.pptx + coverage
python translate_to_editorial.py SOURCE.pptx   # hand-tuned worked example -> SOURCE_editorial.pptx
```

Two approaches: **`auto_translate.py`** for a fast, lossless-as-possible first pass on *any* deck
(then refine the slides the report flags), and **`translate_to_editorial.py`** as a reference for what
a finished, hand-mapped translation looks like.

## What it does (and proves)

1. **Extract** — dumps each slide's text, tables, chart series/values and speaker notes.
2. **Map** — title→cover, about→headline+stats, services→coral-indexed pillars, approach→steps,
   the **table**→a 5-metric stat grid, the **bar chart**→a rebuilt coral column chart (all data kept),
   testimonial→coral pull-quote, contact→black closing. Logos placed per the brand matrix.
3. **Carry notes** — every speaker note moves across verbatim.
4. **Reconcile** — re-extracts the output and checks every critical token from the source survived.
   Current result: **44/44 carried — nothing lost.** (`Mivada` itself is carried by the wordmark logo.)

To translate a real deck, point the translator at it; where a slide doesn't fit an existing
Editorial template, **extend the system** (add a `Deck` method) rather than drop content — see
`.claude/skills/mivada/translate-deck.md`.

## Slide variations — `variations.py`

Three customer-facing versions each of the Micah deck's **Why Mivada**, **Workday Services**, and
**Workday Capability** slides (9 total), in the Editorial look — the original diagrams reimagined as
clean editorial compositions. Output: `slide-variations_editorial.pptx`.

- **Why Mivada** — v1 differentiator grid · v2 black manifesto ("we grew out of a customer") · v3 split (manifesto + capability rows). All close with a *Trusted by* line.
- **Workday Services** — v1 nine-service catalogue · v2 lifecycle flow (Advise→Implement→Optimise→Manage) · v3 three practices.
- **Workday Capability** — v1 KPI grid · v2 hero stat (10+ years) · v3 KPI cards.

```bash
python variations.py     # -> slide-variations_editorial.pptx (9 slides)
```

## Refined final deck — `micah_refined.py`

`First_Meeting_with_Micah_Projects_refined.pptx` (14 slides) — the single-version,
de-duplicated, reordered first-meeting deck (the keeper). Reworked intro: **white logo
slide → framed cover → agenda → Why-Mivada-as-a-statement**; pillars live only in Services,
numbers only in Capability, *trusted by* once; adds a "how we deliver" divider and a
next-steps close.

Flow: Logo · Cover · Agenda · Why Mivada · Your needs · Workday services · Capability ·
Case studies · *How we deliver* · Methodology · Implementation plan · Governance · RACI ·
Next steps.

```bash
python micah_refined.py    # -> First_Meeting_with_Micah_Projects_refined.pptx
```

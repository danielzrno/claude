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
| `translate_to_editorial.py` | the translator: extract → map → rebuild (Editorial `Deck` + logos) → reconcile |
| `legacy-overview_editorial.pptx` | the **after** (on-brand Editorial output) |

## Run

```bash
python make_legacy.py                              # (re)build the sample source
python translate_to_editorial.py SOURCE.pptx       # -> SOURCE_editorial.pptx  (defaults to legacy-overview.pptx)
```

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

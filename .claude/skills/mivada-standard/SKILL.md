---
name: mivada-standard
description: >-
  Build a Mivada presentation, proposal, AMS pack, capability deck, Word document or HTML deck
  in the Mivada STANDARD style, assembled from the 27-slide standard library (cover, who-we-are,
  engagement model, delivery team, AMS coverage/governance/SLAs/commercials, outcomes, …). Use
  whenever the user wants to create/build a Mivada deck or doc, a "standard" presentation, an AMS
  or Workday proposal, or asks for black & white / light & dark versions of slides. Always starts
  from the standard, asks which slides, and alternates black/white. (For the older Editorial
  document look, use the `mivada` skill instead.)
---

# Mivada — the Standard (deck & document builder)

Everything lives in **`design-system/standard/`** (paths below are relative to the repo root).
The **source of truth** is `Mivada_Standard.pptx` (27 curated slides); every other format is
generated from it. Read **`style.md`** (next to this file) for the full visual spec before editing
or building slides.

## The one rule about boxes (read this)

A box is a **filled/bordered rectangle with a coral label inside — nothing else.** Match the
"Four ways we support you" and "Across the whole Workday lifecycle" slides. **Never** add a thin
coral/ink **accent stripe down the left edge** of a card (it's an AI tell), never a drop shadow,
never a per-card header band unless the deck already uses one. Coral is the single accent, carried
by the **label text**, not by a stripe.

## Workflow — always follow this

1. **Read `style.md`** for tokens, the box idiom, components and the slide catalogue.
2. **Ask the user which slides they want** — this is required. Use `AskUserQuestion` with a
   multi-select built from the catalogue (`python design-system/standard/assemble.py --list`),
   plus a question for the **output format** (PowerPoint / Word / HTML / all) and any **content**
   they want changed from the standard. Don't assume — ask.
3. **Assemble from the standard.** Run the assembler with the chosen slide keys, in order:
   ```bash
   cd design-system/standard
   python assemble.py --name <DeckName> --slides cover,who-we-are,engagement-model,delivery-team,outcomes,contact
   ```
   This **always alternates black/white** and emits, in `design-system/standard/`:
   `<DeckName>.pptx` (alternating), `<DeckName>_Light.pptx`, `<DeckName>_Dark.pptx`,
   `<DeckName>-deck.html`, `<DeckName>-a4-landscape.html`, `<DeckName>-a4-portrait.html`,
   `<DeckName>.docx`, `<DeckName>_Dark.docx`. Add `--start light|dark` to force the opening
   background; `--no-web` to skip HTML/Word.
4. **Customise content** if the user gave specifics. The standard slides carry illustrative copy —
   edit the text in `<DeckName>.pptx` (find the textbox, replace its runs; keep colours/sizes), or
   build a **new** slide with `kit.py` (the box idiom is built in — `Deck()` auto-alternates,
   `d.card/d.cards/d.flow/d.grid/d.band` are the components). Then **re-theme and re-render**:
   ```bash
   python theme_deck.py <DeckName>.pptx light <DeckName>_Light.pptx
   python theme_deck.py <DeckName>.pptx dark  <DeckName>_Dark.pptx
   python build_web.py --light <DeckName>_Light.pptx --dark <DeckName>_Dark.pptx --basename <DeckName> --title "<DeckName>"
   ```
5. **Show the result.** Render slides to preview (`python render_slide.py <deck>.pptx <0-based-idx> out.png`)
   and surface a couple to the user. Keep figures illustrative unless the user gave real numbers, and
   keep customer names out unless they asked for a named version.

## Tooling (all in `design-system/standard/`)

| File | Does |
|------|------|
| `assemble.py` | Pick standard slides → alternating deck + B&W + HTML/Word. `--list` shows keys. |
| `theme_deck.py` | Recolour any deck to all-light or all-dark (context-aware; importable `build()/recolor()`). |
| `build_web.py` | Render a light+dark pair → interactive HTML deck (toggle), A4 docs, Word (both themes). |
| `kit.py` | Primitives for **new** on-style slides (`Deck`, `card`, `cards`, `flow`, `grid`, `band`). |
| `render_slide.py` | Faithful PNG render of one slide (Inter, groups) — for previews. |

Defaults: Inter is bundled at `design-system/assets/fonts/`; logos at `design-system/assets/logos/`.
Re-run `theme_deck.py` (×2) then `build_web.py` after **any** edit to a master so all formats stay
in lockstep.

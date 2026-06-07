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

**This skill is self-contained.** Everything it needs is bundled here:

- **`scripts/`** — the tooling (Python). Run scripts by their path in this skill.
- **`assets/`** — the brand resources the scripts read: the master deck `Mivada_Standard.pptx`
  (the **source of truth** — 27 curated slides), the Inter fonts (`assets/fonts/`), and the Mivada
  logos (`assets/logos/`).
- **`references/style.md`** — the full visual spec (tokens, the box idiom, components, slide
  catalogue). **Read it before editing or building slides.**

Scripts **read assets from this skill** and **write outputs to your current working directory** — so
run them from the folder where you want the files. Dependencies: `python-pptx`, `python-docx`,
`Pillow`, `lxml` (see `requirements.txt`).

Two kinds of output: **decks** (slides — PPTX/HTML, assembled from the library) and a **proposal**
(a flowing A4 Word *document* in the same design language). Pick based on what the user wants — a
presentation vs. a written proposal/capability document.

## The one rule about boxes (read this)

A box is a **filled/bordered rectangle with a coral label inside — nothing else.** Match the
"Four ways we support you" and "Across the whole Workday lifecycle" slides. **Never** add a thin
coral/ink **accent stripe down the left edge** of a card (it's an AI tell), never a drop shadow,
never a per-card header band unless the deck already uses one. Coral is the single accent, carried
by the **label text**, not by a stripe. The same rule holds in the proposal document.

## Workflow — always follow this

Paths below (`scripts/…`, `assets/…`) are relative to **this skill's directory**. Run from the
folder where you want the output files.

1. **Read `references/style.md`** for tokens, the box idiom, components and the slide catalogue.
2. **Ask the user which slides they want** — this is required. Use `AskUserQuestion` with a
   multi-select built from the catalogue (`python scripts/assemble.py --list`, which also lists
   presets: `company-overview`, `workday-go`, `ams-proposal`, `full`), plus a question for the
   **output format** (PowerPoint / Word / HTML / A4 proposal / all) and any **content** they want
   changed from the standard. Don't assume — ask.
3. **Assemble from the standard.** Run the assembler with a preset or explicit slide keys:
   ```bash
   python scripts/assemble.py --preset ams-proposal --name <DeckName>          # a named set, or…
   python scripts/assemble.py --name <DeckName> --slides cover,who-we-are,engagement-model,outcomes,contact
   ```
   This **always alternates black/white** and emits, in the current directory:
   `<DeckName>.pptx` (alternating), `<DeckName>_Light.pptx`, `<DeckName>_Dark.pptx`,
   `<DeckName>-deck.html`, `<DeckName>-a4-landscape.html`, `<DeckName>-a4-portrait.html`,
   `<DeckName>.docx`, `<DeckName>_Dark.docx`. Add `--start light|dark` to force the opening
   background; `--no-web` to skip HTML/Word.
4. **Customise content** if the user gave specifics. The standard slides carry illustrative copy —
   edit the text in `<DeckName>.pptx` (find the textbox, replace its runs; keep colours/sizes), or
   build a **new** slide with `scripts/kit.py` (the box idiom is built in — `Deck()` auto-alternates,
   `d.card/d.cards/d.flow/d.grid/d.band` are the components). Then **re-theme and re-render**:
   ```bash
   python scripts/theme_deck.py <DeckName>.pptx light <DeckName>_Light.pptx
   python scripts/theme_deck.py <DeckName>.pptx dark  <DeckName>_Dark.pptx
   python scripts/build_web.py --light <DeckName>_Light.pptx --dark <DeckName>_Dark.pptx --basename <DeckName> --title "<DeckName>"
   ```
5. **Show the result.** Render slides to preview (`python scripts/render_slide.py <deck>.pptx <0-based-idx> out.png`)
   and surface a couple to the user. Keep figures illustrative unless the user gave real numbers, and
   keep customer names out unless they asked for a named version.

## Written A4 proposal (a document, not slides)

When the user wants a **proposal/capability document** rather than a deck, use `scripts/proposal.py` —
it builds a native, editable A4 Word document in the same design language (coral eyebrows, ink
headings, coral-header tables, black callout bands, footer with page numbers — **no** left-edge
stripes). Import the `Proposal` class and compose sections (`masthead`, `lede`, `section`, `body`,
`bullets`, `kpis`, `table`, `callout`); `python scripts/proposal.py` writes a worked sample
(`Mivada_Proposal.docx`) you can adapt. This is the right tool for an AMS/Workday proposal written
as prose + tables.

## Tooling (in `scripts/`)

| File | Does |
|------|------|
| `assemble.py` | Pick standard slides (`--preset` or `--slides`) → alternating deck + B&W + HTML/Word. `--list` shows keys + presets. |
| `proposal.py` | Native A4 **Word proposal document** (`Proposal` class); flowing prose + tables, on-brand. |
| `theme_deck.py` | Recolour any deck to all-light or all-dark (context-aware; importable `build()/recolor()`). |
| `build_web.py` | Render a light+dark pair → interactive HTML deck (toggle), A4 docs, Word (both themes). |
| `kit.py` | Primitives for **new** on-style slides (`Deck`, `card`, `cards`, `flow`, `grid`, `band`). |
| `render_slide.py` | Faithful PNG render of one slide (Inter, groups) — for previews. |

Inter is bundled at `assets/fonts/`; logos at `assets/logos/`; the master deck at
`assets/Mivada_Standard.pptx`. Re-run `theme_deck.py` (×2) then `build_web.py` after **any** edit to
a master so all formats (PPTX, B&W, HTML, Word) stay in lockstep.

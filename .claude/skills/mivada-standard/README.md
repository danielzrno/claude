# mivada-standard

A self-contained Claude skill for building **Mivada Standard** decks and documents — PowerPoint,
Word and HTML — assembled from a 27-slide master and always rendered in alternating black/white
(plus uniform light and dark variants). It also builds a native A4 **proposal** document in the same
design language.

Everything the skill needs is bundled in this folder; nothing outside it is required.

```
mivada-standard/
├── SKILL.md              # how the skill works (read by Claude)
├── README.md             # this file
├── requirements.txt      # Python deps
├── references/
│   └── style.md          # the visual spec — tokens, the box idiom, slide catalogue
├── scripts/              # the tooling
│   ├── assemble.py       # pick slides (--preset / --slides) → deck + B&W + HTML + Word
│   ├── proposal.py       # native A4 Word proposal document (Proposal class)
│   ├── theme_deck.py     # recolour a deck to all-light / all-dark
│   ├── build_web.py      # light+dark pair → interactive HTML deck, A4 docs, Word
│   ├── kit.py            # primitives for new on-style slides
│   └── render_slide.py   # faithful PNG render of one slide (for previews)
└── assets/
    ├── Mivada_Standard.pptx   # the master deck (source of truth)
    ├── fonts/                 # Inter (Regular, Bold)
    └── logos/                 # Mivada logos & icons
```

## Quick start

```bash
pip install -r requirements.txt

# from the folder where you want the output files:
python scripts/assemble.py --list                              # slide keys + presets
python scripts/assemble.py --preset ams-proposal --name Acme   # build a deck (all formats)
python scripts/proposal.py                                     # build the sample A4 proposal
```

Scripts **read assets from this skill** and **write outputs to the current working directory**.
`assemble.py` emits the alternating deck, `_Light`/`_Dark` PPTX, an interactive HTML deck, A4 HTML,
and Word (`.docx`) — all from the chosen slides.

## The one rule about boxes

A box is a filled/bordered rectangle with a coral **label** inside — nothing else. No thin coral
stripe down a card's left edge (an AI tell), no drop shadow, no per-card header band. Coral is the
single accent, carried by the label text. See `references/style.md`.

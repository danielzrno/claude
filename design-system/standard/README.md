# Mivada — the standard

The **house standard**: Company Overview + AMS / Services combined into one long deck, with matching
documents. Built on the shared Editorial engine (`../02-editorial/build_pptx.py` + `slide_library.py`).
This is the canonical style — tweak content here, then we update the `/mivada` skill to point at it.

## Formats (same content, every form)

| File | Format |
|------|--------|
| `standard_deck.py` → `Mivada_Standard.pptx` | **PowerPoint** — the long combined deck (17 slides) |
| `presentation.html` | **Interactive HTML deck** (16:9, ←/→ nav, print 1 slide/page) |
| `standard-a4-portrait.html` | **HTML document — A4 portrait** |
| `standard-a4-landscape.html` | **HTML document — A4 landscape** |
| `build_docx.js` → `Mivada_Standard.docx` | **Word** (native .docx) |

## Build / regenerate

```bash
python standard_deck.py                                      # -> Mivada_Standard.pptx
NODE_PATH=/opt/node22/lib/node_modules node build_docx.js    # -> Mivada_Standard.docx
# HTML: open the .html files; "Print → Save as PDF" (deck = 1 slide/page; A4 docs paginate)
```

## Story (17 sections)

Logo · Cover · Who we are · Trusted by · What we do · How we work · Data & AI ·
Full-stack FIN & HCM · Two engagement models · AMS operating framework · Coverage model ·
Governance · Team · Proven outcomes · Why Mivada · Capability · Contact.

## Notes
- Presentations are **16:9**; documents are **A4** (portrait + landscape). Same brand, different canvas.
- The **client wall** uses a name grid until `design-system/logos.png` exists, then the real logo wall.
- Figures are illustrative samples — replace with verified numbers before real use.

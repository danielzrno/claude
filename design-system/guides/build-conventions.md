# Build conventions (shared across all three concepts)

These keep the three systems consistent in *structure* (so they're easy to compare and
reuse) while leaving each free to look completely different.

## Folder structure (identical per concept)

```
NN-name/
  README.md            # the concept spec: palette, type, motifs, when to use
  name.css             # the reusable theme — design tokens + components (the token-light core)
  presentation.html    # self-contained 16:9 deck, links name.css       (HTML presentation)
  document.html        # A4 print-ready document, links name.css         (Word — HTML A4 version)
  build_docx.js        # docx-js generator for the native Word document
  name.docx            # generated native Word document (committed)      (Word — .docx version)
  build_pptx.py        # python-pptx theme module + builder
  name.pptx            # generated PowerPoint (committed)                 (PowerPoint)
```

All three concepts share the **exact same brand tokens** (see `mivada-brand.md`) — identical palette
and font. They differ only in layout, density and colour-blocking.

## The token-saving idea (why this is "fast + token-managed in Claude")

The expensive part of any artifact is the styling. Here it's written **once** per concept:
- HTML files `<link>` to `name.css` and emit only **content markup** + semantic classes.
- The PPTX builder exposes a small **theme dict + helper functions** (`cover()`, `kpis()`,
  `pillars()`, `quote()`…). A new deck is a list of helper calls, not hand-placed shapes.

So generating a *new* Mivada deck/doc later = write the content, reuse the theme. Claude
emits a few dozen lines, not a few hundred. That is the deliverable's real value.

## Font strategy

Each concept names a **primary webfont** (loaded via Google Fonts `<link>` in HTML — renders
in the reader's browser) and an **Office-safe fallback** (so PPTX still looks right for people
who haven't installed the brand fonts). PPTX files reference the brand font name; PowerPoint
substitutes the fallback automatically if it's missing. The brand fonts are free (Google Fonts);
the concept README tells the user which two to install for pixel-perfect decks.

Fonts available locally for QA rendering: Fraunces, Archivo, Hanken Grotesk, JetBrains Mono,
Gabarito, Source Serif 4.

## HTML presentation rules

- 1280×720 (16:9) slides; `scroll-snap` between them; **keyboard nav** (← / → / Space) +
  on-screen slide counter. No external JS libraries — one small inline script only.
- Works as a web page *and* prints clean: `@page { size: 1280px 720px landscape }` style block
  so "Print → Save as PDF" yields one slide per page.
- One choreographed entrance per slide (CSS only). No fade-up-on-everything.

## A4 document rules

- `@page { size: A4; margin: … }`, `-webkit-print-color-adjust: exact; print-color-adjust: exact;`
- Sensible `page-break` rules; a light on-screen "paper" wrapper for preview.
- Links the concept CSS; adds only print/page-specific styles locally.

## PPTX rules (python-pptx)

- 16:9 deck size = 13.333in × 7.5in.
- A `THEME` dict (hex strings, font names, sizes) + helpers using EMU via `Inches`/`Pt`.
- Set `run.font.name` to the brand font on every run; tuned ink/paper, never pure black/white.
- Respect the slide tells in `anti-ai-tells.md`: **no accent line under titles**, no full-width
  header/footer bars unless it's the concept's deliberate motif, no beige default background,
  left-aligned body, a visual element on every slide, varied layouts.
- After building: convert with LibreOffice and eyeball.
  ```bash
  python build_pptx.py
  python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf name.pptx
  rm -f slide-*.jpg && pdftoppm -jpeg -r 110 name.pdf slide && ls -1 "$PWD"/slide-*.jpg
  ```

## Word (.docx) rules — `build_docx.js` (docx-js)

Native Word document built with **docx-js** (`require('docx')`; run with
`NODE_PATH=/opt/node22/lib/node_modules node build_docx.js`). Follow `/mnt/skills/public/docx/SKILL.md`.

- **A4** page (`size: { width: 11906, height: 16838 }`, DXA), ~1" margins; content width ≈ 9026 DXA.
- Override built-in styles (`Heading1`/`Heading2`/`Normal`) with the brand font (**Inter**, fallback Arial)
  and brand colours (coral `EA493F`, ink `111111`). Title/eyebrow can be coral; keep body ink for readability.
- **Never** unicode bullets — use `numbering` config with `LevelFormat.BULLET`.
- Tables need **dual widths** (`columnWidths` + per-cell `width`, DXA), `ShadingType.CLEAR`, cell margins.
  Use a coral-shaded header row or a coral KPI block. Never use tables as dividers — use a paragraph
  **bottom border** for rules.
- Header/footer: wordmark + contact + page number (tab stops, not tables).
- Validate: `python /mnt/skills/public/docx/scripts/office/validate.py name.docx`, then convert for QA:
  `python /mnt/skills/public/pptx/scripts/office/soffice.py --headless --convert-to pdf name.docx`.
- The `.docx` carries the same capability-statement content as `document.html`, styled on-brand within
  Word's constraints (it won't match the HTML pixel-for-pixel — that's expected).

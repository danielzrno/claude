# Workflow — translate an existing deck/document into Editorial (no information lost)

Goal: take a source PowerPoint (or doc) and **re-express every bit of its content** in the Mivada
Editorial system. Nothing gets dropped — text, tables, chart data, speaker notes and image intent
all carry over. Work in phases; do not skip the extract/reconcile bookends.

## Quick start — automated first pass

`design-system/samples/auto_translate.py` does a strong first pass automatically (classify each
slide → cover / content / table-stat-grid / coral chart / quote / contact; carry all text + notes;
print a reconcile coverage report):

```bash
cd design-system/samples
python auto_translate.py /path/to/source.pptx     # -> source_editorial_auto.pptx + coverage report
```

Then **refine by hand** anything the report flags or that deserves a richer layout (promote a bullet
slide to `pillars`/`steps`, a metrics slide to a stat row, a testimonial to `pullquote`). The phases
below are that full, careful process — and what to do when the source has charts/images/edge cases.
`samples/translate_to_editorial.py` is a worked, hand-tuned example of the finished result.

## Phase 1 — Extract everything (lossless inventory)

Read the source skill first: `/mnt/skills/public/pptx/SKILL.md` (and `docx/SKILL.md` for documents).

- Open the source and dump, **per slide, in order**: title, every text frame/bullet (verbatim),
  table contents, chart series + values, image inventory (what each shows), and **speaker notes**.
  ```python
  from pptx import Presentation
  p = Presentation("SOURCE.pptx")
  for i, s in enumerate(p.slides, 1):
      texts = [sh.text_frame.text for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
      notes = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
      # also: sh.has_table -> rows/cells; sh.has_chart -> plots/series/categories; sh.shape_type==PICTURE
  ```
- Write the inventory to a scratch file (`/tmp/source-inventory.md`) — slide number, type, full text,
  tables, chart data, images, notes. This is the contract: the final deck must account for every line.
- Export embedded images you need to reuse (`shape.image.blob`) to a working folder.

## Phase 2 — Map to the Editorial system

For each source slide, choose the closest Editorial slide type (`editorial-system.md`):

| Source slide | Editorial type |
|---|---|
| Title / section divider | `cover` / a black section slide |
| Agenda / overview | `pillars` or a simple list slide |
| Narrative / concept | `kpis` (headline + standfirst) or `split` |
| Process / timeline | `steps` |
| Data / metrics / table | `kpis`/stats row, or a real table re-coloured to brand; charts → brand palette |
| Quote / testimonial | `pullquote` |
| Comparison / two-up | `split` (two columns) |
| Closing / contact / CTA | `contact` |

Rules:
- **Preserve order and roughly the slide count.** Prefer 1:1. Only split a crammed slide into two,
  or merge two near-duplicates, when it clearly serves clarity — note every such change.
- If content doesn't fit any template, **extend the system** (add a Deck method / CSS block) rather
  than cut content. Long body text → a readable content slide, not deletion.
- Keep the **black↔off-white sandwich** rhythm and apply **logos** per the placement matrix.
- Charts: rebuild with python-pptx native charts using the brand palette (coral series first, then
  graphite/ink/mid), or re-place as a cleanly re-coloured image. Keep all data points and labels.
- Images: re-place with rounded corners on-brand; if an image was purely decorative/off-brand,
  replace with a brand treatment (do not silently lose informational images).
- Numbers/claims from the source are **real** — carry them exactly; do not relabel as illustrative.

## Phase 3 — Build

Build the `.pptx` with the Editorial `Deck` (primary, since the source is a deck). Mirror the
inventory slide-by-slide. Keep speaker notes (`slide.notes_slide.notes_text_frame.text = ...`).
If the user also wants HTML/Word, generate those from the same content after the pptx is signed off.

## Phase 4 — Reconcile (prove nothing was lost)

- Re-extract text from the **new** deck the same way as Phase 1 and diff against the inventory:
  every source line (or its deliberate paraphrase) must be present; flag anything missing.
- Confirm: slide count vs source (and a one-line note for any intentional merge/split), all tables/
  chart values present, all speaker notes carried, logos correct per background.
- **Render & QA** (convert to PDF, `pdftoppm`, READ every slide): no overflow/clipping, black/coral
  contrast OK, no AI tells (`anti-ai-tells.md`), no emoji/placeholder.
- Present a short before→after mapping table and the rendered preview. Then commit.

> If anything about consolidation, missing source assets, or whether a figure is real vs illustrative
> is ambiguous, ask the user before finalising — losing information is the one unacceptable outcome.

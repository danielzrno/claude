#!/usr/bin/env python3
# ============================================================================
# GENERAL deck translator -> Mivada Editorial. Classifies each source slide
# (cover / content / table / chart / quote / contact), rebuilds it with the
# Editorial Deck + real logos, carries speaker notes, and reconciles coverage.
#
# An automated FIRST PASS: it places all the text it can and the reconcile
# report flags anything to finish by hand (see translate-deck.md).
#
#   python auto_translate.py SOURCE.pptx  ->  SOURCE_editorial.pptx
# ============================================================================
import sys, os, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck, THEME
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SRC = sys.argv[1] if len(sys.argv) > 1 else "legacy-overview.pptx"
OUT = re.sub(r"\.pptx$", "_editorial_auto.pptx", os.path.basename(SRC))

FOOT = re.compile(r"^\s*\d+\s*$|confidential", re.I)
EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.\w+")
DATEISH = re.compile(r"statement|\b20\d\d\b", re.I)

def clean_lines(text):
    out = []
    for ln in (text or "").split("\n"):
        ln = re.sub(r"^\s*[•\-–•]\s*", "", ln).strip()
        ln = re.sub(r"^\d+\.\s*", "", ln)            # drop "1. " numbering
        if ln and not FOOT.search(ln):
            out.append(ln)
    return out

def extract(path):
    p = Presentation(path); inv = []
    for s in p.slides:
        d = {"lines": [], "tables": [], "charts": [], "images": [], "notes": ""}
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                d["lines"] += clean_lines(sh.text_frame.text)
            if sh.has_table:
                d["tables"].append([[c.text.strip() for c in r.cells] for r in sh.table.rows])
            if sh.has_chart:
                ch = sh.chart
                d["charts"].append((list(ch.plots[0].categories),
                                    [(se.name, list(se.values)) for se in ch.series]))
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try: d["images"].append(sh.image.blob)
                except Exception: pass
        if s.has_notes_slide:
            d["notes"] = s.notes_slide.notes_text_frame.text
        inv.append(d)
    return inv

def head(title):
    w = (title or "").split()
    return (" ".join(w[:-1]), w[-1]) if len(w) > 1 else (title or "", "")

def numeric(rows):
    vals = [r[-1] for r in rows[1:]] if len(rows) > 1 else []
    return bool(vals) and sum(bool(re.match(r"^[\$£€]?[\d.,]+ ?[%+kKmM]?$", v)) for v in vals) >= len(vals) * 0.6

def last_slide(d):
    return d.prs.slides[len(d.prs.slides._sldIdLst) - 1]

def stat_grid(d, eyebrow, headline, rows):
    pre, acc = head(headline)
    s = d._slide(THEME["off_white"]); d._slug(s, eyebrow[:30])
    d._eyebrow(s, 0.96, 0.92, 10, eyebrow, color="coral")
    _, tf = d._box(s, 0.92, 1.4, 11, 1.0); p = d._para(tf, first=True, line=1.0)
    d._run(p, (pre or headline) + " ", 36, "111111", bold=True, spacing=-0.022)
    if acc: d._run(p, acc, 36, "EA493F", bold=True, spacing=-0.022)
    stats = []
    for r in rows[1:]:
        label, val = r[0], r[-1]
        m = re.match(r"^([\$£€]?[\d.,]+)([%+kKmM]?)", val)
        stats.append((m.group(1) if m else val, (m.group(2) if m and m.group(2) else ""), label))
    d._stat_row(s, 4.35, stats[:6], light=True)

# ---- translate -------------------------------------------------------------
src = extract(SRC); d = Deck(); N = len(src)
for i, sl in enumerate(src):
    lines, tables, charts, images, notes = (sl["lines"], sl["tables"], sl["charts"],
                                            sl["images"], sl["notes"])
    title = lines[0] if lines else f"Slide {i+1}"
    body = lines[1:]
    is_contact = (i == N - 1) and (any(EMAIL.search(l) for l in lines)
                  or any(re.search(r"\b(talk|contact|thank you|get in touch)\b", l, re.I) for l in lines))
    qline = next((l for l in lines if re.search(r'[“"]', l) and len(l) > 20), None)
    is_quote = qline is not None and len(lines) <= 3

    if i == 0 and not charts and not tables:                       # COVER
        cand = [l for l in lines[:3] if len(l.split()) >= 2 and not DATEISH.search(l)]
        h = cand[0] if cand else title
        eb = next((l for l in lines[:3] if DATEISH.search(l)), "Capability overview")
        d.cover_plain(eb, *head(h), "", "")
    elif is_contact:                                              # CONTACT
        contacts = [l for l in lines if EMAIL.search(l) or re.search(r"\.com|sydney|melbourne|india|\+\d", l, re.I)]
        d.contact_plain("Let's talk", "Let's", "talk.", " · ".join(contacts[:2]), "")
    elif charts:                                                  # CHART
        cats, series = charts[0]
        h = body[0] if body else title
        d.chart_coral(title, *head(h), cats, series, slug_txt=title[:30])
        if len(body) > 1: notes = (notes + "\n" + " ".join(body[1:])).strip()
    elif tables and numeric(tables[0]):                           # TABLE -> stat grid
        stat_grid(d, title, body[0] if body else "Key figures", tables[0])
        if body: notes = (notes + "\n" + " ".join(body)).strip()
    elif images:                                                 # IMAGE -> media split
        imgp = f"/tmp/_xlate_{i}.png"
        with open(imgp, "wb") as fh: fh.write(images[0])
        items = list(body)
        for t in tables:
            items += [" — ".join(c for c in r if c) for r in t[1:]]
        d.media(title[:30] or "Section", *head(title), items, imgp, dark=(i % 2 == 1), slug_txt=title[:30])
        if len(images) > 1:
            notes = (notes + f"\n[{len(images)-1} more source image(s) to place]").strip()
    elif is_quote:                                                # QUOTE
        attr = next((l for l in lines if l not in (qline, title)), "")
        d.quote_plain(title if title != qline else "In their words", qline, attr, dark=(i % 2 == 1))
    else:                                                         # GENERAL CONTENT
        items = list(body)
        for t in tables:
            items += [" — ".join(c for c in r if c) for r in t[1:]]
        d.content(title[:30] or "Section", *head(title), items, dark=(i % 2 == 1), slug_txt=title[:30])

    if notes:
        last_slide(d).notes_slide.notes_text_frame.text = notes

d.save(OUT)
print(f"Wrote {OUT} ({N} slides)")

# ---- reconcile -------------------------------------------------------------
def alltext(path):
    p = Presentation(path); buf = []
    for s in p.slides:
        for sh in s.shapes:
            if sh.has_text_frame: buf.append(sh.text_frame.text)
            if sh.has_table: buf += [c.text for r in sh.table.rows for c in r.cells]
            if sh.has_chart:
                ch = sh.chart; buf += [str(x) for x in ch.plots[0].categories]
                buf += [str(v) for se in ch.series for v in se.values]
        if s.has_notes_slide: buf.append(s.notes_slide.notes_text_frame.text)
    return "\n".join(buf).lower()

new = alltext(OUT)
toks = set()
for sl in src:
    for l in sl["lines"] + [c for t in sl["tables"] for r in t for c in r]:
        toks.update(w.lower() for w in re.findall(r"[A-Za-z0-9@.+&%-]{4,}", l))
missing = sorted(t for t in toks if t not in new)
carried = len(toks) - len(missing)
print(f"Reconcile: {carried}/{len(toks)} source tokens carried ({100*carried//max(1,len(toks))}%).")
print("  Finish by hand:", ", ".join(missing) if missing else "none — nothing lost ✓")

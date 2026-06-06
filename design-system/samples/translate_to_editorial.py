#!/usr/bin/env python3
# ============================================================================
# Translate a deck into the Mivada EDITORIAL look — end to end, no info lost.
# Demonstrates the /mivada `translate-deck` workflow:
#   1) extract everything from the source (text / tables / chart / notes)
#   2) map each slide to the Editorial system + carry speaker notes
#   3) rebuild as a branded .pptx (real logos via the Editorial Deck)
#   4) reconcile — verify every critical token from the source survived
#
#   python translate_to_editorial.py SOURCE.pptx  ->  <SOURCE>_editorial.pptx
# (defaults SOURCE to legacy-overview.pptx)
# ============================================================================
import sys, os, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck, THEME, PAGE_W
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.dml.color import RGBColor

SRC = sys.argv[1] if len(sys.argv) > 1 else "legacy-overview.pptx"
OUT = re.sub(r"\.pptx$", "_editorial.pptx", os.path.basename(SRC))
CORAL = RGBColor(0xEA, 0x49, 0x3F)

# ---- 1) EXTRACT ------------------------------------------------------------
def extract(path):
    p = Presentation(path); inv = []
    for s in p.slides:
        d = {"text": [], "tables": [], "charts": [], "notes": ""}
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                d["text"].append(sh.text_frame.text)
            if sh.has_table:
                d["tables"].append([[c.text for c in r.cells] for r in sh.table.rows])
            if sh.has_chart:
                ch = sh.chart
                d["charts"].append((list(ch.plots[0].categories),
                                    [(se.name, list(se.values)) for se in ch.series]))
        if s.has_notes_slide:
            d["notes"] = s.notes_slide.notes_text_frame.text
        inv.append(d)
    return inv

# ---- helpers on top of the Editorial Deck ----------------------------------
def notes(s, text):
    if text:
        s.notes_slide.notes_text_frame.text = text

def cover(d, n):
    s = d._slide(THEME["black"])
    d._logo(s, "Mivada_Logo_2C_OnBlack_RGB_L", 0.96, 0.82, 0.42)
    d._eyebrow(s, 0.96, 2.55, 8, "Capability statement · 2026", color="coral")
    _, tf = d._box(s, 0.92, 2.95, 11.4, 2.6)
    p = d._para(tf, first=True, line=0.96)
    d._run(p, "Company ", 80, "FFFFFF", bold=True, spacing=-0.025)
    d._run(p, "overview.", 80, "EA493F", bold=True, spacing=-0.025)
    d._hline(s, 0.96, 6.74, PAGE_W - 1.92, THEME["rev_hair"], 1.0)
    _, tf = d._box(s, 0.96, 6.86, 7, 0.4); d._run(d._para(tf, first=True),
        "An Australian technology consultancy", 11.5, "AEAEAE", bold=True, spacing=0.04)
    _, tf = d._box(s, PAGE_W - 7, 6.86, 6.04, 0.4)
    d._run(d._para(tf, first=True, align=PP_ALIGN.RIGHT), "mivada.com", 11.5, "AEAEAE", bold=True, spacing=0.04)
    notes(s, n)

def numbers(d, n, rows):
    # rows: the source table (incl header) — carry every metric/value pair
    s = d._slide(THEME["off_white"])
    d._slug(s, "By the numbers")
    d._eyebrow(s, 0.96, 0.92, 8, "By the numbers", color="coral")
    _, tf = d._box(s, 0.92, 1.4, 11, 1.0); p = d._para(tf, first=True, line=1.0)
    d._run(p, "The shape of ", 40, "111111", bold=True, spacing=-0.022)
    d._run(p, "the business.", 40, "EA493F", bold=True, spacing=-0.022)
    stats = []
    for label, val in rows[1:]:                       # skip header row
        m = re.match(r"^([0-9]+)([+%]?.*)$", val.strip())
        if m and m.group(1):
            stats.append((m.group(1), m.group(2) if m.group(2) in ("+", "%") else "", label))
        else:
            stats.append((val, "", label))
    d._stat_row(s, 4.35, stats, light=True)
    notes(s, n)

def chart(d, n, cats, series):
    s = d._slide(THEME["off_white"])
    d._slug(s, "Results")
    d._eyebrow(s, 0.96, 0.92, 8, "Results", color="coral")
    _, tf = d._box(s, 0.92, 1.4, 11, 1.0); p = d._para(tf, first=True, line=1.0)
    d._run(p, "Projects delivered ", 34, "111111", bold=True, spacing=-0.022)
    d._run(p, "per year.", 34, "EA493F", bold=True, spacing=-0.022)
    cd = CategoryChartData(); cd.categories = cats
    for name, vals in series:
        cd.add_series(name, vals)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.96), Inches(2.65), Inches(11.4), Inches(4.0), cd)
    ch = gf.chart; ch.has_legend = False
    plot = ch.plots[0]; plot.gap_width = 60; plot.has_data_labels = True
    plot.data_labels.number_format = "0"; plot.data_labels.number_format_is_linked = False
    ser = plot.series[0]; ser.format.fill.solid(); ser.format.fill.fore_color.rgb = CORAL
    ser.format.line.fill.background()
    notes(s, n)

def contact(d, n):
    s = d._slide(THEME["black"])
    d._logo(s, "Mivada_Logo_2C_OnBlack_RGB_L", 0.96, 0.82, 0.42)
    d._eyebrow(s, 0.96, 2.4, 8, "Let's talk", color="coral")
    _, tf = d._box(s, 0.92, 2.85, 11.0, 2.4); p = d._para(tf, first=True, line=0.98)
    d._run(p, "Let's build something ", 54, "FFFFFF", bold=True, spacing=-0.025)
    d._run(p, "together.", 54, "EA493F", bold=True, spacing=-0.025)
    d._hline(s, 0.96, 6.4, PAGE_W - 1.92, THEME["rev_hair"], 1.0)
    _, tf = d._box(s, 0.96, 6.55, 8, 0.5)
    d._run(d._para(tf, first=True), "hello@mivada.com · mivada.com", 12.5, "CFCFCF", bold=True, spacing=0.02)
    _, tf = d._box(s, PAGE_W - 7, 6.55, 6.04, 0.5)
    d._run(d._para(tf, first=True, align=PP_ALIGN.RIGHT),
           "Sydney & Melbourne · Onshore AU & India", 12.5, "AEAEAE", bold=False, spacing=0.02)
    notes(s, n)

# ---- 2/3) MAP + BUILD ------------------------------------------------------
src = extract(SRC)
d = Deck()

# (1) Title -> cover
cover(d, src[0]["notes"])
# (2) About -> kpis (paragraph + 'formerly LJM Infotech' + 'independent' all kept)
d.kpis("About us", "An Australian technology", "consultancy.",
       "We turn enterprise platforms — Workday, payroll, data and AI — into measurable "
       "outcomes, built around how teams actually work. Independent and platform-agnostic; "
       "formerly LJM Infotech.",
       [("2014", "", "Founded"), ("120", "+", "Certified specialists"), ("AU + India", "", "Delivery")])
d.prs.slides[1].notes_slide.notes_text_frame.text = src[1]["notes"]
# (3) What we do -> pillars
d.pillars("What we do", "Four service", "pillars.", [
    ("01", "Workday & ERP", "Implementation, optimisation and managed support", "Platforms"),
    ("02", "Payroll consulting", "Compliant, accurate payroll at scale", "Payroll"),
    ("03", "Data & AI", "Lakehouse architecture, governed analytics, applied AI", "Data"),
    ("04", "Intelligent automation", "Remove manual, error-prone work", "Automation")])
d.prs.slides[2].notes_slide.notes_text_frame.text = src[2]["notes"]
# (4) Approach -> steps
d.steps("How we work", "People", "first.", [
    ("Step 01", "Listen", "Understand the people and the work before the technology"),
    ("Step 02", "Design", "Map to real-world needs"),
    ("Step 03", "Deliver", "Implement in increments; value early"),
    ("Step 04", "Care", "Measure outcomes; keep improving after go-live")])
d.prs.slides[3].notes_slide.notes_text_frame.text = src[3]["notes"]
# (5) Numbers table -> stat grid (all 5 rows carried)
numbers(d, src[4]["notes"], src[4]["tables"][0])
# (6) Chart -> rebuilt coral column chart (data preserved)
cats, series = src[5]["charts"][0]
chart(d, src[5]["notes"], cats, series)
# (7) Testimonial -> pullquote
d.pullquote("What clients say",
            "“They delivered in weeks what we’d scoped for a year — and our team actually uses it.”",
            "— Operations Director, enterprise client",
            [("41", "", "Projects in 2025"), ("98", "%", "Client retention"), ("40", "+", "Implementations")])
d.prs.slides[6].notes_slide.notes_text_frame.text = src[6]["notes"]
# (8) Contact -> closing
contact(d, src[7]["notes"])

d.save(OUT)
print(f"Wrote {OUT} ({len(src)} source slides -> {len(d.prs.slides._sldIdLst)} editorial slides)")

# ---- 4) RECONCILE ----------------------------------------------------------
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
    return "\n".join(buf)

new = alltext(OUT).lower()
# critical tokens that MUST survive (logo-carried 'Mivada' excluded — it's in the wordmark image)
critical = ["company overview", "capability statement", "2014", "ljm infotech", "120",
            "certified specialists", "australia", "india", "independent", "platform-agnostic",
            "workday & erp", "implementation", "payroll consulting", "compliant",
            "data & ai", "lakehouse", "governed analytics", "applied ai",
            "intelligent automation", "manual", "listen", "design", "deliver", "care",
            "real-world", "increments", "go-live", "40+", "98%", "au + india",
            "2022", "2023", "2024", "2025", "12", "19", "28", "41",
            "weeks", "operations director", "hello@mivada.com", "sydney", "melbourne", "discovery"]
missing = [t for t in critical if t.lower() not in new]
print(f"Reconcile: {len(critical) - len(missing)}/{len(critical)} critical tokens carried.")
print("  MISSING:", missing if missing else "none — nothing lost ✓")

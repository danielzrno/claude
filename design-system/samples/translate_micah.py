#!/usr/bin/env python3
# ============================================================================
# Hand-tuned translation of "First Meeting with Micah Projects" into Editorial.
# Brands every slide; carries the diagrams large (they ARE the content);
# rebuilds the capability stats as a coral KPI grid and the case studies as
# columns; preserves speaker notes. Reconciles at the end.
#   python translate_micah.py  ->  First_Meeting_with_Micah_Projects_editorial.pptx
# ============================================================================
import sys, os, re
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck, THEME, PAGE_W
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image as PILImage

SRC = "First_Meeting_with_Micah_Projects.pptx"
OUT = "First_Meeting_with_Micah_Projects_editorial.pptx"
SOFF = "/mnt/skills/public/pptx/scripts/office/soffice.py"
src = Presentation(SRC)

def imgs_from(idx, n):
    """Return paths to the n largest images on source slide idx (WMF->PNG)."""
    sl = src.slides[idx]
    pics = sorted([sh for sh in sl.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE],
                  key=lambda s: s.width * s.height, reverse=True)
    out = []
    for k, sh in enumerate(pics[:n]):
        ct = sh.image.content_type
        ext = ".wmf" if "wmf" in ct else (".jpg" if "jpeg" in ct else ".png")
        raw = f"/tmp/mic_{idx}_{k}{ext}"
        with open(raw, "wb") as f: f.write(sh.image.blob)
        if ext == ".wmf":
            os.system(f"python3 {SOFF} --headless --convert-to png --outdir /tmp {raw} >/dev/null 2>&1")
            png = f"/tmp/mic_{idx}_{k}.png"
            out.append(png if os.path.exists(png) else None)
        else:
            out.append(raw)
    return [p for p in out if p]

def source_slide_png(idx1):
    """Render a source slide to a cropped PNG (robust for EMF/WMF/vector slides)."""
    import glob
    os.system(f"python3 {SOFF} --headless --convert-to pdf --outdir /tmp '{SRC}' >/dev/null 2>&1")
    pdf = "/tmp/" + os.path.splitext(os.path.basename(SRC))[0] + ".pdf"
    for f in glob.glob("/tmp/mic_src*"): os.remove(f)
    os.system(f"pdftoppm -jpeg -r 150 -f {idx1} -l {idx1} '{pdf}' /tmp/mic_src >/dev/null 2>&1")
    files = sorted(glob.glob("/tmp/mic_src*.jpg"))
    if not files: return None
    im = PILImage.open(files[-1]).convert("RGB"); w, h = im.size
    im = im.crop((int(w * 0.02), int(h * 0.14), int(w * 0.98), int(h * 0.99)))  # drop original title band
    out = f"/tmp/mic_src_{idx1}.png"; im.save(out); return out

def place_images(d, s, paths, x=0.96, y=2.25, w=11.4, h=4.65):
    paths = [p for p in paths if p]
    n = len(paths)
    if n == 0: return
    gap = 0.3 if n > 1 else 0
    cw = (w - gap * (n - 1)) / n
    cx = x
    for p in paths:
        try: iw, ih = PILImage.open(p).size; ar = iw / ih
        except Exception: ar = 1.4
        ww = cw; hh = ww / ar
        if hh > h: hh = h; ww = hh * ar
        try:
            s.shapes.add_picture(p, Inches(cx + (cw - ww) / 2), Inches(y + (h - hh) / 2), width=Inches(ww))
        except Exception as e:
            print("  ! image skipped:", p, e)
        cx += cw + gap

def header(d, s, eyebrow, hpre, hacc, slug=None, dark=False):
    d._slug(s, (slug or eyebrow)[:30], dark=dark)
    d._eyebrow(s, 0.96, 0.92, 11, eyebrow, color="coral")
    tcol = "FFFFFF" if dark else "111111"
    _, tf = d._box(s, 0.92, 1.34, 11.4, 1.0)
    p = d._para(tf, first=True, line=1.0)
    d._run(p, hpre + (" " if hacc else ""), 30, tcol, bold=True, spacing=-0.022)
    if hacc: d._run(p, hacc, 30, "EA493F", bold=True, spacing=-0.022)

def diagram(d, eyebrow, hpre, hacc, idx, n=1, notes="", slug=None):
    s = d._slide(THEME["off_white"]); header(d, s, eyebrow, hpre, hacc, slug)
    place_images(d, s, imgs_from(idx, n))
    if notes: s.notes_slide.notes_text_frame.text = notes

d = Deck()

# (1) Cover
d.cover_plain("Micah Projects · March 2026", "Implementation", "discussion.",
              "Mivada — Workday implementation & managed services", "mivada.com")

# (2) Why Mivada
d.content("Why Mivada", "Your long-term", "Workday partner.", [
    "We grew out of a customer — so we understand your challenges first-hand.",
    "Australian-owned business, creating Australian jobs",
    "Not a Big 4 — direct access to our CEO, with a reduced escalation matrix",
    "Dedicated local engagement manager",
    "People Systems & Transformation",
    "Data, Analytics & Business Intelligence",
    "Business Process Management & Automation",
    "Dedicated offshore capability",
    "We work with notable brands such as Qantas, Guzman y Gomez and Western Sydney International Airport",
], dark=True)

# (3) Workday Services — big diagram
diagram(d, "Mivada's Workday services", "Mivada's Workday", "services.", 2,
        n=1, notes="Daniel & what we are doing with 9")

# (4) Capability in Workday — KPI grid (all figures carried)
s = d._slide(THEME["off_white"]); header(d, s, "Mivada's capability in Workday", "Capability in", "Workday.")
_, tf = d._box(s, 0.92, 2.12, 11.4, 0.7)
d._run(d._para(tf, first=True, line=1.2),
       "Empowering innovation through a Workday partnership and local expertise — our consultants "
       "average 3+ years' Workday experience, across successfully completed implementations and support "
       "engagements.", 13, "323232", bold=False, spacing=0)
d._stat_row(s, 3.6, [("3", "+", "Certifications / consultant"), ("2000", "+", "Workday integrations"),
                     ("3000", "+", "Workday custom reports"), ("30", "+", "Implementations & support")], light=True)
d._stat_row(s, 5.55, [("200", "+", "Employees · AU, US & India"),
                      ("100", "+", "People-systems clients · Payroll, T&A, HRIS"), ("10", "+", "Years")], light=True)
s.notes_slide.notes_text_frame.text = "Daniel"

# (5) Methodology — big diagram
diagram(d, "How we deliver", "Our", "methodology.", 4, n=1)

# (6) Draft Implementation Plan — keep the native EMF timeline (renders in PowerPoint;
#     no Linux tool can rasterize EMF, so it appears blank only in LibreOffice previews)
s = d._slide(THEME["off_white"]); header(d, s, "Indicative timeline", "Draft implementation", "plan.")
_sl6 = src.slides[5]
_pic = max((sh for sh in _sl6.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE),
           key=lambda x: x.width * x.height)
_emf = "/tmp/mic6.emf"
with open(_emf, "wb") as f: f.write(_pic.image.blob)
_ar = _pic.width / _pic.height
_w = 11.4; _h = _w / _ar
if _h > 4.5: _h = 4.5; _w = _h * _ar
s.shapes.add_picture(_emf, Inches(0.96 + (11.4 - _w) / 2), Inches(2.45), width=Inches(_w), height=Inches(_h))

# (7) Governance — two diagrams
diagram(d, "How we run it", "", "Governance.", 6, n=2)

# (8) RACI & Data Migration — two diagrams
diagram(d, "Roles & data", "RACI & data", "migration.", 7, n=2)

# (9) Understanding of your needs
d.content("Understanding your needs", "What you need", "this to do.", [
    "Payroll compliance confidence",
    "Connected HR & payroll",
    "Rostering that enforces rules",
    "A single, reliable user experience",
    "Frontline manager visibility",
    "Funding & cost visibility",
], dark=False)

# (10) Real-world success stories — 3 cases
s = d._slide(THEME["off_white"]); header(d, s, "Real-world success stories", "Proven Workday", "outcomes.")
_, tf = d._box(s, 0.92, 2.12, 11.4, 0.5)
d._run(d._para(tf, first=True, line=1.2),
       "How Mivada enables high-performance Workday operations.", 13, "323232", bold=False, spacing=0)
cases = [
    ("Guzman y Gomez",
     "Mivada partnered with GYG to stabilise and optimise their Workday platform post–go-live across "
     "modules, reporting, security and compliance. Through Agile delivery we onboarded 150+ franchisees, "
     "delivered new reporting and a redesigned security model, strengthened audit readiness and improved "
     "performance across HCM, Absence and Learning — enabling scalable, compliant global operations."),
    ("Qantas",
     "Mivada has supported Qantas since 2014, taking over Workday support in 2017 and rapidly becoming a "
     "strategic partner across HR, Payroll, Digital, Automation, Data and Analytics. By scaling integrations, "
     "modernising architectures and embedding automation and analytics, Mivada evolved from post–go-live "
     "support to Qantas' long-term enterprise partner."),
    ("Western Sydney Intl Airport",
     "Mivada partnered with WSI to deliver a card-management system, then expanded into a strategic role "
     "supporting WSI's Workday program. By resolving functional and integration issues and assuming full "
     "AMS ownership, Mivada stabilised HR, Payroll, Finance and Procurement — becoming WSI's trusted "
     "long-term partner."),
]
x0, total = 0.96, 11.4; col = total / 3
for i, (client, para) in enumerate(cases):
    cx = x0 + i * col
    if i > 0: d._vline(s, cx, 3.0, 3.8, THEME["hairline"], 1.0)
    pad = 0.0 if i == 0 else 0.32
    _, tf = d._box(s, cx + pad, 2.95, col - pad - 0.25, 0.5)
    d._run(d._para(tf, first=True), client.upper(), 12, "EA493F", bold=True, spacing=0.06, caps=True)
    _, tf = d._box(s, cx + pad, 3.5, col - pad - 0.28, 3.4)
    d._run(d._para(tf, first=True, line=1.26), para, 10, "323232", bold=False, spacing=0)
s.notes_slide.notes_text_frame.text = "Rama"

# (11) Questions / closing
d.contact_plain("Over to you", "Questions?", "",
                "hello@mivada.com · mivada.com", "Sydney · Melbourne · AU & India")

d.save(OUT)
print(f"Wrote {OUT} ({len(d.prs.slides._sldIdLst)} slides)")

# ---- reconcile -------------------------------------------------------------
def alltext(path):
    p = Presentation(path); buf = []
    for s in p.slides:
        for sh in s.shapes:
            if sh.has_text_frame: buf.append(sh.text_frame.text)
        if s.has_notes_slide: buf.append(s.notes_slide.notes_text_frame.text)
    return "\n".join(buf).lower()
def srctext():
    buf = []
    for s in src.slides:
        for sh in s.shapes:
            if sh.has_text_frame: buf.append(sh.text_frame.text)
        if s.has_notes_slide: buf.append(s.notes_slide.notes_text_frame.text)
    return "\n".join(buf)
new = alltext(OUT)
toks = set(w.lower() for w in re.findall(r"[A-Za-z0-9@.+&%]{4,}", srctext()))
missing = sorted(t for t in toks if t not in new)
print(f"Reconcile: {len(toks)-len(missing)}/{len(toks)} source tokens carried "
      f"({100*(len(toks)-len(missing))//max(1,len(toks))}%).")
print("  review:", ", ".join(missing) if missing else "none — nothing lost ✓")

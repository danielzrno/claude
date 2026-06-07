#!/usr/bin/env python3
"""Generate the web/print artifacts for the Standard from the curated master decks.
Renders every slide of the light & dark builds (hi-res, Inter) and assembles:
  presentation.html        — interactive 16:9 deck, light/dark toggle, keyboard nav, print 1/page
  standard-a4-landscape.html, standard-a4-portrait.html — print-ready A4, slide per page, toggle
Run from design-system/standard/. Slide PNGs land in previews/{light,dark}/."""
import os, importlib.util
from pptx import Presentation

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("rs", "/tmp/render_slide.py")
rs = importlib.util.module_from_spec(spec); spec.loader.exec_module(rs)
rs.DPI = 144            # 1.5x → 1920x1080, crisp for screen & decent for print
rs.FRAME = False        # no debug frame on deliverable images
rs._fc.clear()
render = rs.render

DECKS = {"light": "Mivada_Standard_Light.pptx", "dark": "Mivada_Standard_Dark.pptx"}
n = len(Presentation(os.path.join(HERE, DECKS["light"])).slides._sldIdLst)

for theme, deck in DECKS.items():
    outdir = os.path.join(HERE, "previews", theme); os.makedirs(outdir, exist_ok=True)
    for i in range(n):
        render(os.path.join(HERE, deck), i, os.path.join(outdir, f"slide-{i+1:02d}.png"))
print(f"rendered {n} slides x2 themes")

def imgs(theme): return [f"previews/{theme}/slide-{i+1:02d}.png" for i in range(n)]

# ---------- interactive deck ----------
slides_html = "\n".join(
    f'  <figure class="slide"><img loading="lazy" data-l="{l}" data-d="{d}" src="{l}" alt="Slide {i+1}"></figure>'
    for i, (l, d) in enumerate(zip(imgs("light"), imgs("dark"))))
PRES = f"""<!doctype html><html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mivada — the Standard · deck</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@500;700&display=swap" rel="stylesheet">
<style>
 :root{{--coral:#EA493F}}
 *{{box-sizing:border-box;margin:0;padding:0}}
 html,body{{background:#4d4d4d;font-family:Inter,system-ui,Arial,sans-serif}}
 .bar{{position:fixed;inset:0 0 auto 0;height:48px;display:flex;align-items:center;gap:16px;
   padding:0 18px;background:#111;color:#fff;z-index:10;font-size:13px;font-weight:600}}
 .bar b{{color:var(--coral)}} .bar .sp{{flex:1}}
 .bar button{{background:#222;color:#fff;border:1px solid #3a3a3a;border-radius:6px;
   padding:7px 14px;font:inherit;font-weight:700;cursor:pointer}}
 .bar button:hover{{border-color:var(--coral)}}
 .deck{{display:flex;flex-direction:column;align-items:center;gap:22px;padding:74px 0 90px}}
 .slide{{width:1280px;max-width:95vw;aspect-ratio:16/9;background:#000;
   box-shadow:0 12px 44px rgba(0,0,0,.4);scroll-snap-align:center}}
 .slide img{{width:100%;height:100%;display:block;object-fit:contain}}
 .count{{position:fixed;right:16px;bottom:14px;background:#111;color:#fff;border-radius:20px;
   padding:6px 14px;font-size:12px;font-weight:700;z-index:10}}
 @media print{{ html,body{{background:#fff}} .bar,.count{{display:none}} .deck{{gap:0;padding:0}}
   .slide{{width:100%;max-width:none;box-shadow:none;break-after:page}} }}
</style></head><body>
<div class="bar"><b>MIVADA</b> · the Standard — {n} slides <span class="sp"></span>
  <button onclick="toggle()" id="tg">◑ Dark</button>
  <button onclick="print()">⎙ Print / PDF</button></div>
<main class="deck" id="deck">
{slides_html}
</main>
<div class="count" id="count">1 / {n}</div>
<script>
 let dark=false;
 function toggle(){{dark=!dark;document.querySelectorAll('.slide img').forEach(im=>im.src=dark?im.dataset.d:im.dataset.l);
   document.getElementById('tg').textContent=dark?'◐ Light':'◑ Dark';
   document.body.style.background=dark?'#1a1a1a':'#4d4d4d';}}
 const sl=[...document.querySelectorAll('.slide')];
 function go(d){{const y=window.scrollY+innerHeight/2;let i=sl.findIndex(s=>s.offsetTop+s.offsetHeight/2>y);
   if(i<0)i=sl.length-1;i=Math.max(0,Math.min(sl.length-1,i+d));sl[i].scrollIntoView({{behavior:'smooth',block:'center'}});}}
 addEventListener('keydown',e=>{{if(['ArrowRight','ArrowDown','PageDown',' '].includes(e.key)){{e.preventDefault();go(1);}}
   if(['ArrowLeft','ArrowUp','PageUp'].includes(e.key)){{e.preventDefault();go(-1);}}
   if(e.key.toLowerCase()==='d')toggle();}});
 const io=new IntersectionObserver(es=>es.forEach(en=>{{if(en.isIntersecting)
   document.getElementById('count').textContent=(sl.indexOf(en.target)+1)+' / {n}';}}),{{threshold:.6}});
 sl.forEach(s=>io.observe(s));
</script></body></html>"""
open(os.path.join(HERE, "presentation.html"), "w").write(PRES)

# ---------- A4 print docs ----------
def a4(orient):
    page = "297mm 210mm" if orient == "landscape" else "210mm 297mm"
    pw = "297mm" if orient == "landscape" else "210mm"
    ph = "210mm" if orient == "landscape" else "297mm"
    figs = "\n".join(
        f'  <section class="pg"><img data-l="{l}" data-d="{d}" src="{l}" alt="Slide {i+1}"></section>'
        for i, (l, d) in enumerate(zip(imgs("light"), imgs("dark"))))
    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<title>Mivada — the Standard · A4 {orient}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@600;700&display=swap" rel="stylesheet">
<style>
 *{{box-sizing:border-box;margin:0;padding:0}}
 html,body{{background:#4d4d4d;font-family:Inter,system-ui,Arial,sans-serif}}
 @page{{size:{page};margin:0}}
 .bar{{position:fixed;inset:0 0 auto 0;height:44px;display:flex;gap:14px;align-items:center;
   padding:0 16px;background:#111;color:#fff;z-index:9;font-size:13px;font-weight:700}}
 .bar button{{background:#222;color:#fff;border:1px solid #3a3a3a;border-radius:6px;padding:6px 12px;font:inherit;cursor:pointer}}
 .doc{{display:flex;flex-direction:column;align-items:center;gap:16px;padding:64px 0}}
 .pg{{width:{pw};height:{ph};background:#000;box-shadow:0 8px 30px rgba(0,0,0,.4);
   display:flex;align-items:center;justify-content:center}}
 .pg img{{width:100%;height:100%;object-fit:contain}}
 @media print{{ html,body{{background:#fff}} .bar{{display:none}} .doc{{gap:0;padding:0}}
   .pg{{box-shadow:none;break-after:page}} }}
</style></head><body>
<div class="bar"><b style="color:#EA493F">MIVADA</b> · A4 {orient} · {n} pages
  <button onclick="t()">Dark</button><button onclick="print()">Print / PDF</button></div>
<div class="doc">
{figs}
</div>
<script>let d=0;function t(){{d=!d;document.querySelectorAll('.pg img').forEach(i=>i.src=d?i.dataset.d:i.dataset.l);}}</script>
</body></html>"""
open(os.path.join(HERE, "standard-a4-landscape.html"), "w").write(a4("landscape"))
open(os.path.join(HERE, "standard-a4-portrait.html"), "w").write(a4("portrait"))
print("wrote presentation.html + A4 landscape/portrait")

# ---------- Word (.docx) — one full-page slide per page, both themes ----------
from PIL import Image
from docx import Document
from docx.shared import Mm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_docx(theme, outfile, bg):
    pad = os.path.join(HERE, "previews", theme + "_page"); os.makedirs(pad, exist_ok=True)
    pages = []
    for i in range(n):
        im = Image.open(os.path.join(HERE, f"previews/{theme}/slide-{i+1:02d}.png")).convert("RGB")
        H = round(im.width * 210 / 297)                 # pad 16:9 onto exact A4-landscape ratio
        canvas = Image.new("RGB", (im.width, H), bg)
        canvas.paste(im, (0, (H - im.height) // 2))
        p = os.path.join(pad, f"p{i+1:02d}.png"); canvas.save(p); pages.append(p)
    doc = Document(); sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Mm(297), Mm(210)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Mm(0)
    for i, p in enumerate(pages):
        if i: doc.add_page_break()
        par = doc.add_paragraph(); par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = par.paragraph_format; pf.space_before = pf.space_after = Mm(0); pf.line_spacing = 1.0
        par.add_run().add_picture(p, width=Mm(297), height=Mm(210))
    doc.save(outfile); print("wrote", outfile)

build_docx("light", os.path.join(HERE, "Mivada_Standard.docx"), (250, 250, 250))
build_docx("dark",  os.path.join(HERE, "Mivada_Standard_Dark.docx"), (0, 0, 0))
# drop the padded intermediates (regenerated each run)
import shutil
for t in ("light", "dark"):
    shutil.rmtree(os.path.join(HERE, "previews", t + "_page"), ignore_errors=True)
print("done")

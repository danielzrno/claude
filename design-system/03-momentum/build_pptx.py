#!/usr/bin/env python3
# ============================================================================
# MIVADA · Concept 03 — MOMENTUM · PowerPoint builder (python-pptx)
#
# THEME dict + a small Deck helper API. A new on-brand deck is a list of
# helper calls, not hand-placed shapes:
#
#     d = Deck()
#     d.cover_coral("Mivada", "Technology, human first.", "Capability overview · 2026")
#     d.kpis_cards("Who we are", "...", [("2014","Founded"), ...])
#     d.save("momentum.pptx")
#
# Momentum = the site's energy: coral hero fields (PNG with a faint darker
# diagonal wave), M-chip tiles on cards, pill labels, rounded photo blocks,
# KPI cards, rounded corners, one tuned soft shadow on cards.
#
# Run:  python build_pptx.py    (regenerates assets if missing, then builds)
# ============================================================================
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))

# ---- ensure embedded assets exist ------------------------------------------
import make_assets
for _f in ("mchip.png", "mchip_white.png", "hero_coral.png", "photo_duotone.png", "photo_neutral.png"):
    if not os.path.exists(os.path.join(HERE, _f)):
        make_assets.build_all()
        break

# ============================================================================
# THEME — every colour/size/font in one place. Swap "font" to rebrand.
# ============================================================================
THEME = {
    "font": "Inter",
    "font_fallback": "Arial",
    "coral":      "EA493F",
    "coral_deep": "C9362B",
    "black":      "000000",
    "ink":        "111111",
    "off_white":  "FAFAFA",
    "white":      "FFFFFF",
    "warm":       "F2F3EE",
    "graphite":   "323232",
    "mid_grey":   "AEAEAE",
    "hairline":   "E4E4E0",
    "coral_soft": "FCE7E4",   # faint coral tint for soft tags
    "on_dark_eyebrow": "FFD9D4",
}

ASSET = {
    "mchip":        os.path.join(HERE, "mchip.png"),
    "mchip_white":  os.path.join(HERE, "mchip_white.png"),
    "hero_coral":   os.path.join(HERE, "hero_coral.png"),
    "photo_duo":    os.path.join(HERE, "photo_duotone.png"),
    "photo_neu":    os.path.join(HERE, "photo_neutral.png"),
}

EMU_IN = 914400


def C(hexstr):
    return RGBColor.from_string(hexstr)


# ============================================================================
# Low-level helpers
# ============================================================================
class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.W = 13.333
        self.H = 7.5
        self.blank = self.prs.slide_layouts[6]

    # ---- shapes ------------------------------------------------------------
    def _slide(self, bg=THEME["off_white"]):
        s = self.prs.slides.add_slide(self.blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = C(bg)
        return s

    def _no_line(self, shp):
        shp.line.fill.background()
        return shp

    def _rect(self, s, x, y, w, h, fill, line=None, line_w=1.0):
        sp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sp.fill.solid(); sp.fill.fore_color.rgb = C(fill)
        if line:
            sp.line.color.rgb = C(line); sp.line.width = Pt(line_w)
        else:
            sp.line.fill.background()
        sp.shadow.inherit = False
        return sp

    def _round(self, s, x, y, w, h, fill, line=None, line_w=1.0, radius=0.085, shadow=False):
        sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        try:
            sp.adjustments[0] = radius
        except Exception:
            pass
        sp.fill.solid(); sp.fill.fore_color.rgb = C(fill)
        if line:
            sp.line.color.rgb = C(line); sp.line.width = Pt(line_w)
        else:
            sp.line.fill.background()
        sp.shadow.inherit = False
        if shadow:
            self._soft_shadow(sp)
        return sp

    def _soft_shadow(self, sp):
        """One tuned soft shadow (card-based concept — allowed here)."""
        spPr = sp._element.spPr
        # remove any existing effectLst
        for e in spPr.findall(qn('a:effectLst')):
            spPr.remove(e)
        ef = spPr.makeelement(qn('a:effectLst'), {})
        sh = ef.makeelement(qn('a:outerShdw'), {
            'blurRad': str(int(0.16 * EMU_IN)),
            'dist':    str(int(0.055 * EMU_IN)),
            'dir':     '5400000',  # straight down (90°)
            'rotWithShape': '0',
        })
        clr = sh.makeelement(qn('a:srgbClr'), {'val': '111111'})
        alpha = clr.makeelement(qn('a:alpha'), {'val': '12000'})  # 12%
        clr.append(alpha); sh.append(clr); ef.append(sh)
        spPr.append(ef)

    def _pic(self, s, path, x, y, w, h):
        return s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))

    def _pic_rounded(self, s, path, x, y, w, h, radius=0.06):
        pic = self._pic(s, path, x, y, w, h)
        # apply a roundRect crop preset so photo blocks read with the radius language
        spPr = pic._element.spPr
        # remove existing prstGeom
        for g in spPr.findall(qn('a:prstGeom')):
            spPr.remove(g)
        geom = spPr.makeelement(qn('a:prstGeom'), {'prst': 'roundRect'})
        av = geom.makeelement(qn('a:avLst'), {})
        gd = av.makeelement(qn('a:gd'), {'name': 'adj', 'fmla': 'val %d' % int(radius * 100000)})
        av.append(gd); geom.append(av); spPr.append(geom)
        return pic

    # ---- text --------------------------------------------------------------
    def _text(self, s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
              line_spacing=1.0, wrap=True, space_after=0):
        """runs: list of paragraphs; each paragraph is a list of (text, size, color,
        bold, tracking, font) run-tuples. tracking in 1/100 pt em*... we pass spc directly."""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = anchor
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        for i, para in enumerate(runs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            if space_after:
                p.space_after = Pt(space_after)
            p.space_before = Pt(0)
            for (txt, size, color, bold, tracking, font) in para:
                r = p.add_run(); r.text = txt
                r.font.size = Pt(size)
                r.font.bold = bold
                r.font.name = font or THEME["font"]
                r.font.color.rgb = C(color)
                if tracking is not None:
                    self._set_tracking(r, tracking)
        return tb

    def _set_tracking(self, run, pts):
        """Letter spacing in points (use small +/- values; caps eyebrows ~ +1.2)."""
        rPr = run._r.get_or_add_rPr()
        rPr.set('spc', str(int(pts * 100)))

    def run(self, txt, size, color, bold=False, tracking=None, font=None):
        return (txt, size, color, bold, tracking, font)

    # ---- composite motifs --------------------------------------------------
    def mchip(self, s, x, y, size, reverse=False):
        path = ASSET["mchip_white"] if reverse else ASSET["mchip"]
        return self._pic(s, path, x, y, size, size)

    def eyebrow(self, s, x, y, text, color=None, w=6.0):
        color = color or THEME["coral"]
        self._text(s, x, y, w, 0.3,
                   [[self.run(text.upper(), 11.5, color, True, 1.6)]],
                   anchor=MSO_ANCHOR.MIDDLE)

    def pill(self, s, x, y, text, kind="line", w=None, h=0.40, size=12.5):
        """kind: line | coral | dark | ghost (ghost = on a coral/dark field)."""
        # estimate width if not given (~0.092in per char + padding)
        if w is None:
            w = max(0.9, 0.40 + len(text) * 0.095)
        fills = {
            "line":  (THEME["white"], THEME["ink"],   THEME["hairline"]),
            "coral": (THEME["coral"], THEME["white"],  None),
            "dark":  (THEME["ink"],   THEME["white"],  None),
            "ghost": (None,           THEME["white"],  None),   # transparent on field
            "white": (THEME["white"], THEME["coral_deep"], None),
        }
        bg, fg, ln = fills[kind]
        if kind == "ghost":
            sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
            sp.adjustments[0] = 0.5
            sp.fill.background()
            sp.line.color.rgb = C(THEME["white"]); sp.line.width = Pt(1.0)
            # 45% white line via alpha
            self._line_alpha(sp, "FFFFFF", 42000)
            sp.shadow.inherit = False
        else:
            sp = self._round(s, x, y, w, h, bg, line=ln, line_w=1.0, radius=0.5)
        tf = sp.text_frame
        tf.word_wrap = False
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = True if kind in ("coral", "dark", "white") else False
        r.font.name = THEME["font"]; r.font.color.rgb = C(fg)
        return sp, w

    def _line_alpha(self, sp, hexval, alpha):
        ln = sp.line._get_or_add_ln()
        for f in ln.findall(qn('a:solidFill')):
            ln.remove(f)
        fill = ln.makeelement(qn('a:solidFill'), {})
        clr = fill.makeelement(qn('a:srgbClr'), {'val': hexval})
        a = clr.makeelement(qn('a:alpha'), {'val': str(alpha)})
        clr.append(a); fill.append(clr)
        ln.insert(0, fill)

    def tag(self, s, x, y, text, soft=False, h=0.32, size=11.5):
        w = max(0.7, 0.36 + len(text) * 0.085)
        bg = THEME["coral_soft"] if soft else THEME["warm"]
        fg = THEME["coral_deep"] if soft else THEME["graphite"]
        ln = "F6CFC9" if soft else THEME["hairline"]
        sp = self._round(s, x, y, w, h, bg, line=ln, line_w=0.75, radius=0.5)
        tf = sp.text_frame; tf.word_wrap = False
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = True
        r.font.name = THEME["font"]; r.font.color.rgb = C(fg)
        return sp, w

    def slide_no(self, s, n, on_dark=False):
        col = "FFFFFF" if on_dark else THEME["mid_grey"]
        alpha = 60000 if on_dark else None
        tb = self._text(s, self.W - 2.35, self.H - 0.52, 1.85, 0.3,
                        [[self.run("%02d / 08" % n, 10.5, col, False, 0.4)]],
                        align=PP_ALIGN.RIGHT)
        if on_dark:
            # dim the white
            for p in tb.text_frame.paragraphs:
                for r in p.runs:
                    self._run_alpha(r, alpha)
        return tb

    def _run_alpha(self, run, alpha):
        rPr = run._r.get_or_add_rPr()
        for f in rPr.findall(qn('a:solidFill')):
            rPr.remove(f)
        fill = rPr.makeelement(qn('a:solidFill'), {})
        clr = fill.makeelement(qn('a:srgbClr'), {'val': 'FFFFFF'})
        a = clr.makeelement(qn('a:alpha'), {'val': str(alpha)})
        clr.append(a); fill.append(clr)
        rPr.append(fill)

    # ====================================================================
    # SLIDE HELPERS (the public API)
    # ====================================================================
    def cover_coral(self, name, headline, meta):
        s = self._slide(THEME["coral"])
        # full-bleed coral hero PNG (faint darker diagonal wave)
        self._pic(s, ASSET["hero_coral"], 0, 0, self.W, self.H)
        m = 0.92
        # wordmark lock-up
        self.mchip(s, m, 0.78, 0.78, reverse=True)
        self._text(s, m + 0.98, 0.78, 4.0, 0.78,
                   [[self.run("Mivada", 23, THEME["white"], True, -0.4)]],
                   anchor=MSO_ANCHOR.MIDDLE)
        # meta pill, top-right
        pw = max(2.6, 0.4 + len(meta) * 0.095)
        self.pill(s, self.W - m - pw, 0.92, meta, kind="ghost", w=pw, h=0.44, size=12.5)
        # eyebrow + big headline
        self.eyebrow(s, m, 3.45, "Australian technology consultancy", color=THEME["on_dark_eyebrow"])
        self._text(s, m - 0.02, 3.78, 10.5, 2.4,
                   [[self.run(headline, 62, THEME["white"], True, -1.2)]],
                   line_spacing=0.98)
        # service pills along the bottom
        pills = ["Workday & ERP", "Payroll", "Data & AI", "Automation"]
        px = m
        for t in pills:
            sp, w = self.pill(s, px, self.H - 1.15, t, kind="ghost", h=0.44, size=12.5)
            px += w + 0.16
        self.slide_no(s, 1, on_dark=True)
        return s

    def kpis_cards(self, eyebrow, headline_runs, lede, kpis):
        """kpis: list of (num, label); the 3rd is highlighted coral by default."""
        s = self._slide(THEME["off_white"])
        m = 0.92
        self.eyebrow(s, m, 0.74, eyebrow)
        self._text(s, m - 0.02, 1.06, 9.2, 1.7, headline_runs, line_spacing=1.02)
        # M-chip top-right
        self.mchip(s, self.W - m - 0.86, 0.74, 0.86)
        # lede
        self._lede(s, m, 2.95, 8.6, lede)
        # KPI cards row
        self._kpi_row(s, m, 5.0, self.W - 2 * m, kpis)
        self.slide_no(s, 2)
        return s

    def _lede(self, s, x, y, w, parts):
        """parts: list of (text, bold) — ink when bold, graphite otherwise."""
        runs = []
        for txt, bold in parts:
            col = THEME["ink"] if bold else THEME["graphite"]
            runs.append(self.run(txt, 17, col, bold, None))
        self._text(s, x, y, w, 1.2, [runs], line_spacing=1.32)

    def _kpi_row(self, s, x, y, w, kpis, h=1.55, gap=0.28):
        n = len(kpis)
        cw = (w - gap * (n - 1)) / n
        styles = {
            "white": (THEME["white"], THEME["coral"], THEME["graphite"], THEME["hairline"]),
            "coral": (THEME["coral"], THEME["white"], "FFE3DF", None),
            "ink":   (THEME["ink"],   THEME["white"], "C9C9C9", None),
        }
        for i, item in enumerate(kpis):
            if len(item) == 3:
                num, label, kind = item
            else:
                num, label = item; kind = "coral" if i == n - 1 else "white"
            bg, numc, labc, ln = styles[kind]
            cx = x + i * (cw + gap)
            card = self._round(s, cx, y, cw, h, bg, line=ln, line_w=1.0, radius=0.085,
                               shadow=(kind == "white" or kind == "coral"))
            self._text(s, cx + 0.28, y + 0.24, cw - 0.5, 0.85,
                       [[self.run(num, 38, numc, True, -1.0)]], anchor=MSO_ANCHOR.TOP)
            self._text(s, cx + 0.30, y + h - 0.52, cw - 0.55, 0.42,
                       [[self.run(label, 12.5, labc, False, None)]],
                       anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.05)

    def pillars_chips(self, eyebrow, headline, pill_text, pillars):
        """pillars: list of (index, title, body). 4 cards, each with an M-chip."""
        s = self._slide(THEME["off_white"])
        m = 0.92
        self.eyebrow(s, m, 0.74, eyebrow)
        self._text(s, m - 0.02, 1.06, 8.2, 0.9, [[self.run(headline, 33, THEME["ink"], True, -0.8)]])
        pw = max(2.4, 0.4 + len(pill_text) * 0.092)
        self.pill(s, self.W - m - pw, 1.16, pill_text, kind="line", w=pw, h=0.44, size=12.5)
        # 2x2 grid of cards
        gx, gy = m, 2.35
        gw = self.W - 2 * m
        gap = 0.30
        cw = (gw - gap) / 2
        ch = 2.18
        for i, (idx, title, body) in enumerate(pillars):
            col = i % 2; row = i // 2
            cx = gx + col * (cw + gap)
            cy = gy + row * (ch + gap)
            self._round(s, cx, cy, cw, ch, THEME["white"], line=THEME["hairline"],
                        line_w=1.0, radius=0.05, shadow=True)
            self.mchip(s, cx + 0.30, cy + 0.30, 0.50)
            self._text(s, cx + 0.95, cy + 0.34, cw - 1.2, 0.4,
                       [[self.run(idx, 12.5, THEME["coral"], True, 0.4)]],
                       anchor=MSO_ANCHOR.MIDDLE)
            self._text(s, cx + 0.30, cy + 0.98, cw - 0.6, 0.5,
                       [[self.run(title, 18, THEME["ink"], True, -0.3)]])
            self._text(s, cx + 0.30, cy + 1.46, cw - 0.6, 0.7,
                       [[self.run(body, 13, THEME["graphite"], False, None)]],
                       line_spacing=1.3)
        self.slide_no(s, 3)
        return s

    def steps(self, eyebrow, headline, pill_text, lede, steps):
        """steps: list of (label, stepno, body). Last is coral."""
        s = self._slide(THEME["off_white"])
        m = 0.92
        self.eyebrow(s, m, 0.74, eyebrow)
        self._text(s, m - 0.02, 1.06, 7.0, 0.8, [[self.run(headline, 33, THEME["ink"], True, -0.8)]])
        pw = max(3.0, 0.4 + len(pill_text) * 0.092)
        self.pill(s, self.W - m - pw, 1.16, pill_text, kind="coral", w=pw, h=0.44, size=12.5)
        self._text(s, m, 2.0, 9.0, 0.5,
                   [[self.run(lede, 17, THEME["graphite"], False, None)]], line_spacing=1.3)
        # 4 cards
        gx, gy = m, 2.95
        gw = self.W - 2 * m
        gap = 0.26
        cw = (gw - gap * 3) / 4
        ch = 3.1
        for i, (label, stepno, body) in enumerate(steps):
            coral = (i == len(steps) - 1)
            cx = gx + i * (cw + gap)
            bg = THEME["coral"] if coral else THEME["white"]
            self._round(s, cx, gy, cw, ch, bg, line=(None if coral else THEME["hairline"]),
                        line_w=1.0, radius=0.06, shadow=True)
            # pill label (dark, or white on coral card)
            if coral:
                self.pill(s, cx + 0.26, gy + 0.30, label, kind="white", w=1.15, h=0.36, size=12.5)
                stepc = "FFE3DF"; bodyc = "FFFFFF"
            else:
                self.pill(s, cx + 0.26, gy + 0.30, label, kind="dark", w=1.25, h=0.36, size=12.5)
                stepc = THEME["coral"]; bodyc = THEME["graphite"]
            self._text(s, cx + 0.28, gy + 0.92, cw - 0.5, 0.35,
                       [[self.run("Step %02d" % stepno, 12, stepc, True, 0.3)]])
            self._text(s, cx + 0.28, gy + 1.40, cw - 0.52, 1.5,
                       [[self.run(body, 13, bodyc, False, None)]], line_spacing=1.32)
        self.slide_no(s, 4)
        return s

    def split_photo(self, eyebrow, headline, caps, cap_label, duotone=True):
        """caps: list of (n, title, desc). Photo block on the right."""
        s = self._slide(THEME["off_white"])
        m = 0.92
        colw = 6.4
        self.eyebrow(s, m, 0.92, eyebrow)
        self._text(s, m - 0.02, 1.26, colw, 1.5,
                   [[self.run(headline, 30, THEME["ink"], True, -0.7)]], line_spacing=1.03)
        # capability list
        cy = 2.95
        for (n, title, desc) in caps:
            self._text(s, m, cy, 0.5, 0.4, [[self.run(n, 13, THEME["coral"], True, 0.2)]])
            self._text(s, m + 0.5, cy - 0.02, colw - 0.5, 0.4,
                       [[self.run(title, 15.5, THEME["ink"], True, -0.2)]])
            self._text(s, m + 0.5, cy + 0.30, colw - 0.5, 0.4,
                       [[self.run(desc, 12.8, THEME["graphite"], False, None)]], line_spacing=1.2)
            cy += 0.80
        # photo block on the right
        px = m + colw + 0.55
        pw = self.W - px - m
        py = 0.92
        ph = self.H - 2 * py
        path = ASSET["photo_duo"] if duotone else ASSET["photo_neu"]
        self._pic_rounded(s, path, px, py, pw, ph, radius=0.05)
        # caption chip at the bottom-left of the photo
        self._photo_caption(s, px + 0.26, py + ph - 0.60, cap_label, on_dark=duotone)
        self.slide_no(s, 5)
        return s

    def _photo_caption(self, s, x, y, text, on_dark=True):
        w = max(2.0, 0.7 + len(text) * 0.085)
        bg = THEME["ink"] if on_dark else THEME["white"]
        fg = THEME["white"] if on_dark else THEME["graphite"]
        cap = self._round(s, x, y, w, 0.40, bg, radius=0.5)
        if on_dark:
            self._fill_alpha(cap, "111111", 40000)
        # small M-chip glyph inside
        self.mchip(s, x + 0.10, y + 0.075, 0.25)
        tf = cap.text_frame; tf.word_wrap = False
        tf.margin_left = Inches(0.42); tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = text
        r.font.size = Pt(11.5); r.font.bold = False
        r.font.name = THEME["font"]; r.font.color.rgb = C(fg)

    def _fill_alpha(self, sp, hexval, alpha):
        spPr = sp.fill._xPr
        # easier: rebuild solidFill with alpha
        for f in spPr.findall(qn('a:solidFill')):
            spPr.remove(f)
        fill = spPr.makeelement(qn('a:solidFill'), {})
        clr = fill.makeelement(qn('a:srgbClr'), {'val': hexval})
        a = clr.makeelement(qn('a:alpha'), {'val': str(alpha)})
        clr.append(a); fill.append(clr)
        # insert before line element if present
        ln = spPr.find(qn('a:ln'))
        if ln is not None:
            ln.addprevious(fill)
        else:
            spPr.append(fill)

    def stats3_quote(self, eyebrow, headline, kpis, quote_parts, name, role):
        s = self._slide(THEME["off_white"])
        m = 0.92
        self.eyebrow(s, m, 0.74, eyebrow)
        self._text(s, m - 0.02, 1.06, 9.0, 0.8, [[self.run(headline, 33, THEME["ink"], True, -0.8)]])
        # KPI row
        self._kpi_row(s, m, 2.15, self.W - 2 * m, kpis)
        # quote card
        qy = 4.25
        qh = 2.35
        self._round(s, m, qy, self.W - 2 * m, qh, THEME["white"], line=THEME["hairline"],
                    line_w=1.0, radius=0.045, shadow=True)
        # the quote text (mixed runs, em word coral)
        runs = []
        for txt, em in quote_parts:
            col = THEME["coral"] if em else THEME["ink"]
            runs.append(self.run(txt, 21, col, False if not em else False, -0.4))
        self._text(s, m + 0.45, qy + 0.40, self.W - 2 * m - 4.0, 1.6, [runs], line_spacing=1.28)
        # attribution: avatar + name/role, right side
        ax = self.W - m - 3.2
        self._pic(s, ASSET["mchip_white"], ax, qy + qh - 0.95, 0.0001, 0.0001)  # noop guard (unused)
        # coral avatar circle
        av = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ax), Inches(qy + 0.78), Inches(0.62), Inches(0.62))
        av.fill.solid(); av.fill.fore_color.rgb = C(THEME["coral"]); av.line.fill.background()
        av.shadow.inherit = False
        self._text(s, ax + 0.78, qy + 0.80, 2.4, 0.35, [[self.run(name, 14, THEME["ink"], True, -0.2)]],
                   anchor=MSO_ANCHOR.MIDDLE)
        self._text(s, ax + 0.78, qy + 1.16, 2.4, 0.3, [[self.run(role, 12, THEME["mid_grey"], False, None)]],
                   anchor=MSO_ANCHOR.MIDDLE)
        self.slide_no(s, 6)
        return s

    def reasons(self, eyebrow, headline, items):
        """items: list of (title, body, [tags], coral?)."""
        s = self._slide(THEME["off_white"])
        m = 0.92
        self.eyebrow(s, m, 0.74, eyebrow)
        self._text(s, m - 0.02, 1.06, 8.2, 0.9, [[self.run(headline, 33, THEME["ink"], True, -0.8)]])
        self.mchip(s, self.W - m - 0.86, 0.74, 0.86)
        gx, gy = m, 2.45
        gw = self.W - 2 * m
        gap = 0.30
        cw = (gw - gap * 2) / 3
        ch = 3.55
        for i, (title, body, tags, coral) in enumerate(items):
            cx = gx + i * (cw + gap)
            bg = THEME["coral"] if coral else THEME["white"]
            self._round(s, cx, gy, cw, ch, bg, line=(None if coral else THEME["hairline"]),
                        line_w=1.0, radius=0.05, shadow=True)
            self.mchip(s, cx + 0.30, gy + 0.32, 0.50)
            tcol = THEME["white"] if coral else THEME["ink"]
            bcol = "FFFFFF" if coral else THEME["graphite"]
            self._text(s, cx + 0.30, gy + 1.02, cw - 0.6, 0.55,
                       [[self.run(title, 17.5, tcol, True, -0.3)]], line_spacing=1.02)
            self._text(s, cx + 0.30, gy + 1.62, cw - 0.6, 1.2,
                       [[self.run(body, 13, bcol, False, None)]], line_spacing=1.3)
            # tags at bottom
            tx = cx + 0.30
            ty = gy + ch - 0.58
            for t in tags:
                if coral:
                    sp, w = self._ghost_tag_on_coral(s, tx, ty, t)
                else:
                    sp, w = self.tag(s, tx, ty, t, soft=True)
                tx += w + 0.12
        self.slide_no(s, 7)
        return s

    def _ghost_tag_on_coral(self, s, x, y, text):
        w = max(0.7, 0.36 + len(text) * 0.085)
        sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.32))
        sp.adjustments[0] = 0.5
        sp.fill.background()
        sp.line.color.rgb = C("FFFFFF"); sp.line.width = Pt(1.0)
        self._line_alpha(sp, "FFFFFF", 45000)
        sp.shadow.inherit = False
        tf = sp.text_frame; tf.word_wrap = False
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = text
        r.font.size = Pt(11.5); r.font.name = THEME["font"]; r.font.color.rgb = C("FFFFFF")
        return sp, w

    def contact(self, headline, email, site, contact_bits):
        s = self._slide(THEME["coral"])
        self._pic(s, ASSET["hero_coral"], 0, 0, self.W, self.H)
        m = 0.92
        self.mchip(s, m, 0.78, 0.78, reverse=True)
        self._text(s, m + 0.98, 0.78, 4.0, 0.78,
                   [[self.run("Mivada", 23, THEME["white"], True, -0.4)]],
                   anchor=MSO_ANCHOR.MIDDLE)
        diff = "Experience the Mivada difference"
        pw = max(3.4, 0.4 + len(diff) * 0.092)
        self.pill(s, self.W - m - pw, 0.92, diff, kind="ghost", w=pw, h=0.44, size=12.5)
        # eyebrow + headline
        self.eyebrow(s, m, 3.05, "Let's talk", color=THEME["on_dark_eyebrow"])
        self._text(s, m - 0.02, 3.42, 9.6, 1.9,
                   [[self.run(headline, 44, THEME["white"], True, -1.0)]], line_spacing=1.0)
        # CTA pills
        self.pill(s, m, 5.55, email, kind="white", w=max(2.6, 0.5 + len(email) * 0.10), h=0.52, size=15)
        ew = max(2.6, 0.5 + len(email) * 0.10)
        self.pill(s, m + ew + 0.18, 5.55, site, kind="ghost", w=max(2.0, 0.5 + len(site) * 0.10), h=0.52, size=15)
        # contact line
        line = "      ·      ".join(contact_bits)
        tb = self._text(s, m, self.H - 0.95, 11.0, 0.4, [[self.run(line, 13.5, "FFFFFF", False, None)]])
        for p in tb.text_frame.paragraphs:
            for r in p.runs:
                self._run_alpha(r, 82000)
        self.slide_no(s, 8, on_dark=True)
        return s

    def save(self, path):
        self.prs.save(path)


# ============================================================================
# BUILD THE DECK — content only; styling lives in the helpers above.
# ============================================================================
def build():
    d = Deck()

    # 1 · Cover
    d.cover_coral("Mivada", "Technology,\nhuman first.", "Capability overview · 2026")

    # 2 · Who we are
    d.kpis_cards(
        "Who we are",
        [[d.run("We turn enterprise platforms into ", 33, THEME["ink"], True, -0.8),
          d.run("measurable human outcomes", 33, THEME["coral"], True, -0.8),
          d.run(".", 33, THEME["ink"], True, -0.8)]],
        [("Mivada is an Australian technology consultancy that turns Workday, payroll, data and AI into ", False),
         ("outcomes people actually feel", True),
         (" — built around how teams really work.", False)],
        [("2014", "Founded — formerly LJM Infotech"),
         ("120+", "Specialists, certified not generalist"),
         ("AU + India", "Onshore leadership, offshore delivery", "coral")],
    )

    # 3 · What we do
    d.pillars_chips(
        "What we do", "Four pillars, one operating idea.", "Human first, end to end",
        [("01", "Workday & ERP", "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."),
         ("02", "Payroll consulting", "Compliant, accurate payroll across complex awards and multi-entity structures."),
         ("03", "Data & AI", "Lakehouse architecture, governed analytics and applied AI on your people and finance data."),
         ("04", "Intelligent automation", "RPA, ML and process design combined to remove low-value work for good.")],
    )

    # 4 · How we work
    d.steps(
        "How we work", "Think human first.", "Weeks, not multi-year programs",
        "A four-step engagement model that puts people before platforms.",
        [("Listen", 1, "Understand the people and the work before the technology."),
         ("Design", 2, "Shape the platform around how teams actually operate."),
         ("Deliver", 3, "Implement in weeks-long increments, not multi-year programs."),
         ("Care", 4, "Stay on after go-live; measure outcomes, not tickets closed.")],
    )

    # 5 · Data & AI
    d.split_photo(
        "Data & AI", "Decisions on current data, not last quarter's.",
        [("01", "Lakehouse architecture", "Designed by certified consultants around your platforms."),
         ("02", "Delta Lake pipelines", "Batch and streaming, built to govern and to scale."),
         ("03", "Governed KPI store", "Databricks integrated with Workday, payroll and finance."),
         ("04", "Real-time insight", "So leaders decide on what is true right now.")],
        "Databricks · Workday · Finance", duotone=True,
    )

    # 6 · Outcomes
    d.stats3_quote(
        "Outcomes", "Proof, in numbers and in practice.",
        [("40+", "Workday deployments delivered", "white"),
         ("98%", "Client retention", "coral"),
         ("Weeks", "To value — not months", "ink")],
        [("They listened first, then built around how our teams actually work. Go-live felt like a ", False),
         ("beginning", True),
         (", not a handover.", False)],
        "People & Systems lead", "Enterprise client · illustrative",
    )

    # 7 · Why Mivada
    d.reasons(
        "Why Mivada", "Built to be used, not just installed.",
        [("Certified specialists", "Workday- and Databricks-certified consultants, not generalists learning on your time.",
          ["Workday", "Databricks"], False),
         ("Onshore + offshore", "Senior onshore leadership paired with cost-effective delivery from India.",
          ["Australia", "India"], False),
         ("People-first change", "Adoption built in from day one, so the platform actually gets used.",
          ["Change", "Adoption"], True)],
    )

    # 8 · Contact
    d.contact(
        "Let's build smarter systems, with people at the centre.",
        "hello@mivada.com", "mivada.com",
        ["Sydney", "Melbourne", "Onshore AU & India", "Figures illustrative"],
    )

    out = os.path.join(HERE, "momentum.pptx")
    d.save(out)
    print("wrote", out)


if __name__ == "__main__":
    build()

#!/usr/bin/env python3
# =============================================================================
# SIGNAL — Mivada design system 02  ·  python-pptx builder
# Data / AI. Dark by default, precise, technical. Engineering-led.
#
# Token-light core: a THEME dict + a Deck class of helpers (cover, kpis,
# pillars, pipeline, stats3, quote, reasons, contact). A new deck is a list of
# helper calls, not hand-placed shapes. Build → signal.pptx.
# =============================================================================
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))

# ----------------------------------------------------------------------------
# THEME  — single source of truth (hex strings, fonts, sizes)
# ----------------------------------------------------------------------------
THEME = {
    # palette
    "ink":     "0E1416",   # near-black, faint cool-green undertone — dark bg
    "surface": "161E21",   # raised panel
    "line":    "28343A",   # hairline / border on dark
    "paper":   "ECEFEC",   # off-white text on dark
    "mint":    "34E2A8",   # signal mint — PRIMARY accent
    "slate":   "7E97A1",   # muted secondary text
    "amber":   "E9B949",   # warm highlight — rare 2nd accent
    "node":    "131B1E",   # slightly-raised pipeline cell on ink

    # type
    "sans":    "Hanken Grotesk",
    "mono":    "JetBrains Mono",

    # geometry (inches)
    "W": 13.333, "H": 7.5,
    "mx": 0.92,            # outer margin x
    "radius": 0.05,        # ~4px corner
}

C = {k: RGBColor.from_string(THEME[k]) for k in
     ("ink", "surface", "line", "paper", "mint", "slate", "amber", "node")}

EMU_IN = 914400


# ----------------------------------------------------------------------------
# dot-grid background PNG — THE signature texture (generated once)
# ----------------------------------------------------------------------------
def make_dotgrid(path, tint="slate"):
    """Render a 1280x720 ink field with faint dots; cached on disk."""
    from PIL import Image, ImageDraw
    w, h, gap = 1280, 720, 44
    img = Image.new("RGB", (w, h), "#" + THEME["ink"])
    d = ImageDraw.Draw(img, "RGBA")
    base = THEME[tint]
    r, g, b = int(base[0:2], 16), int(base[2:4], 16), int(base[4:6], 16)
    a = 46 if tint == "slate" else 40
    for y in range(0, h + gap, gap):
        for x in range(0, w + gap, gap):
            d.ellipse((x - 1, y - 1, x + 1, y + 1), fill=(r, g, b, a))
    img.save(path)
    return path


# ----------------------------------------------------------------------------
# Deck — the helper layer
# ----------------------------------------------------------------------------
class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(THEME["W"])
        self.prs.slide_height = Inches(THEME["H"])
        self.blank = self.prs.slide_layouts[6]
        self.bg_slate = make_dotgrid(os.path.join(HERE, "_dotgrid.png"), "slate")
        self.bg_mint = make_dotgrid(os.path.join(HERE, "_dotgrid_mint.png"), "mint")

    # -- low-level ----------------------------------------------------------
    def _slide(self, dotgrid=True, mint=False):
        s = self.prs.slides.add_slide(self.blank)
        # solid ink underlay (so export never shows white)
        self._rect(s, 0, 0, THEME["W"], THEME["H"], C["ink"], line=None)
        if dotgrid:
            s.shapes.add_picture(self.bg_mint if mint else self.bg_slate,
                                 0, 0, Inches(THEME["W"]), Inches(THEME["H"]))
        return s

    def _rect(self, s, x, y, w, h, fill, line=None, line_w=1.0, radius=False):
        shp_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        sp = s.shapes.add_shape(shp_type, Inches(x), Inches(y), Inches(w), Inches(h))
        sp.shadow.inherit = False
        if radius:
            try:
                sp.adjustments[0] = 0.06
            except Exception:
                pass
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid(); sp.fill.fore_color.rgb = fill
        if line is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb = line; sp.line.width = Pt(line_w)
        return sp

    def _hline(self, s, x, y, w, color=None, weight=1.0):
        color = color or C["line"]
        ln = s.shapes.add_connector(2, Inches(x), Inches(y), Inches(x + w), Inches(y))
        ln.line.color.rgb = color; ln.line.width = Pt(weight)
        ln.shadow.inherit = False
        return ln

    def _vline(self, s, x, y, h, color=None, weight=1.0):
        color = color or C["line"]
        ln = s.shapes.add_connector(2, Inches(x), Inches(y), Inches(x), Inches(y + h))
        ln.line.color.rgb = color; ln.line.width = Pt(weight)
        ln.shadow.inherit = False
        return ln

    def _text(self, s, x, y, w, h, runs, *, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, leading=1.0, space=0.0, wrap=True):
        """runs: list of paragraphs; each paragraph is a list of (text, font, size,
        color, bold, tracking) run-tuples."""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = anchor
        for m in ("left", "right", "top", "bottom"):
            setattr(tf, "margin_" + m, 0)
        for i, para in enumerate(runs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = leading
            if space:
                p.space_before = Pt(space)
            for (txt, font, size, color, bold, track) in para:
                r = p.add_run(); r.text = txt
                r.font.name = font; r.font.size = Pt(size); r.font.bold = bold
                r.font.color.rgb = color
                self._set_face(r, font)
                if track:
                    self._tracking(r, track)
        return tb

    def _set_face(self, run, font):
        """Force latin + east-asian + cs faces so the brand font always wins."""
        rPr = run._r.get_or_add_rPr()
        for tag in ("latin", "ea", "cs"):
            e = rPr.find(qn("a:" + tag))
            if e is None:
                e = rPr.makeelement(qn("a:" + tag), {}); rPr.append(e)
            e.set("typeface", font)

    def _tracking(self, run, pts):
        run._r.get_or_add_rPr().set("spc", str(int(pts * 100)))

    # quick run-tuple builders -------------------------------------------
    def mono(self, t, size=11, color=C["slate"], bold=False, track=0.6):
        return (t, THEME["mono"], size, color, bold, track)

    def sans(self, t, size=18, color=C["paper"], bold=False, track=0.0):
        return (t, THEME["sans"], size, color, bold, track)

    # -- chrome -------------------------------------------------------------
    def _frame(self, s, tag):
        mx = THEME["mx"]
        self._hline(s, mx, 0.52, THEME["W"] - 2 * mx, C["line"], 1.0)
        self._text(s, THEME["W"] - mx - 1.2, 0.30, 1.2, 0.3,
                   [[self.mono(tag, 10.5, C["slate"], track=1.0)]],
                   align=PP_ALIGN.RIGHT)

    def _kicker(self, s, x, y, num, label):
        self._text(s, x, y, 7.0, 0.3,
                   [[self.mono("// %s " % num, 12, C["mint"], bold=True),
                     self.mono("— " + label, 12, C["slate"])]])

    def _h2(self, s, x, y, w, segs, size=33):
        """segs: list of (text, is_mint)."""
        runs = [self.sans(t, size, C["mint"] if mint else C["paper"], bold=True, track=-0.2)
                for (t, mint) in segs]
        self._text(s, x, y, w, 1.4, [runs], leading=1.02)

    # ======================================================================
    # SLIDE HELPERS
    # ======================================================================
    def cover(self, toplabel, title_segs, sub, coord_top, coord_bot):
        s = self._slide(dotgrid=True)
        mx = THEME["mx"]
        self._text(s, mx, 0.46, 9, 0.3,
                   [[self.mono(toplabel, 12, C["slate"], track=0.7)]])
        # big Hanken headline — title_segs: list of lines, each a list of (text, is_mint)
        self._text(s, mx, 2.35, 11, 2.6,
                   [[self.sans(t, 67, C["mint"] if mint else C["paper"], bold=True, track=-1.2)
                     for (t, mint) in line]
                    for line in title_segs],
                   leading=0.96)
        self._text(s, mx, 5.05, 7.6, 0.9,
                   [[self.sans(sub, 17, C["slate"])]], leading=1.45)
        # single mint signal element (the one allowed glow) bottom-left
        self._signal_dot(s, mx, THEME["H"] - 1.06)
        # mono coordinate detail bottom-right
        self._text(s, THEME["W"] - mx - 4.0, THEME["H"] - 1.12, 4.0, 0.8,
                   [[self.mono(coord_top, 11, C["mint"], track=0.4)],
                    [self.mono(coord_bot, 11, C["slate"], track=0.4)]],
                   align=PP_ALIGN.RIGHT, leading=1.5, space=4)
        return s

    def _signal_dot(self, s, x, y):
        # faint outer ring (hairline) + solid mint dot + soft mint glow
        ring = self._rect(s, x - 0.10, y - 0.10, 0.36, 0.36, None,
                          line=C["mint"], line_w=0.75, radius=True)
        ring.adjustments[0] = 1.0
        glow = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                  Inches(0.16), Inches(0.16))
        glow.fill.solid(); glow.fill.fore_color.rgb = C["mint"]
        glow.line.fill.background()
        # the ONE subtle mint glow used once in the deck
        self._glow(glow, "34E2A8")

    def _glow(self, shape, hexcol):
        spPr = shape._element.spPr
        ef = spPr.makeelement(qn("a:effectLst"), {})
        glow = ef.makeelement(qn("a:glow"), {"rad": str(int(0.16 * EMU_IN))})
        clr = glow.makeelement(qn("a:srgbClr"), {"val": hexcol})
        alpha = clr.makeelement(qn("a:alpha"), {"val": "55000"})
        clr.append(alpha); glow.append(clr); ef.append(glow); spPr.append(ef)

    # ---- 2: who we are ----------------------------------------------------
    def who(self, num, label, title_segs, oneliner, kpis):
        s = self._slide(); self._frame(s, "[01]")
        mx = THEME["mx"]
        self._kicker(s, mx, 0.95, num, label)
        self._h2(s, mx, 1.42, 7.2, title_segs, size=33)
        # right-hand one-liner
        self._text(s, 8.55, 1.55, THEME["W"] - mx - 8.55, 1.8,
                   [[self.sans(oneliner, 17, C["paper"])]], leading=1.42)
        self.kpis(s, kpis, y=4.35)
        return s

    # ---- KPI tiles (hairline cells) --------------------------------------
    def kpis(self, s, items, y=4.35):
        mx = THEME["mx"]
        gap = 0.28
        total = THEME["W"] - 2 * mx
        w = (total - gap * (len(items) - 1)) / len(items)
        h = 2.18
        for i, (num, cap) in enumerate(items):
            x = mx + i * (w + gap)
            self._rect(s, x, y, w, h, C["surface"], line=C["line"], line_w=1.0, radius=True)
            self._text(s, x + 0.28, y + 0.30, w - 0.5, 1.1,
                       [[self.sans(num, 50, C["mint"], bold=True, track=-1.0)]],
                       leading=0.9)
            self._hline(s, x + 0.30, y + h - 0.72, w - 0.6, C["line"], 0.75)
            self._text(s, x + 0.28, y + h - 0.56, w - 0.5, 0.5,
                       [[self.mono(cap, 10.5, C["slate"], track=0.4)]], leading=1.2)

    # ---- 3: pillars -------------------------------------------------------
    def pillars(self, num, label, title_segs, items):
        s = self._slide(); self._frame(s, "[02]")
        mx = THEME["mx"]
        self._kicker(s, mx, 0.95, num, label)
        self._h2(s, mx, 1.42, 11, title_segs, size=33)
        gx, gy = 0.28, 0.28
        total = THEME["W"] - 2 * mx
        w = (total - gx) / 2
        top, h = 2.55, 1.95
        for i, (ix, name, line) in enumerate(items):
            col, row = i % 2, i // 2
            x = mx + col * (w + gx)
            y = top + row * (h + gy)
            self._rect(s, x, y, w, h, C["surface"], line=C["line"], line_w=1.0, radius=True)
            # mint left tick
            self._rect(s, x, y + 0.32, 0.045, h - 0.64, C["mint"], line=None)
            self._text(s, x + 0.36, y + 0.30, w - 0.6, 0.3,
                       [[self.mono(ix, 11.5, C["slate"], track=0.6)]])
            self._text(s, x + 0.36, y + 0.62, w - 0.6, 0.5,
                       [[self.sans(name, 21, C["paper"], bold=True, track=-0.2)]])
            self._text(s, x + 0.36, y + 1.12, w - 0.62, 0.7,
                       [[self.sans(line, 13.5, C["slate"])]], leading=1.4)
        return s

    # ---- 4: how we work (track of nodes + mint arrows) -------------------
    def track(self, num, label, title_segs, stops):
        s = self._slide(); self._frame(s, "[03]")
        mx = THEME["mx"]
        self._kicker(s, mx, 0.95, num, label)
        self._h2(s, mx, 1.42, 11, title_segs, size=33)
        total = THEME["W"] - 2 * mx
        n = len(stops)
        colw = total / n
        rail_y = 4.30
        self._hline(s, mx + 0.1, rail_y, total - 0.2, C["line"], 1.0)
        for i, (idx, name, desc) in enumerate(stops):
            cx = mx + i * colw
            # mono index + arrow above
            arrow = "" if i == n - 1 else "  →"
            self._text(s, cx, rail_y - 0.78, colw - 0.3, 0.3,
                       [[self.mono(idx, 11, C["slate"], track=0.6),
                         self.mono(arrow, 13, C["mint"], bold=True)]])
            # node dot on the rail (ink core, mint ring)
            dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx - 0.085),
                                     Inches(rail_y - 0.085), Inches(0.17), Inches(0.17))
            dot.fill.solid(); dot.fill.fore_color.rgb = C["ink"]
            dot.line.color.rgb = C["mint"]; dot.line.width = Pt(2.0)
            dot.shadow.inherit = False
            # name + desc below
            self._text(s, cx, rail_y + 0.30, colw - 0.45, 0.5,
                       [[self.sans(name, 23, C["paper"], bold=True, track=-0.2)]])
            self._text(s, cx, rail_y + 0.84, colw - 0.5, 1.0,
                       [[self.sans(desc, 13, C["slate"])]], leading=1.38)
        return s

    # ---- 5: data & ai pipeline (showpiece) -------------------------------
    def pipeline(self, num, label, title_segs, standfirst, nodes):
        s = self._slide(dotgrid=True, mint=True); self._frame(s, "[04]")
        mx = THEME["mx"]
        # left column
        self._kicker(s, mx, 1.30, num, label)
        self._h2(s, mx, 1.78, 3.7, title_segs, size=31)
        self._text(s, mx, 3.42, 3.5, 2.0,
                   [[self.sans(standfirst, 15, C["slate"])]], leading=1.46)
        # right column pipeline of hairline cells + mint flow arrows
        left = 4.95
        right = THEME["W"] - mx
        span = right - left
        n = len(nodes)
        arrow_w = 0.40
        cw = (span - arrow_w * (n - 1)) / n
        h = 2.05
        cy = (THEME["H"] - h) / 2 + 0.22
        for i, (ix, t, sub) in enumerate(nodes):
            x = left + i * (cw + arrow_w)
            self._rect(s, x, cy, cw, h, C["node"], line=C["line"], line_w=1.0, radius=True)
            self._text(s, x + 0.18, cy + 0.20, cw - 0.32, 0.3,
                       [[self.mono(ix, 10, C["mint"], track=0.4)]])
            self._text(s, x + 0.18, cy + 0.58, cw - 0.30, 1.0,
                       [[self.sans(t, 15, C["paper"], bold=True, track=-0.3)]], leading=1.04)
            self._text(s, x + 0.18, cy + h - 0.62, cw - 0.30, 0.5,
                       [[self.mono(sub, 9, C["slate"], track=0.2)]], leading=1.25)
            if i < n - 1:
                self._text(s, x + cw, cy + h / 2 - 0.20, arrow_w, 0.40,
                           [[self.mono("→", 18, C["mint"], bold=True)]],
                           align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        return s

    # ---- 6: outcomes (3 big numerals + quote) ----------------------------
    def stats3(self, num, label, title_segs, stats, quote_text, attr):
        s = self._slide(); self._frame(s, "[05]")
        mx = THEME["mx"]
        self._kicker(s, mx, 0.95, num, label)
        self._h2(s, mx, 1.42, 11, title_segs, size=33)
        total = THEME["W"] - 2 * mx
        colw = total / 3
        sy = 2.45
        for i, (n, cap) in enumerate(stats):
            x = mx + i * colw
            # scale the numeral to its length so word-values (e.g. "weeks") fit one line
            sz = 84 if len(n) <= 3 else 64
            self._text(s, x, sy, colw - 0.25, 1.3,
                       [[self.sans(n, sz, C["mint"], bold=True, track=-2.0)]],
                       leading=0.86, wrap=False)
            self._text(s, x, sy + 1.30, colw - 0.4, 0.4,
                       [[self.mono(cap, 12, C["slate"], track=0.4)]])
        self.quote(s, quote_text, attr, y=4.85)
        return s

    # ---- quote ------------------------------------------------------------
    def quote(self, s, text, attr, y=4.85):
        mx = THEME["mx"]
        self._hline(s, mx, y, THEME["W"] - 2 * mx, C["line"], 1.0)
        self._text(s, mx, y + 0.10, 0.9, 1.0,
                   [[self.sans("“", 56, C["line"], bold=True)]], leading=0.7)
        self._text(s, mx + 0.95, y + 0.34, 9.3, 1.2,
                   [[self.sans(text, 23, C["paper"], bold=True, track=-0.2)]], leading=1.22)
        self._text(s, mx + 0.95, y + 1.42, 9.3, 0.4,
                   [[self.mono(attr, 12, C["slate"], track=0.4)]])

    # ---- 7: why mivada (hairline rows) -----------------------------------
    def reasons(self, num, label, title_segs, rows):
        s = self._slide(); self._frame(s, "[06]")
        mx = THEME["mx"]
        self._kicker(s, mx, 0.95, num, label)
        self._h2(s, mx, 1.42, 11, title_segs, size=33)
        total = THEME["W"] - 2 * mx
        top = 2.78
        rh = 1.30
        self._hline(s, mx, top, total, C["line"], 1.0)
        for i, (ix, claim, line) in enumerate(rows):
            y = top + i * rh
            self._text(s, mx, y + 0.32, 0.9, 0.4,
                       [[self.mono(ix, 13, C["mint"], track=0.6)]])
            self._text(s, mx + 1.05, y + 0.24, 9.8, 0.5,
                       [[self.sans(claim, 25, C["paper"], bold=True, track=-0.3)]])
            self._text(s, mx + 1.05, y + 0.74, 9.8, 0.5,
                       [[self.sans(line, 15, C["slate"])]], leading=1.35)
            self._hline(s, mx, y + rh, total, C["line"], 1.0)
        return s

    # ---- 8: contact -------------------------------------------------------
    def contact(self, toplabel, title_segs, line_main, line_arrow, meta, coord_top, coord_bot):
        s = self._slide(dotgrid=True)
        mx = THEME["mx"]
        self._text(s, mx, 0.46, 9, 0.3,
                   [[self.mono(toplabel, 12, C["slate"], track=0.7)]])
        self._text(s, mx, 2.30, 11, 2.2,
                   [[self.sans(t, 50, C["mint"] if mint else C["paper"], bold=True, track=-1.0)
                     for (t, mint) in title_segs]], leading=1.0)
        self._text(s, mx, 4.55, 10, 0.5,
                   [[self.mono(line_main, 15, C["mint"], bold=True, track=0.3),
                     self.mono(line_arrow, 15, C["paper"], track=0.3)]])
        self._text(s, mx, 5.15, 10, 0.4,
                   [[self.mono(meta, 11.5, C["slate"], track=0.5)]])
        self._signal_dot(s, mx, THEME["H"] - 1.06)
        self._text(s, THEME["W"] - mx - 4.0, THEME["H"] - 1.12, 4.0, 0.8,
                   [[self.mono(coord_top, 11, C["mint"], track=0.4)],
                    [self.mono(coord_bot, 11, C["slate"], track=0.4)]],
                   align=PP_ALIGN.RIGHT, leading=1.5, space=4)
        return s

    def save(self, path):
        self.prs.save(path)


# ----------------------------------------------------------------------------
# BUILD — content only; the look is inherited from the helpers above.
# ----------------------------------------------------------------------------
def build():
    d = Deck()

    # 1 — cover  (title_segs is a list of lines; each line a list of (text, is_mint))
    d.cover(
        "mivada.com  //  capability overview",
        [[("Technology,", False)],
         [("human", True), (" first.", False)]],
        "An Australian technology consultancy. We turn enterprise platforms into measurable human outcomes.",
        "33.87°S  151.21°E",
        "capability overview / 2026",
    )

    # 2 — who we are
    d.who("01", "who we are",
          [("From ", False), ("my village", True), (" and wisdom.", False)],
          "Mivada turns enterprise platforms—Workday, payroll, data and AI—into outcomes people can measure.",
          [("2014", "founded // formerly LJM Infotech"),
           ("120+", "specialists // certified, not generalist"),
           ("AU·IN", "onshore Australia + offshore India")])

    # 3 — what we do
    d.pillars("02", "what we do",
              [("Four practices, one operating model.", False)],
              [("[01]", "Workday & ERP", "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."),
               ("[02]", "Payroll consulting", "Compliant, accurate payroll across complex awards and multi-entity structures."),
               ("[03]", "Data & AI", "Lakehouse architecture, governed analytics and applied AI on your people and finance data."),
               ("[04]", "Intelligent automation", "RPA, ML and process design combined to remove low-value work.")])

    # 4 — how we work
    d.track("03", "how we work",
            [("Think human first.", False)],
            [("01", "Listen", "Understand the people and the work before the technology."),
             ("02", "Design", "Shape the platform around how teams actually operate."),
             ("03", "Deliver", "Implement in weeks-long increments, not multi-year programs."),
             ("04", "Care", "Stay on after go-live; measure outcomes, not tickets closed.")])

    # 5 — data & ai (showpiece)
    d.pipeline("04", "data & ai",
               [("Governed data, in real time.", False)],
               "Certified consultants integrate Databricks with Workday, payroll and finance—so leaders decide on current data.",
               [("[01]", "Sources", "Workday · payroll · finance"),
                ("[02]", "Delta Lake", "batch + streaming"),
                ("[03]", "Governed KPI store", "one trusted definition"),
                ("[04]", "Real-time insight", "decide on current data")])

    # 6 — outcomes
    d.stats3("05", "outcomes",
             [("Measured, not promised.", False)],
             [("40+", "Workday deployments"),
              ("98%", "client retention"),
              ("weeks", "to value, not months")],
             "They stayed after go-live and measured what changed—our team actually uses the platform.",
             "— Programme lead, enterprise client  //  illustrative")

    # 7 — why mivada
    d.reasons("06", "why mivada",
              [("Three reasons teams stay.", False)],
              [("[01]", "Certified specialists", "Workday- and Databricks-certified consultants, not generalists."),
               ("[02]", "Onshore + offshore", "Senior onshore leadership with cost-effective India delivery."),
               ("[03]", "People-first change", "Adoption built in from day one, so platforms actually get used.")])

    # 8 — contact
    d.contact(
        "mivada.com  //  let's talk",
        [("Let's build smarter systems, with people at the ", False), ("centre.", True)],
        "hello@mivada.com", "  →  let's talk",
        "Sydney & Melbourne  //  onshore AU & India  //  mivada.com",
        "// end", "figures illustrative",
    )

    out = os.path.join(HERE, "signal.pptx")
    d.save(out)
    print("wrote", out, "·", len(d.prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    build()

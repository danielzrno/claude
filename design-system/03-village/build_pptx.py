#!/usr/bin/env python3
"""
VILLAGE — Mivada design system 03 · PowerPoint builder
=======================================================
Human-first, warm, premium. Sage + amber + espresso on a named oat surface.

A `THEME` dict + a `Deck` class of helpers (cover, kpis, pillars, steps,
split, stats3, quote, reasons, contact). A new deck is a short list of helper
calls, not hand-placed shapes:

    from build_pptx import Deck
    d = Deck()
    d.cover("Mivada", "Technology, human first.", "Capability overview · 2026")
    d.kpis("Who we are", [("2014","Founded"), ("120+","Specialists"), ("AU + India","Delivery")])
    d.save("mivada.pptx")

Brand fonts: Gabarito (display) + Source Serif 4 (body). PowerPoint substitutes
the Office-safe fallbacks (Trebuchet MS / Georgia) if they are not installed.

NOTE ON SURFACE: oat (#FAF5EC) is an *intentional brand surface* — a named warm
paper, not the accidental beige default the anti-AI guide warns about. Espresso
ink on oat is tuned for strong contrast; sage and amber are used as accents only.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.oxml.ns import qn
import copy

# ============================================================
# THEME — palette (hex), fonts, sizes. Mirrors village.css.
# ============================================================
THEME = {
    # palette
    "ink":    "241F1B",   # warm espresso brown-black — text
    "oat":    "FAF5EC",   # soft warm oat — primary brand surface (named, intentional)
    "card":   "FFFDF8",   # raised card on oat
    "sage":   "6E7E54",   # sage/olive — PRIMARY accent (the 10%)
    "sage_d": "586647",   # deeper sage
    "amber":  "E0A33E",   # warm amber/gold — the underline motif colour
    "clay":   "B7593C",   # dusty clay — tertiary, sparing
    "ink60":  "6E665C",   # secondary text
    "line":   "E4DBC9",   # warm hairline
    "paper":  "FBF8F0",   # near-white warm for text ON dark/sage fields

    # type
    "display": "Gabarito",        # Office fallback: Trebuchet MS
    "body":    "Source Serif 4",  # Office fallback: Georgia

    # form
    "radius_in": 0.16,            # held card radius (~14px)
}

EMU_IN = 914400


def _c(hexstr):
    return RGBColor.from_string(hexstr)


class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.W = 13.333
        self.H = 7.5
        self.MX = 0.92          # outer margin
        self._blank = self.prs.slide_layouts[6]

    # ---------------------------------------------------------
    # low-level helpers
    # ---------------------------------------------------------
    def _slide(self, bg):
        s = self.prs.slides.add_slide(self._blank)
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = _c(bg)
        return s

    def _no_line(self, shp):
        shp.line.fill.background()

    def _kill_shadow(self, shp):
        """Force-remove any preset/theme shadow by writing an empty effectLst.
        (shadow.inherit=False alone is unreliable for autoshapes in LO.)"""
        spPr = shp._element.spPr
        for tag in ("a:effectLst", "a:effectDag"):
            for el in spPr.findall(qn(tag)):
                spPr.remove(el)
        spPr.append(spPr.makeelement(qn("a:effectLst"), {}))

    def _soft_shadow(self, shp, blur=0.16, dist=0.07, alpha=88000):
        """ONE tuned warm soft shadow: rgba(60,40,20,.08), low and soft."""
        spPr = shp._element.spPr
        # remove any inherited effect list
        for tag in ("a:effectLst", "a:effectDag"):
            for el in spPr.findall(qn(tag)):
                spPr.remove(el)
        eff = spPr.makeelement(qn("a:effectLst"), {})
        sh = eff.makeelement(qn("a:outerShdw"), {
            "blurRad": str(int(blur * EMU_IN)),
            "dist":    str(int(dist * EMU_IN)),
            "dir":     "5400000",   # straight down
            "rotWithShape": "0",
        })
        clr = sh.makeelement(qn("a:srgbClr"), {"val": "3C2814"})  # warm brown 60/40/20
        a = clr.makeelement(qn("a:alpha"), {"val": str(alpha)})
        clr.append(a)
        sh.append(clr)
        eff.append(sh)
        spPr.append(eff)

    def _rect(self, s, x, y, w, h, fill=None, line=None, line_w=1.0,
              radius=False, shadow=False):
        shape_t = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        shp = s.shapes.add_shape(shape_t, Inches(x), Inches(y), Inches(w), Inches(h))
        if radius:
            try:
                shp.adjustments[0] = THEME["radius_in"] / min(w, h)
            except Exception:
                pass
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid(); shp.fill.fore_color.rgb = _c(fill)
        if line is None:
            self._no_line(shp)
        else:
            shp.line.color.rgb = _c(line); shp.line.width = Pt(line_w)
        if shadow:
            self._soft_shadow(shp)
        else:
            self._kill_shadow(shp)
        return shp

    def _oval(self, s, x, y, d, fill=None, line=None, line_w=2.0):
        shp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid(); shp.fill.fore_color.rgb = _c(fill)
        if line is None:
            self._no_line(shp)
        else:
            shp.line.color.rgb = _c(line); shp.line.width = Pt(line_w)
        self._kill_shadow(shp)
        return shp

    def _text(self, s, x, y, w, h, runs, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, line_spacing=None, space_after=None,
              wrap=True):
        """runs: list of (text, font, size, colour, bold, italic, tracking).
        Multiple paragraphs via a tuple whose text is '\\n' -> new paragraph,
        or pass a list of run-lists for multi-paragraph blocks."""
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        # wrap="none" makes LibreOffice auto-size the box and centre the text on
        # its midpoint — which silently defeats left/right alignment. Only allow
        # no-wrap for genuinely centred labels (pills, tokens). Everything that
        # is left/right aligned keeps square wrap so alignment is honoured.
        tf.word_wrap = wrap if align == PP_ALIGN.CENTER else True
        tf.vertical_anchor = anchor
        tf.margin_left = 0; tf.margin_right = 0
        tf.margin_top = 0; tf.margin_bottom = 0
        # never auto-fit/auto-grow the frame
        bodyPr = tf._txBody.find(qn("a:bodyPr"))
        for tag in ("a:normAutofit", "a:spAutoFit"):
            e = bodyPr.find(qn(tag))
            if e is not None:
                bodyPr.remove(e)
        if bodyPr.find(qn("a:noAutofit")) is None:
            bodyPr.append(bodyPr.makeelement(qn("a:noAutofit"), {}))

        paragraphs = runs if (runs and isinstance(runs[0], list)) else [runs]
        for pi, para in enumerate(paragraphs):
            p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
            p.alignment = align
            if line_spacing:
                p.line_spacing = line_spacing
            if space_after is not None:
                p.space_after = Pt(space_after)
            p.space_before = Pt(0)
            for (txt, font, size, col, bold, ital, track) in para:
                r = p.add_run(); r.text = txt
                r.font.name = font
                r.font.size = Pt(size)
                r.font.bold = bold
                r.font.italic = ital
                r.font.color.rgb = _c(col)
                r.font.language_id = MSO_LANGUAGE_ID.ENGLISH_AUS
                # force the brand face for latin + EA + cs slots
                rPr = r._r.get_or_add_rPr()
                for tag in ("a:latin", "a:cs"):
                    e = rPr.find(qn(tag))
                    if e is None:
                        e = rPr.makeelement(qn(tag), {}); rPr.append(e)
                    e.set("typeface", font)
                if track:
                    rPr.set("spc", str(int(track * 100)))
        return tb

    # quick run-tuple builders
    def _d(self, txt, size, col=None, bold=True, ital=False, track=0):
        return (txt, THEME["display"], size, col or THEME["ink"], bold, ital, track)

    def _b(self, txt, size, col=None, bold=False, ital=False, track=0):
        return (txt, THEME["body"], size, col or THEME["ink"], bold, ital, track)

    def _pill(self, s, x, y, text, solid=False, w=None):
        """Sage outline pill (or solid). Returns its width in inches."""
        text = text.upper()
        pw = w if w else (0.118 * len(text) + 0.42)
        ph = 0.34
        fill = THEME["sage"] if solid else None
        line = None if solid else THEME["sage"]
        shp = self._rect(s, x, y, pw, ph, fill=fill, line=line, line_w=1.5, radius=True)
        shp.adjustments[0] = 0.5
        col = THEME["paper"] if solid else THEME["sage_d"]
        tb = self._text(s, x, y - 0.012, pw, ph,
                        [self._d(text, 10.5, col, bold=True, track=0.10)],
                        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        return pw

    def _dots(self, s, x, y, col=None):
        """quiet ••• micro-divider in amber."""
        self._text(s, x, y, 1.2, 0.3,
                   [self._d("• • •", 11, col or THEME["amber"], bold=True, track=0.04)],
                   anchor=MSO_ANCHOR.MIDDLE, wrap=False)

    def _token(self, s, x, y, label, d=0.92, ring=False, amber=False,
               avatar=False, fsize=20):
        """Circular sage token for step / stat / avatar numbers."""
        if ring:
            self._oval(s, x, y, d, fill=THEME["oat"], line=THEME["sage"], line_w=2.0)
            tcol = THEME["sage_d"]
        elif amber:
            self._oval(s, x, y, d, fill=THEME["amber"])
            tcol = THEME["ink"]
        elif avatar:
            self._oval(s, x, y, d, fill=THEME["sage_d"])
            tcol = THEME["paper"]
        else:
            self._oval(s, x, y, d, fill=THEME["sage"])
            tcol = THEME["paper"]
        self._text(s, x, y, d, d,
                   [self._d(label, fsize, tcol, bold=True, track=0)],
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)

    def _underline(self, s, x, y, w):
        """Signature hand-drawn underline: a thin, slightly-rotated rounded
        amber bar under one word. Used ONCE per deck (cover + contact share
        the same motif language; called once each on those bookends)."""
        bar = self._rect(s, x, y, w, 0.085, fill=THEME["amber"], radius=True)
        bar.adjustments[0] = 0.5
        bar.rotation = -1.4

    def _check(self, s, x, y, d=0.30):
        """Custom sage check token: pale sage disc + a tick built from two
        thin rounded bars (the two strokes of a check). Never an emoji."""
        disc = self._oval(s, x, y, d)
        disc.fill.solid(); disc.fill.fore_color.rgb = _c("E7EADD")  # pale sage tint
        self._no_line(disc)
        sw = d * 0.085                      # stroke thickness
        cx, cy = x + d * 0.50, y + d * 0.52
        # short stroke (down-right), ~45°
        a = self._rect(s, cx - d*0.22, cy - sw/2, d*0.26, sw,
                       fill=THEME["sage_d"], radius=True)
        a.adjustments[0] = 0.5; a.rotation = 45
        # long stroke (up-right), ~-50°
        b = self._rect(s, cx - d*0.04, cy - sw/2 - d*0.02, d*0.40, sw,
                       fill=THEME["sage_d"], radius=True)
        b.adjustments[0] = 0.5; b.rotation = -52

    def _foot(self, s, text="Mivada — Technology, human first.", dark=False):
        col = "FBF8F0" if dark else THEME["ink60"]
        if dark:
            col = "FBF8F0"
        self._text(s, self.MX, self.H - 0.52, 8, 0.3,
                   [self._d(text, 9, col if not dark else "FBF8F0",
                            bold=True, track=0.04)], wrap=False)

    def _slidenum(self, s, n, dark=False):
        col = "FBF8F0" if dark else THEME["ink60"]
        self._text(s, self.W - self.MX - 1.0, 0.6, 1.0, 0.3,
                   [self._d(f"{n:02d}", 9.5, col, bold=True, track=0.14)],
                   align=PP_ALIGN.RIGHT, wrap=False)

    def _eyebrow(self, s, pill_text, n, solid=False, extra=None):
        pw = self._pill(s, self.MX, 0.72, pill_text, solid=solid)
        self._dots(s, self.MX + pw + 0.28, 0.74)
        if extra:
            self._text(s, self.MX + pw + 0.95, 0.72, 5, 0.34,
                       [self._d(extra, 11.5, THEME["ink60"], bold=False)],
                       anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        self._slidenum(s, n)
        return pw

    # =========================================================
    # SLIDE HELPERS
    # =========================================================
    def cover(self, brand, tagline, caption):
        s = self._slide(THEME["ink"])
        # faint concentric espresso rings, bottom-left, centred on (-0.6, H+0.4)
        gcx, gcy = -0.6, self.H + 0.4
        for d in (8.4, 6.0, 3.8):
            self._oval(s, gcx - d/2, gcy - d/2, d, fill=None, line="3A332C", line_w=1.0)
        # right-side concentric rings + sage leaf token
        ring = self._oval(s, self.W - 4.0, self.H/2 - 2.5, 5.0,
                          fill=None, line="6E5A3A", line_w=1.25)
        ring2 = self._oval(s, self.W - 3.1, self.H/2 - 1.6, 3.2,
                           fill=None, line="3E372F", line_w=1.0)
        leaf = self._oval(s, self.W - 2.55, self.H/2 - 1.05, 2.1, fill=THEME["sage"])
        self._text(s, self.W - 2.55, self.H/2 - 1.05, 2.1, 2.1,
                   [self._d("m", 60, THEME["paper"], bold=True)],
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)

        # wordmark dot + name
        self._oval(s, self.MX, 1.35, 0.17, fill=THEME["amber"])
        self._text(s, self.MX + 0.3, 1.18, 5, 0.5,
                   [self._d(brand, 24, THEME["paper"], bold=True)], wrap=False)

        # topline
        self._text(s, self.MX, 1.95, 6, 0.3,
                   [self._d("CAPABILITY OVERVIEW · 2026", 11, THEME["amber"],
                            bold=True, track=0.16)], wrap=False)

        # headline with amber underline on "human"
        self._text(s, self.MX, 2.55, 8.4, 2.4,
                   [[self._d("Technology,", 52, THEME["paper"], bold=True, track=-0.01)],
                    [self._d("human", 52, THEME["paper"], bold=True, track=-0.01),
                     self._d(" first.", 52, THEME["paper"], bold=True, track=-0.01)]],
                   line_spacing=0.98)
        # underline sits under "human" on the second line
        self._underline(s, self.MX + 0.02, 4.62, 1.95)

        # caption (serif italic)
        self._text(s, self.MX, 5.15, 6.6, 1.2,
                   [[self._b("From ", 16, "C9C1B4", ital=True),
                     self._b("my village", 16, "C9C1B4", ital=True),
                     self._b(" — community — and wisdom. We turn enterprise platforms into measurable human outcomes.",
                             16, "C9C1B4", ital=True)]],
                   line_spacing=1.3)
        return s

    def kpis(self, title, stats):
        """3 proof stats as sage ring tokens + serif labels."""
        s = self._slide(THEME["oat"])
        self._eyebrow(s, title, 2)
        self._text(s, self.MX, 1.45, 9.6, 1.5,
                   [self._d("An Australian technology consultancy, built around people.",
                            34, THEME["ink"], bold=True, track=-0.015)],
                   line_spacing=1.02)
        self._text(s, self.MX, 2.95, 9.0, 1.2,
                   [self._b("We turn enterprise platforms — Workday, payroll, data and AI — into measurable human outcomes, not just go-lives. Plain advice, senior hands, work delivered in weeks.",
                            16.5, THEME["ink"])],
                   line_spacing=1.32)

        # three ring-token stats along the lower third
        labels_top = [st[0] for st in stats]
        sub = [st[1] for st in stats]
        ring_tokens = ["14", "120", "AU"]
        y = 5.2
        colw = (self.W - 2*self.MX) / 3
        for i, st in enumerate(stats):
            x = self.MX + i*colw
            self._token(s, x, y, ring_tokens[i] if i < 2 else "AU",
                        d=0.92, ring=True, fsize=18 if i < 2 else 15)
            self._text(s, x + 1.12, y - 0.02, colw - 1.2, 0.6,
                       [self._d(st[0], 19, THEME["ink"], bold=True, track=-0.02)],
                       anchor=MSO_ANCHOR.MIDDLE, wrap=False)
            self._text(s, x + 1.12, y + 0.5, colw - 1.2, 0.5,
                       [self._b(st[1], 12.5, THEME["ink60"])], wrap=True)
        self._foot(s)
        return s

    def pillars(self, title, items):
        """4 warm cards, 2×2. items: list of (idx, name, line)."""
        s = self._slide(THEME["oat"])
        self._eyebrow(s, title, 3)
        self._text(s, self.MX, 1.42, 10, 0.9,
                   [self._d("Four practices, one operating model.", 30,
                            THEME["ink"], bold=True, track=-0.018)], wrap=False)

        gx, gy = self.MX, 2.5
        gap = 0.4
        cw = (self.W - 2*self.MX - gap) / 2
        ch = 1.9
        for i, (idx, name, line) in enumerate(items):
            r, col = divmod(i, 2)
            x = gx + col*(cw + gap)
            y = gy + r*(ch + gap)
            self._rect(s, x, y, cw, ch, fill=THEME["card"], line=THEME["line"],
                       line_w=1.0, radius=True, shadow=True)
            self._pill(s, x + 0.32, y + 0.32, idx)
            self._text(s, x + 0.32, y + 0.74, cw - 0.6, 0.55,
                       [self._d(name, 19, THEME["ink"], bold=True, track=-0.015)],
                       wrap=False)
            self._text(s, x + 0.32, y + 1.22, cw - 0.62, 0.62,
                       [self._b(line, 12.5, THEME["ink60"])], line_spacing=1.25)
        self._foot(s)
        return s

    def steps(self, title, sub, steps):
        """4 sage numbered tokens + serif captions + dotted connector."""
        s = self._slide(THEME["oat"])
        self._eyebrow(s, title, 4, solid=True, extra=sub)
        self._text(s, self.MX, 1.45, 10.4, 0.9,
                   [self._d("Four steps, run in weeks — not multi-year programs.",
                            30, THEME["ink"], bold=True, track=-0.018)], wrap=False)

        n = len(steps)
        y = 3.5
        d = 0.92
        colw = (self.W - 2*self.MX) / n
        # dotted connector behind the tokens
        cy = y + d/2
        for i in range(n - 1):
            x0 = self.MX + i*colw + d + 0.18
            x1 = self.MX + (i+1)*colw - 0.05
            self._dotted_line(s, x0, cy, x1)
        for i, (num, name, cap) in enumerate(steps):
            x = self.MX + i*colw
            self._token(s, x, y, num, d=d, fsize=20)
            self._text(s, x, y + 1.12, colw - 0.4, 0.5,
                       [self._d(name, 19, THEME["ink"], bold=True, track=-0.015)],
                       wrap=False)
            self._text(s, x, y + 1.62, colw - 0.45, 1.0,
                       [self._b(cap, 12.5, THEME["ink60"])], line_spacing=1.28)
        self._foot(s)
        return s

    def _dotted_line(self, s, x0, y, x1):
        """A dotted connector built from small dots (reliable across renderers)."""
        gap = 0.16
        x = x0
        while x < x1:
            dot = self._oval(s, x, y - 0.018, 0.036, fill=THEME["line"])
            x += gap

    def split(self, title, headline, standfirst, cap_title, caps):
        """Two-column: serif standfirst left, soft card w/ checks right."""
        s = self._slide(THEME["oat"])
        self._slidenum(s, 5)
        # left
        lx = self.MX
        self._pill(s, lx, 1.5, title)
        self._text(s, lx, 2.1, 5.5, 1.7,
                   [self._d(headline, 30, THEME["ink"], bold=True, track=-0.02)],
                   line_spacing=1.02)
        self._text(s, lx, 4.0, 5.3, 1.7,
                   [self._b(standfirst, 16, THEME["ink"])], line_spacing=1.34)
        self._dots(s, lx, 5.7)

        # right card
        cx = 7.15
        cw = self.W - self.MX - cx
        cy = 1.35
        chh = 4.8
        self._rect(s, cx, cy, cw, chh, fill=THEME["card"], line=THEME["line"],
                   line_w=1.0, radius=True, shadow=True)
        self._text(s, cx + 0.5, cy + 0.42, cw - 1.0, 0.4,
                   [self._d(cap_title, 16, THEME["ink"], bold=True, track=-0.01)],
                   wrap=False)
        iy = cy + 1.05
        for (head, sub) in caps:
            self._check(s, cx + 0.5, iy + 0.03, d=0.30)
            self._text(s, cx + 0.98, iy - 0.04, cw - 1.5, 0.36,
                       [self._b(head, 14.5, THEME["ink"], bold=True)], wrap=False)
            self._text(s, cx + 0.98, iy + 0.32, cw - 1.5, 0.36,
                       [self._b(sub, 11.5, THEME["ink60"])], wrap=True)
            iy += 0.92
        self._foot(s)
        return s

    def stats3(self, title, headline, stats, quote, attrib, attrib_sub, initials):
        """3 stat callouts + client quote with avatar token."""
        s = self._slide(THEME["oat"])
        self._eyebrow(s, title, 6)
        self._text(s, self.MX, 1.42, 11, 0.8,
                   [self._d(headline, 26, THEME["ink"], bold=True, track=-0.018)],
                   wrap=False)

        # 3 stat callouts
        colw = (self.W - 2*self.MX) / 3
        y = 2.45
        for i, (val, unit, lab) in enumerate(stats):
            x = self.MX + i*colw
            if i > 0:
                # thin warm divider rule
                rule = self._rect(s, x - 0.2, y + 0.1, 0.012, 1.15, fill=THEME["line"])
            runs = [self._d(val, 50, THEME["ink"], bold=True, track=-0.03)]
            if unit:
                runs.append(self._d(unit, 50, THEME["amber"], bold=True, track=-0.03))
            self._text(s, x, y, colw - 0.3, 0.9, [runs], wrap=False)
            self._text(s, x, y + 1.0, colw - 0.5, 0.6,
                       [self._b(lab, 12.5, THEME["ink60"])], line_spacing=1.25)

        # divider above quote
        self._rect(s, self.MX, 4.5, self.W - 2*self.MX, 0.012, fill=THEME["line"])
        # quote with avatar token
        self._token(s, self.MX, 4.85, initials, d=1.15, avatar=True, fsize=22)
        self._text(s, self.MX + 1.5, 4.78, 9.4, 1.1,
                   [[self._b("“", 24, THEME["ink"], ital=True),
                     self._b(quote, 19, THEME["ink"], ital=True, bold=True),
                     self._b("”", 24, THEME["ink"], ital=True)]],
                   line_spacing=1.22)
        self._text(s, self.MX + 1.5, 6.25, 9, 0.5,
                   [[self._d(attrib, 12.5, THEME["ink"], bold=True),
                     self._d("   " + attrib_sub, 12.5, THEME["ink60"], bold=False)]],
                   wrap=False)
        self._foot(s, text="Figures illustrative.")
        return s

    def reasons(self, title, headline, items):
        """3 reasons: sage numeral token + claim + line."""
        s = self._slide(THEME["oat"])
        self._eyebrow(s, title, 7)
        self._text(s, self.MX, 1.45, 9, 0.9,
                   [self._d(headline, 30, THEME["ink"], bold=True, track=-0.018)],
                   wrap=False)

        colw = (self.W - 2*self.MX) / 3
        y = 3.0
        for i, (num, claim, line) in enumerate(items):
            x = self.MX + i*colw
            self._token(s, x, y, num, d=0.92, fsize=20)
            self._text(s, x, y + 1.15, colw - 0.45, 0.55,
                       [self._d(claim, 19, THEME["ink"], bold=True, track=-0.015)],
                       wrap=False)
            self._text(s, x, y + 1.68, colw - 0.5, 1.1,
                       [self._b(line, 13, THEME["ink60"])], line_spacing=1.3)
        self._foot(s)
        return s

    def contact(self, headline_a, mark_word, headline_b, lines):
        """Sage bookend; CTA with amber underline on one word; serif contact."""
        s = self._slide(THEME["sage"])
        # warm espresso glow top-right
        for d in (8.0, 5.5):
            self._oval(s, self.W - d + 2.4, -d + 2.2, d, fill=None,
                       line="5C6A45", line_w=1.0)
        # paper-outline pill for contrast on the sage field
        pw = 0.118 * len("LET'S TALK") + 0.42
        pill = self._rect(s, self.MX, 1.5, pw, 0.34, fill=None,
                          line="FBF8F0", line_w=1.5, radius=True)
        pill.adjustments[0] = 0.5
        self._text(s, self.MX, 1.488, pw, 0.34,
                   [self._d("LET'S TALK", 10.5, THEME["paper"], bold=True, track=0.10)],
                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, wrap=False)
        # headline: runs across two lines, underline under mark_word
        self._text(s, self.MX, 2.25, 9.7, 2.4,
                   [[self._d(headline_a, 38, THEME["paper"], bold=True, track=-0.02)],
                    [self._d(mark_word, 38, THEME["paper"], bold=True, track=-0.02),
                     self._d(headline_b, 38, THEME["paper"], bold=True, track=-0.02)]],
                   line_spacing=1.02)
        self._underline(s, self.MX + 0.02, 3.62, 1.85)

        # contact line in serif
        self._text(s, self.MX, 4.85, 11, 1.2,
                   [[self._d("mivada.com", 17, THEME["paper"], bold=True),
                     self._b("    ·    ", 17, "E0A33E", bold=False),
                     self._d("hello@mivada.com", 17, THEME["paper"], bold=True)],
                    [self._b("Sydney & Melbourne", 15, "EAE6DB", ital=False),
                     self._b("    ·    ", 15, "E0A33E"),
                     self._b("Onshore AU & India", 15, "EAE6DB")]],
                   line_spacing=1.5)
        self._foot(s, text="Technology, human first.", dark=True)
        return s

    # ---------------------------------------------------------
    def save(self, path):
        self.prs.save(path)


# ============================================================
# BUILD — the 8-slide Mivada capability deck
# ============================================================
def build():
    d = Deck()

    d.cover("Mivada", "Technology, human first.", "Capability overview · 2026")

    d.kpis("Who we are", [
        ("Since 2014", "Founded as LJM Infotech"),
        ("120+ specialists", "Certified, not generalists"),
        ("AU + India", "Onshore led, offshore scaled"),
    ])

    d.pillars("What we do", [
        ("Practice 01", "Workday & ERP",
         "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."),
        ("Practice 02", "Payroll consulting",
         "Compliant, accurate payroll across complex awards and multi-entity structures."),
        ("Practice 03", "Data & AI",
         "Lakehouse architecture, governed analytics and applied AI on your people and finance data."),
        ("Practice 04", "Intelligent automation",
         "RPA, ML and process design combined to remove low-value, repetitive work."),
    ])

    d.steps("How we work", "Think human first", [
        ("1", "Listen", "Understand the people and the work before the technology."),
        ("2", "Design", "Shape the platform around how teams actually operate."),
        ("3", "Deliver", "Implement in weeks-long increments, measured as we go."),
        ("4", "Care", "Stay on after go-live; measure outcomes, not tickets closed."),
    ])

    d.split("Data & AI",
            "Decisions on current data, not last quarter's.",
            "Certified consultants design Lakehouse architectures and integrate Databricks with Workday, payroll and finance — so leaders see one trusted set of numbers.",
            "Selected capabilities",
            [("Lakehouse architecture", "Designed for governance from day one"),
             ("Delta Lake pipelines", "Batch and streaming, built to scale"),
             ("Governed KPI store", "One trusted definition per metric"),
             ("Real-time insight", "Current data in the hands of leaders")])

    d.stats3("Outcomes",
             "The work shows up as measurable results.",
             [("40", "+", "Workday deployments delivered"),
              ("98", "%", "Client retention, year on year"),
              ("Weeks", "", "To value — not multi-year programs")],
             "They listened first, then built around how our teams actually work. Adoption looked after itself.",
             "Executive sponsor", "Enterprise client · illustrative", "EW")

    d.reasons("Why Mivada",
              "Three reasons clients stay with us.",
              [("1", "Certified specialists",
                "Workday- and Databricks-certified consultants, not generalists learning on your time."),
               ("2", "Onshore + offshore",
                "Senior onshore leadership with cost-effective India delivery, on one team."),
               ("3", "People-first change",
                "Adoption built in from day one, so the platforms you buy actually get used.")])

    d.contact("Let's build smarter systems, with",
              "people", " at the centre.",
              None)

    d.save("village.pptx")
    print("Saved village.pptx —", len(d.prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    build()

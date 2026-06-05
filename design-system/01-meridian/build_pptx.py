#!/usr/bin/env python3
# =============================================================================
# MERIDIAN — Mivada editorial-consulting PPTX builder
# A THEME dict + a Deck class of helper methods, so a whole deck is a short
# list of calls. 16:9, white content slides, ink-navy bookends. Fraunces titles
# + Archivo body. Stats as large Fraunces numerals. Depth from hairline rules
# (thin rectangles placed in MARGINS — never under a title) and bone panels.
# No drop shadows anywhere.
#
#   python build_pptx.py        ->  meridian.pptx
#
# A new deck = reuse THEME + Deck, write a few content calls (see build() at end).
# =============================================================================

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.oxml.ns import qn

# -----------------------------------------------------------------------------
# THEME — the single source of truth (mirrors meridian.css :root tokens)
# -----------------------------------------------------------------------------
THEME = {
    # palette (hex without '#') — 60/30/10: paper+ink dominate, bone supports,
    # teal is the accent, brass is rare.
    "ink":     "14283A",   # deep slate-navy — primary text; bookend background
    "paper":   "FBFAF7",   # barely-warm white — primary background
    "bone":    "ECE7DB",   # warm stone — secondary panel / fill
    "teal":    "18786A",   # deep teal — PRIMARY accent (the sharp 10%)
    "teal_lift": "5BB5A4", # teal lifted for dark grounds
    "brass":   "A87C3D",   # muted brass — secondary accent, very sparing
    "stone":   "C8C1B2",   # hairline / rule colour
    "stone_dim": "3A4B5C", # hairline on dark grounds
    "ink_60":  "5A6675",   # secondary text
    "paper_dim": "8794A2", # muted caption on dark grounds
    "paper_dim2": "A7B2BD",

    # type families (PowerPoint substitutes the Office fallback if missing)
    "display": "Fraunces",                 # fallback: Georgia
    "text":    "Archivo",                  # fallback: Calibri

    # geometry
    "EMU_W": Inches(13.333),
    "EMU_H": Inches(7.5),
    "margin": Inches(0.92),                # wide editorial outer margin
    "tracking_caps": 1.6,                  # Pt of letter-spacing for tracked caps
}

# Convenience -----------------------------------------------------------------
def _rgb(key):
    return RGBColor.from_string(THEME[key])

EMU_PER_INCH = 914400


# =============================================================================
# Deck — helper methods build whole slides from content arguments.
# =============================================================================
class Deck:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = THEME["EMU_W"]
        self.prs.slide_height = THEME["EMU_H"]
        self._blank = self.prs.slide_layouts[6]   # truly blank layout
        self.W = self.prs.slide_width
        self.H = self.prs.slide_height
        self.M = THEME["margin"]

    # ---- low-level primitives ------------------------------------------------
    def _slide(self, bg="paper"):
        s = self.prs.slides.add_slide(self._blank)
        # paint background
        fill = s.background.fill
        fill.solid()
        fill.fore_color.rgb = _rgb(bg)
        return s

    def _rect(self, s, x, y, w, h, color):
        """A flat filled rectangle (used for hairlines + panels). No shadow ever."""
        shp = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(color)
        shp.line.fill.background()
        shp.shadow.inherit = False              # kill the default pptx shadow
        return shp

    def _hairline(self, s, x, y, length, color="stone", vertical=False, weight_emu=None):
        """A 1px-equivalent rule. Placed in margins / between columns — never under a title."""
        thick = weight_emu if weight_emu else Emu(9525)   # ~1px at 96dpi
        if vertical:
            self._rect(s, x, y, thick, length, color)
        else:
            self._rect(s, x, y, length, thick, color)

    def _text(self, s, x, y, w, h, runs, *, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, line_spacing=1.0, space_after=0,
              wrap=True):
        """
        Add a textbox. `runs` is a list of paragraphs; each paragraph is a list of
        (text, font, size_pt, color_key, bold, italic, tracking) run-tuples.
        Missing tuple fields default sensibly.
        """
        tb = s.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = anchor
        # zero the internal margins so text aligns to our grid exactly
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        for pi, para in enumerate(runs):
            p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
            p.alignment = align
            if line_spacing:
                p.line_spacing = line_spacing
            if space_after:
                p.space_after = Pt(space_after)
            p.space_before = Pt(0)
            for run_tuple in para:
                text  = run_tuple[0]
                font  = run_tuple[1] if len(run_tuple) > 1 and run_tuple[1] else THEME["text"]
                size  = run_tuple[2] if len(run_tuple) > 2 and run_tuple[2] else 18
                color = run_tuple[3] if len(run_tuple) > 3 and run_tuple[3] else "ink"
                bold  = run_tuple[4] if len(run_tuple) > 4 else False
                ital  = run_tuple[5] if len(run_tuple) > 5 else False
                track = run_tuple[6] if len(run_tuple) > 6 else None
                r = p.add_run()
                r.text = text
                r.font.name = font
                r.font.size = Pt(size)
                r.font.bold = bold
                r.font.italic = ital
                r.font.color.rgb = _rgb(color)
                r.font.language_id = MSO_LANGUAGE_ID.ENGLISH_AUS
                if track is not None:
                    self._set_tracking(r, track)
        return tb

    @staticmethod
    def _set_tracking(run, pt):
        """Letter-spacing in points -> OOXML `spc` (1/100 pt) on the run."""
        rPr = run._r.get_or_add_rPr()
        rPr.set("spc", str(int(pt * 100)))

    def _kicker(self, s, x, y, number, label, num_color="ink", label_color="ink_60",
                num_size=30, w=Inches(7)):
        """The signature numbered-section kicker: big light Fraunces numeral + tracked caps label.
        Returns the bottom y so callers can stack content beneath it (with air — no rule under)."""
        tb = s.shapes.add_textbox(x, y, w, Inches(0.7))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.line_spacing = 1.0
        rn = p.add_run(); rn.text = number
        rn.font.name = THEME["display"]; rn.font.size = Pt(num_size)
        rn.font.bold = False; rn.font.color.rgb = _rgb(num_color)
        rl = p.add_run(); rl.text = "   " + label.upper()
        rl.font.name = THEME["text"]; rl.font.size = Pt(11)
        rl.font.bold = True; rl.font.color.rgb = _rgb(label_color)
        self._set_tracking(rl, THEME["tracking_caps"])
        return y + Inches(0.62)

    def _slide_no(self, s, n, dark=False):
        self._text(
            s, self.W - self.M - Inches(1.6), self.H - Inches(0.62),
            Inches(1.6), Inches(0.3),
            [[("%02d — 08" % n, THEME["text"], 10,
               "paper_dim" if dark else "ink_60", False, False, THEME["tracking_caps"])]],
            align=PP_ALIGN.RIGHT)

    # =========================================================================
    # SLIDE HELPERS  — one per layout in the outline
    # =========================================================================

    def cover(self, wordmark, title_lines, caption, n=1):
        """Ink-navy cover. Small tracked wordmark top-left; huge light Fraunces hero
        lower-left (one teal italic word); caption; thin BRASS vertical rule in the
        right margin (never under the title)."""
        s = self._slide("ink")
        # wordmark top-left
        self._text(s, self.M, self.M, Inches(6), Inches(0.4),
                   [[(wordmark.upper(), THEME["text"], 12, "paper", True, False, 4.2)]])
        # hero lower-left — title_lines is a list of (text, italic_bool) segments per line
        tb = s.shapes.add_textbox(self.M, Inches(3.55), Inches(9.2), Inches(2.9))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
        for li, line in enumerate(title_lines):
            p = tf.paragraphs[0] if li == 0 else tf.add_paragraph()
            p.line_spacing = 0.98; p.space_after = Pt(0); p.space_before = Pt(0)
            for seg_text, is_em in line:
                r = p.add_run(); r.text = seg_text
                r.font.name = THEME["display"]; r.font.size = Pt(74)
                r.font.bold = False
                r.font.italic = is_em
                r.font.color.rgb = _rgb("teal_lift" if is_em else "paper")
        # caption
        self._text(s, self.M, Inches(6.5), Inches(8), Inches(0.4),
                   [[(caption, THEME["text"], 13, "paper_dim", False, False, 0.6)]])
        # brass vertical rule — right margin element
        rx = self.W - self.M
        self._hairline(s, rx, Inches(1.0), Inches(5.5), color="brass", vertical=True)
        self._slide_no(s, n, dark=True)
        return s

    def kpis(self, number, label, headline_segments, standfirst, stats, n=2):
        """'Who we are' split: left = kicker + headline (one teal italic word) +
        light Fraunces standfirst; right = stacked stats separated by hairlines.
        `headline_segments` = list of (text, italic_bool). `stats` = [(figure, caption), ...]."""
        s = self._slide("paper")
        left_x = self.M
        col_w = Inches(6.6)
        y = self._kicker(s, left_x, self.M, number, label)
        # headline
        tb = s.shapes.add_textbox(left_x, y + Inches(0.25), col_w, Inches(2.1))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        p = tf.paragraphs[0]; p.line_spacing = 1.04
        for seg_text, is_em in headline_segments:
            r = p.add_run(); r.text = seg_text
            r.font.name = THEME["display"]; r.font.size = Pt(34)
            r.font.bold = (not is_em)
            r.font.italic = is_em
            r.font.color.rgb = _rgb("teal" if is_em else "ink")
        # standfirst
        self._text(s, left_x, y + Inches(2.35), col_w, Inches(2.0),
                   [[(standfirst, THEME["display"], 19, "ink", False, False)]],
                   line_spacing=1.3)
        # right column: stacked stats with hairlines between
        rx = Inches(8.55)
        rw = self.W - self.M - rx
        # vertically center the stack
        block_h = Inches(1.55) * len(stats)
        sy = (self.H - block_h) / 2
        for i, (fig, cap) in enumerate(stats):
            if i > 0:
                self._hairline(s, rx, sy - Inches(0.18), rw, color="stone")
            self._stat(s, rx, sy, rw, fig, cap, fig_size=40)
            sy += Inches(1.55)
        self._slide_no(s, n)
        return s

    def _stat(self, s, x, y, w, figure, caption, fig_size=40, fig_color="ink",
              em_figure=False):
        """A stat unit: big Fraunces numeral + small tracked Archivo caps caption."""
        tb = s.shapes.add_textbox(x, y, w, Inches(1.4))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        p = tf.paragraphs[0]; p.line_spacing = 0.95
        r = p.add_run(); r.text = figure
        r.font.name = THEME["display"]; r.font.size = Pt(fig_size)
        r.font.bold = False; r.font.italic = em_figure
        r.font.color.rgb = _rgb("teal" if em_figure else fig_color)
        # caption paragraph
        pc = tf.add_paragraph(); pc.space_before = Pt(6); pc.line_spacing = 1.1
        rc = pc.add_run(); rc.text = caption.upper()
        rc.font.name = THEME["text"]; rc.font.size = Pt(10.5)
        rc.font.bold = True; rc.font.color.rgb = _rgb("ink_60")
        self._set_tracking(rc, THEME["tracking_caps"])
        return tb

    def pillars(self, number, label, items, n=3):
        """'What we do' 2x2 separated by hairline rules (NO boxes, NO shadows).
        Each item = (index, name, line). Brass index numeral + Fraunces name + Archivo line."""
        s = self._slide("paper")
        self._kicker(s, self.M, self.M, number, label)
        # grid geometry
        gx = self.M
        gy = Inches(2.35)
        gw = self.W - 2 * self.M
        gh = Inches(4.0)
        col_w = gw / 2
        row_h = gh / 2
        gap = Inches(0.5)
        # interior hairlines: one vertical between columns, one horizontal between rows,
        # both INSIDE the content block (these are column/row rules, not title underlines)
        self._hairline(s, gx + col_w, gy, gh, color="stone", vertical=True)
        self._hairline(s, gx, gy + row_h, gw, color="stone")
        for i, (ix, name, line) in enumerate(items):
            r = i // 2; c = i % 2
            cell_x = gx + c * col_w + (gap if c == 1 else Emu(0))
            cell_y = gy + r * row_h + (gap if r == 1 else Inches(0.1))
            cell_w = col_w - gap
            # brass index
            self._text(s, cell_x, cell_y, cell_w, Inches(0.4),
                       [[(ix, THEME["display"], 16, "brass", False, False)]])
            # name
            self._text(s, cell_x, cell_y + Inches(0.4), cell_w, Inches(0.6),
                       [[(name, THEME["display"], 21, "ink", True, False)]],
                       line_spacing=1.05)
            # line
            self._text(s, cell_x, cell_y + Inches(1.02), cell_w, Inches(1.0),
                       [[(line, THEME["text"], 12.5, "ink_60", False, False)]],
                       line_spacing=1.35)
        self._slide_no(s, n)
        return s

    def steps(self, number, label, items, n=4):
        """'How we work' horizontal Listen->Design->Deliver->Care: big Fraunces step
        numerals hung from a connecting hairline; short captions. items=(num,name,cap).
        The rule sits ABOVE the numerals (a clean connector), matching the HTML deck."""
        s = self._slide("paper")
        self._kicker(s, self.M, self.M, number, label)
        gx = self.M
        gw = self.W - 2 * self.M
        n_steps = len(items)
        col_w = gw / n_steps
        # connecting hairline above the numerals, spanning the full step band
        rule_y = Inches(2.7)
        self._hairline(s, gx, rule_y, gw, color="stone")
        for i, (num, name, cap) in enumerate(items):
            cx = gx + i * col_w
            # numeral hangs just beneath the rule
            self._text(s, cx, Inches(2.95), Inches(1.2), Inches(1.0),
                       [[(num, THEME["display"], 50, "ink", False, False)]],
                       line_spacing=1.0)
            # name
            self._text(s, cx, Inches(4.0), col_w - Inches(0.4), Inches(0.5),
                       [[(name, THEME["display"], 20, "ink", True, False)]])
            # caption
            self._text(s, cx, Inches(4.5), col_w - Inches(0.5), Inches(1.4),
                       [[(cap, THEME["text"], 12, "ink_60", False, False)]],
                       line_spacing=1.35)
        self._slide_no(s, n)
        return s

    def split(self, number, label, standfirst_segments, list_items, n=5):
        """'Data & AI' two columns: left = light Fraunces standfirst (one teal italic
        word); right = tracked-caps labelled list. standfirst_segments=[(text,em)],
        list_items=[(term, desc), ...]."""
        s = self._slide("paper")
        left_x = self.M
        col_w = Inches(5.6)
        y = self._kicker(s, left_x, self.M, number, label)
        # standfirst, vertically settled in the left column
        tb = s.shapes.add_textbox(left_x, Inches(2.7), col_w, Inches(3.0))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        p = tf.paragraphs[0]; p.line_spacing = 1.32
        for seg_text, is_em in standfirst_segments:
            r = p.add_run(); r.text = seg_text
            r.font.name = THEME["display"]; r.font.size = Pt(23)
            r.font.bold = False; r.font.italic = is_em
            r.font.color.rgb = _rgb("teal" if is_em else "ink")
        # right column labelled list, hairline between entries
        rx = Inches(7.7)
        rw = self.W - self.M - rx
        sy = Inches(2.0)
        for i, (term, desc) in enumerate(list_items):
            if i > 0:
                self._hairline(s, rx, sy - Inches(0.22), rw, color="stone")
            self._text(s, rx, sy, rw, Inches(0.35),
                       [[(term.upper(), THEME["text"], 11, "teal", True, False, THEME["tracking_caps"])]])
            self._text(s, rx, sy + Inches(0.42), rw, Inches(0.8),
                       [[(desc, THEME["text"], 12.5, "ink_60", False, False)]],
                       line_spacing=1.35)
            sy += Inches(1.4)
        self._slide_no(s, n)
        return s

    def stats3(self, number, label, stats, quote, attrib, n=6):
        """'Outcomes' — 3 big Fraunces stat callouts across the top, then an italic
        Fraunces pull-quote with small-caps attribution. stats=[(figure,caption,em)]."""
        s = self._slide("paper")
        self._kicker(s, self.M, self.M, number, label)
        gx = self.M
        gw = self.W - 2 * self.M
        col_w = gw / 3
        for i, st in enumerate(stats):
            fig, cap = st[0], st[1]
            em = st[2] if len(st) > 2 else False
            self._stat(s, gx + i * col_w, Inches(2.2), col_w - Inches(0.4),
                       fig, cap, fig_size=60, em_figure=em)
        # pull-quote, lower band
        self._text(s, gx, Inches(4.55), Inches(9.6), Inches(1.4),
                   [[("“" + quote + "”", THEME["display"], 23, "ink", False, True)]],
                   line_spacing=1.32)
        self._text(s, gx, Inches(6.0), Inches(9.6), Inches(0.4),
                   [[(attrib.upper(), THEME["text"], 10.5, "ink_60", True, False, THEME["tracking_caps"])]])
        self._slide_no(s, n)
        return s

    def reasons(self, number, label, items, n=7):
        """'Why Mivada' — numbered reasons in a column, teal Fraunces numerals,
        generous leading. items=(num, title, body)."""
        s = self._slide("paper")
        self._kicker(s, self.M, self.M, number, label)
        gx = self.M
        # vertically center the three reasons in the body area
        n_items = len(items)
        row_h = Inches(1.4)
        block_h = row_h * n_items
        sy = Inches(2.4)
        num_w = Inches(1.1)
        body_x = gx + num_w + Inches(0.3)
        body_w = self.W - self.M - body_x
        for (num, title, body) in items:
            # teal numeral
            self._text(s, gx, sy, num_w, Inches(1.0),
                       [[(num, THEME["display"], 46, "teal", False, False)]],
                       line_spacing=1.0)
            # title + body
            self._text(s, body_x, sy + Inches(0.05), body_w, Inches(0.5),
                       [[(title, THEME["display"], 21, "ink", True, False)]])
            self._text(s, body_x, sy + Inches(0.58), body_w, Inches(0.8),
                       [[(body, THEME["text"], 13.5, "ink_60", False, False)]],
                       line_spacing=1.4)
            sy += row_h
        self._slide_no(s, n)
        return s

    def contact(self, wordmark, cta_segments, contact_line, n=8):
        """Ink-navy closing bookend. Large light Fraunces CTA (one teal italic word);
        contact line in Archivo small caps; a brass hairline above it."""
        s = self._slide("ink")
        self._text(s, self.M, self.M, Inches(6), Inches(0.4),
                   [[(wordmark.upper(), THEME["text"], 12, "paper", True, False, 4.2)]])
        # CTA, centred-left block, vertically settled
        tb = s.shapes.add_textbox(self.M, Inches(2.5), Inches(9.4), Inches(2.6))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.line_spacing = 1.04
        for seg_text, is_em in cta_segments:
            r = p.add_run(); r.text = seg_text
            r.font.name = THEME["display"]; r.font.size = Pt(40)
            r.font.bold = False; r.font.italic = is_em
            r.font.color.rgb = _rgb("teal_lift" if is_em else "paper")
        # brass hairline above contact
        self._hairline(s, self.M, Inches(5.85), Inches(3.2), color="brass")
        # contact line
        self._text(s, self.M, Inches(6.1), Inches(11), Inches(0.4),
                   [[(contact_line, THEME["text"], 12, "paper_dim2", False, False, 1.0)]])
        self._slide_no(s, n, dark=True)
        return s

    # ---- output --------------------------------------------------------------
    def save(self, path):
        self.prs.save(path)
        return path


# =============================================================================
# BUILD — the same 8-slide deck as the HTML/outline, as a short list of calls.
# =============================================================================
def build():
    d = Deck()

    # 1 · Cover
    d.cover(
        "Mivada",
        [[("Technology,", False)],
         [("human ", False), ("first.", True)]],
        "Capability overview · 2026",
        n=1)

    # 2 · Who we are
    d.kpis(
        "01", "Who we are",
        [("We turn enterprise platforms into measurable ", False),
         ("human", True), (" outcomes.", False)],
        "Mivada is an Australian technology consultancy — Workday, payroll, "
        "data and AI — built around the people who use the systems we deliver.",
        [("2014", "Founded · formerly LJM Infotech"),
         ("120+", "Specialists, not generalists"),
         ("AU + India", "Onshore leadership · offshore delivery")],
        n=2)

    # 3 · What we do
    d.pillars(
        "02", "What we do · four service pillars",
        [("01", "Workday & ERP",
          "Implementation, optimisation and managed support for Workday HCM, Financials and adjacent ERP."),
         ("02", "Payroll consulting",
          "Compliant, accurate payroll across complex awards and multi-entity structures."),
         ("03", "Data & AI",
          "Lakehouse architecture, governed analytics and applied AI on your people and finance data."),
         ("04", "Intelligent automation",
          "RPA, ML and process design combined to remove low-value work.")],
        n=3)

    # 4 · How we work
    d.steps(
        "03", "How we work · think human first",
        [("1", "Listen", "Understand the people and the work before the technology."),
         ("2", "Design", "Shape the platform around how teams actually operate."),
         ("3", "Deliver", "Implement in weeks-long increments, not multi-year programs."),
         ("4", "Care", "Stay on after go-live; measure outcomes, not tickets closed.")],
        n=4)

    # 5 · Data & AI
    d.split(
        "04", "Data & AI",
        [("Certified consultants design ", False), ("Lakehouse", True),
         (" architectures and integrate Databricks with Workday, payroll and "
          "finance — so leaders decide on current data.", False)],
        [("Governed KPI store", "One trusted source for people and finance measures, defined once."),
         ("Delta Lake pipelines", "Batch and streaming ingestion, versioned and reliable by design."),
         ("Real-time insight", "Current figures in the hands of leaders, not last quarter's export.")],
        n=5)

    # 6 · Outcomes
    d.stats3(
        "05", "Outcomes",
        [("40+", "Workday deployments", False),
         ("98%", "Client retention", False),
         ("Weeks", "Not months, to value", True)],
        "They started with our people, not the platform — and the rollout stuck because of it.",
        "Programme lead · enterprise services (illustrative)",
        n=6)

    # 7 · Why Mivada
    d.reasons(
        "06", "Why Mivada",
        [("1", "Certified specialists",
          "Workday- and Databricks-certified consultants, not generalists assigned to fill a seat."),
         ("2", "Onshore + offshore",
          "Senior onshore leadership paired with cost-effective delivery from India."),
         ("3", "People-first change",
          "Adoption built in from day one, so the platforms we deliver actually get used.")],
        n=7)

    # 8 · Contact
    d.contact(
        "Mivada",
        [("Let's build smarter systems, with people at the ", False),
         ("centre.", True)],
        "mivada.com   ·   hello@mivada.com   ·   Sydney & Melbourne   ·   Onshore AU & India",
        n=8)

    return d.save("meridian.pptx")


if __name__ == "__main__":
    out = build()
    print("Saved", out)

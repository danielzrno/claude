#!/usr/bin/env python3
# ============================================================================
# Three compelling, customer-facing versions each of three Micah-deck slides —
# Why Mivada (2), Mivada's Workday Services (3), Workday Capability (4) — in the
# Editorial look. Diagrams reimagined as clean editorial compositions.
#   python variations.py  ->  slide-variations_editorial.pptx  (9 slides)
# ============================================================================
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck, THEME, PAGE_W
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CORAL, INK, WHITE = "EA493F", "111111", "FFFFFF"
M = 0.96                      # page margin
CW = PAGE_W - 2 * M           # content width = 11.413
d = Deck()
BRANDS = "Qantas   ·   Guzman y Gomez   ·   Western Sydney International Airport"


def head(s, eyebrow, hpre, hacc, tag, dark=False, hsize=34, y=1.32):
    d._slug(s, tag, dark=dark)
    d._eyebrow(s, M, 0.92, 11, eyebrow, color="coral")
    _, tf = d._box(s, 0.92, y, CW, 1.2)
    p = d._para(tf, first=True, line=1.0)
    d._run(p, hpre + (" " if hacc else ""), hsize, (WHITE if dark else INK), bold=True, spacing=-0.022)
    if hacc:
        d._run(p, hacc, hsize, CORAL, bold=True, spacing=-0.022)


def standfirst(s, text, y, w=10.8, color="323232", size=14):
    _, tf = d._box(s, M, y, w, 0.9)
    d._run(d._para(tf, first=True, line=1.3), text, size, color, bold=False, spacing=0)


def trusted(s, y=6.92, dark=False):
    d._hline(s, M, y - 0.16, CW, THEME["rev_hair"] if dark else THEME["hairline"], 1.0)
    _, tf = d._box(s, M, y, CW, 0.4)
    p = d._para(tf, first=True)
    d._run(p, "TRUSTED BY   ", 10.5, CORAL, bold=True, spacing=0.14, caps=True)
    d._run(p, BRANDS, 10.5, THEME["mid_grey"] if not dark else THEME["rev_soft"], bold=True, spacing=0.04)


def card(s, x, y, w, h, num, title, desc=None, dark=False):
    d._rect(s, x, y, w, h, fill_hex=("141414" if dark else WHITE),
            line_hex=(THEME["rev_hair"] if dark else THEME["hairline"]), line_w=1.0)
    _, tf = d._box(s, x + 0.26, y + 0.2, w - 0.45, 0.4)
    d._run(d._para(tf, first=True), num, 13.5, CORAL, bold=True, spacing=0.06)
    _, tf = d._box(s, x + 0.26, y + 0.54, w - 0.45, 0.62)
    d._run(d._para(tf, first=True, line=1.04), title, 14.5, (WHITE if dark else INK), bold=True, spacing=-0.01)
    if desc:
        _, tf = d._box(s, x + 0.26, y + 1.14, w - 0.48, h - 1.2)
        d._run(d._para(tf, first=True, line=1.16), desc, 10.5,
               (THEME["rev_soft"] if dark else "323232"), bold=False, spacing=0)


def grid(items, cols, x0, y0, w, rowh, gapx, gapy, dark=False, desc=True):
    cw = (w - gapx * (cols - 1)) / cols
    for i, it in enumerate(items):
        r, c = divmod(i, cols)
        x = x0 + c * (cw + gapx); y = y0 + r * (rowh + gapy)
        card(s, x, y, cw, rowh, it[0], it[1], (it[2] if desc and len(it) > 2 else None), dark=dark)


# ============================================================================
# WHY MIVADA — 3 versions
# ============================================================================
WHY = [
    ("01", "Australian-owned", "Creating Australian jobs, accountable locally."),
    ("02", "Not a Big 4", "Senior practitioners on the tools — no pyramid, no juniors learning on you."),
    ("03", "Direct access to our CEO", "A reduced escalation matrix — decisions in hours, not weeks."),
    ("04", "Dedicated engagement manager", "One local owner across your whole engagement."),
    ("05", "People Systems & Transformation", "HCM, Payroll, Time & Attendance and HRIS."),
    ("06", "Data, Analytics & BI", "From reporting to governed, trusted analytics."),
    ("07", "Process & Automation", "Remove the manual, error-prone work."),
    ("08", "Dedicated offshore capability", "Scale without losing local accountability."),
]

# v1 — differentiator grid (off-white)
s = d._slide(THEME["off_white"])
head(s, "Why Mivada", "Your long-term", "Workday partner.", "Why Mivada · v1")
standfirst(s, "We grew out of a customer — so we understand your challenges from the inside.", 2.18)
grid(WHY, 4, M, 2.95, CW, 1.66, 0.22, 0.2, desc=True)
trusted(s)

# v2 — editorial manifesto (black)
s = d._slide(THEME["black"])
d._slug(s, "Why Mivada · v2", dark=True)
d._eyebrow(s, M, 0.92, 10, "Why Mivada", color="coral")
_, tf = d._box(s, 0.92, 1.5, 11.6, 2.2)
p = d._para(tf, first=True, line=0.98)
d._run(p, "We grew out of a customer.", 44, WHITE, bold=True, spacing=-0.025)
p2 = d._para(tf, line=0.98)
d._run(p2, "So we ", 44, WHITE, bold=True, spacing=-0.025)
d._run(p2, "get it.", 44, CORAL, bold=True, spacing=-0.025)
_, tf = d._box(s, 0.96, 3.55, 11, 0.5)
d._run(d._para(tf, first=True), "Your long-term partner in Workday implementations & managed services.",
       15, THEME["rev_body"], bold=True, spacing=0)
# two columns of numbered differentiators
col2 = [WHY[:4], WHY[4:]]
for ci, group in enumerate(col2):
    cx = M + ci * (CW / 2)
    for ri, it in enumerate(group):
        ry = 4.35 + ri * 0.62
        _, tf = d._box(s, cx, ry, 0.7, 0.5)
        d._run(d._para(tf, first=True), it[0], 14, CORAL, bold=True, spacing=0)
        _, tf = d._box(s, cx + 0.62, ry, CW / 2 - 0.9, 0.5)
        d._run(d._para(tf, first=True), it[1], 14, WHITE, bold=True, spacing=-0.01)
trusted(s, dark=True)

# v3 — split: manifesto left, capability rows right (off-white)
s = d._slide(THEME["off_white"])
d._slug(s, "Why Mivada · v3")
d._eyebrow(s, M, 1.15, 5, "Why Mivada", color="coral")
_, tf = d._box(s, 0.92, 1.6, 5.3, 1.9)
p = d._para(tf, first=True, line=1.0)
d._run(p, "Built by people who ", 34, INK, bold=True, spacing=-0.022)
d._run(p, "ran it.", 34, CORAL, bold=True, spacing=-0.022)
_, tf = d._box(s, M, 3.65, 5.0, 2.6)
d._run(d._para(tf, first=True, line=1.34),
       "We grew out of a customer, so we understand your challenges first-hand. Australian-owned and "
       "senior-led — not a Big 4 — with direct access to our CEO and a dedicated local engagement "
       "manager who owns your outcome.", 13.5, "323232", bold=False, spacing=0)
rx, rw = 6.95, PAGE_W - M - 6.95
rows = [("People Systems & Transformation", "HCM · Payroll · T&A · HRIS"),
        ("Data, Analytics & BI", "Reporting · dashboards · governed analytics"),
        ("Process & Automation", "Redesign · integrations · automation"),
        ("Dedicated offshore capability", "Scale with local accountability")]
y = 1.3; rh = 1.12
d._hline(s, rx, y, rw, THEME["hairline"], 1.0)
for k, sub in rows:
    _, tf = d._box(s, rx, y + 0.2, rw, 0.5)
    d._run(d._para(tf, first=True), k, 16, INK, bold=True, spacing=-0.01)
    _, tf = d._box(s, rx, y + 0.62, rw, 0.4)
    d._run(d._para(tf, first=True), sub, 11.5, CORAL, bold=True, spacing=0.02)
    y += rh
    d._hline(s, rx, y, rw, THEME["hairline"], 1.0)
trusted(s)

# ============================================================================
# MIVADA'S WORKDAY SERVICES — 3 versions
# ============================================================================
SERVICES9 = [
    ("01", "Advisory & roadmap"), ("02", "Implementation & deployment"), ("03", "Optimisation & phase 2"),
    ("04", "Managed services (AMS)"), ("05", "Integrations"), ("06", "Data, analytics & BI"),
    ("07", "Reporting & dashboards"), ("08", "Security & compliance"), ("09", "Automation"),
]

# v1 — service catalogue 3x3 (off-white)
s = d._slide(THEME["off_white"])
head(s, "What we do", "Workday,", "end to end.", "Workday services · v1")
standfirst(s, "One partner across the whole platform — from first roadmap to long-term managed services.", 2.18)
grid(SERVICES9, 3, M, 2.95, CW, 1.16, 0.22, 0.2, desc=False)

# v2 — lifecycle flow (black)
s = d._slide(THEME["black"])
head(s, "Workday services", "Across the whole", "Workday lifecycle.", "Workday services · v2", dark=True)
stages = [("Advise", "Roadmap, business case, design"),
          ("Implement", "Deploy HCM, Payroll, Finance & more"),
          ("Optimise", "Phase 2, integrations, reporting"),
          ("Manage", "AMS, automation, continuous value")]
total = CW; col = total / 4
ytop = 3.4
for i, (name, sub) in enumerate(stages):
    cx = M + i * col
    if i > 0:
        d._vline(s, cx, ytop, 2.4, THEME["rev_hair"], 1.0)
    pad = 0.0 if i == 0 else 0.34
    _, tf = d._box(s, cx + pad, ytop, col - pad - 0.25, 0.5)
    d._run(d._para(tf, first=True), f"STAGE 0{i+1}", 11, CORAL, bold=True, spacing=0.16, caps=True)
    _, tf = d._box(s, cx + pad, ytop + 0.5, col - pad - 0.25, 0.7)
    d._run(d._para(tf, first=True), name, 26, WHITE, bold=True, spacing=-0.025)
    _, tf = d._box(s, cx + pad, ytop + 1.25, col - pad - 0.3, 1.2)
    d._run(d._para(tf, first=True, line=1.25), sub, 12, THEME["rev_soft"], bold=False, spacing=0)
_, tf = d._box(s, M, 6.5, CW, 0.4)
d._run(d._para(tf, first=True), "Wrapped in managed services & support at every stage.",
       12.5, THEME["mid_grey"], bold=True, spacing=0.02)

# v3 — three pillars (off-white)
s = d._slide(THEME["off_white"])
head(s, "Services", "Three practices,", "one platform.", "Workday services · v3")
pillars = [("People Systems & Transformation",
            ["HCM & core HR", "Payroll & compliance", "Time & Attendance", "HRIS implementation & optimisation"]),
           ("Data, Analytics & BI",
            ["Reporting & dashboards", "Prism & data lakehouse", "Governed analytics", "Insight at scale"]),
           ("Business Process & Automation",
            ["Process redesign", "Integrations", "Automation", "Controls & assurance"])]
total = CW; col = total / 3; ytop = 2.7
for i, (name, subs) in enumerate(pillars):
    cx = M + i * col
    if i > 0:
        d._vline(s, cx, ytop, 3.9, THEME["hairline"], 1.0)
    pad = 0.0 if i == 0 else 0.34
    _, tf = d._box(s, cx + pad, ytop, col - pad - 0.3, 0.9)
    d._run(d._para(tf, first=True, line=1.05), name, 18, INK, bold=True, spacing=-0.015)
    yy = ytop + 1.15
    for sub in subs:
        _, tf = d._box(s, cx + pad, yy, col - pad - 0.3, 0.5)
        p = d._para(tf, first=True)
        d._run(p, "—  ", 13, CORAL, bold=True, spacing=0)
        d._run(p, sub, 13, "323232", bold=False, spacing=0)
        yy += 0.55
_, tf = d._box(s, M, 6.62, CW, 0.4)
d._run(d._para(tf, first=True), "All wrapped in managed services & support.", 12.5, CORAL, bold=True, spacing=0.02)

# ============================================================================
# WORKDAY CAPABILITY — 3 versions
# ============================================================================
STATS = [("3", "+", "Certifications / consultant"), ("2000", "+", "Workday integrations"),
         ("3000", "+", "Workday custom reports"), ("30", "+", "Implementations & support"),
         ("200", "+", "Employees · AU, US & India"), ("100", "+", "People-systems clients"),
         ("10", "+", "Years delivering Workday")]

# v1 — KPI grid (off-white)
s = d._slide(THEME["off_white"])
head(s, "Capability", "Capability in", "Workday.", "Capability · v1")
standfirst(s, "Empowering innovation through a Workday partnership and local expertise — consultants average "
              "3+ years' Workday experience.", 2.18)
d._stat_row(s, 3.7, STATS[:4], light=True)
d._stat_row(s, 5.6, STATS[4:], light=True)

# v2 — hero stat (black)
s = d._slide(THEME["black"])
d._slug(s, "Capability · v2", dark=True)
d._eyebrow(s, M, 1.0, 10, "Proven capability", color="coral")
_, tf = d._box(s, 0.9, 1.5, 7.2, 2.6)
p = d._para(tf, first=True, line=0.82)
d._run(p, "10", 150, WHITE, bold=True, spacing=-0.04)
d._run(p, "+", 150, CORAL, bold=True, spacing=-0.04)
_, tf = d._box(s, 1.0, 4.05, 6.6, 0.6)
d._run(d._para(tf, first=True), "YEARS DELIVERING WORKDAY", 14, THEME["rev_body"], bold=True, spacing=0.14, caps=True)
_, tf = d._box(s, 1.0, 4.7, 6.4, 1.6)
d._run(d._para(tf, first=True, line=1.3),
       "Empowering innovation through a true Workday partnership and deep local expertise.",
       15, THEME["rev_soft"], bold=False, spacing=0)
# supporting stats, right column
rx = 8.35; ry = 1.6
support = [("30+", "Implementations & support"), ("200+", "Employees · AU, US & India"),
           ("2000+", "Workday integrations"), ("3000+", "Custom reports")]
d._hline(s, rx, ry - 0.16, PAGE_W - M - rx, THEME["rev_hair"], 1.0)
for num, lab in support:
    _, tf = d._box(s, rx, ry, PAGE_W - M - rx, 0.7)
    p = d._para(tf, first=True, line=0.9)
    d._run(p, num, 34, WHITE, bold=True, spacing=-0.03)
    _, tf = d._box(s, rx, ry + 0.66, PAGE_W - M - rx, 0.4)
    d._run(d._para(tf, first=True), lab.upper(), 10, THEME["mid_grey"], bold=True, spacing=0.1, caps=True)
    ry += 1.25
    d._hline(s, rx, ry - 0.18, PAGE_W - M - rx, THEME["rev_hair"], 1.0)

# v3 — KPI cards (off-white)
s = d._slide(THEME["off_white"])
head(s, "Capability", "The numbers behind", "the partnership.", "Capability · v3")
cards = [("3+", "Certifications / consultant"), ("2000+", "Workday integrations"),
         ("3000+", "Custom reports"), ("30+", "Implementations & support"),
         ("200+", "Employees · AU, US & India"), ("100+", "People-systems clients")]
cols = 3; gx, gy = 0.24, 0.24; cw = (CW - gx * (cols - 1)) / cols; ch = 1.55; y0 = 2.95
for i, (num, lab) in enumerate(cards):
    r, c = divmod(i, cols)
    x = M + c * (cw + gx); y = y0 + r * (ch + gy)
    d._rect(s, x, y, cw, ch, fill_hex=WHITE, line_hex=THEME["hairline"], line_w=1.0)
    _, tf = d._box(s, x + 0.3, y + 0.26, cw - 0.5, 0.8)
    p = d._para(tf, first=True, line=0.9)
    d._run(p, num.rstrip("+"), 40, INK, bold=True, spacing=-0.035)
    d._run(p, "+" if num.endswith("+") else "", 40, CORAL, bold=True, spacing=-0.035)
    _, tf = d._box(s, x + 0.3, y + 1.04, cw - 0.5, 0.45)
    d._run(d._para(tf, first=True), lab.upper(), 10.5, THEME["mid_grey"], bold=True, spacing=0.1, caps=True)
_, tf = d._box(s, M, y0 + 2 * (ch + gy) + 0.05, CW, 0.5)
d._run(d._para(tf, first=True), "10+ years delivering Workday · consultants average 3+ years' experience.",
       13, CORAL, bold=True, spacing=0.02)

d.save("slide-variations_editorial.pptx")
print(f"Wrote slide-variations_editorial.pptx ({len(d.prs.slides._sldIdLst)} slides)")

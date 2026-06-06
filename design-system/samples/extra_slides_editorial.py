#!/usr/bin/env python3
# ============================================================================
# Mivada — AMS & Services pack, rebuilt in the Editorial look from extra_slides.
# Six source slides re-expressed (no content lost): FIN & HCM coverage, the two
# engagement models, the AMS operating framework, the coverage model, the
# governance model, and the team structure. Diagrams reimagined as clean
# editorial compositions.
#   python extra_slides_editorial.py -> Mivada_AMS_Services_editorial.pptx
# ============================================================================
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck, THEME, PAGE_W
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

CORAL, INK, WHITE = "EA493F", "111111", "FFFFFF"
M = 0.96; CW = PAGE_W - 2 * M
d = Deck()

def head(s, hpre, hacc, tag, dark=False, y=1.0):
    d._slug(s, tag, dark=dark)
    _, tf = d._box(s, 0.92, y, CW, 1.1); p = d._para(tf, first=True, line=1.0)
    d._run(p, hpre + " ", 33, (WHITE if dark else INK), bold=True, spacing=-0.022)
    d._run(p, hacc, 33, CORAL, bold=True, spacing=-0.022)

def standfirst(s, text, y=1.92, w=11.0, size=13, dark=False):
    _, tf = d._box(s, M, y, w, 0.7)
    d._run(d._para(tf, first=True, line=1.25), text, size, THEME["rev_soft"] if dark else "323232", bold=False, spacing=0)

def bullets(s, x, y, w, items, size=11, dark=False, gap=0.42, dash=True):
    cy = y
    for it in items:
        _, tf = d._box(s, x, cy, w, gap)
        p = d._para(tf, first=True, line=1.12)
        if dash:
            d._run(p, "—  ", size, CORAL, bold=True, spacing=0)
        d._run(p, it, size, (THEME["rev_body"] if dark else "323232"), bold=False, spacing=0)
        cy += gap
    return cy

def card(s, x, y, w, h, label, sub, items, dark=False):
    d._rect(s, x, y, w, h, fill_hex=("141414" if dark else WHITE),
            line_hex=(THEME["rev_hair"] if dark else THEME["hairline"]), line_w=1.0)
    _, tf = d._box(s, x + 0.3, y + 0.26, w - 0.55, 0.4)
    d._run(d._para(tf, first=True), label.upper(), 12.5, CORAL, bold=True, spacing=0.12, caps=True)
    _, tf = d._box(s, x + 0.3, y + 0.66, w - 0.55, 0.5)
    d._run(d._para(tf, first=True, line=1.05), sub, 14.5, (WHITE if dark else INK), bold=True, spacing=-0.01)
    bullets(s, x + 0.3, y + 1.32, w - 0.6, items, size=10.5, dark=dark, gap=0.46)

# ============================================================================
# 0 · COVER
# ============================================================================
d.cover_plain("Mivada · AMS & services", "Run, supported,", "improved.",
              "Application Management Services", "mivada.com")

# ============================================================================
# 1 · FIN & HCM COVERAGE (off-white)
# ============================================================================
s = d._slide(THEME["off_white"]); head(s, "Full-stack Workday,", "FIN & HCM.", "Consultant coverage")
standfirst(s, "Full-stack Workday functional and integration expertise across both suites.")
HCM = ["Core HCM", "Absence", "Time Tracking", "Recruiting", "Talent & Performance", "Compensation",
       "Benefits", "Payroll (AU / global)", "Learning", "Security & BP config",
       "Integrations (EIB, CC, Studio)", "Advanced & Composite Reporting"]
FIN = ["Financial Accounting (GL)", "Accounting Center", "Accounts Payable / Suppliers",
       "Accounts Receivable / Customers", "Procurement", "Expenses", "Banking & Settlement",
       "Business Assets", "Financial Reporting", "Prism Analytics", "Security & BP config",
       "Integrations (EIB, CC, Studio)"]
def suite(x, label, items):
    _, tf = d._box(s, x, 2.7, 5.0, 0.5)
    d._run(d._para(tf, first=True), label, 22, INK, bold=True, spacing=-0.02)
    # two sub-columns of 6 items each
    half = (len(items) + 1) // 2
    bullets(s, x, 3.35, 2.5, items[:half], size=10.5, gap=0.38)
    bullets(s, x + 2.55, 3.35, 2.5, items[half:], size=10.5, gap=0.38)
d._vline(s, 6.66, 2.7, 3.5, THEME["hairline"], 1.0)
suite(M, "HCM", HCM)
suite(7.0, "FIN", FIN)
_, tf = d._box(s, M, 6.5, CW, 0.5); p = d._para(tf, first=True)
d._run(p, "Adaptive Planning — all modules.   ", 12.5, CORAL, bold=True, spacing=0.01)
d._run(p, "Cross-skilled in reporting, security and integrations across both suites — one team, two suites.",
       12.5, "323232", bold=False, spacing=0)

# ============================================================================
# 2 · TWO ENGAGEMENT MODELS (black)
# ============================================================================
s = d._slide(THEME["black"]); head(s, "Two ways we", "support you.", "Engagement models", dark=True)
standfirst(s, "Augmentation and Application Management Services — flex between them as your needs change.", dark=True)
cw = (CW - 0.4) / 2
card(s, M, 2.55, cw, 3.6, "Staff augmentation", "Our certified consultants, embedded in your team", [
    "You direct the work; we provide certified Workday FIN & HCM consultants",
    "Scale capacity up or down by the day, week or sprint",
    "Ideal for project surge, BAU backlog, leave cover or niche skills",
    "Onshore (AU), offshore (India) or a blended model",
    "Time-and-materials or a pre-agreed block of hours",
], dark=True)
card(s, M + cw + 0.4, 2.55, cw, 3.6, "Managed services (AMS)", "We own the outcome, end to end", [
    "Mivada owns incident, problem, change & release management",
    "SLA-backed and governed, with 24×7 P1/P2 on-call",
    "Continuous improvement and bi-annual release management",
    "Ideal for steady-state operations and long-term partnership",
    "Fixed monthly hours with carry-forward flexibility",
], dark=True)
_, tf = d._box(s, M, 6.35, CW, 0.5)
d._run(d._para(tf, first=True), "Many clients start with augmentation and move into AMS as their environment stabilises — the same team carries the knowledge across.",
       12, THEME["mid_grey"], bold=True, spacing=0)

# ============================================================================
# 3 · AMS OPERATING FRAMEWORK (off-white) — 23 capabilities, 3 themes
# ============================================================================
s = d._slide(THEME["off_white"]); head(s, "AMS operating", "framework.", "AMS framework")
standfirst(s, "Driving operational excellence through governance, integration and continuous improvement.")
themes = [
    ("Operational excellence · ITIL", ["Incident management", "Problem management", "Change management",
        "Release management", "Service requests", "SLAs & metrics", "Change Advisory Board",
        "Service integration (ITSM)", "Capacity management"]),
    ("Operations as a priority", ["Governance", "BP & security config management", "Integration health checks",
        "Cyber security", "Backlog management", "Prioritisation forum", "WD bi-annual release management"]),
    ("Business value realisation", ["Continuous improvement", "Enhancements", "Knowledge base",
        "Thought leadership", "Technical-debt reduction", "Product & vendor relationship"]),
]
col = CW / 3
for i, (title, items) in enumerate(themes):
    cx = M + i * col
    if i > 0: d._vline(s, cx, 2.7, 4.0, THEME["hairline"], 1.0)
    pad = 0.0 if i == 0 else 0.34
    _, tf = d._box(s, cx + pad, 2.7, col - pad - 0.3, 0.7)
    d._run(d._para(tf, first=True, line=1.05), title, 14.5, INK, bold=True, spacing=-0.01)
    bullets(s, cx + pad, 3.45, col - pad - 0.3, items, size=11, gap=0.4)

# ============================================================================
# 4 · COVERAGE MODEL (black) — onshore/offshore timeline
# ============================================================================
s = d._slide(THEME["black"]); head(s, "Coverage model —", "always on.", "Coverage model", dark=True)
standfirst(s, "Seamless onshore–offshore collaboration for continuous support.", dark=True)
ax0, axw = M, CW
def t(h): return ax0 + (h - 4) / 18.0 * axw            # 4:00 .. 22:00 AEST
# warm-handover band (overlap 12:30–17:30)
d._rect(s, t(12.5), 2.95, t(17.5) - t(12.5), 1.7, fill_hex="2A1714", line_hex=None)
# bars
def bar(y, x1, x2, label, hours):
    _, tf = d._box(s, M, y - 0.34, 8, 0.3)
    p = d._para(tf, first=True)
    d._run(p, label + "   ", 11.5, WHITE, bold=True, spacing=0.02)
    d._run(p, hours, 11.5, THEME["mid_grey"], bold=True, spacing=0.04)
    d._rect(s, x1, y, x2 - x1, 0.5, fill_hex=CORAL, line_hex=None)
bar(3.15, t(8.5), t(17.5), "ONSHORE · Australia", "8:30 – 17:30 AEST")
bar(4.15, t(12.5), t(21.5), "OFFSHORE · India", "8:00 – 17:00 IST")
# hour ticks
for h, lab in [(8.5, "8:30"), (12.5, "12:30"), (17.5, "17:30"), (21.5, "21:30")]:
    _, tf = d._box(s, t(h) - 0.4, 4.78, 0.8, 0.3)
    d._run(d._para(tf, first=True, align=PP_ALIGN.CENTER), lab, 9.5, THEME["mid_grey"], bold=True, spacing=0.04)
# 24x7 on-call strip
d._rect(s, M, 5.35, CW, 0.46, fill_hex="141414", line_hex=THEME["rev_hair"], line_w=1.0)
_, tf = d._box(s, M + 0.25, 5.42, CW - 0.5, 0.34); p = d._para(tf, first=True)
d._run(p, "24×7 P1/P2 ON-CALL   ", 11, CORAL, bold=True, spacing=0.12, caps=True)
d._run(p, "always-on safety net for P1 & P2 incidents, in line with monthly hours.", 11, THEME["rev_body"], bold=True, spacing=0)
# descriptions
bullets(s, M, 6.05, 5.4, ["Governance & stakeholder engagement", "Resolution of P1 / P2 incidents",
        "App health check & start-of-day triage"], size=10.5, dark=True, gap=0.36)
bullets(s, 7.0, 6.05, 5.4, ["Day-to-day incident management & support", "Routine incidents & service requests",
        "Monitoring & backlog resolution per SLAs"], size=10.5, dark=True, gap=0.36)

# ============================================================================
# 5 · GOVERNANCE MODEL (off-white) — 3 tiers
# ============================================================================
s = d._slide(THEME["off_white"]); head(s, "Governance —", "aligned at every tier.", "Governance")
standfirst(s, "Strong governance. Aligned teams. Confident execution.")
tiers = [
    ("Strategic", "Quarterly", "Executive Sponsor · CIO / IT Exec · Head of HR / Payroll",
     "Strategic oversight & direction · innovation & thought leadership · roadmap & investment · relationship health",
     "CEO · Client Partner"),
    ("Tactical", "Monthly", "Product Owner – Workday · IT Service Owner – AMS · Risk, Security & Compliance Lead",
     "Service performance & SLA review · risk & compliance · release planning & adoption · improvement backlog",
     "Client Partner · Service Delivery Lead"),
    ("Operational", "Weekly / daily", "IT App Support / AMS team · Change & Release Coordinators · Super Users / SMEs",
     "Daily incident triage & resolution · change & release execution · user experience · operational performance",
     "Service Delivery Lead · Workday Consultants"),
]
y = 2.55; bandh = 1.35
d._hline(s, M, y, CW, THEME["hairline"], 1.0)
# column headers
for hx, htxt in [(2.95, "CUSTOMER"), (6.5, "ACTIVITIES"), (10.55, "MIVADA")]:
    _, tf = d._box(s, hx, y + 0.08, 3.4, 0.3)
    d._run(d._para(tf, first=True), htxt, 9, THEME["mid_grey"], bold=True, spacing=0.12, caps=True)
y += 0.42
for name, cad, cust, act, miv in tiers:
    _, tf = d._box(s, M, y + 0.06, 1.9, 0.6)
    d._run(d._para(tf, first=True), name, 15, INK, bold=True, spacing=-0.01)
    _, tf = d._box(s, M, y + 0.42, 1.9, 0.3)
    d._run(d._para(tf, first=True), cad.upper(), 9, CORAL, bold=True, spacing=0.1, caps=True)
    _, tf = d._box(s, 2.95, y + 0.04, 3.45, 1.1); d._run(d._para(tf, first=True, line=1.2), cust, 10, "323232", bold=False, spacing=0)
    _, tf = d._box(s, 6.5, y + 0.04, 3.9, 1.1); d._run(d._para(tf, first=True, line=1.2), act, 10, "323232", bold=False, spacing=0)
    _, tf = d._box(s, 10.55, y + 0.04, 1.85, 1.1); d._run(d._para(tf, first=True, line=1.2), miv, 10, INK, bold=True, spacing=0)
    y += bandh
    d._hline(s, M, y, CW, THEME["hairline"], 1.0)
_, tf = d._box(s, M, y + 0.12, CW, 0.4)
d._run(d._para(tf, first=True), "Principles: transparent reporting · escalation by exception · decisions at the right tier · improvement built in.",
       11, CORAL, bold=True, spacing=0.01)

# ============================================================================
# 6 · TEAM STRUCTURE (black)
# ============================================================================
s = d._slide(THEME["black"]); head(s, "One integrated team —", "onshore + offshore.", "Team structure", dark=True)
standfirst(s, "Integrated onshore–offshore teams delivering as one.", dark=True)
cw3 = (CW - 0.8) / 3
lead = [("Client Partner", "Strategic relationship & sponsorship"),
        ("Service Delivery Lead", "Day-to-day AMS operations & SLAs")]
for i, (role, desc) in enumerate(lead):
    x = M + i * (cw3 + 0.4)
    card_h = 1.5
    d._rect(s, x, 2.5, cw3, card_h, fill_hex="141414", line_hex=THEME["rev_hair"], line_w=1.0)
    _, tf = d._box(s, x + 0.28, 2.72, cw3 - 0.5, 0.5); d._run(d._para(tf, first=True), role, 16, WHITE, bold=True, spacing=-0.01)
    _, tf = d._box(s, x + 0.28, 3.2, cw3 - 0.5, 0.6); d._run(d._para(tf, first=True, line=1.2), desc, 11, THEME["rev_soft"], bold=False, spacing=0)
# third leadership card — thought leadership / CEO access
x = M + 2 * (cw3 + 0.4)
d._rect(s, x, 2.5, cw3, 1.5, fill_hex=CORAL, line_hex=None)
_, tf = d._box(s, x + 0.28, 2.72, cw3 - 0.5, 0.4); d._run(d._para(tf, first=True), "THOUGHT LEADERSHIP", 10.5, WHITE, bold=True, spacing=0.1, caps=True)
_, tf = d._box(s, x + 0.28, 3.12, cw3 - 0.5, 0.8); d._run(d._para(tf, first=True, line=1.18), "Direct CEO access for executive escalation and strategic sponsorship throughout the engagement.", 11, WHITE, bold=True, spacing=0)
# consultant rows
cwh = (CW - 0.4) / 2
team = [("Onshore · Australia", "Functional / Integration Consultants", "Core HCM · Financials · Integrations (inbound & outbound)"),
        ("Offshore · India", "Functional / Integration Consultants", "Core HCM · Financials · Integrations (inbound & outbound)")]
for i, (loc, role, scope) in enumerate(team):
    x = M + i * (cwh + 0.4)
    d._rect(s, x, 4.2, cwh, 1.35, fill_hex="141414", line_hex=THEME["rev_hair"], line_w=1.0)
    _, tf = d._box(s, x + 0.3, 4.4, cwh - 0.5, 0.4); d._run(d._para(tf, first=True), loc.upper(), 11, CORAL, bold=True, spacing=0.1, caps=True)
    _, tf = d._box(s, x + 0.3, 4.78, cwh - 0.5, 0.4); d._run(d._para(tf, first=True), role, 14.5, WHITE, bold=True, spacing=-0.01)
    _, tf = d._box(s, x + 0.3, 5.18, cwh - 0.5, 0.4); d._run(d._para(tf, first=True), scope, 10.5, THEME["rev_soft"], bold=False, spacing=0)
_, tf = d._box(s, M, 5.75, CW, 0.5)
d._run(d._para(tf, first=True), "Dedicated Workday practice · innovation — automation, new Workday features and process optimisation, flowing from Mivada's wider practice.",
       11, THEME["mid_grey"], bold=True, spacing=0)

OUT = "Mivada_AMS_Services_editorial.pptx"
d.save(OUT)
print(f"Wrote {OUT} ({len(d.prs.slides._sldIdLst)} slides)")

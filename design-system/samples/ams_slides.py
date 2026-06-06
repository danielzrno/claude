#!/usr/bin/env python3
# ============================================================================
# Mivada AMS & Services slides (Editorial) — shared builders used by both the
# standalone AMS pack and the merged overview+services deck. The four diagram
# slides keep the ORIGINAL visual structures (layered framework, onshore/offshore
# gantt with hour grid, governance pyramid, team org-chart), restyled to brand.
#   python ams_slides.py  -> Mivada_AMS_Services_editorial.pptx  (standalone)
# ============================================================================
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck, THEME, PAGE_W
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

CORAL, INK, WHITE = "EA493F", "111111", "FFFFFF"
DARK = "141414"
M = 0.96; CW = PAGE_W - 2 * M

# ---------------- helpers ----------------------------------------------------
def head(d, s, hpre, hacc, tag, dark=False, y=1.0, hsize=33):
    d._slug(s, tag, dark=dark)
    _, tf = d._box(s, 0.92, y, CW, 1.1); p = d._para(tf, first=True, line=1.0)
    d._run(p, hpre + " ", hsize, (WHITE if dark else INK), bold=True, spacing=-0.022)
    d._run(p, hacc, hsize, CORAL, bold=True, spacing=-0.022)

def standfirst(d, s, text, y=1.9, w=11.2, size=13, dark=False):
    _, tf = d._box(s, M, y, w, 0.6)
    d._run(d._para(tf, first=True, line=1.25), text, size, THEME["rev_soft"] if dark else "323232", bold=False, spacing=0)

def bullets(d, s, x, y, w, items, size=11, dark=False, gap=0.42):
    cy = y
    for it in items:
        _, tf = d._box(s, x, cy, w, gap); p = d._para(tf, first=True, line=1.12)
        d._run(p, "—  ", size, CORAL, bold=True, spacing=0)
        d._run(p, it, size, (THEME["rev_body"] if dark else "323232"), bold=False, spacing=0)
        cy += gap

def tile(d, s, x, y, w, h, text, fill=WHITE, line=None, tc=INK, size=9.5):
    d._rect(s, x, y, w, h, fill_hex=fill, line_hex=line, line_w=1.0)
    _, tf = d._box(s, x + 0.08, y, w - 0.16, h, anchor=MSO_ANCHOR.MIDDLE)
    p = d._para(tf, first=True, align=PP_ALIGN.CENTER, line=1.0)
    d._run(p, text, size, tc, bold=True, spacing=-0.005)

def band(d, s, x, y, w, h, label, items=None, fill=INK, tc=WHITE):
    d._rect(s, x, y, w, h, fill_hex=fill)
    _, tf = d._box(s, x + 0.25, y, w - 0.5, h, anchor=MSO_ANCHOR.MIDDLE)
    p = d._para(tf, first=True)
    d._run(p, label.upper(), 11.5, tc, bold=True, spacing=0.1, caps=True)
    if items:
        d._run(p, "    " + items, 10, tc, bold=False, spacing=0)

def card(d, s, x, y, w, h, label, sub, items, dark=True):
    d._rect(s, x, y, w, h, fill_hex=(DARK if dark else WHITE),
            line_hex=(THEME["rev_hair"] if dark else THEME["hairline"]), line_w=1.0)
    _, tf = d._box(s, x + 0.3, y + 0.26, w - 0.55, 0.4)
    d._run(d._para(tf, first=True), label.upper(), 12.5, CORAL, bold=True, spacing=0.12, caps=True)
    _, tf = d._box(s, x + 0.3, y + 0.66, w - 0.55, 0.5)
    d._run(d._para(tf, first=True, line=1.05), sub, 14.5, (WHITE if dark else INK), bold=True, spacing=-0.01)
    bullets(d, s, x + 0.3, y + 1.32, w - 0.6, items, size=10.5, dark=dark, gap=0.46)

# ---------------- slides -----------------------------------------------------
def fin_hcm(d):
    s = d._slide(THEME["off_white"]); head(d, s, "Full-stack Workday,", "FIN & HCM.", "Consultant coverage")
    standfirst(d, s, "Full-stack Workday functional and integration expertise across both suites.")
    HCM = ["Core HCM", "Absence", "Time Tracking", "Recruiting", "Talent & Performance", "Compensation",
           "Benefits", "Payroll (AU / global)", "Learning", "Security & BP config",
           "Integrations (EIB, CC, Studio)", "Advanced & Composite Reporting"]
    FIN = ["Financial Accounting (GL)", "Accounting Center", "Accounts Payable / Suppliers",
           "Accounts Receivable / Customers", "Procurement", "Expenses", "Banking & Settlement",
           "Business Assets", "Financial Reporting", "Prism Analytics", "Security & BP config",
           "Integrations (EIB, CC, Studio)"]
    def suite(x, label, items):
        _, tf = d._box(s, x, 2.7, 5.0, 0.5); d._run(d._para(tf, first=True), label, 22, INK, bold=True, spacing=-0.02)
        half = (len(items) + 1) // 2
        bullets(d, s, x, 3.35, 2.5, items[:half], size=10.5, gap=0.38)
        bullets(d, s, x + 2.55, 3.35, 2.5, items[half:], size=10.5, gap=0.38)
    d._vline(s, 6.66, 2.7, 3.5, THEME["hairline"], 1.0)
    suite(M, "HCM", HCM); suite(7.0, "FIN", FIN)
    _, tf = d._box(s, M, 6.5, CW, 0.5); p = d._para(tf, first=True)
    d._run(p, "Adaptive Planning — all modules.   ", 12.5, CORAL, bold=True, spacing=0.01)
    d._run(p, "Cross-skilled in reporting, security and integrations across both suites — one team, two suites.", 12.5, "323232", bold=False, spacing=0)

def engagement_models(d):
    s = d._slide(THEME["black"]); head(d, s, "Two ways we", "support you.", "Engagement models", dark=True)
    standfirst(d, s, "Augmentation and Application Management Services — flex between them as your needs change.", dark=True)
    cw = (CW - 0.4) / 2
    card(d, s, M, 2.55, cw, 3.6, "Staff augmentation", "Our certified consultants, embedded in your team", [
        "You direct the work; we provide certified Workday FIN & HCM consultants",
        "Scale capacity up or down by the day, week or sprint",
        "Ideal for project surge, BAU backlog, leave cover or niche skills",
        "Onshore (AU), offshore (India) or a blended model",
        "Time-and-materials or a pre-agreed block of hours"])
    card(d, s, M + cw + 0.4, 2.55, cw, 3.6, "Managed services (AMS)", "We own the outcome, end to end", [
        "Mivada owns incident, problem, change & release management",
        "SLA-backed and governed, with 24×7 P1/P2 on-call",
        "Continuous improvement and bi-annual release management",
        "Ideal for steady-state operations and long-term partnership",
        "Fixed monthly hours with carry-forward flexibility"])
    _, tf = d._box(s, M, 6.35, CW, 0.5)
    d._run(d._para(tf, first=True), "Many clients start with augmentation and move into AMS as their environment stabilises — the same team carries the knowledge across.",
           12, THEME["mid_grey"], bold=True, spacing=0)

def ams_framework(d):
    """Layered framework stack: Continuous Improvement / Governance bands over a
    tile grid of capabilities, on Service Integration, on the 3 foundations."""
    s = d._slide(THEME["off_white"]); head(d, s, "AMS operating", "framework.", "AMS framework")
    standfirst(d, s, "Driving operational excellence through governance, integration and continuous improvement.")
    fx, fw = M, CW
    # top: continuous improvement (coral) then governance (ink)
    band(d, s, fx, 2.5, fw, 0.34, "Continuous improvement", fill=CORAL)
    band(d, s, fx, 2.88, fw, 0.46, "Governance",
         "Thought leadership · Product & vendor relationship · Technical-debt reduction · WD bi-annual release management", fill=INK)
    # tile grid 4x4 — the operational capabilities
    tiles = ["Incident management", "Problem management", "Change management", "Release management",
             "Service requests", "Change Advisory Board", "SLAs", "Metrics",
             "Integration health checks", "Backlog management", "Prioritisation forum", "Capacity management",
             "Enhancements", "Knowledge base", "Cyber security", "BP & security config"]
    gx, gy = 0.14, 0.14; cols = 4; tw = (fw - gx * (cols - 1)) / cols; th = 0.5; y0 = 3.5
    for i, t in enumerate(tiles):
        r, c = divmod(i, cols); x = fx + c * (tw + gx); y = y0 + r * (th + gy)
        tile(d, s, x, y, tw, th, t, fill=WHITE, line=THEME["hairline"], tc=INK, size=9.5)
    yb = y0 + 4 * (th + gy)
    band(d, s, fx, yb, fw, 0.42, "Service integration · ITSM", fill=CORAL)
    # 3 foundations
    fy = yb + 0.52; fwidth = (fw - 0.3) / 3
    for i, f in enumerate(["Operations as a priority", "Business value realisation", "Operational excellence · ITIL"]):
        x = fx + i * (fwidth + 0.15)
        tile(d, s, x, fy, fwidth, 0.4, f.upper(), fill="F2F3EE", line=THEME["hairline"], tc=CORAL, size=10)

def coverage_model(d):
    """Onshore/offshore gantt over an AEST/IST hour grid, warm-handover overlap, 24×7 strip."""
    s = d._slide(THEME["off_white"]); head(d, s, "Coverage model —", "always on.", "Coverage model")
    standfirst(d, s, "Seamless onshore–offshore collaboration for continuous support.")
    hours_aest = ["8:30", "9:30", "10:30", "11:30", "12:30", "13:30", "14:30", "15:30", "16:30",
                  "17:30", "18:30", "19:30", "20:30", "21:30"]
    hours_ist = ["4:00", "5:00", "6:00", "7:00", "8:00", "9:00", "10:00", "11:00", "12:00",
                 "13:00", "14:00", "15:00", "16:00", "17:00"]
    n = len(hours_aest); gx0 = M; gw = CW; cw = gw / n
    gy = 2.95
    # AEST header row
    _, tf = d._box(s, gx0 - 0.0, gy - 0.3, 1.2, 0.25)
    for i, h in enumerate(hours_aest):
        _, tf = d._box(s, gx0 + i * cw, gy - 0.28, cw, 0.24)
        d._run(d._para(tf, first=True, align=PP_ALIGN.CENTER), h, 8, THEME["mid_grey"], bold=True, spacing=0)
    # grid cells (faint)
    for i in range(n + 1):
        d._vline(s, gx0 + i * cw, gy, 1.5, THEME["hairline"], 0.75)
    d._hline(s, gx0, gy, gw, THEME["hairline"], 0.75)
    d._hline(s, gx0, gy + 1.5, gw, THEME["hairline"], 0.75)
    # warm-handover band (overlap cols 4..9 -> 12:30..17:30)
    d._rect(s, gx0 + 4 * cw, gy, (10 - 4) * cw, 1.5, fill_hex="FBE7E4", line_hex=None)
    # bars
    d._rect(s, gx0 + 0 * cw + 0.04, gy + 0.18, 10 * cw - 0.08, 0.46, fill_hex=CORAL)      # onshore 8:30-17:30 (cols 0-9)
    _, tf = d._box(s, gx0 + 0.12, gy + 0.22, 5, 0.36); d._run(d._para(tf, first=True), "ONSHORE · Australia", 11, WHITE, bold=True, spacing=0.02)
    d._rect(s, gx0 + 4 * cw + 0.04, gy + 0.86, 10 * cw - 0.08, 0.46, fill_hex=INK)         # offshore 12:30-21:30 (cols 4-13)
    _, tf = d._box(s, gx0 + 4 * cw + 0.12, gy + 0.9, 5, 0.36); d._run(d._para(tf, first=True), "OFFSHORE · India", 11, WHITE, bold=True, spacing=0.02)
    # IST footer row
    for i, h in enumerate(hours_ist):
        _, tf = d._box(s, gx0 + i * cw, gy + 1.54, cw, 0.24)
        d._run(d._para(tf, first=True, align=PP_ALIGN.CENTER), h, 8, THEME["mid_grey"], bold=True, spacing=0)
    _, tf = d._box(s, gx0, gy + 1.78, 3, 0.22); d._run(d._para(tf, first=True), "← AEST   ·   IST →", 8.5, CORAL, bold=True, spacing=0.06, caps=True)
    # descriptions
    bullets(d, s, M, 5.15, 5.3, ["Governance & stakeholder engagement", "Resolution of P1 / P2 incidents",
            "App health check & start-of-day triage"], size=10.5, gap=0.36)
    bullets(d, s, 6.5, 5.15, 4.0, ["Day-to-day incident management & support", "Routine incidents & service requests",
            "Monitoring & backlog resolution per SLAs"], size=10.5, gap=0.36)
    # 24x7 callout (coral) bottom strip
    d._rect(s, M, 6.5, CW, 0.46, fill_hex=CORAL)
    _, tf = d._box(s, M + 0.25, 6.57, CW - 0.5, 0.34); p = d._para(tf, first=True)
    d._run(p, "24×7 P1/P2 ON-CALL   ", 11, WHITE, bold=True, spacing=0.12, caps=True)
    d._run(p, "always-on safety net for P1 & P2 incidents, in line with monthly hours.", 11, WHITE, bold=False, spacing=0)

def governance_model(d):
    """Central stepped pyramid (Strategic/Tactical/Operational) flanked by
    Customer · Activities (left) and Mivada · Cadence (right)."""
    s = d._slide(THEME["off_white"]); head(d, s, "Governance —", "aligned at every tier.", "Governance")
    standfirst(d, s, "Strong governance. Aligned teams. Confident execution.")
    # column headers
    for hx, w, htxt in [(M, 2.0, "CUSTOMER"), (3.05, 2.0, "ACTIVITIES"), (8.3, 2.0, "MIVADA"), (10.4, 2.0, "CADENCE")]:
        _, tf = d._box(s, hx, 2.4, w, 0.3); d._run(d._para(tf, first=True), htxt, 9, THEME["mid_grey"], bold=True, spacing=0.12, caps=True)
    tiers = [
        ("Strategic", "Quarterly", "Executive Sponsor · CIO / IT Exec · Head of HR / Payroll",
         "Strategic oversight · innovation & thought leadership · roadmap & investment · relationship health", "CEO · Client Partner"),
        ("Tactical", "Monthly", "Product Owner · IT Service Owner – AMS · Risk & Compliance Lead",
         "Service & SLA review · risk & compliance · release planning & adoption · improvement backlog", "Client Partner · Service Delivery Lead"),
        ("Operational", "Weekly / daily", "AMS team · Change & Release Coordinators · Super Users / SMEs",
         "Daily triage & resolution · change & release execution · UX · operational performance", "Service Delivery Lead · Consultants"),
    ]
    y0 = 2.78; bh = 1.18
    pyr_cx = PAGE_W / 2; pyr_w = [1.5, 2.3, 3.1]; pyr_fill = [CORAL, "D8392E", INK]
    for i, (name, cad, cust, act, miv) in enumerate(tiers):
        y = y0 + i * bh
        # pyramid step
        w = pyr_w[i]; d._rect(s, pyr_cx - w / 2, y + 0.12, w, bh - 0.26, fill_hex=pyr_fill[i])
        _, tf = d._box(s, pyr_cx - w / 2, y + 0.12, w, bh - 0.26, anchor=MSO_ANCHOR.MIDDLE)
        d._run(d._para(tf, first=True, align=PP_ALIGN.CENTER), name.upper(), 12, WHITE, bold=True, spacing=0.08, caps=True)
        # left: customer + activities
        _, tf = d._box(s, M, y + 0.05, 2.0, bh - 0.1); d._run(d._para(tf, first=True, line=1.16), cust, 9.5, "323232", bold=False, spacing=0)
        _, tf = d._box(s, 3.05, y + 0.05, 1.95, bh - 0.1); d._run(d._para(tf, first=True, line=1.16), act, 9.5, "323232", bold=False, spacing=0)
        # right: mivada + cadence
        _, tf = d._box(s, 8.3, y + 0.05, 2.0, bh - 0.1); d._run(d._para(tf, first=True, line=1.16), miv, 9.5, INK, bold=True, spacing=0)
        _, tf = d._box(s, 10.4, y + 0.16, 2.0, 0.4); d._run(d._para(tf, first=True), cad.upper(), 9.5, CORAL, bold=True, spacing=0.08, caps=True)
        d._hline(s, M, y, CW, THEME["hairline"], 0.75)
    d._hline(s, M, y0 + 3 * bh, CW, THEME["hairline"], 0.75)
    _, tf = d._box(s, M, y0 + 3 * bh + 0.12, CW, 0.4)
    d._run(d._para(tf, first=True), "Principles: transparent reporting · escalation by exception · decisions at the right tier · improvement built in.",
           11, CORAL, bold=True, spacing=0.01)

def team_structure(d):
    """Org chart: leadership row → onshore/offshore delivery, with two coral
    framing banners (Thought Leadership / Dedicated Workday Practice)."""
    s = d._slide(THEME["off_white"]); head(d, s, "One integrated team —", "onshore + offshore.", "Team structure")
    standfirst(d, s, "Integrated onshore–offshore teams delivering as one.")
    cwL = (CW - 0.5) / 2
    def node(x, y, w, h, title, desc, fill=WHITE, tc=INK, dc="323232"):
        d._rect(s, x, y, w, h, fill_hex=fill, line_hex=(None if fill != WHITE else THEME["hairline"]), line_w=1.0)
        _, tf = d._box(s, x + 0.28, y + 0.2, w - 0.5, 0.4); d._run(d._para(tf, first=True), title, 15, tc, bold=True, spacing=-0.01)
        _, tf = d._box(s, x + 0.28, y + 0.62, w - 0.5, 0.6); d._run(d._para(tf, first=True, line=1.18), desc, 10.5, dc, bold=False, spacing=0)
    # leadership row
    node(M, 2.5, cwL, 1.05, "Client Partner", "Strategic relationship & sponsorship")
    node(M + cwL + 0.5, 2.5, cwL, 1.05, "Service Delivery Lead", "Day-to-day AMS operations & SLAs")
    # connectors
    d._vline(s, M + cwL / 2, 3.55, 0.35, THEME["mid_grey"], 1.0)
    d._vline(s, M + cwL + 0.5 + cwL / 2, 3.55, 0.35, THEME["mid_grey"], 1.0)
    d._hline(s, M + cwL / 2, 3.9, cwL + 0.5, THEME["mid_grey"], 1.0)
    d._vline(s, PAGE_W / 2, 3.9, 0.3, THEME["mid_grey"], 1.0)
    # delivery row
    node(M, 4.2, cwL, 1.15, "Onshore · Australia", "Functional / Integration Consultants — Core HCM · Financials · Integrations")
    node(M + cwL + 0.5, 4.2, cwL, 1.15, "Offshore · India", "Functional / Integration Consultants — Core HCM · Financials · Integrations")
    # coral framing banners
    bw = (CW - 0.5) / 2
    d._rect(s, M, 5.65, bw, 1.0, fill_hex=CORAL)
    _, tf = d._box(s, M + 0.28, 5.78, bw - 0.5, 0.4); d._run(d._para(tf, first=True), "THOUGHT LEADERSHIP", 11, WHITE, bold=True, spacing=0.1, caps=True)
    _, tf = d._box(s, M + 0.28, 6.14, bw - 0.5, 0.45); d._run(d._para(tf, first=True, line=1.12), "Strategic advisory · direct CEO access for escalation & sponsorship.", 10.5, WHITE, bold=True, spacing=0)
    d._rect(s, M + bw + 0.5, 5.65, bw, 1.0, fill_hex=INK)
    _, tf = d._box(s, M + bw + 0.78, 5.78, bw - 0.5, 0.4); d._run(d._para(tf, first=True), "DEDICATED WORKDAY PRACTICE", 11, CORAL, bold=True, spacing=0.1, caps=True)
    _, tf = d._box(s, M + bw + 0.78, 6.14, bw - 0.5, 0.45); d._run(d._para(tf, first=True, line=1.12), "Innovation — automation, new Workday features & process optimisation.", 10.5, THEME["rev_body"], bold=True, spacing=0)

def build_standalone():
    d = Deck()
    d.cover_plain("Mivada · AMS & services", "Run, supported,", "improved.",
                  "Application Management Services", "mivada.com")
    fin_hcm(d); engagement_models(d); ams_framework(d); coverage_model(d); governance_model(d); team_structure(d)
    out = "Mivada_AMS_Services_editorial.pptx"; d.save(out)
    print(f"Wrote {out} ({len(d.prs.slides._sldIdLst)} slides)")

if __name__ == "__main__":
    build_standalone()

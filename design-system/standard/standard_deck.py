#!/usr/bin/env python3
# ============================================================================
# Mivada — STANDARD deck. Company Overview + AMS/Services combined into one long
# PowerPoint. This is the house standard; HTML (interactive + A4 portrait/
# landscape) and Word templates mirror this content & style.
#   python standard_deck.py -> Mivada_Standard.pptx
# ============================================================================
import sys, os
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "02-editorial"))
sys.path.insert(0, os.path.join(HERE, "..", "samples"))
from build_pptx import Deck, THEME, PAGE_W
import slide_library as L
from ams_slides import fin_hcm, engagement_models, ams_framework, coverage_model, governance_model, team_structure
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN

CORAL, INK, WHITE = "EA493F", "111111", "FFFFFF"
M = 0.96; CW = PAGE_W - 2 * M
CLIENT_IMG = next((p for p in [os.path.join(HERE, "..", "logos.png"),
                               os.path.join(HERE, "..", "assets", "clients", "clients-logos.png")]
                   if os.path.exists(p)), None)
CLIENTS = ["Qantas", "Jetstar", "Guzman y Gomez", "Western Sydney Airport", "Canva", "NAB",
           "Rio Tinto", "ResMed", "HCF", "IAG", "University of Sydney", "Macquarie University",
           "Nine", "Seven West Media", "Amart", "dnata", "Queensland Airports", "Kennards",
           "Anglicare", "ProPharma"]

d = Deck()

# ---- OVERVIEW -------------------------------------------------------------
L.logo_slide(d)
d.cover_plain("Mivada · Company overview & services · 2026", "Technology,", "human first.",
              "Australian technology consultancy", "mivada.com")
d.kpis("Who we are", "An Australian technology", "consultancy.",
       "We turn Workday, people systems, data and automation into outcomes teams actually feel — "
       "where human understanding becomes system advantage.",
       [("2014", "", "Founded"), ("200", "+", "Specialists · AU, US & India"), ("100", "+", "People-systems clients")])
L.logo_wall(d, "In good", "company.", "Trusted by", names=CLIENTS, image=CLIENT_IMG)
d.pillars("What we do", "Across your", "people platform.", [
    ("01", "People Systems & Transformation", "HCM, Payroll, Time & Attendance and HRIS — implemented and optimised.", "Workday"),
    ("02", "Data, Analytics & BI", "Reporting, dashboards and governed analytics leaders trust.", "Insight"),
    ("03", "Process & Automation", "Redesign, integrations and automation that remove manual work.", "Efficiency"),
    ("04", "Managed Services & Support", "Long-term AMS that keeps delivering value after go-live.", "Partnership")])
d.steps("How we work", "People first,", "start to finish.", [
    ("01", "Listen", "We learn your people, processes and constraints before touching the platform."),
    ("02", "Design", "We shape Workday around how your teams really work."),
    ("03", "Deliver", "Agile delivery, senior practitioners, clear governance."),
    ("04", "Care", "Long-term managed services that keep improving outcomes.")])
d.split("Data & AI", "Decisions on", "current data.",
        "From reporting to governed analytics and automation — we turn Workday data into trusted, timely insight.",
        [("Report", [("Dashboards & reporting people actually use", True)]),
         ("Analyse", [("Prism & data lakehouse, governed analytics", True)]),
         ("Automate", [("Integrations & automation across the stack", True)]),
         ("Assure", [("Security, controls and audit readiness", True)])])

# ---- AMS / SERVICES (shared builders) -------------------------------------
fin_hcm(d)
engagement_models(d)
ams_framework(d)
coverage_model(d)
governance_model(d)
team_structure(d)

# ---- PROOF & CLOSE --------------------------------------------------------
L.cases(d, "Proven Workday", "outcomes.", "Outcomes",
        items=[("Qantas", "A partner since 2014 — taking over Workday support in 2017 and growing into a strategic partner across HR, Payroll, Digital, Automation, Data and Analytics. We scaled integrations, modernised architectures and embedded automation, evolving into a long-term enterprise partner."),
               ("Guzman y Gomez", "Stabilised and optimised Workday post–go-live across modules, reporting, security and compliance. Agile delivery onboarded 150+ franchisees, shipped new reporting and a redesigned security model, and lifted performance across HCM, Absence and Learning."),
               ("Western Sydney Intl Airport", "Delivered a card-management system, then expanded into a strategic role on WSI's Workday program. By resolving functional and integration issues and assuming full AMS ownership, we stabilised HR, Payroll, Finance and Procurement.")],
        sub="Long-term partnerships with some of Australia's most demanding operations.")
d.reasons("Why Mivada", "A partner,", "not a vendor.", [
    ("01", "We grew out of a customer", "So we understand your challenges from the inside — and we're not a Big 4."),
    ("02", "Senior, local, accountable", "Australian-owned, with direct access to our CEO and a dedicated engagement manager."),
    ("03", "In it for the long term", "From go-live into managed services — we grow with you.")])
L.kpi_grid(d, "Proven in", "Workday.", "Capability",
           stats=[("3", "+", "Certifications / consultant"), ("2000", "+", "Workday integrations"), ("3000", "+", "Custom reports"),
                  ("30", "+", "Implementations & support"), ("200", "+", "Employees · AU, US & India"), ("100", "+", "People-systems clients"),
                  ("10", "+", "Years delivering Workday")],
           sub="A decade of Workday delivery — consultants average 3+ years on the platform.")
d.contact_plain("Let's talk", "Let's build something", "human.",
                "hello@mivada.com · mivada.com", "Sydney · Melbourne · AU & India")

OUT = os.path.join(HERE, "Mivada_Standard.pptx")
d.save(OUT)
print(f"Wrote Mivada_Standard.pptx ({len(d.prs.slides._sldIdLst)} slides)  clients_image={'yes' if CLIENT_IMG else 'name-grid'}")

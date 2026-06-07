#!/usr/bin/env python3
# ============================================================================
# Mivada AMS & Services slides — Mivada CONTENT on top of the standard slide
# library (../02-editorial/slide_library.py). Thin wrappers so the standalone
# AMS pack and the merged deck share one implementation of each slide type.
#   python ams_slides.py  -> Mivada_AMS_Services_editorial.pptx  (standalone)
# ============================================================================
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck
import slide_library as L
from slide_library import head, standfirst  # re-export for decks that import them here

CORAL, INK = "EA493F", "111111"

def fin_hcm(d):
    HCM = ["Core HCM", "Absence", "Time Tracking", "Recruiting", "Talent & Performance", "Compensation",
           "Benefits", "Payroll (AU / global)", "Learning", "Security & BP config",
           "Integrations (EIB, CC, Studio)", "Advanced & Composite Reporting"]
    FIN = ["Financial Accounting (GL)", "Accounting Center", "Accounts Payable / Suppliers",
           "Accounts Receivable / Customers", "Procurement", "Expenses", "Banking & Settlement",
           "Business Assets", "Financial Reporting", "Prism Analytics", "Security & BP config",
           "Integrations (EIB, CC, Studio)"]
    L.columns(d, "Full-stack Workday,", "FIN & HCM.", "Consultant coverage",
              cols=[("HCM", HCM), ("FIN", FIN)],
              sub="Full-stack Workday functional and integration expertise across both suites.",
              footer="Adaptive Planning — all modules.  ·  Cross-skilled across both suites — one team, two suites.")

def engagement_models(d):
    # Workday lifecycle (3 stages) with the engagement model wrapped underneath:
    # staff augmentation across Implement -> Optimise; managed services (incl. Managed Payroll) at Manage.
    L.lifecycle(d, "Across the whole", "Workday lifecycle.", "Workday services",
                stages=[("Implement", "Deploy Workday HCM, Payroll & Finance."),
                        ("Optimise", "Phase 2, integrations & reporting."),
                        ("Manage", "Managed Services & Managed Payroll — 24×7, continuous improvement.")],
                wrap=(("Staff augmentation", "Our certified consultants, embedded in your team — through Implement and Optimise."),
                      ("Managed services", "We own the outcome at Manage")),
                footer="Flex between engagement models as you move from build to run.")

def ams_framework(d):
    # AMS framed at the benefit level — built on ITIL: what it is, why it matters, how we apply it.
    L.cards3(d, "AMS, built on", "ITIL.", "AMS framework",
             sub="We run Workday on ITIL — the global standard for IT service management.",
             cards=[("What it is", "The proven framework for running IT services — incident, problem, change and release management, done to a recognised standard."),
                    ("Why it matters", "Predictable, SLA-backed support: fewer incidents, faster resolution, controlled change and continuous improvement."),
                    ("How we apply it", "Governed run operations · 24×7 P1/P2 on-call · change & release control · continuous improvement.")],
             footer="Operational excellence, by design.")

def coverage_model(d):
    AEST = ["8:30", "9:30", "10:30", "11:30", "12:30", "13:30", "14:30", "15:30", "16:30", "17:30", "18:30", "19:30", "20:30", "21:30"]
    IST = ["4:00", "5:00", "6:00", "7:00", "8:00", "9:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00"]
    L.timeline_gantt(d, "Coverage model —", "always on.", "Coverage model", AEST, IST,
                     bars=[("ONSHORE · Australia", 0, 10, CORAL), ("OFFSHORE · India", 4, 14, INK)],
                     overlap=(4, 10), callout=("24×7 P1/P2 on-call", "always-on safety net for P1 & P2 incidents, in line with monthly hours."),
                     descs=[["Governance & stakeholder engagement", "Resolution of P1 / P2 incidents", "App health check & start-of-day triage"],
                            ["Day-to-day incident management & support", "Routine incidents & service requests", "Monitoring & backlog resolution per SLAs"]],
                     sub="Seamless onshore–offshore collaboration for continuous support.")

def governance_model(d):
    L.pyramid_tiers(d, "Governance —", "aligned at every tier.", "Governance",
                    tiers=[("Strategic", "Quarterly", "Executive Sponsor · CIO / IT Exec · Head of HR / Payroll",
                            "Strategic oversight · innovation & thought leadership · roadmap & investment · relationship health", "CEO · Client Partner"),
                           ("Tactical", "Monthly", "Product Owner · IT Service Owner – AMS · Risk & Compliance Lead",
                            "Service & SLA review · risk & compliance · release planning & adoption · improvement backlog", "Client Partner · Service Delivery Lead"),
                           ("Operational", "Weekly / daily", "AMS team · Change & Release Coordinators · Super Users / SMEs",
                            "Daily triage & resolution · change & release execution · UX · operational performance", "Service Delivery Lead · Consultants")],
                    footer="Principles: transparent reporting · escalation by exception · decisions at the right tier · improvement built in.")

def team_structure(d):
    L.org_chart(d, "One integrated team —", "onshore + offshore.", "Team structure",
                leaders=[("Client Partner", "Strategic relationship & sponsorship"),
                         ("Service Delivery Lead", "Day-to-day AMS operations & SLAs")],
                delivery=[("Onshore · Australia", "Functional / Integration Consultants — Core HCM · Financials · Integrations"),
                          ("Offshore · India", "Functional / Integration Consultants — Core HCM · Financials · Integrations")],
                banners=[("Thought leadership", "Strategic advisory · direct CEO access for escalation & sponsorship.", CORAL),
                         ("Dedicated Workday practice", "Innovation — automation, new Workday features & process optimisation.", INK)])

def build_standalone():
    d = Deck()
    d.cover_plain("Mivada · AMS & services", "Run, supported,", "improved.",
                  "Application Management Services", "mivada.com")
    fin_hcm(d); engagement_models(d); ams_framework(d); coverage_model(d); governance_model(d); team_structure(d)
    out = "Mivada_AMS_Services_editorial.pptx"; d.save(out)
    print(f"Wrote {out} ({len(d.prs.slides._sldIdLst)} slides)")

if __name__ == "__main__":
    build_standalone()

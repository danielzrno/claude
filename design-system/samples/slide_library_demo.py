#!/usr/bin/env python3
# ============================================================================
# Editorial slide library — CATALOGUE. One of every standard slide type, with
# sample content, so you can see the building blocks at a glance.
#   python slide_library_demo.py -> Editorial_Slide_Library.pptx
# ============================================================================
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "02-editorial"))
from build_pptx import Deck
import slide_library as L

CORAL, INK = "EA493F", "111111"
d = Deck()

# --- brand / structural ---
L.logo_slide(d)
d.cover_plain("Mivada · Slide library", "The Editorial", "slide library.", "Standard slides", "mivada.com")
L.divider(d, "Section divider", "How we", "deliver.", "A proven methodology, a clear plan, and governance that keeps everyone aligned.")
L.statement(d, "We grew out of a customer.", "So we speak your language.", "Statement",
            sub="Your long-term partner in Workday implementations & managed services.",
            points=[("Australian-owned, senior-led", "Not a Big 4"), ("Direct CEO access", "Reduced escalation matrix"),
                    ("Dedicated engagement manager", "One owner end to end"), ("Onshore + offshore", "Local accountability")],
            footer="TRUSTED BY  ·  Qantas · Guzman y Gomez · Western Sydney International Airport")

# --- overview blocks (Deck methods) ---
d.kpis("Who we are", "An Australian technology", "consultancy.",
       "We turn Workday, people systems, data and automation into outcomes teams actually feel.",
       [("2014", "", "Founded"), ("200", "+", "Specialists"), ("100", "+", "Clients")])
d.pillars("What we do", "Across your", "people platform.",
          [("01", "People Systems & Transformation", "HCM, Payroll, Time & Attendance, HRIS.", "Workday"),
           ("02", "Data, Analytics & BI", "Reporting to governed analytics.", "Insight"),
           ("03", "Process & Automation", "Remove manual, error-prone work.", "Efficiency"),
           ("04", "Managed Services & Support", "AMS that keeps delivering value.", "Partnership")])
d.steps("How we work", "People first,", "start to finish.",
        [("01", "Listen", "Learn your people, processes and constraints."),
         ("02", "Design", "Shape Workday around how teams really work."),
         ("03", "Deliver", "Agile delivery, senior practitioners, clear governance."),
         ("04", "Care", "Long-term managed services that keep improving.")])
d.split("Data & AI", "Decisions on", "current data.",
        "From reporting to governed analytics and automation — Workday data made trusted and timely.",
        [("Report", [("Dashboards & reporting people actually use", True)]),
         ("Analyse", [("Prism & data lakehouse, governed analytics", True)]),
         ("Automate", [("Integrations & automation across the stack", True)]),
         ("Assure", [("Security, controls and audit readiness", True)])])

# --- comparison / lists ---
L.two_cards(d, "Two ways we", "support you.", "Engagement models",
            cards=[("Staff augmentation", "Certified consultants, embedded in your team",
                    ["You direct the work; certified Workday FIN & HCM consultants", "Scale up or down by day, week or sprint",
                     "Onshore (AU), offshore (India) or blended", "Time-and-materials or a block of hours"]),
                   ("Managed services (AMS)", "We own the outcome, end to end",
                    ["We own incident → release management", "SLA-backed and governed, 24×7 P1/P2 on-call",
                     "Continuous improvement & bi-annual releases", "Fixed monthly hours, carry-forward flexibility"])],
            sub="Flex between them as your needs change.",
            footer="Many clients start with augmentation and move into AMS as their environment stabilises.")
L.columns(d, "Full-stack Workday,", "FIN & HCM.", "Coverage",
          cols=[("HCM", ["Core HCM", "Absence", "Time Tracking", "Recruiting", "Talent & Performance", "Compensation",
                         "Benefits", "Payroll (AU / global)", "Learning", "Security & BP config", "Integrations", "Reporting"]),
                ("FIN", ["Financial Accounting (GL)", "Accounting Center", "AP / Suppliers", "AR / Customers", "Procurement",
                         "Expenses", "Banking & Settlement", "Business Assets", "Financial Reporting", "Prism Analytics", "Security & BP config", "Integrations"])],
          sub="Full-stack functional and integration expertise across both suites.",
          footer="Adaptive Planning — all modules.")

# --- diagram slides ---
L.framework_stack(d, "AMS operating", "framework.", "Framework",
                  top_label="Continuous improvement", gov_label="Governance",
                  gov_items="Thought leadership · Product & vendor relationship · Technical-debt reduction · WD bi-annual release",
                  tiles=["Incident management", "Problem management", "Change management", "Release management",
                         "Service requests", "Change Advisory Board", "SLAs", "Metrics",
                         "Integration health checks", "Backlog management", "Prioritisation forum", "Capacity management",
                         "Enhancements", "Knowledge base", "Cyber security", "BP & security config"],
                  mid_label="Service integration · ITSM",
                  foundations=["Operations as a priority", "Business value realisation", "Operational excellence · ITIL"],
                  sub="Driving operational excellence through governance, integration and continuous improvement.")
AEST = ["8:30", "9:30", "10:30", "11:30", "12:30", "13:30", "14:30", "15:30", "16:30", "17:30", "18:30", "19:30", "20:30", "21:30"]
IST = ["4:00", "5:00", "6:00", "7:00", "8:00", "9:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00"]
L.timeline_gantt(d, "Coverage model —", "always on.", "Coverage", AEST, IST,
                 bars=[("ONSHORE · Australia", 0, 10, CORAL), ("OFFSHORE · India", 4, 14, INK)],
                 overlap=(4, 10), callout=("24×7 P1/P2 on-call", "always-on safety net for P1 & P2 incidents."),
                 descs=[["Governance & stakeholder engagement", "Resolution of P1 / P2 incidents", "Start-of-day health check & triage"],
                        ["Day-to-day incident management & support", "Routine incidents & service requests", "Monitoring & backlog resolution per SLAs"]],
                 sub="Seamless onshore–offshore collaboration for continuous support.")
L.pyramid_tiers(d, "Governance —", "aligned at every tier.", "Governance",
                tiers=[("Strategic", "Quarterly", "Executive Sponsor · CIO / IT Exec · Head of HR / Payroll",
                        "Strategic oversight · innovation · roadmap & investment · relationship health", "CEO · Client Partner"),
                       ("Tactical", "Monthly", "Product Owner · IT Service Owner – AMS · Risk & Compliance Lead",
                        "Service & SLA review · risk & compliance · release planning · improvement backlog", "Client Partner · Service Delivery Lead"),
                       ("Operational", "Weekly / daily", "AMS team · Change & Release Coordinators · Super Users / SMEs",
                        "Daily triage & resolution · change & release execution · UX · performance", "Service Delivery Lead · Consultants")],
                footer="Principles: transparent reporting · escalation by exception · decisions at the right tier · improvement built in.")
L.org_chart(d, "One integrated team —", "onshore + offshore.", "Team",
            leaders=[("Client Partner", "Strategic relationship & sponsorship"), ("Service Delivery Lead", "Day-to-day AMS operations & SLAs")],
            delivery=[("Onshore · Australia", "Functional / Integration Consultants — HCM · Financials · Integrations"),
                      ("Offshore · India", "Functional / Integration Consultants — HCM · Financials · Integrations")],
            banners=[("Thought leadership", "Strategic advisory · direct CEO access for escalation & sponsorship.", CORAL),
                     ("Dedicated Workday practice", "Innovation — automation, new features & process optimisation.", INK)])

# --- proof / social proof ---
L.logo_wall(d, "In good", "company.", "Trusted by",
            names=["Qantas", "Jetstar", "Guzman y Gomez", "Western Sydney Airport", "Canva", "NAB", "Rio Tinto", "ResMed",
                   "HCF", "IAG", "University of Sydney", "Macquarie University", "Nine", "Seven West Media", "Amart", "dnata",
                   "Queensland Airports", "Kennards", "Anglicare", "ProPharma"],
            image=("../logos.png" if os.path.exists("../logos.png") else None))
L.cases(d, "Proven Workday", "outcomes.", "Outcomes",
        items=[("Qantas", "A partner since 2014 — from Workday support into a strategic partner across HR, Payroll, Data and Automation."),
               ("Guzman y Gomez", "Stabilised and optimised Workday post–go-live; onboarded 150+ franchisees with new reporting and security."),
               ("Western Sydney Intl Airport", "Assumed full AMS ownership, stabilising HR, Payroll, Finance and Procurement.")],
        sub="Long-term partnerships with some of Australia's most demanding operations.")
L.kpi_grid(d, "Proven in", "Workday.", "Capability",
           stats=[("3", "+", "Certifications / consultant"), ("2000", "+", "Workday integrations"), ("3000", "+", "Custom reports"),
                  ("30", "+", "Implementations & support"), ("200", "+", "Employees · AU, US & India"), ("100", "+", "People-systems clients"),
                  ("10", "+", "Years delivering Workday")],
           sub="A decade of Workday delivery — consultants average 3+ years on the platform.")
d.pullquote("Outcomes", "They delivered in weeks what we'd scoped for a year — and our team actually uses it.",
            "Enterprise Workday client", [("30", "+", "Implementations"), ("2000", "+", "Integrations"), ("3000", "+", "Reports")])
d.reasons("Why Mivada", "A partner,", "not a vendor.",
          [("01", "We grew out of a customer", "Inside-out understanding — and we're not a Big 4."),
           ("02", "Senior, local, accountable", "Direct CEO access and a dedicated engagement manager."),
           ("03", "In it for the long term", "From go-live into managed services — we grow with you.")])
try:
    d.chart_coral("Growth", "Engagements", "by year.", ["2019", "2020", "2021", "2022", "2023"], [("Engagements", [12, 18, 24, 30, 38])], "Chart")
except Exception as e:
    print("chart_coral skipped:", e)
d.contact_plain("Let's talk", "Let's build something", "human.", "hello@mivada.com · mivada.com", "Sydney · Melbourne · AU & India")

OUT = "Editorial_Slide_Library.pptx"
d.save(OUT)
print(f"Wrote {OUT} ({len(d.prs.slides._sldIdLst)} slides)")

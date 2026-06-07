# Mivada — the standard

The **house standard**: Company Overview + AMS / Services in one combined deck, with matching
documents. This is the canonical style — tweak content here, then we point the `/mivada` skill at it.

## Source of truth

`Mivada_Standard.pptx` is now a **hand-curated master** (21 slides) — edited directly in PowerPoint,
on the shared Editorial look (black canvas, coral accent, Inter). It is no longer regenerated from a
script; edit the `.pptx` itself.

> `standard_deck.py` is the **legacy generator** that seeded the first version. It has **diverged** from
> the curated deck above — don't run it expecting to reproduce this deck. Kept for reference/history.

## Formats

| File | Format | Status |
|------|--------|--------|
| `Mivada_Standard.pptx` | **PowerPoint** — the long combined deck (21 slides) | **Curated master** |
| `presentation.html` | Interactive HTML deck (16:9, ←/→ nav, print 1 slide/page) | Pending re-sync to master |
| `standard-a4-portrait.html` | HTML document — A4 portrait | Pending re-sync to master |
| `standard-a4-landscape.html` | HTML document — A4 landscape | Pending re-sync to master |
| `build_docx.js` → `Mivada_Standard.docx` | Word (native .docx) | Pending re-sync to master |

The HTML and Word formats still carry the earlier content; bring them in line with the curated PPTX
before using them as a set.

## Story (21 slides)

1. Logo open · 2. Cover — *Technology, human first.* · 3. Who we are — *ANZ's largest locally owned
Workday partner.* · 4. Capability — *Proven in Workday.* · 5. Trusted by — *In good company.* ·
6. What we do — *Across your ERP platform.* · 7. Engagement model — *Across the whole Workday lifecycle.* ·
8. Engagement models — *Four ways we support you.* · 9. Delivery approach — *Velocity & Value — Workday GO.* ·
10. Implementation timeline — *Your implementation timeline.* · 11. Delivery approach — *The Data Handshake:
ensuring certainty.* · 12. The one team — *Your delivery team.* · 13. Getting payroll right — *Payroll testing
rigour.* · 14. Consultant coverage — *Full-stack Workday, FIN & HCM.* · 15. AMS framework — *AMS, built on
ITIL.* · 16. AMS framework — *AMS operating framework.* · 17. Governance · 18. Team structure · 19. Outcomes
— *Proven Workday outcomes.* · 20. Why Mivada — *A partner, not a vendor.* · 21. Let's talk.

### Slide 7 — the Workday lifecycle

A four-column grid on a consistent gutter: **ADVISE › IMPLEMENT › OPTIMISE › MANAGE** across the top,
with the engagement bands aligned beneath the phases they cover — **Advisory** (under Advise),
**Lead Implementation Partner** + **Staff Augmentation** (spanning Implement/Optimise), and
**Managed Services** (under Manage). The two tall bookend cards frame the stacked middle bands.

### Slides 9–14 — the delivery cluster (light slides)

Re-skinned from project proposals into the deck's light-slide language (off-white canvas, white cards,
coral accents, Inter) and made **generic** (no customer name).

- **9 · Velocity & Value — Workday GO** — the 90/10 split (pre-configured vs unique to you) over the
  five-phase delivery model flow. Intro prose and per-phase descriptions dropped to keep it light.
- **10 · Your implementation timeline** — a 20-week Workday GO Gantt (Mobilise → Architect → Configure
  → Test → GO-LIVE → Hypercare), phase bars in coral, the milestone in black.
- **11 · The Data Handshake** — a two-card "handshake" (Your team ↔ Mivada) over three load-cycle cards
  (First load · Second load · Gold build) for iterative, no-"big-bang" validation.
- **12 · Your delivery team** — two org charts as top node + role chips (Mivada: Client Partner →
  Engagement Manager · Functional Lead · Functional Consultants · Data & Integration; Customer: Executive
  Sponsor → Project Manager · Change Manager · Functional Leads & SMEs · Testing Lead · Data Champions),
  with a note on shared/combined roles beneath.
- **13 · Payroll testing rigour** — the dense proposal handout distilled to a slide: three highlight
  cards (2+ parallel cycles · line-level variance · $0 net-pay target) over the parallel-testing cycle
  flow, with an Australian-complexity footer band.
- **14 · Full-stack Workday, FIN & HCM** — coverage as a boxed module grid: two suite panels (HCM, FIN),
  twelve module chips each, replacing the earlier dashed lists.

## Notes

- Presentations are **16:9**; documents are **A4** (portrait + landscape). Same brand, different canvas.
- The **client wall** (Trusted by) uses a name grid until a real logo wall image is dropped in.
- Figures are illustrative samples — replace with verified numbers before real use.

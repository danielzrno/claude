# Mivada — the standard

The **house standard**: Company Overview + AMS / Services in one combined deck, with matching
documents. This is the canonical style — tweak content here, then we point the `/mivada` skill at it.

## Source of truth

`Mivada_Standard.pptx` is now a **hand-curated master** (27 slides) — edited directly in PowerPoint,
on the shared Editorial look (black canvas, coral accent, Inter). It is no longer regenerated from a
script; edit the `.pptx` itself.

> `standard_deck.py` is the **legacy generator** that seeded the first version. It has **diverged** from
> the curated deck above — don't run it expecting to reproduce this deck. Kept for reference/history.

## Formats

| File | Format | Status |
|------|--------|--------|
| `Mivada_Standard.pptx` | **PowerPoint** — the long combined deck (27 slides) | **Curated master** |
| `presentation.html` | Interactive HTML deck (16:9, ←/→ nav, print 1 slide/page) | Pending re-sync to master |
| `standard-a4-portrait.html` | HTML document — A4 portrait | Pending re-sync to master |
| `standard-a4-landscape.html` | HTML document — A4 landscape | Pending re-sync to master |
| `build_docx.js` → `Mivada_Standard.docx` | Word (native .docx) | Pending re-sync to master |

The HTML and Word formats still carry the earlier content; bring them in line with the curated PPTX
before using them as a set.

### Light & dark — every slide in both

The curated master alternates light and dark slides. Two **single-theme** builds make every slide
available on one background:

| File | Every slide on… |
|------|-----------------|
| `Mivada_Standard_Light.pptx` | off-white (`#FAFAFA`) |
| `Mivada_Standard_Dark.pptx` | black (`#000000`) |

They're generated from the master by `theme_deck.py`, a context-aware recolour: card fills, borders,
connector lines and **text colours flip to suit the background** (text colour is chosen from the
luminance of the shape behind it, so it works both directions). Coral stays coral. Logos swap per theme
(coral/white M-icon; on-white vs on-black wordmark) and the governance pyramid gets a white backing
panel on dark. Regenerate after editing the master:

```bash
python theme_deck.py Mivada_Standard.pptx light Mivada_Standard_Light.pptx
python theme_deck.py Mivada_Standard.pptx dark  Mivada_Standard_Dark.pptx
```

## Story (27 slides)

1. Logo open · 2. Cover — *Technology, human first.* · 3. Who we are · 4. Capability — *Proven in
Workday.* · 5. Trusted by · 6. What we do · 7. Engagement model — *the Workday lifecycle* · 8. Engagement
models · 9. *Velocity & Value — Workday GO* · 10. *Your implementation timeline* · 11. *The Data Handshake*
· 12. *Your delivery team* · 13. *Payroll testing rigour* · 14. *Full-stack Workday, FIN & HCM* · 15. AMS,
built on ITIL · 16. AMS operating framework · **17. Coverage model** · **18. Governance model** ·
**19. Transition approach** · **20. Service levels (SLAs)** · **21. Commercials** · **22. Augmentation
capability** · 23. Governance (pyramid) · 24. Team structure · 25. Outcomes · 26. Why Mivada · 27. Let's talk.

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

### Slides 17–22 — the AMS deep-dive (light slides)

Rebuilt from an AMS proposal into the standard's light language (off-white, white cards, coral/black,
Inter), generic, and **box-based rather than native tables** so the light/dark recolour works cleanly.

- **17 · Coverage model** — onshore (AU) / offshore (IN) coverage with a handover bar and a 24×7 on-call band.
- **18 · Governance model** — a Strategic / Tactical / Operational matrix (customer · activities · Mivada · cadence).
- **19 · Transition approach** — a four-week transition (week cards + milestone chips).
- **20 · Service levels** — the P1–P4 SLA grid (definition · acknowledge · target resolution).
- **21 · Commercials** — pricing stat cards, a blended-rate band, and key considerations (illustrative figures).
- **22 · Augmentation capability** — FIN/HCM availability cards + a bench grid, kept as fillable `[placeholders]`.

> Overlap to resolve: the new **Governance model (18)** covers similar ground to the existing **Governance
> pyramid (23)**, and **Coverage model (17)** overlaps **Team structure (24)**. Kept both for now — say the
> word and I'll drop the older ones.

## Notes

- Presentations are **16:9**; documents are **A4** (portrait + landscape). Same brand, different canvas.
- The **client wall** (Trusted by) uses a name grid until a real logo wall image is dropped in.
- Figures are illustrative samples — replace with verified numbers before real use.

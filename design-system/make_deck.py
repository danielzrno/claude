#!/usr/bin/env python3
# ============================================================================
# make_deck — list the Editorial standard slide library and scaffold a new deck.
#   python make_deck.py list            # show every standard slide + its call
#   python make_deck.py new <name>      # write <name>.py (edit, then run it)
# ============================================================================
import sys, os

HERE = os.path.dirname(os.path.abspath(__file__))
ED = os.path.join(HERE, "02-editorial")

CATALOGUE = [
    ("logo slide",        "L.logo_slide(d, pre, accent)"),
    ("cover",             "d.cover_plain(eyebrow, pre, accent, footL, footR)"),
    ("section divider",   "L.divider(d, eyebrow, pre, accent, sub)"),
    ("statement",         "L.statement(d, pre, accent, tag, sub, points, footer, dark=True)"),
    ("who we are (stats)","d.kpis(eyebrow, pre, accent, standfirst, stats)"),
    ("pillars",           "d.pillars(eyebrow, pre, accent, items)"),
    ("steps",             "d.steps(eyebrow, pre, accent, steps)"),
    ("two-column list",   "d.split(eyebrow, pre, accent, standfirst, caps)"),
    ("two cards",         "L.two_cards(d, pre, accent, tag, cards, sub, footer, dark=True)"),
    ("columns",           "L.columns(d, pre, accent, tag, cols, sub, footer)"),
    ("framework stack",   "L.framework_stack(d, pre, accent, tag, top_label, gov_label, gov_items, tiles, mid_label, foundations)"),
    ("coverage gantt",    "L.timeline_gantt(d, pre, accent, tag, hours_top, hours_bottom, bars, overlap, callout, descs)"),
    ("governance pyramid","L.pyramid_tiers(d, pre, accent, tag, tiers, footer)"),
    ("org chart",         "L.org_chart(d, pre, accent, tag, leaders, delivery, banners)"),
    ("logo / client wall","L.logo_wall(d, pre, accent, tag, names, image)"),
    ("case studies",      "L.cases(d, pre, accent, tag, items, sub)"),
    ("KPI grid",          "L.kpi_grid(d, pre, accent, tag, stats, sub)"),
    ("pull-quote",        "d.pullquote(eyebrow, quote, attr, stats)"),
    ("reasons",           "d.reasons(eyebrow, pre, accent, items)"),
    ("chart",             "d.chart_coral(eyebrow, pre, accent, categories, series, slug)"),
    ("contact",           "d.contact_plain(eyebrow, pre, accent, footL, footR)"),
]

SCAFFOLD = '''#!/usr/bin/env python3
# {name} — Mivada Editorial deck. Edit the content, then:  python {name}.py
import sys
sys.path.insert(0, {ed!r})
from build_pptx import Deck
import slide_library as L

d = Deck()

L.logo_slide(d)
d.cover_plain("{name} · 2026", "Technology,", "human first.", "Mivada", "mivada.com")

d.kpis("Who we are", "An Australian technology", "consultancy.",
       "We turn Workday, people systems, data and automation into outcomes teams actually feel.",
       [("2014", "", "Founded"), ("200", "+", "Specialists"), ("100", "+", "Clients")])

# Add standard slides here. Run `python make_deck.py list` to see every option, e.g.:
# L.statement(d, "We grew out of a customer.", "So we get it.", "Why us",
#             points=[("Australian-owned","Not a Big 4"), ("Direct CEO access","Fast decisions")])
# L.framework_stack(d, "AMS operating", "framework.", "Framework",
#                   top_label="Continuous improvement", gov_label="Governance", gov_items="...",
#                   tiles=["Incident", "Problem", "Change", "Release"], mid_label="Service integration",
#                   foundations=["Operations", "Value", "Excellence"])

d.contact_plain("Let's talk", "Let's build something", "human.",
                "hello@mivada.com · mivada.com", "Sydney · Melbourne · AU & India")

d.save("{name}.pptx")
print("wrote {name}.pptx")
'''


def do_list():
    print("\nMivada — Editorial standard slide library\n" + "=" * 52)
    print("Compose a deck from these (d = Deck(), import slide_library as L):\n")
    for label, call in CATALOGUE:
        print(f"  {label:<22} {call}")
    print("\nScaffold a new deck:  python make_deck.py new <name>")
    print("Catalogue decks:      samples/slide_library_demo.py (PPTX)  ·  02-editorial/slide-library.html (HTML)\n")


def do_new(name):
    name = name.replace(" ", "_")
    path = os.path.abspath(f"{name}.py")
    if os.path.exists(path):
        print(f"refusing to overwrite {path}"); return
    with open(path, "w") as f:
        f.write(SCAFFOLD.format(name=name, ed=ED))
    print(f"wrote {path}\n  edit the content, then:  python {name}.py   ->  {name}.pptx")


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "new" and len(args) > 1:
        do_new(args[1])
    else:
        do_list()

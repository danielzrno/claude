# Mivada logos — drop-in assets + placement rules

Put the official Mivada logo files in this folder using the names below. Every template
(HTML / PowerPoint / Word) reads from here, picks the **right variant for the background**,
and falls back to the CSS/Pillow **"M" chip** when a file is missing — so things work before
the real art arrives and auto-upgrade the moment you add it.

> If your brand kit uses different filenames, that's fine — just tell me the mapping and I'll
> wire it; or rename to match these.

## Expected files

Vector (**SVG**) is preferred for HTML; transparent **PNG** (≥ 1000px wide / 4× the placed size)
is needed for PowerPoint & Word embedding. Provide both where you can.

| File | What it is | For |
|------|-----------|-----|
| `mivada-wordmark-white.svg` / `.png` | Full wordmark, **white** | dark / black / coral backgrounds |
| `mivada-wordmark-coral.svg` / `.png` | Full wordmark, **coral `#EA493F`** | light / off-white backgrounds (primary) |
| `mivada-wordmark-black.svg` / `.png` | Full wordmark, **black/ink** | light backgrounds where coral would clash |
| `mivada-mark-white.svg` / `.png` | The **"M" mark** only, white | dark/coral, tight spaces (slugs, footers) |
| `mivada-mark-coral.svg` / `.png` | The **"M" mark** only, coral | light, tight spaces / card chips |
| `mivada-mark-black.svg` / `.png` | The **"M" mark** only, black | light, mono contexts |

(Transparent backgrounds, please. Keep some built-in clear space; templates won't crop.)

## Placement matrix — Editorial (concept 02)

| Surface | Background | Logo used |
|---------|-----------|-----------|
| Deck **cover** | black | **white wordmark**, top-left |
| Deck **closing / contact** | black | **white wordmark** |
| Deck content slides | off-white | **coral M mark** in the slug/footer (small) |
| Deck content slides | black (sandwich) | **white M mark** in the slug/footer |
| **Document masthead** | off-white | **coral wordmark** (or black if coral over-saturates beside other coral) |
| Document **footer** | off-white | **coral M mark** (small) + contact text |
| Any **coral field** | coral | **white wordmark** / white mark |

**Rules of thumb:** dark or coral background → **white** variant. Light background → **coral**
(primary), or **black** when coral sits next to other coral. Use the **wordmark** for covers and
mastheads; use the **mark** for tight spots (slide slugs, footers, card chips). Never place the
coral logo on coral or the white logo on white. Respect the logo's own clear space; don't stretch —
templates scale by height and preserve aspect ratio.

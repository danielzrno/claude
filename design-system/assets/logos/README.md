# Mivada logos — the kit + how it's wired

The official Mivada logo files (provided by the brand). Templates read from this folder and pick
the **right variant for the background**; if a file is missing they fall back to the CSS/Pillow
**"M" chip**. All files are RGB PNGs with transparent backgrounds.

## The kit

**Horizontal wordmark** ("mıvada", 1002×152):

| File | Composition | Use on |
|------|-------------|--------|
| `Mivada_Logo_Master_RGB_L.png` | coral M + ink text (`#071621`) | **light** backgrounds (primary) |
| `Mivada_Logo_Black_RGB_L.png` | all black | light, mono |
| `Mivada_Logo_White_RGB_L.png` | all white | dark / coral, mono |
| `Mivada_Logo_2C_OnBlack_RGB_L.png` | coral M + white text | **black** backgrounds |
| `Mivada_Logo_2C_OnRed_RGB_L.png` | ink M + white text | **coral / red** backgrounds |

**Vertical lockup** (M stacked over wordmark, 1018×880): `Mivada_Logo_Vertical_{Master… }` —
`_2C_OnBlack`, `_2C_OnRed`, `_Black`, `_Melon`, `_White`. Use where height is available (covers, title cards).

**Icon / "M" mark only** (1018×602):

| File | Colour | Use on |
|------|--------|--------|
| `Mivada_Icon_Melon_RGB_L.png` | coral (Melon) | light backgrounds, slugs, footers, chips |
| `Mivada_Icon_White_RGB_L.png` | white | dark / coral backgrounds |
| `Mivada_Icon_Black_RGB_L.png` | black | light, mono |
| `Mivada_Icon_Contrast_RGB_L.png` | dark ink (`#071621`) | light, contrast |

## Rule of thumb

Dark/coral background → **white** (or the `2C_On…`) variant. Light background → **Master** (coral),
or **Black** beside other coral. **Wordmark** for covers/mastheads; **icon** for tight spots
(slide slugs, footers, card chips). Never coral-on-coral or white-on-white; preserve aspect ratio.

## How the Editorial templates use it (concept 02)

| Surface | Background | Logo |
|---------|-----------|------|
| Deck cover & closing | black | `Mivada_Logo_2C_OnBlack` (wordmark, top-left) |
| Deck content slides | off-white | `Mivada_Icon_Melon` (small M, top-left) |
| Deck content slides | black (sandwich) | `Mivada_Icon_White` (small M, top-left) |
| Document masthead | off-white | `Mivada_Logo_Master` (wordmark) |
| Document footer | off-white | `Mivada_Icon_Melon` (small M) |

Wired in `02-editorial/`: `build_pptx.py` (`_logo()` helper), `presentation.html` (`.logo` img +
content-slide `::before` mark), `document.html` (masthead/footer), `build_docx.js` (`ImageRun`).

> **Colour note:** the logo's coral ("Melon") samples as `#EA5454`; the design-system accent is
> currently `#EA493F` (sampled from the site's JPEG). They're very close. Say the word and I'll
> standardise the accent to the official `#EA5454` across the system.

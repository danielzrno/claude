# The "Doesn't Look AI-Generated" Checklist

Every Mivada artifact is checked against this list. These are the patterns that make
a deck, doc, or page read as "a language model made this." Avoid the tell, apply the fix.

The rule of thumb: **default settings are the tell.** Anything you get for free —
the default font, the default gray, the centered hero, the emoji bullet — is what
everyone else's AI output also got for free. Make a deliberate choice instead.

---

## Typography

| Tell | Fix |
|------|-----|
| `Inter`, `Roboto`, `Open Sans`, `Arial`, `system-ui` as the *brand* face | Use a display face with a point of view + a distinct text face. Each concept here names one. |
| **`Poppins` + `Lora`** pairing | This is the literal Anthropic brand-guidelines pairing — an instant giveaway. Never use it. |
| One weight everywhere | Use weight *contrast* (e.g. 300 standfirst vs 800 headline) and let it carry hierarchy. |
| Flat all-caps with no tracking | Caps need `letter-spacing: 0.08–0.16em`. Lowercase labels need none. |
| Gradient-filled headline text | Solid ink. Color the *one* word that matters, not the whole line. |

## Colour

| Tell | Fix |
|------|-----|
| Purple→indigo / purple→pink gradients (esp. on white) | The #1 AI-slop signature. Banned. |
| `#faf9f5` bone + `#d97757` coral-orange + `#6a9bcc` blue | This *is* the Anthropic palette. Don't reproduce it. (Our warm concept uses sage + amber + espresso, deliberately not this.) |
| Equal-weight rainbow ("one of each colour") | Enforce 60 / 30 / 10 dominance: one colour rules, one supports, one sharp accent. |
| Pure `#000` on pure `#fff`; default Tailwind grays (`#6B7280`) | Tune the ink and paper (warm or cool near-black/near-white) and a real secondary text tone. |
| The same soft drop shadow on every card | Pick a shadow language (hard offset, or none, or one tuned soft) and commit. |

## Layout & composition

| Tell | Fix |
|------|-----|
| Centered hero: gradient headline + two pill buttons | Asymmetry. Left-anchor. Let the grid show. Use the margin. |
| 3 identical feature cards: rounded box + soft shadow + circle icon on top | Vary block sizes; use rules/numbers/columns instead of three twins. |
| Everything evenly spaced, no tension | Dominant element + quiet support. Generous negative space *or* deliberate density — not mush. |
| Uniform 16px radius on everything | Choose a radius language per concept (sharp 0–2px, or one considered curve) and hold it. |

## Icons, emoji, ornament

| Tell | Fix |
|------|-----|
| Emoji as bullets / section markers (🚀 ✨ 📊 ✅) | No emoji in professional artifacts. Use numerals, rules, or a single custom motif. |
| Lucide/Heroicons line-icon in a colored circle, repeated | Use sparingly and restyled, or replace with type/number-led hierarchy. |
| Glassmorphism / blurred translucent cards | Solid surfaces with intentional contrast. |

## Motion (HTML)

| Tell | Fix |
|------|-----|
| Generic `fade-up` on every element | One choreographed load (staggered reveal) beats scattered micro-animations. Keep it CSS-only and subtle. |

## Slides specifically (PowerPoint / HTML deck)

| Tell | Fix |
|------|-----|
| **Accent line directly under the title** | Hallmark of AI slides. Use whitespace or a background field instead — never the underline. |
| Full-width colored header/footer bars, side ribbons, stripes | Reads as slop. Only use a band if it's a deliberate, repeated motif of the system. |
| Cream/beige *default* backgrounds (`#F5F5DC`, `#FAF0E6`…) | Use white or dark, or a *named brand* surface — never the accidental warm-neutral default. |
| Centered body text | Left-align all paragraphs and lists. Center only short titles/cover lines. |
| Identical layout on every slide; text-only slides | Vary: two-column, stat callout, quote, grid. Every slide earns one visual element. |

## Copy / voice

| Tell | Fix |
|------|-----|
| "In today's fast-paced world…", "Let's dive in", "unlock", "elevate", "seamless", "robust", "leverage", "supercharge" | Cut them. Say the specific thing. |
| Em-dash every other sentence | Use them rarely. Prefer a period. |
| Everything in threes ("fast, simple, and powerful") | Vary list length. Sometimes one strong claim beats three weak ones. |
| Title Case Headlines On Everything | Sentence case reads human. Reserve caps for short tracked labels. |
| Exclamation marks; hype adjectives | Confidence is quiet. State the result, attach a number. |

---

**Final gate before shipping any artifact:** *If I swapped this company's name out, would
this look like it was generated for anyone? If yes, it isn't finished.*

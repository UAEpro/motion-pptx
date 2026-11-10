# Design reference

The kit supplies mechanics (layout units, animation, RTL, checks). The look
is yours to decide for every deck: palette, background, typography, motif
and pacing should come from the subject, the audience and any brand the
user has — never carried over from a previous deck or from `example.js`.

## 1. Decide a design direction first

Before writing code, write down (in your head or a comment at the top of
the build script) a short brief:

- **Mood** in three words, from the content: e.g. "calm, precise, trusted"
  for a security review; "bold, fast, playful" for a product launch;
  "warm, human, hopeful" for a community programme.
- **Light or dark.** Dark suits cinematic, technical, evening keynotes and
  makes glows read well. Light suits editorial, data-heavy, print-adjacent,
  daytime rooms with bright projectors. Both work with every kit effect.
- **Brand.** If the user has brand colours, fonts or a logo, use them and
  build the rest of the palette around them. Ask if it's unclear and it
  matters; otherwise choose.
- **Motif**: one repeated visual idea that fits the subject — icons in
  circles, numbered markers, outlined cards, photo frames, a diagonal cut,
  connector lines. One motif, used everywhere.

Test the palette: if you could swap it into an unrelated deck and it would
still "work", it is too generic; push it toward this topic.

## 2. Build the palette as roles

Pass these to `new Deck({ palette })` (hex, no `#`):

| Role | Use | Notes |
|---|---|---|
| `bg` | slide background colour (also under the backdrop image) | carries ~60–70% of the visual weight |
| `ink` | headings, key text | contrast with `bg` ≥ 7:1 |
| `muted` | body text, captions, connectors | contrast with `bg` ≥ 4.5:1 |
| `accent` | the one colour that means "look here" | badges, highlights, active states |
| `accent2` | optional second accent | use sparingly, or for a second category |
| `onAccent` | text/icons placed on accent fills | must contrast with both accents |
| `surface`, `surfaceAlpha` | card fill and its transparency % | glass: ink at ~94%; solid: white/near-bg at 0 |
| `border`, `borderAlpha` | card outline | faint on glass; 0 alpha for crisp light cards |
| others | anything else (e.g. `danger`, `success`) | name by meaning, not by hue |

Dominance over equality: one hue dominates, one accent is sharp, everything
else is neutral. Reserve warning colours (reds) for meaning (risk, isolation,
errors), not decoration. Check text contrast; low-contrast muted text is the
most common failure on both light and dark themes.

## 3. Background

Options, pick what fits the direction:
- flat `bg` colour (cleanest; fine for light editorial decks);
- `scripts/backdrop.js`: radial gradient + optional grid + glows, any colours
  (`--base`, `--base2`, `--glow`, `--glow2`, `--grid`, `--focus`);
- your own SVG rasterised with sharp (diagonal shapes, stripes, dot fields,
  paper grain, duotone photo) — anything a gradient or texture needs, since
  pptxgenjs fills can't do gradients or blur;
- a variant per section (e.g. a different glow colour on divider slides).
Keep backgrounds quiet; content must stay the brightest/most contrasted thing.

## 4. Typography
- Use fonts installed on the presenter's machine; web fonts get substituted.
  Safe with Office: Arial, Calibri, Segoe UI (Windows), Cambria, Georgia;
  with Arabic coverage: Dubai (Windows 10+/Office), Arial, Tahoma, Segoe UI.
- Scale on the 1920 px canvas (px; ÷2 = pt): hero 120–170, slide title
  64–80, section header 40–48, body 24–34, captions ≥ 22.
- Title alignment follows reading direction; body text is never centred
  except inside small cards or numbers.

## 5. Layout grid (1920 × 1080 px)
- 1 px = 1/144 in. Margins ≥ 110 px; title band y≈60–220; content from
  y≈300; footer at y≈1012 if used.
- Gaps 24–34 px between siblings; padding 30–34 px inside cards; radius to
  taste (0 for sharp editorial, 22–28 for soft).
- Vary layouts between slides: grid of cards, flow diagram, big stat, split
  text/visual, timeline. One message per slide.
- RTL: mirror x so the first item sits at the right edge.

## 6. Motion grammar (30 fps frames)
- Frame 0: title (word by word suits short titles; plain float suits long ones).
- +8: subtitle. +14–30: primary content, siblings staggered 6–14 frames.
- Connectors draw after both ends exist; travelling dots begin after the line.
- Emphasis (quote, note, callout) last, around 80–110.
- Keep each slide's motion under ~4 s, and match energy to the mood: calm
  decks use fade/float with longer durations; energetic decks use zoom,
  wipe and tighter staggers.

## 7. RTL / Arabic
- `new Deck({ rtl: true })` sets presentation RTL, paragraph `rtl="1"`,
  right alignment and `ar-AE` language.
- pptxgenjs reads paragraph direction from the first run's options: build
  every run of a mixed paragraph with `s.run()` / `s.bold()`.
- Latin tokens mid-sentence are fine; at the start or end of an Arabic line,
  especially with punctuation (`.env`, `v0.1.x`, `(Linux)`), bidi reordering
  can misplace them — reword.
- Stand-alone numbers (stats, step numbers): pass `ltr: true`.

## 8. QA rendering
LibreOffice renders the final layout (no animation). Arial and Calibri map
to metric-compatible Liberation Sans / Carlito, so their previews are
accurate. If the deck font is missing in the sandbox, alias it to a
metrically similar font so text widths are trustworthy — e.g. for Dubai use
IBM Plex Sans Arabic (from `@fontsource/ibm-plex-sans-arabic`, woff2 → ttf
with fontTools; merge the `arabic` and `latin` subsets, or digits render
spaced and grey):

```xml
<!-- ~/.config/fontconfig/fonts.conf -->
<fontconfig><alias binding="same"><family>Dubai</family>
<prefer><family>IBM Plex Sans Arabic</family></prefer></alias></fontconfig>
```
Then `fc-cache -f`. When you used a stand-in like this, tell the user the
preview font was a substitute.

## Worked direction (for reference only)
The deck this skill was derived from used: dark navy base with a teal
highlight glow and faint grid, glass cards, icons in glowing circles, teal
primary + gold secondary accents, red only for "isolated/no network", Dubai
font, RTL. It suited a security/platform briefing. It is one direction among
many — don't reuse it unless the user asks for that look.

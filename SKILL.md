---
name: motion-pptx
description: Build motion-design PowerPoint decks (.pptx) where every element is a native, editable shape or text box animated with native PowerPoint entrance effects — word-by-word titles, spring pop-ins, rising cards, self-drawing connectors and rings, travelling dots, sliding toggles — that play automatically when each slide opens. Guides the design process (palette, background, typography, motif and pacing chosen fresh for each deck's topic or brand) and supplies the animation engine, Arabic/RTL and English/LTR support, and checks that keep PowerPoint from reporting the file as damaged. Use this whenever someone wants an animated, "video-like", motion-graphics, keynote-style, Remotion-style or "beautiful" PowerPoint, wants animations added to a deck, or asks to turn a video/storyboard design into an editable PowerPoint — even if they only say "make it look like a video" or "add animation". Prefer it over plain static deck generation whenever motion is requested.
---

# Motion PowerPoint decks

Native shapes + native PowerPoint animations, built with pptxgenjs and a
post-processing step that writes each slide's `<p:timing>`. The result is
fully editable in PowerPoint (text, colours, positions, the Animation Pane),
small (≈1 MB), and plays like a motion-graphics video.

The skill steers **how** to design and animate; it does not prescribe a
look. You choose the palette, background, fonts and motif for each deck
from its subject, audience and brand. The kit has no default colours — it
refuses to build until you pass a palette — and the example's light theme
is a placeholder, not a template.

Why this design: pptxgenjs cannot write animations, and hand-written
animation XML that is schema-valid can still be rejected by PowerPoint as
"damaged". `scripts/anim.js` encodes the rules PowerPoint actually enforces
(see `references/animation-xml.md`), so use it rather than writing timing XML
yourself.

## Files

| File | Use |
|---|---|
| `scripts/motion_kit.js` | `Deck` / `Slide` API: text, runs, rect, card (glass or solid, from the palette), oval, badge (icon in circle), arc (percentage ring), line (self-drawing arrow, optional looping dot), title, footer, `save()` |
| `scripts/anim.js` | Timeline generator (used by `save()`); `SAFE=1` keeps only basic effects |
| `scripts/backdrop.js` | One background technique (gradient + optional grid + glows), any colours you choose |
| `scripts/check_deck.py` | Checks the PowerPoint-only rules (+ XSD with `--xsd`) — run on every deck |
| `scripts/example.js` | Mechanics demo: five slide patterns (title, cards, flow diagram, stat ring, timeline), RTL or `--ltr`, full + safe decks |
| `references/design.md` | **Read before designing**: choosing a direction, palette roles, backgrounds, typography, grid, motion grammar, RTL, QA fonts |
| `references/animation-xml.md` | Effect catalogue and the PowerPoint rules, for debugging or new effects |

## Workflow

1. **Set up** a working directory with the dependencies, and point Node at
   them (the scripts live in the skill folder, so they resolve modules via
   `NODE_PATH`):
   ```bash
   npm install pptxgenjs sharp react react-dom react-icons
   export NODE_PATH="$(npm root)"
   ```
   If pptxgenjs is preinstalled globally, `NODE_PATH` may already be enough.
   Your build script `require`s the kit by its path in the skill folder
   (`<skill>/scripts/motion_kit.js`); the npm modules come from `NODE_PATH`.

2. **Plan the story** before any code: one message per slide, ~6 s of motion
   each, title → summary → sections (with divider slides) → close.

3. **Decide the design direction** for this deck — read
   `references/design.md` first. From the content, audience and any brand
   the user gave you, choose: mood, light or dark, a palette expressed as
   roles (`bg, ink, muted, accent, accent2, onAccent, surface…`), the
   background treatment, fonts the presenter will have, one visual motif,
   and how energetic the motion should be. Write it as a short comment at
   the top of the build script. Don't reuse colours from earlier decks or
   from `example.js` unless the user asks for them; if the user supplies
   brand colours or a style, those win.

4. **Read `scripts/example.js`** for the mechanics and adapt the patterns
   you need to your own design (not its colours or layout). Lay out in
   px on a 1920×1080 canvas; in RTL, reading order starts at the right edge,
   so mirror x (`X(left, w)` in the example). Give each element an
   animation with `s.at(frame)` (30 fps frames) — stagger siblings by 6–14
   frames, let titles lead, details follow. Units: **start times are
   frames** (`s.at(frame)`), **durations are ms** (`s.at(frame, durMs)`);
   `s.atMs(ms)` exists for computed offsets:
   ```js
   const { Deck } = require('<skill>/scripts/motion_kit.js');
   // palette = the one you designed in step 3 (hex, no #)
   const deck = new Deck({ rtl: true, font: 'Dubai', palette });
   const P = deck.palette;
   const s = deck.slide({ bg: 'bg.png', notes: 'speaker notes' });
   s.title('العنوان', 'سطر فرعي');                      // words rise in, subtitle follows
   s.card(110, 320, 540, 400, s.at(20).float());        // card rises in at frame 20
   await s.badge(1640, 340, 96, 'FaShieldHalved', P.accent, P.onAccent, 24);
   s.line(1500, 400, 1100, 400, P.muted, 30, { dot: true }); // draws, then a dot loops along it
   await deck.save('deck.pptx');
   ```
   Effects: `fade float drop slide zoom grow wipe(dir) wheel words() move(dx,dy,{loop})`.
   Helpers beyond the basics: `s.arc(x, y, d, pct, color, frame)` draws a
   partial ring for a percentage; `line(..., { dot, dotDelay })` adds a dot
   that waits at the line's start until `dotDelay` ms (default 650) after the
   line begins drawing, then loops along it.
   **Escape hatch** for any other pptxgenjs object: add it with
   `s.s.addShape(...)` / `s.s.addText(...)` passing `objectName: s.name()`
   (keep the returned name), then `s.anim(name, s.at(frame).fade())`. Raw
   pptxgenjs takes inches, so convert with `I(px)` exported by the kit.

5. **Background**, as decided in step 3: a flat `bg` colour, or an image.
   `scripts/backdrop.js` is one technique (radial gradient, optional grid,
   glows — every colour is an argument, e.g.
   `node <skill>/scripts/backdrop.js bg.png --base <bg> --base2 <tone> --glow <accent> --grid 0.15`);
   for other treatments write a small SVG and rasterise it with sharp.
   Section dividers can use a variant (e.g. a different glow).

6. **Build, then check** — both every time:
   ```bash
   node build.js
   python3 <skill>/scripts/check_deck.py deck.pptx --xsd <pptx-skill>/scripts/office/schemas/ISO-IEC29500-4_2016/pml.xsd
   ```
   (`--xsd` is optional; use it when the pptx skill's schemas are available.)
   Fix anything it reports in the build script, never by hand-editing the zip.

7. **Visual QA** of the final frame of every slide (animations don't render
   in LibreOffice, but layout does):
   ```bash
   soffice --headless --convert-to pdf deck.pptx && pdftoppm -jpeg -r 60 deck.pdf qa
   ```
   If the pptx skill is installed, its `scripts/office/soffice.py` wrapper
   avoids sandbox hangs. Look for text overflow, RTL runs in the wrong order,
   overlapping cards, weak contrast, a palette that drifted from the
   brief. Arial/Calibri decks preview with metric-compatible
   substitutes (Liberation Sans, Carlito), so widths are trustworthy; for
   other fonts (e.g. Dubai) see the substitution note in
   `references/design.md`. Travelling dots appear parked at the start of
   their lines in previews; that is their resting position, not a bug.

8. **Deliver two files** when animations matter: the full deck and a fallback,
   saved from the same script as `example.js` does:
   `await deck.save('deck.pptx'); await deck.save('deck-safe.pptx', { safe: true });`
   Safe mode keeps fade/float/drop/zoom/grow/wipe; by-word titles animate as
   one block, rings fade in instead of drawing, and motion paths are dropped
   (travelling dots and sliding toggles stay still). Tell the user the
   animations could not be played in this environment, so they should run
   the slideshow once, and that the safe version exists if PowerPoint
   complains.

## Things that go wrong

- **"PowerPoint found a problem / repaired the file", images missing**: a
  timing rule was broken. Run `check_deck.py`; the rules are in
  `references/animation-xml.md`. Don't bypass `anim.js`.
- **Mixed Arabic + English text in the wrong order**: every run must carry
  the paragraph direction — build runs with `s.run()` / `s.bold()`, not raw
  `{text, options}`. Avoid starting or ending an Arabic line with a Latin
  token plus punctuation (`.env`, `0.1.x`); reword ("ملف env").
- **Fonts**: the font must exist on the presenter's machine. Dubai (Windows
  10+/Office) and Arial are safe for Arabic; Arial/Calibri for English.
  A web font used in a storyboard (e.g. IBM Plex) will be substituted.
- **What native animation can't do**: number count-ups (use zoom), blur,
  colour morphs, back-and-forth toggles. Say so instead of faking it.

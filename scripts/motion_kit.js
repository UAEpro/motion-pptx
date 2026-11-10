// Motion-design PowerPoint decks from native shapes + native animations.
//
//   const { Deck } = require('<skill>/scripts/motion_kit.js');
//   const deck = new Deck({ rtl: true, font: 'Dubai', palette: { bg, ink, muted, accent, surface, onAccent } });
//   const s = deck.slide({ bg: 'out/bg.png', notes: '...' });
//
// The kit has no look of its own: every colour comes from the palette the
// deck passes in (see references/design.md for how to choose one).
//   s.title('العنوان', 'العنوان الفرعي');
//   s.glass(110, 320, 540, 400, s.at(20).float());
//   await deck.save('deck.pptx');
//
// Coordinates are px on a 1920x1080 canvas (1 px = 1/144 in, font px/2 = pt),
// so a layout drawn in HTML/Remotion transfers 1:1. Delays are in frames at
// 30 fps (s.at(frame)) so motion timing can be copied from a storyboard.
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');
const { timingXml, shapeIds } = require('./anim');

const I = (px) => px / 144;
const msFromFrames = (f) => Math.round(f * 33.33);

// Palette roles. Required ones have no default on purpose: the deck's
// designer decides them for its topic or brand.
const REQUIRED_ROLES = ['bg', 'ink', 'muted', 'accent'];
function resolvePalette(p = {}) {
  const missing = REQUIRED_ROLES.filter((k) => !p[k]);
  if (missing.length) {
    throw new Error(`Deck palette is missing ${missing.join(', ')}. Choose colours for this deck ` +
      '(references/design.md) and pass { bg, ink, muted, accent, ... } as hex without #.');
  }
  return {
    accent2: p.accent,          // secondary accent; defaults to the main one
    onAccent: p.bg,             // text/icons drawn on accent fills
    surface: p.ink,             // card fill (used with surfaceAlpha)
    surfaceAlpha: 94,           // 94% transparent tint of ink reads as glass on dark and light alike
    border: p.ink,
    borderAlpha: 86,
    ...p,
  };
}

let _icons = null;
async function iconPng(name, color) {
  // react-icons (fa6) -> PNG via sharp. Lazy so decks without icons need neither.
  if (!_icons) {
    _icons = { cache: {}, React: require('react'), DOM: require('react-dom/server'), sharp: require('sharp'), fa: require('react-icons/fa6') };
  }
  const k = name + color;
  if (!_icons.cache[k]) {
    const C = _icons.fa[name];
    if (!C) throw new Error(`unknown react-icons/fa6 icon ${name}`);
    const svg = _icons.DOM.renderToStaticMarkup(_icons.React.createElement(C, { color: '#' + color, size: 256 }));
    const png = await _icons.sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
    _icons.cache[k] = 'image/png;base64,' + png.toString('base64');
  }
  return _icons.cache[k];
}

/** Animation spec builder: s.at(frame, durMs).float() etc. */
class At {
  constructor(delayMs, dur) { this.delay = delayMs; this.dur = dur; }
  _(type, extra = {}) { return { type, delay: this.delay, dur: this.dur ?? extra.dur, ...extra }; }
  fade() { return this._('fade', { dur: this.dur ?? 500 }); }
  float() { return this._('float', { dur: this.dur ?? 600 }); }          // fade + rise
  drop() { return this._('drop', { dur: this.dur ?? 550 }); }            // fade + fall into place
  slide() { return this._('slide', { dur: this.dur ?? 550 }); }          // fade + slide sideways
  zoom() { return this._('zoom', { dur: this.dur ?? 550 }); }            // spring pop 0 -> 108% -> 100%
  grow() { return this._('grow', { dur: this.dur ?? 600 }); }            // 85% -> 100%
  wipe(dir = 'right') { return this._('wipe', { dir, dur: this.dur ?? 650 }); } // dir = edge it starts from
  wheel() { return this._('wheel', { dur: this.dur ?? 1300 }); }         // clockwise reveal (rings)
  words(pct = 12000) { return this._('float', { byWord: true, wordPct: pct, dur: this.dur ?? 450 }); }
  /** Motion path in px, relative to where the shape sits. loop = forever. */
  move(dxPx, dyPx, { fromPx = [0, 0], loop = false } = {}) {
    return this._('path', {
      from: [(fromPx[0] / 1920).toFixed(4), (fromPx[1] / 1080).toFixed(4)],
      to: [((fromPx[0] + dxPx) / 1920).toFixed(4), ((fromPx[1] + dyPx) / 1080).toFixed(4)],
      loop, dur: this.dur ?? (loop ? 1500 : 400),
    });
  }
}

class Slide {
  constructor(deck, { bg, bgColor, notes } = {}) {
    this.deck = deck;
    this.p = deck.palette;
    this.s = deck.pres.addSlide();
    if (bg) this.s.background = { data: bg.startsWith('image/') ? bg : 'image/png;base64,' + fs.readFileSync(bg).toString('base64') };
    else this.s.background = { color: bgColor || this.p.bg };
    if (notes) this.s.addNotes(notes);
    this.anims = [];
    this.n = 0;
  }
  at(frame, durMs) { return new At(msFromFrames(frame), durMs); }
  atMs(ms, durMs) { return new At(ms, durMs); }
  name() { return `o${++this.n}`; }
  anim(name, a) { if (a) this.anims.push({ name, ...a }); return name; }

  /** Text box. t: string, or runs from s.run()/s.bold(). o.size in px. */
  text(t, x, y, w, h, o = {}, a) {
    const name = this.name();
    const rtl = o.ltr ? false : this.deck.rtl;
    this.s.addText(t, {
      x: I(x), y: I(y), w: I(w), h: I(h), fontFace: this.deck.font, fontSize: (o.size || 30) / 2, bold: !!o.bold,
      color: o.color || this.p.ink, align: o.align || (this.deck.rtl ? 'right' : 'left'), valign: o.valign || 'top',
      rtlMode: rtl, lang: o.ltr ? 'en-US' : this.deck.lang, margin: 0, isTextBox: true, objectName: name,
      lineSpacingMultiple: o.lh || 1.0, fit: 'none', wrap: true,
    });
    return this.anim(name, a);
  }
  /** A run for mixed-style paragraphs. Carries the paragraph direction (needed: pptxgenjs reads it from runs). */
  run(t, size, color = this.p.muted, bold = false, extra = {}) {
    return { text: t, options: { rtlMode: this.deck.rtl, lang: this.deck.lang, fontFace: this.deck.font, fontSize: size / 2, color, bold, ...extra } };
  }
  bold(t, size, color = this.p.ink) { return this.run(t, size, color, true); }

  rect(x, y, w, h, o = {}, a) {
    const name = this.name();
    const square = o.radius === 0;
    this.s.addShape(square ? this.deck.pres.shapes.RECTANGLE : this.deck.pres.shapes.ROUNDED_RECTANGLE, {
      x: I(x), y: I(y), w: I(w), h: I(h), rectRadius: square ? undefined : I(Math.min(o.radius ?? 28, h / 2, w / 2)),
      fill: o.fill ? { color: o.fill, transparency: o.alpha ?? 0 } : { type: 'none' },
      line: o.line ? { color: o.line, width: o.lw ?? 0.75, transparency: o.lineAlpha ?? 0, ...(o.dash ? { dashType: o.dash } : {}) } : { type: 'none' },
      shadow: o.glow ? { type: 'outer', color: o.glow, opacity: o.glowOp ?? 0.45, blur: o.blur ?? 18, offset: 0, angle: 90 } : undefined,
      objectName: name,
    });
    return this.anim(name, a);
  }
  /**
   * Card in the palette's surface style (glass by default: a faint tint of
   * ink with a faint border). tint = [color, transparency%, borderColor?]
   * for a coloured card. For solid cards set palette surfaceAlpha: 0.
   */
  glass(x, y, w, h, a, tint) {
    const p = this.p;
    return this.rect(x, y, w, h, tint
      ? { fill: tint[0], alpha: tint[1], line: tint[2] || tint[0], lineAlpha: tint[2] ? 0 : 60 }
      : { fill: p.surface, alpha: p.surfaceAlpha, line: p.border, lineAlpha: p.borderAlpha }, a);
  }
  card(...args) { return this.glass(...args); }
  oval(x, y, d, o = {}, a) {
    const name = this.name();
    this.s.addShape(this.deck.pres.shapes.OVAL, {
      x: I(x), y: I(y), w: I(d), h: I(o.h ?? d),
      fill: o.fill ? { color: o.fill, transparency: o.alpha ?? 0 } : { type: 'none' },
      line: o.line ? { color: o.line, width: o.lw ?? 1, transparency: o.lineAlpha ?? 0 } : { type: 'none' },
      shadow: o.glow ? { type: 'outer', color: o.glow, opacity: 0.5, blur: 20, offset: 0, angle: 90 } : undefined,
      objectName: name,
    });
    return this.anim(name, a);
  }
  async image(data, x, y, w, h, a, alt = '') {
    const name = this.name();
    this.s.addImage({ data, x: I(x), y: I(y), w: I(w), h: I(h), objectName: name, altText: alt });
    return this.anim(name, a);
  }
  /** Icon in a filled circle; both pop in together. */
  async badge(x, y, d, icon, color = this.p.accent, fg = this.p.onAccent, frame = 0, glow = true) {
    const a = this.at(frame).zoom();
    this.oval(x, y, d, { fill: color, glow: glow ? color : undefined }, a);
    const pad = d * 0.26;
    await this.image(await iconPng(icon, fg), x + pad, y + pad, d - 2 * pad, d - 2 * pad, { ...a }, icon);
  }
  /** Partial ring for a percentage: arc from 12 o'clock clockwise, revealed by wheel. */
  arc(x, y, d, pct, color, frame, { thickness = 0.06, dur = 2000 } = {}) {
    const name = this.name();
    // PowerPoint angles: 0 = 3 o'clock, clockwise; 270 = 12 o'clock.
    const end = (270 + Math.max(0.5, Math.min(99.5, pct)) * 3.6) % 360;
    this.s.addShape(this.deck.pres.shapes.BLOCK_ARC, {
      x: I(x), y: I(y), w: I(d), h: I(d), angleRange: [270, end], arcThicknessRatio: thickness,
      fill: { color }, line: { type: 'none' }, objectName: name,
    });
    return this.anim(name, this.at(frame, dur).wheel());
  }
  /**
   * Arrow that draws itself in its direction of travel, starting at `frame`.
   * o.dot: a dot then loops along it, starting o.dotDelay ms (default 650)
   * after the line starts. Before that the dot sits at the line's start.
   */
  line(x1, y1, x2, y2, color, frame, o = {}) {
    const name = this.name();
    this.s.addShape(this.deck.pres.shapes.LINE, {
      x: I(Math.min(x1, x2)), y: I(Math.min(y1, y2)), w: I(Math.abs(x2 - x1) || 0.5), h: I(Math.abs(y2 - y1) || 0.5),
      flipH: x2 < x1, flipV: y2 < y1, objectName: name,
      line: { color, width: o.w ?? 1.5, transparency: 10, ...(o.noArrow ? {} : { endArrowType: 'triangle' }) },
    });
    const dx = x2 - x1, dy = y2 - y1;
    const dir = Math.abs(dx) >= Math.abs(dy) ? (dx < 0 ? 'right' : 'left') : (dy > 0 ? 'top' : 'bottom');
    this.anim(name, this.at(frame, o.dur ?? 650).wipe(dir));
    if (o.dot) {
      const dn = this.oval(x1 - 8, y1 - 8, 16, { fill: color });
      this.anim(dn, this.atMs(msFromFrames(frame) + (o.dotDelay ?? 650), o.dotDur ?? 1500).move(dx, dy, { loop: true }));
    }
  }
  /** Standard title block: words rise in, subtitle follows. */
  title(t, sub, { x = 110, y = 62, w = 1700, size = 76 } = {}) {
    this.text(t, x, y, w, 110, { size, bold: true, valign: 'middle' }, this.at(0).words());
    if (sub) this.text(sub, x, y + 110, w, 50, { size: 32, color: this.p.muted }, this.at(8).float());
  }
  footer(label, n) {
    const rtl = this.deck.rtl;
    this.text(label, rtl ? 1110 : 110, 1012, 700, 32, { size: 22, color: this.p.muted });
    this.text(String(n).padStart(2, '0'), rtl ? 110 : 1610, 1012, 200, 32, { size: 22, color: this.p.muted, align: rtl ? 'left' : 'right', ltr: true });
  }
}

class Deck {
  constructor({ rtl = false, font = 'Arial', lang, palette = {}, title = '', author = '' } = {}) {
    this.rtl = rtl;
    this.font = font;
    this.lang = lang || (rtl ? 'ar-AE' : 'en-US');
    this.palette = resolvePalette(palette);
    this.pres = new pptxgen();
    this.pres.layout = 'LAYOUT_WIDE';
    this.pres.rtlMode = rtl;
    this.pres.title = title;
    this.pres.author = author;
    this.pres.theme = { headFontFace: font, bodyFontFace: font };
    this.slides = [];
  }
  slide(opts) { const s = new Slide(this, opts); this.slides.push(s); return s; }

  /**
   * Write the deck, then add transitions + timelines and clean pptxgenjs XML.
   * Safe to call more than once (e.g. full + { safe: true }): pptxgenjs
   * converts option values in place on its first write, so the raw package
   * is generated once and every save post-processes a copy of it.
   */
  async save(file, { transition = 'fade', safe = false } = {}) {
    if (!this._raw) this._raw = await this.pres.write({ outputType: 'nodebuffer' });
    const zip = await JSZip.loadAsync(this._raw);
    if (safe) process.env.SAFE = '1';
    for (let i = 0; i < this.slides.length; i++) {
      const name = `ppt/slides/slide${i + 1}.xml`;
      let xml = await zip.file(name).async('string');
      // pptxgenjs repeats <a:pPr> before every run of a multi-run paragraph;
      // only one, first, is allowed.
      xml = xml.replace(/(<\/a:r>)<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, '$1');
      xml = xml.replace(/<p:transition[\s\S]*?(<\/p:transition>|\/>)/, '').replace(/<p:timing>[\s\S]*?<\/p:timing>/, '');
      const anims = this.slides[i].anims;
      const timing = anims.length ? timingXml(anims, shapeIds(xml)) : '';
      const trans = transition ? `<p:transition spd="med"><p:${transition}/></p:transition>` : '';
      xml = xml.replace('</p:clrMapOvr>', '</p:clrMapOvr>' + trans + timing);
      zip.file(name, xml);
    }
    if (safe) delete process.env.SAFE;
    fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    return path.resolve(file);
  }
}

module.exports = { Deck, iconPng, I, msFromFrames };

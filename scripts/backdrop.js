#!/usr/bin/env node
// One technique for a slide background image: a radial gradient between two
// tones, an optional faint grid that fades at the edges, and up to two
// blurred glows. Use it when that suits the deck's design direction; write
// your own SVG (shapes, stripes, paper texture, photo duotone...) when it
// doesn't. Gradients and blur are not native pptxgenjs fills, which is why
// the background is an image while everything on top stays native.
//
//   node backdrop.js out.png --base HEX [--base2 HEX] [--glow HEX] [--glow2 HEX]
//                     [--grid 0..1] [--line HEX] [--focus x,y]
//   --base    main background colour (required)
//   --base2   tone of the highlight (default: --base, i.e. flat)
//   --glow    colour of the large soft glow (omit for none)
//   --glow2   colour of the second, smaller glow (omit for none)
//   --grid    grid opacity 0..1 (default 0; try 0.12-0.2)
//   --line    grid line colour (default: white on dark bases, black on light)
//   --focus   where the highlight sits, as fractions (default 0.85,0.1 —
//             top right, suits RTL; use 0.15,0.1 for LTR)
const sharp = require('sharp');

const args = process.argv.slice(2);
const out = args[0];
const opt = (k, d) => { const i = args.indexOf('--' + k); return i > -1 ? args[i + 1] : d; };
const base = opt('base');
if (!out || !base) {
  console.error('usage: backdrop.js out.png --base HEX [--base2 HEX] [--glow HEX] [--glow2 HEX] [--grid 0..1] [--line HEX] [--focus x,y]');
  process.exit(2);
}
const base2 = opt('base2', base);
const glow = opt('glow'), glow2 = opt('glow2');
const grid = Number(opt('grid', '0'));
const [fx, fy] = opt('focus', '0.85,0.1').split(',').map(Number);
const lum = (h) => { const n = parseInt(h, 16); return (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255; };
const line = opt('line', lum(base) < 0.5 ? 'FFFFFF' : '000000');

const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080">
  <defs>
    <radialGradient id="bg" cx="${fx}" cy="${fy}" r="0.9">
      <stop offset="0" stop-color="#${base2}"/><stop offset="0.6" stop-color="#${base}"/><stop offset="1" stop-color="#${base}"/>
    </radialGradient>
    <pattern id="grid" width="80" height="80" patternUnits="userSpaceOnUse">
      <path d="M80 0H0V80" fill="none" stroke="#${line}" stroke-width="1"/>
    </pattern>
    <radialGradient id="fadeMask" cx="0.5" cy="0.5" r="0.6">
      <stop offset="0.3" stop-color="#fff" stop-opacity="1"/><stop offset="1" stop-color="#fff" stop-opacity="0"/>
    </radialGradient>
    <mask id="m"><rect width="1920" height="1080" fill="url(#fadeMask)"/></mask>
    <filter id="blur" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="120"/></filter>
  </defs>
  <rect width="1920" height="1080" fill="url(#bg)"/>
  ${grid > 0 ? `<rect width="1920" height="1080" fill="url(#grid)" mask="url(#m)" opacity="${grid}"/>` : ''}
  ${glow ? `<circle cx="${fx * 1920}" cy="${fy * 1080}" r="450" fill="#${glow}" opacity="0.10" filter="url(#blur)"/>` : ''}
  ${glow2 ? `<circle cx="${(1 - fx) * 1920}" cy="${(1 - fy) * 1080}" r="350" fill="#${glow2}" opacity="0.07" filter="url(#blur)"/>` : ''}
</svg>`;

sharp(Buffer.from(svg)).png().toFile(out).then(() => console.log('wrote', out));

#!/usr/bin/env node
// Worked example of the MECHANICS: five slide patterns (title, cards, flow
// diagram, stat ring, timeline), RTL or LTR, full + safe decks.
//   node example.js out.pptx [--ltr]
//
// The palette, background and layout here are placeholders chosen only to
// show the API on a light theme. Do not reuse them: design each deck's
// look for its own topic or brand (references/design.md).
const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { Deck } = require('./motion_kit');

const OUT = process.argv[2] || 'example.pptx';
const LTR = process.argv.includes('--ltr');
const BG = path.join(path.dirname(path.resolve(OUT)), 'bg.png');
// Placeholder palette for this demo only.
const PALETTE = { bg: 'F4F6F8', ink: '1F2A37', muted: '5B6675', accent: 'E4572E', accent2: '2E86AB', onAccent: 'FFFFFF',
  surface: 'FFFFFF', surfaceAlpha: 0, border: 'D8DEE6', borderAlpha: 0 };
if (!fs.existsSync(BG)) {
  execFileSync('node', [path.join(__dirname, 'backdrop.js'), BG, '--base', PALETTE.bg, '--base2', 'FFFFFF',
    '--glow', PALETTE.accent2, '--grid', '0.06', '--focus', LTR ? '0.15,0.1' : '0.85,0.1'], { stdio: 'inherit' });
}

const T = LTR ? {
  title: 'Project Atlas', tagline: 'One platform for every team', sum: 'Executive summary', sumSub: 'Three things to remember',
  cards: [['Faster', 'Ship features in days, not months'], ['Safer', 'Isolation and audit on every layer'], ['Shared', 'One platform every team can reuse']],
  flow: 'How it fits together', flowSub: 'Every agent goes through one gateway',
  nodes: ['App A', 'App B', 'Gateway', 'Models', 'Tools'], stat: 'Uptime this quarter', statN: '99.9', road: 'Roadmap', roadSub: 'Start small, grow safely',
  steps: ['Pilot', 'Expand', 'Govern', 'Scale'], footer: 'Atlas · internal',
} : {
  title: 'مشروع أطلس', tagline: 'منصة واحدة لكل الفرق', sum: 'الملخص التنفيذي', sumSub: 'ثلاث رسائل رئيسية',
  cards: [['أسرع', 'إطلاق الميزات في أيام بدل أشهر'], ['أكثر أماناً', 'عزل وتدقيق على كل طبقة'], ['مشترك', 'منصة واحدة يعيد كل فريق استخدامها']],
  flow: 'كيف تترابط الأجزاء', flowSub: 'كل وكيل يمر عبر بوابة واحدة',
  nodes: ['تطبيق أ', 'تطبيق ب', 'البوابة', 'النماذج', 'الأدوات'], stat: 'نسبة التشغيل هذا الربع', statN: '99.9', road: 'خارطة الطريق', roadSub: 'ابدأ صغيراً وتوسّع بأمان',
  steps: ['تجربة', 'توسّع', 'حوكمة', 'انتشار'], footer: 'أطلس · للاستخدام الداخلي',
};

// In RTL the reading order starts at the right edge: mirror x for layouts.
const X = (left, w) => (LTR ? left : 1920 - left - w);

(async () => {
  const deck = new Deck({ rtl: !LTR, font: LTR ? 'Arial' : 'Dubai', title: T.title, palette: PALETTE });
  const P = deck.palette;

  // 1. Title: ring draws, badge pops, title word by word.
  let s = deck.slide({ bg: BG, notes: 'Opening.' });
  s.oval(X(150, 192), 204, 192, { line: P.accent, lw: 2.5 }, s.at(0).wheel());
  await s.badge(X(174, 144), 228, 144, 'FaRocket', P.accent, P.onAccent, 10);
  s.text(T.title, X(400, 1270), 175, 1270, 250, { size: 150, bold: true, valign: 'middle' }, s.at(14, 600).words(30000));
  s.text(T.tagline, X(150, 1620), 500, 1620, 90, { size: 56, bold: true }, s.at(40).float());

  // 2. Cards: rise in, staggered 12 frames.
  s = deck.slide({ bg: BG });
  s.title(T.sum, T.sumSub);
  const icons = ['FaBolt', 'FaShieldHalved', 'FaCubes'];
  for (let i = 0; i < 3; i++) {
    const w = 544, x = X(110 + i * (w + 34), w), f = 20 + i * 12;
    s.glass(x, 330, w, 440, s.at(f).float());
    await s.badge(LTR ? x + 34 : x + w - 34 - 110, 364, 110, icons[i], i === 2 ? P.accent2 : P.accent, P.onAccent, f + 8);
    s.text(T.cards[i][0], x + 34, 518, w - 68, 64, { size: 46, bold: true }, s.at(f + 3).float());
    s.text(T.cards[i][1], x + 34, 592, w - 68, 150, { size: 30, color: P.muted }, s.at(f + 5).float());
  }
  s.footer(T.footer, 2);

  // 3. Flow diagram: boxes zoom, connectors draw, dots travel.
  s = deck.slide({ bg: BG });
  s.title(T.flow, T.flowSub);
  const box = (left, y, w, label, fill, f, dark) => {
    const x = X(left, w);
    s.rect(x, y, w, 120, { fill, radius: 22, glow: fill, glowOp: 0.3, blur: 24 }, s.at(f).zoom());
    s.text(label, x, y, w, 120, { size: 32, bold: true, align: 'center', valign: 'middle', color: P.onAccent }, s.at(f).zoom());
    return x;
  };
  box(150, 330, 320, T.nodes[0], P.ink, 10);
  box(150, 620, 320, T.nodes[1], P.ink, 16);
  box(800, 460, 320, T.nodes[2], P.accent, 26, true);
  box(1450, 330, 320, T.nodes[3], P.accent2, 40, true);
  box(1450, 620, 320, T.nodes[4], P.accent2, 46, true);
  const ex = (left, w, side) => (LTR ? (side === 'end' ? left + w : left) : (side === 'end' ? 1920 - left - w : 1920 - left));
  s.line(ex(150, 320, 'end'), 390, ex(800, 320, 'start'), 520, P.muted, 30, { dot: true });
  s.line(ex(150, 320, 'end'), 680, ex(800, 320, 'start'), 520, P.muted, 34, { dot: true });
  s.line(ex(800, 320, 'end'), 520, ex(1450, 320, 'start'), 390, P.accent2, 50, { dot: true });
  s.line(ex(800, 320, 'end'), 520, ex(1450, 320, 'start'), 680, P.accent2, 54, { dot: true });
  s.footer(T.footer, 3);

  // 4. Stat ring: track fades, the arc for the value draws clockwise, number pops.
  s = deck.slide({ bg: BG });
  s.title(T.stat);
  s.oval(830, 340, 260, { line: P.border, lw: 8 }, s.at(10).fade());
  s.arc(822, 332, 276, Number(T.statN), P.accent, 16);
  s.text(T.statN + '%', 780, 410, 360, 120, { size: 80, bold: true, align: 'center', valign: 'middle', ltr: true }, s.at(20, 700).zoom());
  s.footer(T.footer, 4);

  // 5. Timeline: line wipes in reading direction, steps pop in sequence.
  s = deck.slide({ bg: BG });
  s.title(T.road, T.roadSub);
  s.rect(300, 397, 1320, 6, { fill: P.accent, radius: 0 }, s.at(10, 2000).wipe(LTR ? 'left' : 'right'));
  for (let i = 0; i < 4; i++) {
    const w = 402, x = X(110 + i * (w + 30), w), f = 18 + i * 14;
    s.oval(x + w / 2 - 60, 340, 120, { fill: i < 2 ? P.accent : P.accent2, line: P.bg, lw: 3 }, s.at(f).zoom());
    s.text(String(i + 1), x + w / 2 - 60, 340, 120, 120, { size: 56, bold: true, color: P.onAccent, align: 'center', valign: 'middle', ltr: true }, s.at(f).zoom());
    s.glass(x, 496, w, 160, s.at(f + 6).float());
    s.text(T.steps[i], x + 30, 520, w - 60, 100, { size: 40, bold: true, valign: 'middle', align: 'center' }, s.at(f + 9).float());
  }
  s.footer(T.footer, 5);

  // One build script, two decks: the full one and a fallback with basic effects only.
  console.log('wrote', await deck.save(OUT));
  console.log('wrote', await deck.save(OUT.replace(/\.pptx$/, '-safe.pptx'), { safe: true }));
})().catch((e) => { console.error(e); process.exit(1); });

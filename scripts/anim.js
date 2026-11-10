// Native PowerPoint entrance animations, written as <p:timing> XML.
// Every effect starts on slide load ("after previous" for the first, the
// rest "with previous" + their own delay), so one slide is one timeline.
//
// PowerPoint enforces rules the XSD does not; breaking them makes it report
// the file as damaged and "repair" it (dropping animations, sometimes
// objects). This module keeps to them:
//   1. cTn ids are sequential in document order (renumber()).
//   2. bldP: text boxes get no animBg; other shapes get animBg="1";
//      pictures get no bldP at all.
//   3. repeatCount sits on the effect's own cTn; motion paths carry rAng.

const tgt = (spid) => `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>`;

function effect(a, spid, next) {
  const id = () => next();
  const d = a.dur || 500;
  const vis = () => `<p:set><p:cBhvr><p:cTn id="${id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt(spid)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
  const fade = (dur = d) => `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${id()}" dur="${dur}"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
  const anim = (attr, vals, dur = d, decel = true) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${id()}" dur="${dur}" fill="hold"${decel ? ' decel="100000"' : ''}/>${tgt(spid)}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>${vals.map(([tm, v]) => `<p:tav tm="${tm}"><p:val><p:strVal val="${v}"/></p:val></p:tav>`).join('')}</p:tavLst></p:anim>`;
  let preset, cls = 'entr', sub = 0, body;
  switch (a.type) {
    case 'fade': preset = 10; body = vis() + fade(); break;
    case 'float': // fade + rise (Float In)
    case 'drop': {
      const dy = a.type === 'float' ? '+0.045' : '-0.04';
      preset = 42; body = vis() + fade() + anim('ppt_x', [[0, '#ppt_x'], [100000, '#ppt_x']]) + anim('ppt_y', [[0, `#ppt_y${dy}`], [100000, '#ppt_y']]);
      break;
    }
    case 'slide': // fade + slide in from the left (RTL lists)
      preset = 42; body = vis() + fade() + anim('ppt_x', [[0, '#ppt_x-0.03'], [100000, '#ppt_x']]) + anim('ppt_y', [[0, '#ppt_y'], [100000, '#ppt_y']]);
      break;
    case 'zoom': // spring pop: 0 → 108% → 100%
      preset = 53; sub = 16;
      body = vis() + anim('ppt_w', [[0, '0'], [70000, '#ppt_w*1.08'], [100000, '#ppt_w']], d, false) + anim('ppt_h', [[0, '0'], [70000, '#ppt_h*1.08'], [100000, '#ppt_h']], d, false) + fade(Math.round(d * 0.6));
      break;
    case 'grow': // gentle scale from 85%
      preset = 53; sub = 16;
      body = vis() + anim('ppt_w', [[0, '#ppt_w*0.85'], [100000, '#ppt_w']]) + anim('ppt_h', [[0, '#ppt_h*0.85'], [100000, '#ppt_h']]) + fade();
      break;
    case 'wipe': { // dir = edge the reveal starts from
      const map = { right: [2, 'wipe(right)'], left: [8, 'wipe(left)'], top: [1, 'wipe(up)'], bottom: [4, 'wipe(down)'] };
      const [s, f] = map[a.dir || 'right'];
      preset = 22; sub = s;
      body = vis() + `<p:animEffect transition="in" filter="${f}"><p:cBhvr><p:cTn id="${id()}" dur="${d}"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
      break;
    }
    case 'wheel':
      preset = 21; sub = 1;
      body = vis() + `<p:animEffect transition="in" filter="wheel(1)"><p:cBhvr><p:cTn id="${id()}" dur="${d}"/>${tgt(spid)}</p:cBhvr></p:animEffect>`;
      break;
    case 'path': { // motion path, fractions of slide width/height; optional loop
      cls = 'path'; preset = 0;
      body = `<p:animMotion origin="layout" path="M ${a.from[0]} ${a.from[1]} L ${a.to[0]} ${a.to[1]} E" pathEditMode="relative" rAng="0" ptsTypes="AA"><p:cBhvr><p:cTn id="${id()}" dur="${d}" fill="hold"/>${tgt(spid)}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
      break;
    }
    default: throw new Error('unknown effect ' + a.type);
  }
  const iterate = a.byWord ? `<p:iterate type="wd"><p:tmPct val="${a.wordPct || 12000}"/></p:iterate>` : '';
  const extra = a.type === 'path' ? (a.loop ? ' repeatCount="indefinite"' : ' accel="50000" decel="50000"') : '';
  return { cls, xml: (nodeType) => `<p:par><p:cTn id="0" presetID="${preset}" presetClass="${cls}" presetSubtype="${sub}"${extra} fill="hold" grpId="0" nodeType="${nodeType}"><p:stCondLst><p:cond delay="${Math.round(a.delay || 0)}"/></p:stCondLst>${iterate}<p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>` };
}

/** anims: [{name, type, delay(ms), dur, ...}], idsByName: name -> {id, isSp} */
// SAFE=1: only the most basic presets (no by-word, wheel or motion paths).
function safe(a) {
  if (!process.env.SAFE) return a;
  if (a.type === 'path') return null;
  const b = { ...a, byWord: false };
  if (b.type === 'wheel') b.type = 'fade';
  if (b.type === 'drop' || b.type === 'slide') b.type = 'float';
  return b;
}

function timingXml(anims, idsByName) {
  anims = anims.map(safe).filter(Boolean);
  const sorted = [...anims].sort((a, b) => (a.delay || 0) - (b.delay || 0));
  const pars = [];
  const built = new Map();
  sorted.forEach((a, i) => {
    const ref = idsByName[a.name];
    if (!ref) throw new Error('no shape named ' + a.name);
    pars.push(effect(a, ref.id, () => 0).xml(i === 0 ? 'afterEffect' : 'withEffect'));
    // Text boxes build their text; other shapes build as one object.
    if (ref.isSp) built.set(ref.id, ref.isTxBox ? '' : ' animBg="1"');
  });
  const bld = [...built].map(([id, bg]) => `<p:bldP spid="${id}" grpId="0"${bg}/>`).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
    `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
    `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>` +
    `<p:par><p:cTn id="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${pars.join('')}</p:childTnLst></p:cTn></p:par>` +
    `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>` +
    `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>` +
    `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
    `</p:childTnLst></p:cTn></p:par></p:tnLst>${bld ? `<p:bldLst>${bld}</p:bldLst>` : ''}</p:timing>`;
}

/** Sequential cTn ids in document order, as PowerPoint writes them. */
function renumber(timing) {
  let n = 0;
  return timing.replace(/<p:cTn id="\d+"/g, () => `<p:cTn id="${++n}"`);
}

/** Map objectName -> {id, isSp} from a slide's XML. */
function shapeIds(xml) {
  const out = {};
  const re = /<p:(sp|pic|cxnSp)>\s*<p:nv(?:Sp|Pic|CxnSp)Pr>\s*<p:cNvPr id="(\d+)" name="([^"]+)"/g;
  let m;
  while ((m = re.exec(xml))) {
    const tail = xml.slice(m.index, m.index + 400);
    out[m[3]] = { id: m[2], isSp: m[1] === 'sp', isTxBox: m[1] === 'sp' && /<p:cNvSpPr txBox="1"/.test(tail) };
  }
  return out;
}

module.exports = { timingXml: (a, ids) => renumber(timingXml(a, ids)), shapeIds };

# PowerPoint animation XML

Read this when debugging a deck PowerPoint reports as damaged, or when
adding a new effect to `scripts/anim.js`.

## Structure `anim.js` writes (per slide)

```
p:timing
 └ tnLst / par / cTn(tmRoot)
    └ seq(mainSeq) / cTn
       └ par / cTn  stCond: delay=indefinite + onBegin of mainSeq  (starts on slide load)
          └ par / cTn  delay 0                                      (one group = one timeline)
             ├ par / cTn effect 1  nodeType=afterEffect  delay=d1
             ├ par / cTn effect 2  nodeType=withEffect   delay=d2
             └ …
 └ bldLst / bldP …
```
Then `<p:transition spd="med"><p:fade/></p:transition>` precedes `p:timing`
right after `p:clrMapOvr` (CT_Slide order: cSld, clrMapOvr, transition,
timing, extLst).

## Rules PowerPoint enforces beyond the XSD
Breaking any of these gives "PowerPoint found a problem with content… repair",
and the repair may drop animations or objects (objects then look missing).

1. **cTn ids are 1..N in document order.** No gaps, no out-of-order ids.
   `anim.js` builds with id 0 placeholders and renumbers at the end.
2. **Build list (`bldP`)**: one per animated `p:sp`, `grpId` matching the
   effect's `grpId`.
   - text boxes (`txBox="1"`), including by-word animations: `<p:bldP spid grpId="0"/>`
   - other shapes: `<p:bldP spid grpId="0" animBg="1"/>`
   - pictures: no bldP.
3. **Looping**: `repeatCount="indefinite"` on the effect's own cTn, not on
   a behaviour cTn. Motion paths carry `rAng="0"` and `ptsTypes` matching
   the point count ("AA" for M…L…).
4. **One `a:pPr` per paragraph**, first. pptxgenjs writes one per run in
   multi-run paragraphs; `motion_kit.save()` strips the extras.
5. Every `spid` must be an existing `cNvPr id` on that slide.

`scripts/check_deck.py` checks all five.

## Effect catalogue (presetID / class / subtype → behaviours)
| Kit name | Preset | Behaviours |
|---|---|---|
| fade | 10 entr 0 | set visibility + animEffect fade |
| float / drop / slide | 42 entr 0 | + fade + anim ppt_x / ppt_y from `#ppt_y±0.045` (decel) |
| zoom | 53 entr 16 | + anim ppt_w/ppt_h 0 → 108% → 100% + short fade |
| grow | 53 entr 16 | + anim ppt_w/ppt_h from 85% + fade |
| wipe | 22 entr 1/2/4/8 | animEffect `wipe(up/right/down/left)` |
| wheel | 21 entr 1 | animEffect `wheel(1)` (clockwise reveal, good for rings) |
| words() | any entr | `<p:iterate type="wd"><p:tmPct val="12000"/>` before childTnLst |
| move | 0 path 0 | animMotion relative path in slide fractions; loop → repeatCount |

Values like `#ppt_w*1.08` are PowerPoint formula strings; `tm` is
0–100000 (percent × 1000).

## Not possible natively
Count-up numbers, blur, gradient/colour morphs on most shapes, CSS-like
easing curves beyond accel/decel. Approximate (zoom, fade) and say so.

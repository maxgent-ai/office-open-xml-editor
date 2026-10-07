import { describe, it, expect } from 'vitest';
import {
  computeFrameBox,
  pushFloatRect,
  registerFrameFloat,
} from './frame-geometry.js';
import type { FrameBox } from './frame-geometry.js';
import type { FramePr } from './types.js';
import type { FloatRect } from './float-layout.js';

// Table-driven geometry assertions for paragraph frames / drop caps
// (ECMA-376 §17.3.1.11). The VRT only exercises dropCap="drop" wrap="around"
// (private/sample-11 page 5), so the other (wrap, dropCap, hAnchor, vAnchor)
// combinations are pinned here — this is their only regression guard. Each case
// asserts the resolved FrameBox AND the FloatRect that registerFrameFloat emits
// (xLeft/xRight/yTop/yBottom/mode/side), which is what resolveLineFloatWindow
// consumes to wrap the following body text.
//
// Geometry is exercised directly in canonical points. A representative page:
//   pageWidth=600, margins L/R/T/B = 100/100/72/72  ⇒ content band [100,500].
//   A multi-column run would set a narrower contentX/contentW; we model a single
//   column here as contentX=100, contentW=400 and assert hAnchor="text" snaps to
//   it (the #513 column-relative contract).

interface MinState {
  contentX: number;
  contentW: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  pageWidth: number;
  pageH: number;
  floats: FloatRect[];
  floatParaSeq: number;
}

function makeState(over: Partial<MinState> = {}): MinState {
  return {
    contentX: 100,
    contentW: 400,
    marginLeft: 100,
    marginRight: 100,
    marginTop: 72,
    marginBottom: 72,
    pageWidth: 600,
    pageH: 800,
    floats: [],
    floatParaSeq: 0,
    ...over,
  };
}

// Full FramePr with the spec defaults; tests override only the axis under test.
function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none',
    lines: 1,
    wrap: 'around',
    hAnchor: 'text',
    vAnchor: 'text',
    hRule: 'auto',
    hSpace: 0,
    vSpace: 0,
    ...over,
  };
}

// Cast helper: computeFrameBox/registerFrameFloat read only the MinState subset
// of BodyAcquisitionState exercised here.
const box = (fp: FramePr, st: MinState, paraTop: number, cW: number, cH: number, anchorH: number): FrameBox =>
  computeFrameBox(fp, st as never, paraTop, cW, cH, anchorH);
const registerFloat = (b: FrameBox, fp: FramePr, st: MinState): void =>
  registerFrameFloat(b, fp, st as never);

describe('frame geometry (§17.3.1.11) — drop cap placement', () => {
  it('dropCap="drop": frame at column left, height = lines × anchor line height', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'drop', lines: 3, hAnchor: 'text', vAnchor: 'text' });
    // paraTop=200 (in-flow Y), measured cap width 42, anchor line height 14.
    const b = box(fp, st, 200, 42, 50, 14);
    expect(b.x).toBe(100); // column left (contentX)
    expect(b.y).toBe(200); // paragraph top (vAnchor="text")
    expect(b.w).toBe(42); // auto width = measured cap advance
    expect(b.h).toBe(3 * 14); // lines × anchor line height (y/yAlign ignored)
  });

  it('dropCap="margin": frame hangs into the left margin (left = band left − width)', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'margin', lines: 2, hAnchor: 'text' });
    const b = box(fp, st, 150, 30, 40, 12);
    expect(b.x).toBe(100 - 30); // outside the column margin
    expect(b.w).toBe(30);
    expect(b.h).toBe(2 * 12);
  });

  it('drop cap emits a right-side square float (text wraps to the right only)', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'drop', lines: 3, wrap: 'around', hSpace: 8 });
    const b = box(fp, st, 200, 42, 50, 14);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    const f = st.floats[0];
    expect(f.mode).toBe('square');
    expect(f.side).toBe('right');
    // hSpace=8 pads L/R for wrap="around".
    expect(f.xLeft).toBe(100 - 8);
    expect(f.xRight).toBe(100 + 42 + 8);
    expect(f.yTop).toBe(200);
    expect(f.yBottom).toBe(200 + 42); // h=3×14, vSpace=0
  });
});

describe('legacy float transport facts', () => {
  it('fails closed when a floating-table transport omits tblOverlap', () => {
    const st = makeState();

    expect(() => pushFloatRect(st as never, {
      x: 0,
      y: 0,
      w: 10,
      h: 10,
      dl: 0,
      dr: 0,
      dt: 0,
      db: 0,
      kind: 'table',
      mode: 'square',
      side: 'bothSides',
      imageKey: '',
      paraId: 0,
      avoidOverlap: true,
    })).toThrow('Floating-table transport omitted tblOverlap');
  });

  it('keeps a displaced frame wrap band within the page-right boundary', () => {
    const st = makeState({
      pageWidth: 100,
      floats: [{
        kind: 'frame',
        mode: 'square',
        imageKey: 'blocker',
        imageX: 0,
        imageY: 0,
        imageW: 50,
        imageH: 50,
        xLeft: 0,
        xRight: 50,
        yTop: 0,
        yBottom: 50,
        side: 'bothSides',
        distLeft: 0,
        distRight: 0,
        distTop: 0,
        distBottom: 0,
        paraId: 1,
      }],
    });

    const placed = pushFloatRect(st as never, {
      x: 20,
      y: 10,
      w: 45,
      h: 10,
      dl: 0,
      dr: 8,
      dt: 0,
      db: 0,
      kind: 'frame',
      mode: 'square',
      side: 'bothSides',
      imageKey: 'moving',
      paraId: 2,
      avoidOverlap: true,
    });

    expect(placed).toMatchObject({ imageX: 20, imageY: 50 });
    expect(placed.xRight).toBeLessThanOrEqual(100.5);
  });

  it('keeps an overlap-permitted DrawingML anchor at its resolved position (issue #1623)', () => {
    const st = makeState({
      pageWidth: 100,
      floats: [{
        kind: 'shape',
        mode: 'square',
        imageKey: 'blocker',
        imageX: 0,
        imageY: 0,
        imageW: 50,
        imageH: 50,
        xLeft: 0,
        xRight: 50,
        yTop: 0,
        yBottom: 50,
        side: 'bothSides',
        distLeft: 0,
        distRight: 0,
        distTop: 0,
        distBottom: 0,
        paraId: 1,
      }],
    });

    const placed = pushFloatRect(st as never, {
      x: 20,
      y: 10,
      w: 45,
      h: 10,
      dl: 0,
      dr: 8,
      dt: 0,
      db: 0,
      kind: 'shape',
      mode: 'square',
      side: 'bothSides',
      imageKey: 'moving',
      paraId: 2,
      allowOverlap: true,
      avoidOverlap: true,
    });

    expect(placed).toMatchObject({ imageX: 20, imageY: 10 });
  });
});

describe('frame geometry (§17.3.1.11) — wrap modes', () => {
  const st0 = () => makeState();
  const dc = (wrap: FramePr['wrap']) => frame({ dropCap: 'drop', lines: 3, wrap });

  it('wrap="notBeside" → topAndBottom float (text never beside the frame)', () => {
    const st = st0();
    const fp = dc('notBeside');
    const b = box(fp, st, 200, 42, 50, 14);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    expect(st.floats[0].mode).toBe('topAndBottom');
  });

  it('wrap="around" and "auto" → square float (auto ≡ around in Word)', () => {
    for (const w of ['around', 'auto'] as const) {
      const st = st0();
      const fp = dc(w);
      const b = box(fp, st, 200, 42, 50, 14);
      registerFloat(b, fp, st);
      expect(st.floats, `wrap=${w}`).toHaveLength(1);
      expect(st.floats[0].mode, `wrap=${w}`).toBe('square');
    }
  });

  it('wrap="tight" and "through" → square float (rectangle, no contour follow)', () => {
    for (const w of ['tight', 'through'] as const) {
      const st = st0();
      const fp = dc(w);
      const b = box(fp, st, 200, 42, 50, 14);
      registerFloat(b, fp, st);
      expect(st.floats, `wrap=${w}`).toHaveLength(1);
      expect(st.floats[0].mode, `wrap=${w}`).toBe('square');
    }
  });

  it('wrap="none" → topAndBottom float, as notBeside (§17.18.104: next line clear of the frame)', () => {
    // ST_Wrap of w:framePr, not DrawingML wrapNone: text neither wraps
    // beside nor overlaps the frame.
    const st = st0();
    const fp = dc('none');
    const b = box(fp, st, 200, 42, 50, 14);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    expect(st.floats[0].mode).toBe('topAndBottom');
  });
});

describe('frame geometry (§17.3.1.11) — hAnchor / vAnchor containers', () => {
  // A generic (non-drop-cap) frame at an absolute x/y from each anchor base.
  it('hAnchor="text" anchors x against the COLUMN band (contentX)', () => {
    const st = makeState({ contentX: 250, contentW: 200 }); // a right-hand column
    const fp = frame({ hAnchor: 'text', x: 10 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.x).toBe(250 + 10); // column left + x offset
  });

  it('hAnchor="margin" anchors x against the page content margin', () => {
    const st = makeState();
    const fp = frame({ hAnchor: 'margin', x: 10 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.x).toBe(100 + 10); // marginLeft + x
  });

  it('hAnchor="page" anchors x against the physical page edge', () => {
    const st = makeState();
    const fp = frame({ hAnchor: 'page', x: 10 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.x).toBe(0 + 10); // page left + x
  });

  it('vAnchor="text" anchors y at the paragraph top', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'text', y: 5 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.y).toBe(300 + 5); // paraTop + y (yAlign ignored when vAnchor=text)
  });

  it('vAnchor="margin" anchors y at the top content margin', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', y: 5 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.y).toBe(72 + 5); // marginTop + y
  });

  it('vAnchor="page" anchors y at the physical page top', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'page', y: 5 });
    const b = box(fp, st, 300, 60, 40, 12);
    expect(b.y).toBe(0 + 5); // page top + y
  });
});

describe('frame geometry (§17.3.1.11 / §22.9.2.20) — yAlign is vAnchor-band relative', () => {
  // ST_YAlign positions the frame relative to the ANCHOR OBJECT (the vAnchor
  // band), NOT the physical page (§22.9.2.20: "relative position … relative to
  // the vertical anchor"). The band per §17.18.100:
  //   page   → [0, pageH]                       (page edges)
  //   margin → [marginTop, pageH−marginBottom]  (text margins)
  //   text   → [paraTop, paraTop+contentH]      (anchor paragraph text extents)
  // yAlign is ignored for vAnchor="text" (relative positioning not allowed).

  it('vAnchor="margin" + yAlign="center" centers in the MARGIN band, not the page', () => {
    const st = makeState(); // margin band [72, 728], height 656
    const fp = frame({ vAnchor: 'margin', yAlign: 'center', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    // start + (end − start − h)/2 = 72 + (728 − 72 − 60)/2 = 72 + 298 = 370
    expect(b.y).toBe(72 + (728 - 72 - 60) / 2);
  });

  it('vAnchor="margin" + yAlign="center": ASYMMETRIC margins center in the margin band, NOT the page', () => {
    // §22.9.2.20: with vAnchor="margin", yAlign="center" centers in the margin
    // BAND [marginTop, pageH−marginBottom], which only equals the page centre
    // when margins are symmetric. The symmetric cases above (marginTop=
    // marginBottom=72) cannot distinguish "margin band centre" from "page
    // centre"; this asymmetric case pins the band-relative behaviour.
    const st = makeState({ marginTop: 40, marginBottom: 120 }); // band [40, 680], height 640
    const fp = frame({ vAnchor: 'margin', yAlign: 'center', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    // band centre = marginTop + (pageH − marginTop − marginBottom − h)/2
    //             = 40 + (800 − 40 − 120 − 60)/2 = 40 + 290 = 330
    expect(b.y).toBe(40 + (800 - 40 - 120 - 60) / 2); // 330 — NOT page centre (800−60)/2 = 370
    expect(b.y).not.toBe((800 - 60) / 2); // explicit: this is the band centre, not the page centre
  });

  it('vAnchor="margin" + yAlign="bottom" sits flush to the bottom margin', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', yAlign: 'bottom', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(728 - 60); // (pageH − marginBottom) − h
  });

  it('vAnchor="margin" + yAlign="outside" sits flush to the bottom margin', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', yAlign: 'outside', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(728 - 60);
  });

  it('vAnchor="margin" + yAlign="top"/"inside" sits at the margin top (band start)', () => {
    const st = makeState();
    for (const ya of ['top', 'inside'] as const) {
      const fp = frame({ vAnchor: 'margin', yAlign: ya, hRule: 'exact', h: 60 });
      const b = box(fp, st, 300, 40, 100, 12);
      expect(b.y, ya).toBe(72); // band start = marginTop
    }
  });

  it('vAnchor="page" + yAlign="center" centers over the FULL page (no margin offset)', () => {
    const st = makeState(); // page band [0, 800]
    const fp = frame({ vAnchor: 'page', yAlign: 'center', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe((800 - 60) / 2); // 370 — NOT (800−60)/2 − marginTop
  });

  it('vAnchor="page" + yAlign="bottom" sits flush to the physical page bottom', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'page', yAlign: 'bottom', hRule: 'exact', h: 60 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(800 - 60); // pageH − h
  });

  it('explicit y is measured from the vAnchor band start (margin top)', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'margin', y: 5 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(72 + 5); // band start (marginTop) + y
  });

  it('yAlign is ignored for vAnchor="text" (relative positioning not allowed)', () => {
    const st = makeState();
    const fp = frame({ vAnchor: 'text', yAlign: 'center', y: 7 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(300 + 7); // paraTop + y; yAlign ignored
  });
});

// Word ground truth (private/sample-17 Sec B, Word-exported PDF via pdftotext):
// a vAnchor="page" frame whose requested `y` would push its BOTTOM past the
// physical page edge is shifted UP so its bottom sits flush on the page bottom
// (measured top 741.9pt = 841.9 − 100 for a 100pt frame), NOT left overflowing.
// computeFrameBox clamps it via clampAbsBoxIntoContainer — identical to the
// floating-table clamp (float-table-geometry.test.ts). vAnchor="text" is NOT
// clamped (its overflow is the paginator's keep-with-anchor).
describe('frame geometry (§17.3.1.11) — page/margin-anchored clamp (Word ground truth)', () => {
  it('vAnchor="page": a frame overflowing the page bottom is clamped up to pageH − frameH', () => {
    // The sample-17 Sec B geometry at A4 (pageH 841.9, frame 100pt): y=775 would
    // put the bottom at 875 > 841.9 ⇒ clamp to y = 841.9 − 100 = 741.9.
    const st = makeState({ pageH: 841.9 });
    const fp = frame({ vAnchor: 'page', hRule: 'exact', h: 100, y: 775 });
    const b = box(fp, st, 300, 250, 100, 12);
    expect(b.y).toBeCloseTo(841.9 - 100, 6); // 741.9 — clamped to the page bottom
  });

  it('vAnchor="page": a frame that already fits is NOT moved (clamp is idempotent)', () => {
    const st = makeState(); // pageH 800
    const fp = frame({ vAnchor: 'page', hRule: 'exact', h: 60, y: 5 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(5); // 5 + 60 = 65 ≤ 800 ⇒ unchanged
  });

  it('vAnchor="page": a frame TALLER than the page pins to the top (floor = page top)', () => {
    // frameH 900 > pageH 800: pageH − frameH = −100 would push it above the page
    // top, so the floor (containerStart = 0) wins — it overflows the bottom instead.
    const st = makeState(); // pageH 800
    const fp = frame({ vAnchor: 'page', hRule: 'exact', h: 900, y: 50 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(0); // clamped to the page top, not −100
  });

  it('vAnchor="margin": a frame overflowing the bottom margin is clamped to (pageH − marginBottom) − frameH', () => {
    // margin band [72, 728]: y=700 ⇒ frameY=772, bottom 872 > 728 ⇒ clamp to
    // 728 − 100 = 628 (container end = bottom text margin, ASSUMED — no fixture
    // pins where Word clamps a margin-anchored overflow; symmetric with page).
    const st = makeState(); // margin band [72, 728]
    const fp = frame({ vAnchor: 'margin', hRule: 'exact', h: 100, y: 700 });
    const b = box(fp, st, 300, 40, 100, 12);
    expect(b.y).toBe(728 - 100); // 628 — clamped to the margin band bottom
  });

  it('vAnchor="text": an overflowing frame is NOT clamped (paginator keep-with-anchor handles it)', () => {
    // A vAnchor="text" frame rides the flow cursor; the paginator relocates it (and
    // its anchor context) when it overflows, so the geometry must leave the box at
    // paraTop + y even past the page — clamping here would fight that. paraTop 780,
    // y 0, h 100 ⇒ y stays 780 (bottom 880).
    const st = makeState(); // pageH 800
    const fp = frame({ vAnchor: 'text', hRule: 'exact', h: 100, y: 0 });
    const b = box(fp, st, 780, 40, 100, 12);
    expect(b.y).toBe(780); // NOT clamped — text-anchored is the paginator's job
  });
});

describe('frame geometry (§17.3.1.11) — generic frame (dropCap="none") sizing', () => {
  it('hRule="exact" forces the frame height to h regardless of content', () => {
    const st = makeState();
    const fp = frame({ hRule: 'exact', h: 30 });
    const b = box(fp, st, 200, 60, 100, 12); // contentH 100 ignored
    expect(b.h).toBe(30);
  });

  it('hRule="atLeast" takes max(h, content height)', () => {
    const st = makeState();
    expect(box(frame({ hRule: 'atLeast', h: 30 }), st, 200, 60, 100, 12).h).toBe(100);
    expect(box(frame({ hRule: 'atLeast', h: 200 }), st, 200, 60, 100, 12).h).toBe(200);
  });

  it('hRule="auto" uses the content height; explicit w forces exact width', () => {
    const st = makeState();
    const b = box(frame({ hRule: 'auto', w: 150 }), st, 200, 60, 80, 12);
    expect(b.h).toBe(80); // content height
    expect(b.w).toBe(150); // explicit width supersedes natural content width
  });

  it('xAlign="center" centers the frame in the hAnchor band, superseding x', () => {
    const st = makeState(); // text band [100,500], width 400
    const fp = frame({ hAnchor: 'text', x: 999, xAlign: 'center', w: 100 });
    const b = box(fp, st, 200, 100, 40, 12);
    expect(b.x).toBe(100 + (400 - 100) / 2); // centered, x ignored
  });

  it('xAlign="right" right-aligns the frame in the hAnchor band', () => {
    const st = makeState();
    const fp = frame({ hAnchor: 'text', xAlign: 'right', w: 100 });
    const b = box(fp, st, 200, 100, 40, 12);
    expect(b.x).toBe(500 - 100); // band right − width
  });

  it('a generic frame emits a bothSides square float (text wraps either side)', () => {
    const st = makeState();
    const fp = frame({ dropCap: 'none', hRule: 'exact', h: 40, w: 120, x: 50 });
    const b = box(fp, st, 200, 120, 40, 12);
    registerFloat(b, fp, st);
    expect(st.floats).toHaveLength(1);
    expect(st.floats[0].side).toBe('bothSides');
  });
});

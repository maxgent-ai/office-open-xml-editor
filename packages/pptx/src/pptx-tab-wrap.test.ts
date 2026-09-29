import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import type { TextBody, Paragraph } from './types';
import type { TextRunData, TabStop } from '@silurus/ooxml-core';

// issue #1006 — a tab-led line MUST wrap (PowerPoint wraps tab-led paragraphs in
// every script and alignment). Adjudicated against the PowerPoint ground-truth
// PDF (sample-21, 6 slides): a tab is a horizontal pen JUMP to its stop within
// the visual line where the tab occurs; content overflowing the line's right
// edge wraps at a normal break opportunity; and every CONTINUATION line
// re-anchors at the leading text-inset edge (text-left) — identical to a no-tab
// paragraph. The tab-stop ORIGIN (explicit a:tabLst vs default defTabSz grid)
// does not change the continuation anchor. Latin and Thai behave identically.
//
// Deterministic linear metrics (like pptx-multi-tab.test.ts): every glyph is a
// fixed `FONT_PX`-wide cell, so "t01" = 60px, a space = 20px. Box widths are
// chosen so the wrap point PROVES tab-awareness (line 1 fits FEWER tokens than a
// tab-blind budget would, because the tab jump consumes the leading space).

const FONT_PX = 20;
const SCALE = 1 / 12700; // emuToPx(emu, SCALE) = emu·SCALE; 1pt → 1px
const px = (n: number): number => n * 12700; // px → EMU at SCALE

function mockCtx(): {
  ctx: CanvasRenderingContext2D;
} {
  let font = `${FONT_PX}px serif`;
  let letterSpacing = '0px';
  let fillStyle = '';
  let direction: CanvasDirection = 'ltr';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? String(FONT_PX));
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; }, set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; }, set direction(v: CanvasDirection) { direction = v; },
    get letterSpacing() { return letterSpacing; }, set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = pxOf();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText: () => {},
    strokeText: () => {},
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
    setLineDash: () => {}, closePath: () => {}, arc: () => {},
    strokeStyle: '#000', lineWidth: 1, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D };
}

function run(text: string): TextRunData {
  return {
    type: 'text', text, bold: null, italic: null, underline: false,
    strikethrough: false, fontSize: FONT_PX, color: '000000', fontFamily: 'Serif',
  } as TextRunData;
}

function body(
  runs: TextRunData[],
  opts: {
    tabStops?: TabStop[];
    defTabSz?: number;
    rtl?: boolean;
    algn?: string;
    marL?: number;
    indent?: number;
  } = {},
): TextBody {
  const rtl = opts.rtl ?? false;
  const para: Paragraph = {
    alignment: opts.algn ?? (rtl ? 'r' : 'l'),
    marL: opts.marL ?? 0, marR: 0, indent: opts.indent ?? 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null, defBold: null, defItalic: null,
    defFontFamily: null, tabStops: opts.tabStops ?? [], defTabSz: opts.defTabSz, rtl, runs,
  } as unknown as Paragraph;
  return {
    verticalAnchor: 't', paragraphs: [para], defaultFontSize: FONT_PX,
    defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0,
    wrap: 'square', vert: 'horz', autoFit: 'none',
  } as unknown as TextBody;
}

type RunInfo = { text: string; x: number; y: number; w: number };

/** Render and return drawn runs grouped into visual lines (by inShapeY). Each
 *  line lists its runs left-to-right; `startX` is the first run's inShapeX. */
function renderLines(tb: TextBody, boxW: number): {
  runs: RunInfo[];
  lines: { y: number; startX: number; text: string; runs: RunInfo[] }[];
} {
  const { ctx } = mockCtx();
  const runs: RunInfo[] = [];
  renderTextBody(
    ctx, tb, 0, 0, boxW, 400, SCALE,
    null, 0, false, false, '#000000', undefined,
    { themeMajorFont: null, themeMinorFont: null, dpr: 1 },
    (r) => runs.push({ text: r.text, x: r.inShapeX, y: r.inShapeY, w: r.w }),
  );
  const byY = new Map<number, RunInfo[]>();
  for (const r of runs) {
    const key = Math.round(r.y);
    (byY.get(key) ?? byY.set(key, []).get(key)!).push(r);
  }
  const lines = [...byY.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([y, rs]) => {
      const sorted = [...rs].sort((a, b) => a.x - b.x);
      // Consecutive same-format tokens in a cell merge into ONE drawn run, so a
      // line's text is the concatenation of its runs (e.g. "t01 t02 ").
      return {
        y,
        startX: sorted[0].x,
        text: sorted.map((r) => r.text).join('').trim(),
        runs: sorted,
      };
    });
  return { runs, lines };
}

const LATIN = '\tt01 t02 t03 t04 t05 t06 t07 t08 t09 t10 t11 t12';

describe('issue #1006 — tab-led lines wrap; continuation re-anchors at text-left', () => {
  it('keeps an authored Latin run seam inside the wrapped word (control C01)', () => {
    const { lines } = renderLines(body([run('abc'), run('def')]), 90);
    expect(lines.map((line) => line.text)).toEqual(['abcd', 'ef']);
  });
  it('Latin leading tab wraps (tab-aware budget) and continuation starts at text-left', () => {
    // Left stop @100. Budget 280. Tab jumps to 100, leaving 180px on line 1:
    // "t01"(60)+" "(20)+"t02"(60) = 140 fits, +" "(20)=160, "t03" overflows.
    // A TAB-BLIND budget would fit t01,t02,t03 (60/140/220) — so exactly TWO
    // content tokens on line 1 proves the jump is charged to the wrap budget.
    const { lines } = renderLines(
      body([run(LATIN)], { tabStops: [{ pos: px(100), algn: 'l' }] }),
      280,
    );
    expect(lines.length, 'tab-led paragraph must wrap into multiple lines').toBeGreaterThan(1);
    // Line 1 begins at the tab stop; holds exactly t01, t02 (a tab-blind budget
    // of 280 would have fit t01 t02 t03 — proving the jump is charged).
    expect(lines[0].startX).toBeCloseTo(100, 3);
    expect(lines[0].text).toBe('t01 t02');
    // Continuation re-anchors at text-left (x≈0) and leads with t03.
    expect(lines[1].startX).toBeCloseTo(0, 3);
    expect(lines[1].text.startsWith('t03')).toBe(true);
  });

  it('wide box fits more line-1 tokens but still re-anchors continuation at text-left', () => {
    // Same stop @100, budget 520. Natural extent = 100 (jump) + content: five
    // tokens "t01…t05" = 5·60+4·20 = 380 (+jump 480 ≤520); the trailing space →
    // 500 ≤520; "t06" → 560 > 520 wraps. So exactly five tokens on line 1.
    const { lines } = renderLines(
      body([run(LATIN)], { tabStops: [{ pos: px(100), algn: 'l' }] }),
      520,
    );
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[0].startX).toBeCloseTo(100, 3);
    expect(lines[0].text).toBe('t01 t02 t03 t04 t05');
    expect(lines[1].startX).toBeCloseTo(0, 3);
    expect(lines[1].text.startsWith('t06')).toBe(true);
  });

  it('Thai no-space leading tab reaches SEA fill: wraps and continuation at text-left', () => {
    // A long no-space Thai run after a leading tab. Before the fix this stayed on
    // ONE line forever (never reached the SEA branch); now it wraps and every
    // continuation line starts at text-left.
    const thai = '\tการเขียนภาษาไทยไม่ใช้ช่องว่างระหว่างคำแต่ใช้ช่องว่างเฉพาะเมื่อจบประโยคหรือวลี';
    const { lines } = renderLines(
      body([run(thai)], { tabStops: [{ pos: px(100), algn: 'l' }] }),
      280,
    );
    expect(lines.length, 'tab-led Thai must reach SEA fill and wrap').toBeGreaterThan(1);
    expect(lines[0].startX).toBeCloseTo(100, 3); // line 1 at the tab stop
    for (let i = 1; i < lines.length; i++) {
      expect(lines[i].startX, `Thai continuation line ${i} at text-left`).toBeCloseTo(0, 3);
    }
  });

  it('multi-tab: cells sit on their stops; the overflowing cell tail re-anchors at text-left', () => {
    // A1<tab>B2<tab>{t01…} with left stops @100 and @200, budget 260.
    // A1@0, B2@100, t01@200 (ends 260); the tail wraps to text-left.
    const r = run('A1\tB2\tt01 t02 t03 t04 t05 t06 t07 t08 t09 t10');
    const { lines } = renderLines(
      body([r], { tabStops: [{ pos: px(100), algn: 'l' }, { pos: px(200), algn: 'l' }] }),
      260,
    );
    expect(lines.length).toBeGreaterThan(1);
    const first = lines[0];
    const a1 = first.runs.find((x) => x.text === 'A1')!;
    const b2 = first.runs.find((x) => x.text === 'B2')!;
    const t01 = first.runs.find((x) => x.text === 't01')!;
    expect(a1.x).toBeCloseTo(0, 3);
    expect(b2.x).toBeCloseTo(100, 3);
    expect(t01.x).toBeCloseTo(200, 3);
    // The overflowing 3rd cell's tail continues at text-left.
    expect(lines[1].startX).toBeCloseTo(0, 3);
    expect(lines[1].text.startsWith('t02')).toBe(true);
  });

  it('default tab grid (defTabSz, no a:tabLst): leading tab lands on the 1-grid stop', () => {
    // No explicit stops. defTabSz = 100px grid → the leading tab jumps to 100
    // (NOT to a single space width). Continuation re-anchors at text-left.
    const { lines } = renderLines(
      body([run(LATIN)], { defTabSz: px(100) }),
      280,
    );
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[0].startX, 'leading tab lands on the default 1-inch grid stop').toBeCloseTo(100, 3);
    expect(lines[0].text).toBe('t01 t02');
    expect(lines[1].startX).toBeCloseTo(0, 3);
  });

  it('a right-aligned tab cell that FITS never wraps (preserves the demo "22%" invariant)', () => {
    // "ab"<tab>"value", right stop @200, tight budget 220. The cell ends AT 200,
    // so it fits; a tab-blind "jump-then-add" budget would compute 300 and wrongly
    // wrap. Exactly one line; value ends on the stop.
    const { lines } = renderLines(
      body([run('ab\tvalue')], { tabStops: [{ pos: px(200), algn: 'r' }] }),
      220,
    );
    expect(lines.length).toBe(1);
    const value = lines[0].runs.find((x) => x.text === 'value')!;
    expect(value.x + value.w).toBeCloseTo(200, 3);
  });

  it('a right-tab CJK cell that fits by shrinking its gap does NOT wrap early', () => {
    // Leading right tab @200, then a 9-glyph CJK run (180px). The cell ends AT
    // 200 (occupying [20,200]) and fits the 220 box. The CJK branch must use the
    // tab-aware available width (200px absorbable via the shrinking gap), not the
    // additive 20px remainder — otherwise it would falsely wrap after one glyph.
    const { lines } = renderLines(
      body([run('\t日本語日本語日本語')], { tabStops: [{ pos: px(200), algn: 'r' }] }),
      220,
    );
    expect(lines.length, 'fitting right-tab CJK cell stays on one line').toBe(1);
    const cjk = lines[0].runs[lines[0].runs.length - 1];
    expect(cjk.x + cjk.w).toBeCloseTo(200, 1); // right-aligned to the stop
  });

  it('a leading tab whose jump exceeds the box does not emit a tab-only CJK line', () => {
    // Left stop @300 in a 260 box: the jump alone overflows, so no CJK glyph fits
    // past it. The tab-led line must still carry a glyph (overflow) rather than be
    // finalised empty, and the remaining glyphs wrap to text-left.
    const { lines } = renderLines(
      body([run('\t日本語日本語日本語日本語')], { tabStops: [{ pos: px(300), algn: 'l' }] }),
      260,
    );
    expect(lines.length).toBeGreaterThan(1);
    expect(lines[0].text.length, 'the tab-led line carries a glyph, not tab-only').toBeGreaterThan(0);
    expect(lines[1].startX).toBeCloseTo(0, 3); // continuation at text-left
  });
});

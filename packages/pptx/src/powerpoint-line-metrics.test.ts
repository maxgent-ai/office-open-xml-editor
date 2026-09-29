import { describe, expect, it } from 'vitest';
import type { TextRunData } from '@silurus/ooxml-core';
import { renderTextBody } from './renderer.js';
import {
  powerPointAscentShare, powerPointShareCacheSize, SHARE_CACHE_LIMIT,
} from './powerpoint-line-metrics.js';
import type { Paragraph, TextBody } from './types.js';

// Expected values are PowerPoint 16.113.2 PDF baselines of the #1610
// controls, counted in whole 1/100 in from the text-area top (tIns 0 here).
// At this scale 1 pt = 1 canvas unit, so N units of 1/100 in are N × 0.72.
// The export quantizes each baseline to that device unit; layout keeps
// continuous positions, so each must lie within half a unit of the export.
const SCALE = 1 / 12700;
const U = 0.72;

type Spacing = Paragraph['spaceLine'];
interface RunSpec { text: string; font: string; size: number }

function context() {
  const draws: Array<{ text: string; y: number }> = [];
  let font = '';
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    fillStyle: '', strokeStyle: '', direction: 'ltr', textAlign: 'left', textBaseline: 'alphabetic',
    measureText: (text: string) => ({ width: [...text].length * 10, actualBoundingBoxAscent: 7, actualBoundingBoxDescent: 2 }),
    fillText: (text: string, _x: number, y: number) => draws.push({ text, y }),
    fillRect: () => {}, drawImage: () => {}, save: () => {}, restore: () => {},
    translate: () => {}, rotate: () => {}, scale: () => {}, beginPath: () => {},
    moveTo: () => {}, lineTo: () => {}, stroke: () => {}, clip: () => {}, rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, draws };
}

function paragraph(runs: RunSpec[], extra: Partial<Paragraph> = {}): Paragraph {
  return {
    alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null, spaceAfter: null,
    spaceLine: { type: 'pct', val: 100000 }, lvl: 0, bullet: { type: 'none' },
    defFontSize: null, defColor: null, defBold: null, defItalic: null, defFontFamily: null,
    tabStops: [], eaLnBrk: true,
    runs: runs.map((r): TextRunData => ({
      type: 'text', text: r.text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: r.size, color: '000000', fontFamily: r.font, fontFamilyEa: r.font,
    } as TextRunData)),
    ...extra,
  } as Paragraph;
}

function baselines(paragraphs: Paragraph[], over: Partial<TextBody> = {}, boxHeight = 800): number[] {
  const body = {
    verticalAnchor: 't', paragraphs, defaultFontSize: 18, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'none', vert: 'horz', autoFit: 'none', ...over,
  } as TextBody;
  const { ctx, draws } = context();
  renderTextBody(ctx, body, 0, 0, 2000, boxHeight, SCALE);
  const ys: number[] = [];
  for (const d of draws) if (!ys.some((y) => Math.abs(y - d.y) < 1e-6)) ys.push(d.y);
  return ys;
}

const lines = (font: string, size: number, n: number, spaceLine: Spacing = { type: 'pct', val: 100000 }, text = 'Hg') =>
  Array.from({ length: n }, (_, i) => paragraph([{ text: `${text}${i}`, font, size }], { spaceLine }));

/** Continuous baselines expressed in export units, snapped the way the
 * export snaps them (nearest unit), so they compare equal to the PDF. */
const units = (ys: number[]) => ys.map((y) => Math.floor(y / U + 0.5));

describe('PowerPoint text-box line metrics (#1610)', () => {
  it('splits each face by its usWin (or USE_TYPO_METRICS typo) ascent share', () => {
    // A-* controls: first and second baselines of two same-size lines.
    expect(units(baselines(lines('Arial', 72, 2)))).toEqual([97, 217]);
    expect(units(baselines(lines('Calibri', 72, 2)))).toEqual([94, 214]);
    expect(units(baselines(lines('Yu Gothic', 72, 2)))).toEqual([92, 212]);
    expect(units(baselines(lines('Gabriola', 72, 2)))).toEqual([98, 218]);
    expect(units(baselines(lines('Baskerville Old Face', 200, 2)))).toEqual([258, 591]);
    expect(units(baselines(lines('Stencil', 200, 2)))).toEqual([255, 589]);
    expect(units(baselines(lines('Times New Roman', 200, 2)))).toEqual([268, 602]);
    expect(units(baselines(lines('Georgia', 200, 2)))).toEqual([269, 602]);
    expect(units(baselines(lines('Courier New', 200, 2)))).toEqual([245, 578]);
  });

  it('follows PowerPoint substitutions for faces only macOS itself provides', () => {
    expect(units(baselines(lines('Palatino', 200, 2)))).toEqual([259, 593]);
    expect(units(baselines(lines('Helvetica', 200, 2)))).toEqual([270, 603]);
    expect(powerPointAscentShare('Helvetica Neue', false, false))
      .toBe(powerPointAscentShare('Arial', false, false));
    // Avenir, Menlo and Hiragino Sans fall back to deck-dependent faces: unresolved.
    expect(powerPointAscentShare('Avenir', false, false)).toBeUndefined();
    expect(powerPointAscentShare('Hiragino Sans', false, false)).toBeUndefined();
  });

  it('rounds each baseline, not the line pitch, to a whole 1/100 inch', () => {
    // 100 pt lines are 166.67 units high: Arial steps 167, Meiryo 166.
    expect(units(baselines(lines('Arial', 100, 2)))).toEqual([135, 302]);
    expect(units(baselines(lines('Meiryo', 100, 2, null, '日')))).toEqual([118, 284]);
    // Fractional sizes keep an exact 1.2 × size line (Z-*).
    expect(units(baselines(lines('Arial', 10.5, 8)))).toEqual([14, 32, 49, 67, 84, 102, 119, 137]);
    expect(units(baselines(lines('Meiryo', 13.33, 8, null, '日')))).toEqual([16, 38, 60, 82, 105, 127, 149, 171]);
  });

  it('unions the runs of a mixed line and rescales them into 1.2 × the largest size', () => {
    const mixed = (a: string, b: string) => baselines([
      paragraph([{ text: 'H1', font: a, size: 100 }]),
      paragraph([{ text: 'H', font: a, size: 100 }, { text: '日g2', font: b, size: 100 }]),
      paragraph([{ text: 'H3', font: a, size: 100 }]),
    ]);
    expect(units(mixed('Arial', 'Meiryo'))).toEqual([135, 289, 468]);
    expect(units(baselines([
      paragraph([{ text: 'H1', font: 'Arial', size: 40 }]),
      paragraph([{ text: 'H', font: 'Arial', size: 40 }, { text: 'Hg2', font: 'Arial', size: 100 }]),
      paragraph([{ text: 'H3', font: 'Arial', size: 40 }]),
    ]))).toEqual([54, 202, 287]);
  });

  it('re-divides spaced lines with the shared DrawingML rule and rounds spcPts to whole points', () => {
    const pts = (val: number): Spacing => ({ type: 'pts', val });
    expect(units(baselines(lines('Arial', 100, 2, pts(114))))).toEqual([127, 285]);
    expect(units(baselines(lines('Arial', 100, 2, pts(126))))).toEqual([131, 306]);
    // H equal to the natural 120 pt keeps the natural split (MS Gothic).
    expect(units(baselines(lines('MS Gothic', 100, 2, pts(120), '日')))).toEqual([143, 310]);
    expect(units(baselines(lines('Meiryo', 100, 2, pts(144), '日')))).toEqual([143, 343]);
    expect(units(baselines(lines('Meiryo', 100, 2, { type: 'pct', val: 150000 }, '日')))).toEqual([180, 430]);
    // P-*: 40.5 pt spacing lays out as 41 pt, 48.33 pt as 48 pt.
    expect(units(baselines(lines('Arial', 40, 8, pts(40.5))))).toEqual([44, 101, 158, 215, 272, 329, 386, 443]);
    expect(units(baselines(lines('Arial', 40, 8, pts(48.33))))).toEqual([54, 121, 187, 254, 321, 387, 454, 521]);
  });

  it('applies edge spacing only with spcFirstLastPara and anchors the unrounded block', () => {
    const edge = [
      paragraph([{ text: 'H1', font: 'Arial', size: 100 }], { spaceBefore: 7200 }),
      paragraph([{ text: 'H2', font: 'Arial', size: 100 }]),
    ];
    expect(units(baselines(edge))).toEqual([135, 302]);
    expect(units(baselines(edge, { spcFirstLastPara: true }))).toEqual([235, 402]);
    // G-Arial-ctr: two 100 pt lines centred in a 420 pt box, tIns/bIns 3.6 pt.
    const ctr = baselines(lines('Arial', 100, 2), { verticalAnchor: 'ctr', tIns: 45720, bIns: 45720 }, 420);
    // The anchored block top is not on the unit grid; the export counts from it.
    const top = 3.6 + (420 - 7.2 - 2 * 120) / 2;
    expect(units(ctr.map((y) => y - top)).map((n) => n + Math.floor((top - 3.6) / U + 0.5)))
      .toEqual([255, 422]);
  });

  it('sizes a line by the run latin face even when it draws no glyph, never by an unused ea face', () => {
    const probe = (mid: Paragraph) => units(baselines([
      paragraph([{ text: 'Hg1', font: 'Arial', size: 100 }]), mid,
      paragraph([{ text: 'Hg3', font: 'Arial', size: 100 }]),
    ]));
    const run = (text: string, latin: string, ea: string) => {
      const p = paragraph([{ text, font: latin, size: 100 }]);
      (p.runs[0] as { fontFamilyEa: string }).fontFamilyEa = ea;
      return p;
    };
    // supplement-3 U-*: ideographs only, latin face unused but counted.
    expect(probe(run('日本語', 'Calibri', 'MS PGothic'))).toEqual([135, 299, 468]);
    expect(probe(run('日本語', 'Meiryo', 'MS PGothic'))).toEqual([135, 291, 468]);
    // supplement-2 R-*: Latin only, the unused Meiryo ea face does not count.
    expect(probe(run('Hg2', 'Arial', 'Meiryo'))).toEqual([135, 302, 468]);
  });

  it('keeps each run latin slot when adjacent runs share a drawn face (colour never moves a line)', () => {
    // Two ideograph runs drawn in Meiryo, latin slots Calibri and Gabriola.
    const pair = (secondColor: string) => {
      const p = paragraph([
        { text: '日本', font: 'Calibri', size: 60 },
        { text: '語学', font: 'Gabriola', size: 60 },
      ]);
      for (const r of p.runs) (r as { fontFamilyEa: string }).fontFamilyEa = 'Meiryo';
      (p.runs[1] as { color: string }).color = secondColor;
      return baselines([p, paragraph([{ text: 'Hg', font: 'Arial', size: 60 }])]);
    };
    const same = pair('000000');
    const recoloured = pair('FF0000');
    expect(same).toHaveLength(2);
    expect(recoloured[0]).toBeCloseTo(same[0], 9);
    expect(recoloured[1]).toBeCloseTo(same[1], 9);
    // Gabriola's slot contributes: the pair differs from the Calibri-only line.
    const calibriOnly = paragraph([{ text: '日本語学', font: 'Calibri', size: 60 }]);
    (calibriOnly.runs[0] as { fontFamilyEa: string }).fontFamilyEa = 'Meiryo';
    const single = baselines([calibriOnly, paragraph([{ text: 'Hg', font: 'Arial', size: 60 }])]);
    expect(Math.abs(single[0] - same[0])).toBeGreaterThan(0.5);
  });

  it('bounds the share cache while many distinct face names are queried', () => {
    for (let i = 0; i < SHARE_CACHE_LIMIT * 4; i++) {
      powerPointAscentShare(`Unresolved Face ${i}`, i % 2 === 0, i % 3 === 0);
      expect(powerPointShareCacheSize()).toBeLessThanOrEqual(SHARE_CACHE_LIMIT);
    }
    expect(powerPointShareCacheSize()).toBe(SHARE_CACHE_LIMIT);
    // Known faces still resolve after eviction.
    expect(powerPointAscentShare('Arial', false, false)).toBeCloseTo(1854 / 2288, 12);
  });

  it('anchors by the last line natural descent and applies normAutofit to the split', () => {
    const box = { tIns: 45720, bIns: 45720 };
    const rel = (ys: number[]) => ys.map((y) => (y - 3.6) / U);
    // Q-Arial-b-p150: two 100 pt lines at 150 %, bottom of a 420 pt box.
    const q = rel(baselines(lines('Arial', 100, 2, { type: 'pct', val: 150000 }), { ...box, verticalAnchor: 'b' }, 420));
    expect(Math.abs(q[0] - 292.23)).toBeLessThanOrEqual(0.52);
    expect(Math.abs(q[1] - 542.21)).toBeLessThanOrEqual(0.52);
    // N-Arial-90-10-t: normAutofit fontScale 90 %, lnSpcReduction 10 %.
    const n = baselines(lines('Arial', 40 * 0.9, 4), { ...box, autoFit: 'norm', lnSpcReduction: 0.1 }, 260);
    expect(units(n.map((y) => y - 3.6))).toEqual([43, 97, 151, 205]);
  });

  it('keeps the ordinary line model when any run has an unresolved face', () => {
    const ys = baselines([paragraph([{ text: 'H', font: 'Arial', size: 20 }, { text: 'g', font: 'Avenir', size: 20 }])]);
    expect(ys[0]).toBeCloseTo(20 * 1.2 * 0.8, 5);
  });
});

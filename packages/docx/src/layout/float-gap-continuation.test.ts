import { describe, expect, it } from 'vitest';
import { layoutLines, lineBoxHeight, type LayoutTextSeg, type WrapLayoutCtx } from '../line-layout.js';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { runLineBreakerPass } from '../line-breaker/pass-driver.js';
import { convergeLineWrap } from './line-wrap-convergence.js';
import type { FloatRect } from './float-wrap.js';

function context(): CanvasRenderingContext2D {
  return { measureText: (text: string) => ({ width: text.length * 5,
    fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }) } as unknown as CanvasRenderingContext2D;
}
function token(text: string): LayoutTextSeg {
  return { text, fontSize: 10, fontFamily: 'Arial', bold: false, italic: false,
    underline: false, strikethrough: false, color: null, vertAlign: null, measuredWidth: 0 };
}
function obstacle(left: number, right: number, authoredWrap: 'square' | 'tight' | 'through'): FloatRect {
  return { kind: 'shape', mode: 'square', authoredWrap, imageKey: 'test',
    imageX: left, imageY: 0, imageW: right - left, imageH: 60,
    xLeft: left, xRight: right, yTop: 0, yBottom: 60, side: 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0,
    ...(authoredWrap === 'square' ? {} : { wrapPolygon: [
      { xPt: left, yPt: 0 }, { xPt: right, yPt: 0 },
      { xPt: right, yPt: 60 }, { xPt: left, yPt: 60 },
    ] }) };
}
function wrap(floats: FloatRect[], rtl = false): WrapLayoutCtx {
  return { floats, paraX: 0, startPageY: 0, columnXPt: 0, columnWidthPt: 200,
    readingDirection: rtl ? 'rtl' : 'ltr', pageH: 800, lineBoxH: () => 10 };
}
describe('Word measured gap continuation (#1670)', () => {
  it.each(['square', 'tight', 'through'] as const)('fills three %s gaps on one baseline before advancing', (mode) => {
    const lines = layoutLines(context(), ['AAAA ', 'BBBB ', 'CCCC ', 'DDDD'].map(token),
      200, 0, 1, [], wrap([obstacle(40, 70, mode), obstacle(110, 160, mode)]));
    expect(lines.map(l => [l.topY, l.xOffset, l.availWidth])).toEqual([
      [0, 0, 40], [0, 70, 40], [0, 160, 40], [10, 0, 40],
    ]);
    expect(lines.flatMap(l => l.segments.map(s => 'text' in s ? s.text : '')).join('')).toBe('AAAA BBBB CCCC DDDD');
  });
  it('uses right-to-left gap order and restarts at the right on the next baseline', () => {
    const lines = layoutLines(context(), ['AAAA ', 'BBBB ', 'CCCC'].map(token),
      200, 0, 1, [], wrap([obstacle(40, 160, 'square')], true), {}, 0,
      undefined, undefined, undefined, undefined, true);
    expect(lines.map(l => [l.topY, l.xOffset])).toEqual([[0, 160], [0, 0], [10, 160]]);
  });
  it('skips gaps that cannot hold the next atom and keeps an unbroken word whole below the float', () => {
    const lines = layoutLines(context(), [token('UNBROKENWORDUNBROKENWORD')], 200, 0, 1, [],
      wrap([obstacle(100, 160, 'square')]));
    expect(lines).toHaveLength(1);
    expect(lines[0].topY).toBe(60);
    expect(lines[0].segments[0]).toMatchObject({ text: 'UNBROKENWORDUNBROKENWORD' });
  });
  it('admits a word across formatting seams as one atom', () => {
    const lines = layoutLines(context(), [token('UNBROKEN'), { ...token('WORD'), joinPrev: true }],
      200, 0, 1, [], wrap([obstacle(40, 160, 'square')]));
    expect(lines.map(line => line.topY)).toEqual([60]);
    expect(lines[0].segments.map(segment => 'text' in segment ? segment.text : '').join('')).toBe('UNBROKENWORD');
  });
  it('shares the tallest gap metrics and advances one whole physical line', () => {
    const ctx = context();
    ctx.measureText = (text: string) => ({ width: text.length * 5,
      fontBoundingBoxAscent: text.startsWith('AAAA') ? 16 : 8,
      fontBoundingBoxDescent: text.startsWith('AAAA') ? 4 : 2 }) as TextMetrics;
    const w = wrap([obstacle(40, 160, 'square')]);
    w.lineBoxH = (ascent, descent) => ascent + descent;
    const lines = layoutLines(ctx, ['AAAA ', 'BBBB ', 'CCCC'].map(token), 200, 0, 1, [], w);
    expect(lines.map(l => [l.topY, l.ascent, l.descent])).toEqual([[0, 16, 4], [0, 16, 4], [20, 8, 2]]);
  });
  it('manual breaks advance vertically instead of continuing in the next gap', () => {
    const lines = layoutLines(context(), [token('AAAA'), { lineBreak: true, fontSize: 10, measuredWidth: 0 }, token('BBBB')],
      200, 0, 1, [], wrap([obstacle(40, 160, 'square')]));
    expect(lines.map(l => [l.topY, l.xOffset])).toEqual([[0, 0], [10, 0]]);
  });
});


describe('physical line allocation regressions', () => {
  it.each([20, 80])('probes all %i float gaps in the same pass', (count) => {
    const width = count * 40 + 20;
    const w = { ...wrap(Array.from({ length: count }, (_, i) => obstacle(i * 40 + 20, i * 40 + 40, 'square'))), columnWidthPt: width };
    let passes = 0;
    const lines = convergeLineWrap(probeHeights => {
      passes += 1;
      return runLineBreakerPass({ ctx: context(), segs: Array.from({ length: count + 1 }, () => token('AAA ')),
        maxWidth: width, firstIndent: 0, scale: 1, tabStops: [], wrapCtx: w,
        fontFamilyClasses: {}, tabOriginPx: 0, kinsoku: DEFAULT_KINSOKU_RULES,
        defaultTabPt: 36, marginRightPx: width, baseRtl: false, isJustified: false,
        stretchLastLine: false, widthPolicy: 'bounded', overflowPunct: true,
        passContext: { probeHeights },
      });
    }, () => 10);
    expect(passes).toBe(3);
    expect(lines.map(line => [line.topY, line.xOffset])).toEqual(Array.from({ length: count + 1 }, (_, i) => [0, i * 40]));
  });

  it.each([false, true])('applies picture auto leading with later-gap picture=%s', (laterGap) => {
    const w = wrap([obstacle(40, 160, 'square')]);
    w.lineBoxH = (a, d, r, i, e, g, u, p, l) => lineBoxHeight(
      { rule: 'auto', value: 1.5, explicit: true }, a, d, 1, undefined, r, i, e, g, undefined, u, p, l);
    const picture = { imagePath: 'picture', mimeType: 'image/png', anchor: false, anchorXPt: 0,
      anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false, inlinePicture: true as const,
      widthPt: 30, heightPt: 25, paragraphMarkSinglePx: 12, measuredWidth: 0 };
    const lines = layoutLines(context(), [
      ...(laterGap ? [token('BBBB '), picture] : [picture, token('BBBB ')]), token('CCCC'),
    ], 200, 0, 1, [], w);
    expect(lines.map(line => line.topY)).toEqual([0, 0, 33]);
    expect(lines.slice(0, 2).map(line => line.inlinePictureTextSingle)).toEqual([12, 12]);
  });

  it.each(['font', 'picture'])('keeps taller %s content beside admitted Latin grid text', (kind) => {
    const ctx = context();
    ctx.measureText = (text: string) => ({ width: text.length * 5,
      fontBoundingBoxAscent: text.startsWith('TALL') ? 32 : 8,
      fontBoundingBoxDescent: text.startsWith('TALL') ? 8 : 2 }) as TextMetrics;
    const w = { ...wrap([]), hasExclusions: false };
    w.lineBoxH = (a, d, r, i, e, g, u, p, l) => lineBoxHeight(null, a, d, 1,
      { type: 'lines', linePitchPt: 20 }, r, i, e, g, undefined, u, p, l);
    const lines = layoutLines(ctx, [
      { ...token('small '), resolvedLatinGridCellAllocation: true, resolvedLineHeightRatio: 1.15 },
      ...(kind === 'font' ? [{ ...token('TALL '), fontSize: 40 }] : [
        { imagePath: 'picture', mimeType: 'image/png', anchor: false, anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false, inlinePicture: true as const, widthPt: 20, heightPt: 100, paragraphMarkSinglePx: 12, measuredWidth: 0 },
      ]),
      { lineBreak: true, fontSize: 10, measuredWidth: 0 }, token('NEXT'),
    ], 200, 0, 1, [], w);
    expect(lines[1].topY).toBeGreaterThanOrEqual(kind === 'font' ? 40 : 100);
  });
});

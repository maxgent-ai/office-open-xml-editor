import { describe, expect, it } from 'vitest';
import {
  authoredExtraLinesHeight,
  authoredTextLines,
  fillAuthoredText,
  fillAuthoredTextDownward,
  measureAuthoredTextWidth,
} from './chart-text-lines.js';
import { drawRotatedCatLabel } from './shared/axis.js';

interface Call { op: string; args: unknown[] }

// One unit per code point (10 px) keeps widths deterministic.
function fakeCtx(baseline: CanvasTextBaseline = 'alphabetic'): { ctx: CanvasRenderingContext2D; calls: Call[] } {
  const calls: Call[] = [];
  const state = { textBaseline: baseline, textAlign: 'start' as CanvasTextAlign };
  const ctx = {
    get textBaseline() { return state.textBaseline; },
    set textBaseline(value: CanvasTextBaseline) { state.textBaseline = value; },
    get textAlign() { return state.textAlign; },
    set textAlign(value: CanvasTextAlign) { state.textAlign = value; },
    measureText: (text: string) => ({ width: Array.from(text).length * 10 }),
    fillText: (...args: unknown[]) => calls.push({ op: 'fillText', args }),
    save: () => calls.push({ op: 'save', args: [] }),
    restore: () => calls.push({ op: 'restore', args: [] }),
    translate: (...args: unknown[]) => calls.push({ op: 'translate', args }),
    rotate: (...args: unknown[]) => calls.push({ op: 'rotate', args }),
  } as unknown as CanvasRenderingContext2D;
  return { ctx, calls };
}

const texts = (calls: Call[]) => calls.filter(call => call.op === 'fillText').map(call => call.args);

describe('authored chart text lines', () => {
  it('splits at every authored break, keeping interior blank lines and trimming outer breaks', () => {
    expect(authoredTextLines('FY2024\n1Q営業利益')).toEqual(['FY2024', '1Q営業利益']);
    expect(authoredTextLines('A\r\nB\rC')).toEqual(['A', 'B', 'C']);
    expect(authoredTextLines('\nA\n\nB\n')).toEqual(['A', '', 'B']);
    expect(authoredTextLines('  North  ')).toEqual(['  North  ']);
  });

  it('draws unbroken text with the exact single fillText call', () => {
    const { ctx, calls } = fakeCtx();
    fillAuthoredText(ctx, '  North  ', 3, 4, 12);
    expect(texts(calls)).toEqual([['  North  ', 3, 4]]);
  });

  it('stacks lines according to the baseline: top grows down, middle centres, bottom grows up', () => {
    for (const [baseline, ys] of [
      ['top', [100, 112]],
      ['hanging', [100, 112]],
      ['middle', [94, 106]],
      ['bottom', [88, 100]],
      ['alphabetic', [88, 100]],
    ] as const) {
      const { ctx, calls } = fakeCtx(baseline);
      fillAuthoredText(ctx, 'A\nB', 5, 100, 12);
      expect(texts(calls)).toEqual([['A', 5, ys[0]], ['B', 5, ys[1]]]);
    }
  });

  it('keeps an interior blank line as vertical space and elides each line separately', () => {
    const { ctx, calls } = fakeCtx('top');
    fillAuthoredText(ctx, 'ABCDEF\n\nAB', 0, 0, 10, 40);
    expect(texts(calls)).toEqual([['ABC…', 0, 0], ['AB', 0, 20]]);
  });

  it('downward variant keeps the first line on the anchor whatever the baseline', () => {
    const { ctx, calls } = fakeCtx('middle');
    fillAuthoredTextDownward(ctx, 'A\nB\nC', 0, 50, 10);
    expect(texts(calls).map(args => args[2])).toEqual([50, 60, 70]);
  });

  it('measures the widest line and the extra block height', () => {
    const { ctx } = fakeCtx();
    expect(measureAuthoredTextWidth(ctx, 'AB\nABCD')).toBe(40);
    expect(measureAuthoredTextWidth(ctx, 'AB ABCD')).toBe(70);
    expect(authoredExtraLinesHeight(['A', 'B\nC\nD'], 12)).toBe(24);
    expect(authoredExtraLinesHeight(['A', 'B'], 12)).toBe(0);
    expect(authoredExtraLinesHeight(['A\nB'], 12, Math.PI / 3)).toBeCloseTo(6, 10);
  });

  it('rotated category labels stack their lines perpendicular to the text, centred on the pivot', () => {
    const { ctx, calls } = fakeCtx('top');
    drawRotatedCatLabel(ctx, 'FY2024\n1Q', 50, 60, -Math.PI / 4, 12);
    expect(calls.find(call => call.op === 'translate')?.args).toEqual([50, 60]);
    expect(texts(calls)).toEqual([['FY2024', 0, -6], ['1Q', 0, 6]]);

    const flat = fakeCtx('top');
    drawRotatedCatLabel(flat.ctx, 'FY2024\n1Q', 50, 60, 0, 12);
    expect(texts(flat.calls)).toEqual([['FY2024', 50, 60], ['1Q', 50, 72]]);

    const unbroken = fakeCtx('top');
    drawRotatedCatLabel(unbroken.ctx, 'ABCDEFG', 50, 60, -Math.PI / 4, 12, 45);
    expect(texts(unbroken.calls)).toEqual([['ABC…', 0, 0]]);
  });
});

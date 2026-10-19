import { describe, expect, it } from 'vitest';
import { wrapMeasuredText } from './axis.js';

// One unit per code point keeps widths deterministic.
const ctx = {
  measureText: (text: string) => ({ width: Array.from(text).length * 10 }),
} as unknown as CanvasRenderingContext2D;

describe('wrapMeasuredText', () => {
  it('starts a new line at each authored break even when the text would fit', () => {
    expect(wrapMeasuredText(ctx, 'FY2024\n1Q営業利益', 1000)).toEqual(['FY2024', '1Q営業利益']);
    expect(wrapMeasuredText(ctx, 'A\r\nB\rC', 1000)).toEqual(['A', 'B', 'C']);
  });

  it('keeps an interior blank line and trims outer breaks like outer spaces', () => {
    expect(wrapMeasuredText(ctx, '\nA\n\nB\n', 1000)).toEqual(['A', '', 'B']);
  });

  it('still wraps each authored line by width', () => {
    expect(wrapMeasuredText(ctx, 'aa bb\ncc', 30)).toEqual(['aa', 'bb', 'cc']);
  });

  it('keeps unbroken text unchanged', () => {
    expect(wrapMeasuredText(ctx, '  North  region ', 1000)).toEqual(['North region']);
    expect(wrapMeasuredText(ctx, '', 1000)).toEqual(['']);
  });
});

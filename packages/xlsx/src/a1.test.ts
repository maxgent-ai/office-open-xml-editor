import { describe, it, expect } from 'vitest';
import { parseA1, formatA1 } from './a1.js';

describe('parseA1', () => {
  it('rejects cells outside the worksheet grid', () => {
    expect(parseA1('A0')).toBeNull();
    expect(parseA1('XFE2')).toBeNull();
    expect(parseA1('A1048577')).toBeNull();
    expect(parseA1('$XFD$1048576')).toEqual({ row: 1_048_576, col: 16_384 });
  });
});

describe('formatA1', () => {
  it('formats single-letter columns', () => {
    expect(formatA1(1, 1)).toBe('A1');
    expect(formatA1(7, 2)).toBe('B7');
    expect(formatA1(10, 26)).toBe('Z10');
  });

  it('formats multi-letter columns (bijective base-26)', () => {
    expect(formatA1(1, 27)).toBe('AA1');
    expect(formatA1(1, 28)).toBe('AB1');
    expect(formatA1(1, 52)).toBe('AZ1');
    expect(formatA1(1, 53)).toBe('BA1');
    expect(formatA1(5, 702)).toBe('ZZ5');
    expect(formatA1(1, 703)).toBe('AAA1');
  });

  it('round-trips with parseA1', () => {
    for (const [row, col] of [
      [1, 1],
      [7, 2],
      [100, 26],
      [3, 27],
      [42, 703],
    ] as const) {
      const ref = formatA1(row, col);
      expect(parseA1(ref)).toEqual({ row, col });
    }
  });
});

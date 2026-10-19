import { describe, expect, it } from 'vitest';
import { NoteStack } from './note-stack.js';

describe('note stack arithmetic', () => {
  it.each([
    { areaTopPt: 84, advancesPt: Array<number>(50).fill(10.3) },
    { areaTopPt: 84, advancesPt: Array<number>(20).fill(25.056) },
    { areaTopPt: 588.72, advancesPt: [31.056, 25.056, 0, 25.056, 25.056] },
  ])('uses reduced prefixes for every boundary at $areaTopPt', ({ areaTopPt, advancesPt }) => {
    const stack = new NoteStack();
    expect(stack.advancePt).toBe(0);
    expect(stack.topAt(areaTopPt)).toBe(areaTopPt);
    advancesPt.forEach((advancePt, index) => {
      const prefixPt = advancesPt.slice(0, index).reduce((sum, advance) => sum + advance, 0);
      // Acquisition can ask for the next top before its advance is known.
      expect(stack.topAt(areaTopPt)).toBe(areaTopPt + prefixPt);
      stack.append(advancePt);
    });
    advancesPt.forEach((_, index) => {
      const prefixPt = advancesPt.slice(0, index).reduce((sum, advance) => sum + advance, 0);
      // Final placement can rebase those same prefixes to a page-final area.
      expect(stack.topAt(areaTopPt, index)).toBe(areaTopPt + prefixPt);
    });
    const totalPt = advancesPt.reduce((sum, advance) => sum + advance, 0);
    expect(stack.advancePt).toBe(totalPt);
    // The last bottom is the terminal prefix, bit-exactly the area bottom.
    expect(stack.topAt(areaTopPt, advancesPt.length)).toBe(areaTopPt + totalPt);
  });
});

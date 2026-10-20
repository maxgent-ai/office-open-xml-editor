import { describe, it, expect } from 'vitest';
import {
  docGridLineCells,
  eastAsianGridCountSinglePx,
  isGridLineRule,
  lineBoxHeight,
} from './line-layout.js';
import type { LineSpacing } from './types.js';

// Regression: some generators emit <w:spacing w:line="0" w:lineRule="exact"/> on
// table cells (e.g. private/sample-7.docx). A literal exact-0 line collapsed every
// line to 0px, so table rows fell back to the 10px minimum and cell content
// overlapped. ECMA-376 §17.3.1.33 leaves a zero line undefined; Word's native
// model ([MS-DOC] LSPD) cannot represent an exact 0 (exact = negative dyaLine)
// and resolves a non-negative dyaLine as max(dyaLine, single spacing) — so 0
// means exactly single spacing. lineBoxHeight must do the same.

const exact = (value: number): LineSpacing => ({ value, rule: 'exact', explicit: true });
const auto = (value: number): LineSpacing => ({ value, rule: 'auto', explicit: true });
const atLeast = (value: number): LineSpacing => ({
  value,
  rule: 'atLeast',
  explicit: true,
});

// Windows-adjudicated Yu Mincho: 1.3 × OS/2 usWin, with the added
// leading split equally above/below. The line-box contract receives these
// selected design sides and the intended single-line height independently.
const YU_ASC = (2038 + .15 * 2636) / 2048;
const YU_DESC = (598 + .15 * 2636) / 2048;
const YU = YU_ASC + YU_DESC;

describe('lineBoxHeight — degenerate zero line spacing', () => {
  // ascent 12 + descent 3 = 15px natural single-line height (no grid).
  it('treats exact line=0 as single spacing (natural), not 0', () => {
    expect(lineBoxHeight(exact(0), 12, 3, 1, undefined)).toBe(15);
  });
  it('treats auto value=0 as single spacing (natural), not 0', () => {
    expect(lineBoxHeight(auto(0), 12, 3, 1, undefined)).toBe(15);
  });
  it('still honors a positive exact line value', () => {
    // 12pt exact (w:line="240" twips / 20 = 12pt; LineSpacing.value is pt).
    expect(lineBoxHeight(exact(12), 12, 3, 1, undefined)).toBe(12);
  });
  it('treats negative exact/auto values as single spacing too', () => {
    expect(lineBoxHeight(exact(-5), 12, 3, 1, undefined)).toBe(15);
    expect(lineBoxHeight(auto(-1), 12, 3, 1, undefined)).toBe(15);
  });
  it('snaps a degenerate line to the grid pitch in docGrid sections', () => {
    // Same fallback as unspecified spacing: on-grid, the pitch governs.
    expect(
      lineBoxHeight(exact(0), 12, 3, 1, { type: 'lines', linePitchPt: 18 }),
    ).toBe(18);
  });
  it('still applies the auto multiplier for a positive value', () => {
    expect(lineBoxHeight(auto(2), 12, 3, 1, undefined)).toBe(30); // natural 15 × 2
  });
  it('unspecified spacing is single (natural)', () => {
    expect(lineBoxHeight(null, 12, 3, 1, undefined)).toBe(15);
  });
});

describe('lineBoxHeight — snapToChars line-grid participation', () => {
  const grid20 = { type: 'snapToChars', linePitchPt: 20 } as const;

  it('recognizes snapToChars as an active line-grid rule', () => {
    expect(isGridLineRule(grid20)).toBe(true);
  });

  it('applies the line pitch as well as the character-grid behavior', () => {
    expect(lineBoxHeight(null, 8, 2, 1, grid20, false, 0, true)).toBe(20);
  });
});

describe('lineBoxHeight — atLeast with an active line grid', () => {
  const grid18 = { type: 'lines', linePitchPt: 18 } as const;
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  it('retains the active grid minimum when the authored minimum is smaller', () => {
    expect(lineBoxHeight(atLeast(18), 10, 2, 1, grid20)).toBe(20);
  });

  it('retains the authored minimum when it is larger than the grid minimum', () => {
    expect(lineBoxHeight(atLeast(24), 10, 2, 1, grid20)).toBe(24);
  });

  // ECMA-376 defines the authored atLeast minimum and the active grid pitch,
  // but not this precise tall-line interaction. The unsnapped result below is
  // retained as observed Windows Word compatibility behavior.
  it('takes the maximum of content, authored minimum, and one grid pitch', () => {
    expect(lineBoxHeight(atLeast(18), 12 * YU_ASC, 12 * YU_DESC, 1, grid18, false, 12 * YU, true)).toBeCloseTo(12 * YU, 12);
  });
  it('does not round tall East Asian content up to an additional grid cell', () => {
    // The generic atLeast policy keeps the 33.46px design line. The
    // separately admitted specified-text path has its own grid contract.
    expect(lineBoxHeight(atLeast(18), 20 * YU_ASC, 20 * YU_DESC, 1, grid18, false, 20 * YU, true))
      .toBeCloseTo(20 * YU, 12);
  });
  it('preserves automatic whole-cell rounding for inherited-only spacing', () => {
    const inherited = { ...atLeast(18), explicit: false };
    expect(lineBoxHeight(inherited, 20 * YU_ASC, 20 * YU_DESC, 1, grid18, false, 20 * YU, true))
      .toBe(36);
  });
  it('keeps ruby atLeast lines rounded to cells that contain the annotation', () => {
    // glyph box 41px (base + furigana reserve) → ceil(41/18) = 3 cells = 54.
    expect(lineBoxHeight(atLeast(18), 33, 8, 1, grid18, true, 0, true)).toBe(54);
  });
});

// ECMA-376 §17.6.5 / §17.3.1.32 define the line pitch as the height of ONE
// single-spaced line; how many whole grid CELLS a single-spaced East Asian line
// occupies is Word runtime behaviour. That cell count is a function of the
// line's SINGLE-LINE HEIGHT (the document font's design line height): a line
// occupies ceil(singleLineHeight / pitch) cells — the smallest number of whole
// cells that CONTAINS it. It is NOT a function of the raw em, and NOT of the
// substituted-font Canvas glyph box (which can overstate the design height).
//
// Mac controls established whole-cell allocation from the normal line box.
// Windows grid controls adjudicate the metric source as usWin; the generic
// rounding contract also covers boundaries beyond the measured combinations.
describe('docGridLineCells — natural-height East Asian grid cell count (§17.6.5)', () => {
  it('a 12pt Windows FE line exceeds an 18pt pitch', () => {
    expect(docGridLineCells(12 * YU, 18)).toBe(2); // 20.08 > 18
  });
  it('a 10.5pt body line on an 18pt pitch is a single cell', () => {
    expect(docGridLineCells(10.5 * YU, 18)).toBe(1); // 17.57
  });
  it('a 14pt line on an 18pt pitch spills into TWO cells', () => {
    expect(docGridLineCells(14 * YU, 18)).toBe(2); // 23.43 > 18
  });
  it('a 16pt line on an 18pt pitch is two cells', () => {
    expect(docGridLineCells(16 * YU, 18)).toBe(2); // 26.77
  });
  it('a 16pt Windows FE line exceeds a 24pt pitch', () => {
    expect(docGridLineCells(16 * YU, 24)).toBe(2); // 26.77 > 24
  });
  it('a 20pt line on a 24pt pitch is two cells', () => {
    expect(docGridLineCells(20 * YU, 24)).toBe(2); // 33.46
  });
  it('a 20pt title on a 20pt pitch is two cells (sample-9)', () => {
    expect(docGridLineCells(20 * YU, 20)).toBe(2); // 33.46
  });
  it('a line exactly filling k pitches occupies k cells (ceil; boundary unmeasured)', () => {
    expect(docGridLineCells(18, 18)).toBe(1);
    expect(docGridLineCells(36, 18)).toBe(2);
  });
  it('a line taller than two pitches occupies three cells (extrapolated)', () => {
    expect(docGridLineCells(30 * YU, 20)).toBe(3); // 50.20 → ceil = 3
  });
  it('returns at least one cell for degenerate heights and pitches', () => {
    expect(docGridLineCells(0, 18)).toBe(1);
    expect(docGridLineCells(5, 0)).toBe(1);
  });
});

describe('lineBoxHeight — docGrid line-cell rounding (East Asian vs Latin)', () => {
  const grid18 = { type: 'lines', linePitchPt: 18 } as const;
  const grid20 = { type: 'lines', linePitchPt: 20 } as const;

  it('snaps a sub-pitch EA line to one cell', () => {
    expect(lineBoxHeight(null, 10.5 * YU_ASC, 10.5 * YU_DESC, 1, grid18, false, 10.5 * YU, true)).toBe(18);
  });
  it('uses the design grid-count height instead of a substituted Canvas box', () => {
    // The selected 10.5pt face fits one pitch; the measured substitute does not.
    const boxAsc = 18.02 * 0.8;
    const boxDesc = 18.02 * 0.2;
    expect(lineBoxHeight(null, boxAsc, boxDesc, 1, grid18, false, 10.5 * YU, true, 10.5 * YU)).toBe(18);
  });
  it('counts a tall UNTABLED run on a mixed line from its 1.3em FE height', () => {
    // A mixed docGrid line: a small 12pt tabled Yu Mincho run (design 20.08px)
    // plus a 24pt untabled CJK run. The latter contributes 31.2px regardless of
    // its substituted 30px Canvas box, so the line still claims two cells.
    const untabledDesign = eastAsianGridCountSinglePx(0, 24);
    expect(lineBoxHeight(null, 24, 6, 1, grid18, false, 12 * YU, true, untabledDesign)).toBe(36);
  });
  it('rounds a 20pt EA line on a 20pt pitch UP to two cells (sample-9)', () => {
    // design box 33.46px → ceil(33.46/20) = 2 cells = 40.
    expect(lineBoxHeight(null, 20 * YU_ASC, 20 * YU_DESC, 1, grid20, false, 20 * YU, true)).toBe(40);
  });
  it('uses a scale-linear 1.3em fallback for an untabled EA line', () => {
    // MS Mincho's usWin box is 1.0em and Word's FE single-line height is 1.3 ×
    // usWin. The substituted Canvas box is exactly 1.0em here, but a 20pt line
    // on a 20pt grid still occupies two cells in the sample-9 Word PDF.
    expect(lineBoxHeight(null, 18, 2, 1, grid20, false, 0, true, undefined, 20)).toBe(40);
    expect(lineBoxHeight(null, 43, 5, 2.3529, grid20, false, 0, true, undefined, 20 * 2.3529)).toBeCloseTo(40 * 2.3529, 8);
  });
  it('uses the tabled design floor when no per-line grid-count height is supplied', () => {
    // A direct caller without arg 9 still has a deterministic tabled-font input
    // in intendedSinglePx. The substituted 30px box is not consulted for the
    // cell count; Yu Mincho's 33.46px design height takes two 20px cells.
    expect(lineBoxHeight(null, 24, 6, 1, grid20, false, 20 * YU, true)).toBe(40);
  });
  it('allocates admitted Latin single-line design heights to whole grid cells while exact and atLeast retain their paths', () => {
    const single = { rule: 'auto' as const, value: 1, explicit: true };
    const height = (rule: typeof single | { rule: 'exact' | 'atLeast'; value: number; explicit: true }, design: number) =>
      lineBoxHeight(rule, 18, 5, 1, grid20, false, design, false, 23,
        undefined, undefined, 0, design);
    expect(height(single, 23)).toBe(40);
    expect(height(single, 34.5)).toBe(40);
    expect(height({ rule: 'exact', value: 30, explicit: true }, 23)).toBe(30);
    expect(height({ rule: 'atLeast', value: 24, explicit: true }, 23)).toBe(24);
  });
  it('keeps a Latin line at natural height (one-cell floor), NOT cell-rounded', () => {
    // 22px natural, Latin → max(22, 20) = 22px (Word does not cell-round it).
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20, false, 0, false)).toBe(22);
  });
  it('combines inherited spacing with useFELayout-classified Far East grid metrics', () => {
    const inherited = { value: 1.15, rule: 'auto', explicit: false } as const;
    // An 18pt hinted Latin title has a 1.3em FE design height and therefore
    // needs two 18pt cells. The allocation itself exceeds the 1.15-pitch floor.
    expect(lineBoxHeight(
      inherited, 17, 4, 1, grid18, false, 0, true,
      eastAsianGridCountSinglePx(0, 18),
    )).toBe(36);
    // A smaller hinted Latin label stays in one cell, but inherited 1.15 line
    // spacing raises the one-pitch minimum to 20.7pt.
    expect(lineBoxHeight(
      inherited, 7, 2, 1, grid18, false, 0, true,
      eastAsianGridCountSinglePx(0, 8),
    )).toBeCloseTo(20.7, 12);
    // Face hinting alone never changes an empty/Latin mark's grid class.
    expect(lineBoxHeight(inherited, 14, 2, 1, grid18, false, 0, false)).toBe(18);
  });
  it('a RUBY EA line can grow beyond its design cell count to contain furigana', () => {
    // At scale 1, a 13.5pt Yu Mincho base has a usWin single-line height of
    // 22.59px → 2 cells on an 18pt pitch. A measured base + annotation glyph
    // box of 41px (asc 33 + desc 8) needs 3 cells = 54px instead, so the
    // annotation can grow the allocation beyond that design minimum.
    expect(lineBoxHeight(null, 33, 8, 1, grid18, true, 13.5 * YU, true)).toBe(54);
  });
  it('does not cell-round when no docGrid is active (East Asian, off grid)', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, undefined, false, 0, true)).toBe(22);
  });
  it('defaults to Latin (no cell rounding) when the flag is omitted', () => {
    expect(lineBoxHeight(null, 17.6, 4.4, 1, grid20)).toBe(22);
  });
});

import { describe, it, expect } from 'vitest';
import {
  resolveLineFloatWindow,
  skipPastTopAndBottom,
  type FloatRect,
} from './float-layout.js';

// ─────────────────────────────────────────────────────────────────────────────
// A `wrapTopAndBottom` float (ECMA-376 §20.4.2.20, = §20.4.2.16 in the 1st
// edition) must affect text ONLY where the object is horizontally placed.
//
// GROUND TRUTH (spec):
//   §20.4.2.20 wrapTopAndBottom: "This element specifies that text shall wrap
//   around the top and bottom of THIS OBJECT, but not its left or right edges."
//   The exclusion is relative to the object's own extent; distT/distB are "the
//   minimum distance … between the [top/bottom] edge of THIS drawing object and
//   any subsequent text." Within a band the object overlaps, "text must wrap
//   around neither side of this object" — i.e. it blocks the FULL column width
//   (no side-by-side text).
//   §17.6.4 cols: a section's columns are distinct text regions of defined
//   width. Text laid out in a newspaper column the object does NOT horizontally
//   overlap is a separate flow region and must be unaffected.
//
// Floats are registered in ABSOLUTE page coordinates and `state.floats` is
// page-scoped (shared across columns — see measure-column-geometry.test.ts). The
// square-wrap path (§20.4.2.17) already gates on horizontal overlap so a float
// in one column does not reshape another column's line (float-line-start-one-
// inch.test.ts "ignores square wrap rectangles wholly outside … the column").
// topAndBottom must share that horizontal-overlap semantics: a topAndBottom
// float anchored in column 1 must NOT push down a line laid out in column 2.
//
// A 2-column page modeled at scale 1 (px == pt), mirroring
// measure-column-geometry.test.ts:
//   content band [60, 540]; col 0 = [60, 288] (w=228), col 1 = [312, 540] (w=228).
// ─────────────────────────────────────────────────────────────────────────────

const COL0 = { xLeft: 60, xRight: 288, width: 228 };
const COL1 = { xLeft: 312, xRight: 540, width: 228 };

/** A topAndBottom exclusion band (no dist padding) over an absolute rectangle. */
function topAndBottomBand(r: {
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
}): FloatRect {
  return {
    kind: 'shape',
    mode: 'topAndBottom',
    imageKey: 'x',
    imageX: r.xLeft,
    imageY: r.yTop,
    imageW: r.xRight - r.xLeft,
    imageH: r.yBottom - r.yTop,
    xLeft: r.xLeft,
    xRight: r.xRight,
    yTop: r.yTop,
    yBottom: r.yBottom,
    side: 'bothSides',
    distLeft: 0,
    distRight: 0,
    distTop: 0,
    distBottom: 0,
    paraId: 1,
  } as FloatRect;
}

// A line whose vertical range [120, 130] falls inside the float band [100, 200].
const LINE_TOP = 120;
const PROBE_H = 10;

/** resolveLineFloatWindow topY for a line laid out in the given column band. */
function windowTopY(col: { xLeft: number; width: number }, floats: FloatRect[]): number {
  return resolveLineFloatWindow(
    LINE_TOP,
    10,
    PROBE_H,
    col.xLeft,
    col.width,
    floats,
  ).topY;
}

describe('resolveLineFloatWindow — topAndBottom float column scope (§20.4.2.20 / §17.6.4)', () => {
  const floatInCol0 = topAndBottomBand({
    xLeft: COL0.xLeft,
    xRight: COL0.xRight,
    yTop: 100,
    yBottom: 200,
  });

  it('does NOT push a line in the non-overlapped column (column 2 unaffected)', () => {
    // The float is wholly inside column 1's horizontal band; a column-2 line at
    // the same vertical range does not intersect the object horizontally, so per
    // §20.4.2.20 it is not "this object's" top/bottom and must stay put.
    expect(windowTopY(COL1, [floatInCol0])).toBe(LINE_TOP);
  });

  it('DOES push a line in the overlapped column below the band (column 1 blocked)', () => {
    // Same object, a line in its own column: pushed to the band bottom.
    expect(windowTopY(COL0, [floatInCol0])).toBe(200);
  });

  it('blocks the FULL column width even when the float is narrower than the column', () => {
    // §20.4.2.20 "text must wrap around neither side of this object": a narrow
    // topAndBottom float that overlaps only part of column 1 still pushes the
    // whole column-1 line below it (no side-by-side text), and still leaves
    // column 2 untouched.
    const narrow = topAndBottomBand({ xLeft: COL0.xLeft, xRight: 150, yTop: 100, yBottom: 200 });
    expect(windowTopY(COL0, [narrow])).toBe(200);
    expect(windowTopY(COL1, [narrow])).toBe(LINE_TOP);
  });

  it('a float straddling BOTH columns pushes lines in both', () => {
    const straddle = topAndBottomBand({ xLeft: 200, xRight: 400, yTop: 100, yBottom: 200 });
    expect(windowTopY(COL0, [straddle])).toBe(200);
    expect(windowTopY(COL1, [straddle])).toBe(200);
  });

  it('single full-width column: a full-band topAndBottom float still pushes (unchanged)', () => {
    const single = { xLeft: 60, width: 480 };
    const fullBand = topAndBottomBand({ xLeft: 60, xRight: 540, yTop: 100, yBottom: 200 });
    expect(windowTopY(single, [fullBand])).toBe(200);
  });
});

describe('resolveLineFloatWindow — topAndBottom gate uses the COLUMN band, not the indented text band', () => {
  // An indented paragraph's text band starts INSIDE its column (indentLeft > 0).
  // Step 1 (topAndBottom, §20.4.2.20 "text must wrap around neither side of this
  // object" — a full-COLUMN block) must be gated by the raw COLUMN band, while
  // step 2 (square side-gap, §20.4.2.17) keeps gating by the narrower indented
  // text band. A topAndBottom float sitting in the column's LEFT indent margin
  // (inside the column, left of the indented text) therefore still pushes the
  // paragraph's lines below it even though it does not overlap the indented band.
  const COLUMN = { xLeft: 60, xRight: 288 };
  const INDENT_LEFT = 160; // indentLeft = 100 within the column band [60, 288]
  const INDENT_WIDTH = 128; // indented text band [160, 288]
  // [60, 140] ⊂ column band, ∩ indented text band = ∅ (140 < 160).
  const floatInIndentMargin = topAndBottomBand({
    xLeft: 60,
    xRight: 140,
    yTop: 100,
    yBottom: 200,
  });

  it('pushes an indented line below a topAndBottom float in the column indent margin', () => {
    // Reviewer repro: line at y=95 (band [100,200]); column band [60,288] passed
    // as columnXLeft/columnXRight, indented band [160,288] as paraX/maxWidth.
    const topY = resolveLineFloatWindow(
      95,
      0,
      PROBE_H,
      INDENT_LEFT,
      INDENT_WIDTH,
      [floatInIndentMargin],
      COLUMN.xLeft,
      COLUMN.xRight,
    ).topY;
    expect(topY).toBe(200);
  });

  it('leaves the line put when no column band is supplied (defaults to the indented band)', () => {
    // Safe default for direct unit callers: without a column band the float —
    // outside the indented band — is not seen by step 1, so nothing is pushed.
    const topY = resolveLineFloatWindow(
      95,
      0,
      PROBE_H,
      INDENT_LEFT,
      INDENT_WIDTH,
      [floatInIndentMargin],
    ).topY;
    expect(topY).toBe(95);
  });
});

describe('resolveLineFloatWindow — topAndBottom float in the inter-column gutter pushes neither column', () => {
  // Content band [60, 540]; col 0 = [60, 288], col 1 = [312, 540]; gutter
  // (288, 312). A float wholly inside the gutter overlaps neither column.
  const gutterFloat = topAndBottomBand({ xLeft: 292, xRight: 308, yTop: 100, yBottom: 200 });

  it('does not push a line in column 0', () => {
    const topY = resolveLineFloatWindow(
      LINE_TOP,
      10,
      PROBE_H,
      COL0.xLeft,
      COL0.width,
      [gutterFloat],
      COL0.xLeft,
      COL0.xRight,
    ).topY;
    expect(topY).toBe(LINE_TOP);
  });

  it('does not push a line in column 1', () => {
    const topY = resolveLineFloatWindow(
      LINE_TOP,
      10,
      PROBE_H,
      COL1.xLeft,
      COL1.width,
      [gutterFloat],
      COL1.xLeft,
      COL1.xRight,
    ).topY;
    expect(topY).toBe(LINE_TOP);
  });
});

describe('skipPastTopAndBottom — topAndBottom float column scope (§20.4.2.20 / §17.6.4)', () => {
  const floatInCol0 = topAndBottomBand({
    xLeft: COL0.xLeft,
    xRight: COL0.xRight,
    yTop: 100,
    yBottom: 200,
  });

  it('does NOT skip a cursor in the non-overlapped column (column 2 unaffected)', () => {
    expect(skipPastTopAndBottom(LINE_TOP, [floatInCol0], COL1.xLeft, COL1.xRight)).toBe(LINE_TOP);
  });

  it('DOES skip a cursor in the overlapped column past the band', () => {
    expect(skipPastTopAndBottom(LINE_TOP, [floatInCol0], COL0.xLeft, COL0.xRight)).toBe(200);
  });

  it('a float straddling both columns skips a cursor in either column', () => {
    const straddle = topAndBottomBand({ xLeft: 200, xRight: 400, yTop: 100, yBottom: 200 });
    expect(skipPastTopAndBottom(LINE_TOP, [straddle], COL0.xLeft, COL0.xRight)).toBe(200);
    expect(skipPastTopAndBottom(LINE_TOP, [straddle], COL1.xLeft, COL1.xRight)).toBe(200);
  });

  it('single full-width column: a full-band float still skips (unchanged)', () => {
    const fullBand = topAndBottomBand({ xLeft: 60, xRight: 540, yTop: 100, yBottom: 200 });
    expect(skipPastTopAndBottom(LINE_TOP, [fullBand], 60, 540)).toBe(200);
  });
});

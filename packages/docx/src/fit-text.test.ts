import { describe, it, expect } from 'vitest';
import { groupFitTextRegions } from './fit-text.js';
import type { FitTextRun } from './fit-text.js';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.2.14 `<w:fitText>` — "Manual Run Width".
//
// A run's contents are displayed within the width given by `w:val` (twips). When
// several CONSECUTIVE runs carry the SAME `w:id`, they are treated as ONE fit
// region whose combined contents fill `w:val` together (the spec's example: a
// 1-inch region split across three runs occupies one inch in total, not one inch
// each). An id-less fitText run has id 0 (WORD_FIT_TEXT_OMITTED_ID_LINKING). A non-fitText run
// breaks adjacency, so two same-id runs separated by a plain run do NOT merge.
//
// Word's observed geometry (private fixture, Word-generated PDF): the glyphs keep
// their natural advance (incl. §17.3.2.43 `w:w` glyph scaling) and the extra /
// negative slack `(w:val − Σnatural)` is distributed EVENLY as inter-character
// gaps — one gap between every adjacent glyph pair across run boundaries, and NO
// gap after the region's last glyph. Any cached §17.3.2.35 `w:spacing` on the runs
// is IGNORED inside a fit region (Word recomputes the spacing from `w:val`).
//
// This kernel is the pure numeric heart. Each run reports its natural glyph-advance
// SUM in px (pre-`w:w`, without any `w:spacing`), its code-point count, its `w:w`
// fraction, and the target `w:val` in twips. It groups consecutive linked runs and
// returns the resolved regions + the per-gap the renderer feeds into the existing
// justify (`justifiedPiecePositions` / `ctx.letterSpacing`) draw path.
// ─────────────────────────────────────────────────────────────────────────────

/** A CJK glyph's natural advance in px at scale 1 for `sz=24` (12 pt). */
const EM = 12;

/** Build a fitText run of `n` glyphs each `EM` px wide (scale-1 default). */
function fitRun(
  n: number,
  fitTextValTwips: number,
  fitTextId: number | string | undefined,
  extra: Partial<FitTextRun> = {},
): FitTextRun {
  return {
    fitTextValTwips,
    fitTextId,
    charCount: n,
    naturalWidthPx: n * EM,
    ...extra,
  };
}

/** A plain (non-fitText) run of `n` glyphs — breaks fit-region adjacency. */
function plainRun(n: number): FitTextRun {
  return { charCount: n, naturalWidthPx: n * EM };
}

describe('groupFitTextRegions — ECMA-376 §17.3.2.14 fit-region math', () => {
  it('distributes (val − Σnatural)/(n−1) as the inter-character gap, no trailing gap', () => {
    // One run, 6 glyphs (12 px each ⇒ Σ = 72), val = 2400 twips = 120 pt.
    // per-gap = (120 − 72) / (6 − 1) = 9.6 px.
    const regions = groupFitTextRegions([fitRun(6, 2400, undefined)], 1);
    expect(regions).toHaveLength(1);
    const r = regions[0];
    expect(r.start).toBe(0);
    expect(r.end).toBe(1);
    expect(r.targetPx).toBeCloseTo(120, 9);
    expect(r.naturalPx).toBeCloseTo(72, 9);
    expect(r.charCount).toBe(6);
    expect(r.perGapPx).toBeCloseTo(9.6, 9);
    // No trailing gap: natural + (n−1)·perGap === target exactly.
    expect(r.naturalPx + (r.charCount - 1) * r.perGapPx).toBeCloseTo(r.targetPx, 9);
  });

  it('merges CONSECUTIVE same-id runs into one region (氏名/名/称 → 6 glyphs, per-gap 9.6)', () => {
    // Line 1 shape: three runs with the same id, 4 + 1 + 1 glyphs = 6, each val=2400.
    const id = -1431456512;
    const regions = groupFitTextRegions(
      [fitRun(4, 2400, id), fitRun(1, 2400, id), fitRun(1, 2400, id)],
      1,
    );
    expect(regions).toHaveLength(1);
    const r = regions[0];
    expect(r.start).toBe(0);
    expect(r.end).toBe(3);
    expect(r.charCount).toBe(6);
    expect(r.naturalPx).toBeCloseTo(72, 9);
    // Word's recomputed per-gap is 9.6 px — NOT the cached w:spacing=96 twips (4.8 pt)
    // on the runs, which a fit region ignores (§17.3.2.14 recomputes from w:val).
    expect(r.perGapPx).toBeCloseTo(9.6, 9);
  });

  it('applies the general formula to a 4-glyph region (及び/住/所 → per-gap 24)', () => {
    // Line 2 shape: 2 + 1 + 1 glyphs = 4, val=2400. per-gap = (120 − 48)/3 = 24 px.
    const id = -1431456511;
    const regions = groupFitTextRegions(
      [fitRun(2, 2400, id), fitRun(1, 2400, id), fitRun(1, 2400, id)],
      1,
    );
    expect(regions).toHaveLength(1);
    const r = regions[0];
    expect(r.charCount).toBe(4);
    expect(r.naturalPx).toBeCloseTo(48, 9);
    expect(r.perGapPx).toBeCloseTo(24, 9);
  });

  it('composes §17.3.2.43 w:w glyph scaling into the natural width (並びに… line, charScale 0.66)', () => {
    // Line 3 shape: 15 glyphs, all runs val=2400 AND w:w=66% (0.66). Natural width =
    // 15 × 12 × 0.66 = 118.8 px, so per-gap = (120 − 118.8)/14 ≈ 0.0857 px — the same
    // general formula, no special branch. Cached w:spacing (0.9 pt on some runs) is
    // ignored. Model as two same-id runs (10 + 5) to also exercise the merge.
    const id = -1431456510;
    const regions = groupFitTextRegions(
      [
        fitRun(10, 2400, id, { charScale: 0.66 }),
        fitRun(5, 2400, id, { charScale: 0.66 }),
      ],
      1,
    );
    expect(regions).toHaveLength(1);
    const r = regions[0];
    expect(r.charCount).toBe(15);
    expect(r.naturalPx).toBeCloseTo(118.8, 6);
    expect(r.perGapPx).toBeCloseTo(1.2 / 14, 9);
    // Region advance stays pinned to the target (≈ 120 pt) regardless of w:w.
    expect(r.naturalPx + (r.charCount - 1) * r.perGapPx).toBeCloseTo(120, 6);
  });

  it('does NOT merge two same-id runs separated by a plain run (adjacency required)', () => {
    // Same id on runs 0 and 2, but run 1 is plain ⇒ two independent regions.
    const id = 42;
    const regions = groupFitTextRegions(
      [fitRun(3, 2400, id), plainRun(2), fitRun(3, 2400, id)],
      1,
    );
    expect(regions).toHaveLength(2);
    expect(regions[0].start).toBe(0);
    expect(regions[0].end).toBe(1);
    expect(regions[1].start).toBe(2);
    expect(regions[1].end).toBe(3);
  });

  it('does NOT merge adjacent runs with DIFFERENT ids', () => {
    const regions = groupFitTextRegions(
      [fitRun(3, 2400, 7), fitRun(3, 2400, 8)],
      1,
    );
    expect(regions).toHaveLength(2);
    expect(regions[0].end).toBe(1);
    expect(regions[1].start).toBe(1);
  });

  it('links contiguous id-less and id-0 fitText runs as id 0 (WORD_FIT_TEXT_OMITTED_ID_LINKING)', () => {
    const linked = groupFitTextRegions(
      [fitRun(3, 2400, undefined), fitRun(3, 2400, undefined), fitRun(3, 2400, '0')],
      1,
    );
    expect(linked.map((r) => [r.start, r.end])).toEqual([[0, 3]]);
    // Another id, or a plain run between them, still separates regions.
    const separate = groupFitTextRegions(
      [fitRun(3, 2400, undefined), fitRun(3, 2400, 7), fitRun(3, 2400, undefined)],
      1,
    );
    expect(separate.map((r) => [r.start, r.end])).toEqual([[0, 1], [1, 2], [2, 3]]);
  });

  it('compares integer ids by value at arbitrary precision', () => {
    const regions = groupFitTextRegions(
      [fitRun(3, 2400, '0012345678901234567890'), fitRun(3, 2400, '12345678901234567890'),
        fitRun(3, 2400, '12345678901234567891')],
      1,
    );
    expect(regions.map((r) => [r.start, r.end])).toEqual([[0, 2], [2, 3]]);
  });

  it('converts w:val twips → px through the layout scale', () => {
    // scale = 2 px/pt: 6 glyphs at 24 px natural (Σ = 144), val=2400 (120 pt) ⇒
    // target = 240 px, per-gap = (240 − 144)/5 = 19.2 px.
    const runs: FitTextRun[] = [
      { fitTextValTwips: 2400, charCount: 6, naturalWidthPx: 144 },
    ];
    const regions = groupFitTextRegions(runs, 2);
    expect(regions).toHaveLength(1);
    expect(regions[0].targetPx).toBeCloseTo(240, 9);
    expect(regions[0].naturalPx).toBeCloseTo(144, 9);
    expect(regions[0].perGapPx).toBeCloseTo(19.2, 9);
  });

  it('yields a NEGATIVE per-gap when the natural width exceeds the target (compression)', () => {
    // 6 glyphs (Σ = 72 px) into a 1000-twip (50 pt) region ⇒ per-gap = (50 − 72)/5 = −4.4 px.
    const regions = groupFitTextRegions([fitRun(6, 1000, undefined)], 1);
    expect(regions).toHaveLength(1);
    expect(regions[0].perGapPx).toBeCloseTo(-4.4, 9);
    expect(regions[0].perGapPx).toBeLessThan(0);
  });

  it('uses per-gap 0 for a single-glyph region (no division by zero)', () => {
    const regions = groupFitTextRegions([fitRun(1, 2400, undefined)], 1);
    expect(regions).toHaveLength(1);
    expect(regions[0].charCount).toBe(1);
    expect(regions[0].perGapPx).toBe(0);
  });

  it('returns no regions when no run carries fitText', () => {
    const regions = groupFitTextRegions([plainRun(3), plainRun(2)], 1);
    expect(regions).toHaveLength(0);
  });
});

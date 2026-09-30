import { describe, expect, it } from 'vitest';
import {
  BOX_WHISKER_SLOT_GUTTER_FRACTION,
  boxWhiskerGeometry,
  boxWhiskerPointCount,
  computeBoxWhiskerStats,
} from './box-whisker.js';

describe('computeBoxWhiskerStats', () => {
  const quartiles = (values: number[], method: string) => {
    const stats = computeBoxWhiskerStats(values, method);
    return [stats?.q1, stats?.median, stats?.q3];
  };

  it('interpolates exclusive quartiles at (n + 1)p, clamped to the sample', () => {
    expect(quartiles([0, 100, 104, 200], 'exclusive')).toEqual([25, 102, 176]);
    expect(quartiles([0, 20, 60, 100, 120, 200], 'exclusive')).toEqual([15, 80, 140]);
    // Positions outside 1..n clamp to the smallest/largest observation.
    expect(quartiles([0, 200], 'exclusive')).toEqual([0, 100, 200]);
    expect(quartiles([0, 60, 200], 'exclusive')).toEqual([0, 60, 200]);
  });

  it('interpolates inclusive quartiles at (n - 1)p + 1', () => {
    expect(quartiles([0, 100, 104, 200], 'inclusive')).toEqual([75, 102, 128]);
    expect(quartiles([0, 20, 60, 100, 120, 200], 'inclusive')).toEqual([30, 80, 115]);
    expect(quartiles([0, 200], 'inclusive')).toEqual([50, 100, 150]);
    expect(quartiles([0, 60, 200], 'inclusive')).toEqual([30, 60, 130]);
  });

  it('drops missing/non-finite observations and preserves finite repeats', () => {
    const stats = computeBoxWhiskerStats(
      [null, 5, Number.NaN, 5, Infinity, 5, -Infinity, undefined, 5],
      'exclusive',
    );
    expect(stats).toMatchObject({
      q1: 5, median: 5, q3: 5,
      lowerFence: 5, upperFence: 5,
      whiskerLo: 5, whiskerHi: 5, mean: 5,
      outliers: [], inner: [5, 5, 5, 5],
    });
  });

  it('keeps equality inside strict fences and rejects a value just above', () => {
    const atFence = computeBoxWhiskerStats([0, 1, 2, 3, 4, 5, 6, 7, 12], 'inclusive');
    const aboveFence = computeBoxWhiskerStats([0, 1, 2, 3, 4, 5, 6, 7, 12.0001], 'inclusive');
    expect(atFence).toMatchObject({ upperFence: 12, whiskerHi: 12, outliers: [] });
    expect(aboveFence).toMatchObject({ upperFence: 12, whiskerHi: 7, outliers: [12.0001] });
  });

  it('returns no statistics for an empty retained sample', () => {
    expect(computeBoxWhiskerStats([null, NaN, Infinity], 'inclusive')).toBeNull();
  });

  it.each([1e-310, Number.MIN_VALUE])('preserves tied subnormal quartiles at %s', value => {
    const edges = quartiles([value, value, value], 'inclusive');
    expect(edges[0]).toBe(value);
    expect(edges[2]).toBe(value);
  });

  it('interpolates inclusive quartiles between opposite-signed finite extremes', () => {
    const edges = quartiles([-Number.MAX_VALUE, Number.MAX_VALUE], 'inclusive');
    expect((edges[0] as number) / Number.MAX_VALUE).toBeCloseTo(-0.5, 12);
    expect((edges[2] as number) / Number.MAX_VALUE).toBeCloseTo(0.5, 12);
  });

  it('keeps a fence representable when the IQR spread overflows', () => {
    // 1.5 × IQR overflows, but the lower fence 2.5 × Q1 − 1.5 × Q3 does not:
    // exactly, Q1 = 0.25S, Q3 = 1.65625S and the fence is −1.859375S, so the
    // lowest observation (−1.875S) is an outlier. Mirrored for the upper side.
    const S = 2 ** 1023;
    const low = [-1.875, -0.5, 0.5, 1.5, 1.5, 1.625, 1.75, 1.875].map(v => v * S);
    const lower = computeBoxWhiskerStats(low, 'inclusive');
    expect((lower?.lowerFence as number) / S).toBe(-1.859375);
    expect(lower?.outliers).toEqual([-1.875 * S]);
    const upper = computeBoxWhiskerStats(low.map(v => -v), 'inclusive');
    expect((upper?.upperFence as number) / S).toBe(1.859375);
    expect(upper?.outliers).toEqual([1.875 * S]);
  });

  it('keeps exact and maximum-magnitude fences when the spread overflows', () => {
    // Opposite-signed maximum quartiles: both fences saturate outward, so no
    // observation becomes an outlier.
    expect(computeBoxWhiskerStats([-Number.MAX_VALUE, Number.MAX_VALUE], 'exclusive'))
      .toMatchObject({ lowerFence: -Number.MAX_VALUE, upperFence: Number.MAX_VALUE, outliers: [] });
    // On the overflow path the fence keeps the direct expression's rounding
    // steps: these quartiles give −27/16·S, whereas the reassociated
    // 2.5·Q1 − 1.5·Q3 rounds to one ULP above it. An observation exactly on
    // the fence stays inside (strict fences); the next double beyond does not.
    const S = 2 ** 1023;
    const U = 2 ** 971;
    const q1 = (3 / 8) * S + 3 * (U / 8);
    const q3 = (7 / 4) * S;
    const fence = (-27 / 16) * S;
    const middle = [q1, q1, q1, q1, q3, q3, q3, q3];
    const onFence = computeBoxWhiskerStats([fence, ...middle], 'inclusive');
    expect(onFence?.lowerFence).toBe(fence);
    expect(onFence?.outliers).toEqual([]);
    const beyond = computeBoxWhiskerStats([fence - U, ...middle], 'inclusive');
    expect(beyond?.outliers).toEqual([fence - U]);
    const mirrored = computeBoxWhiskerStats([-fence, ...middle.map(v => -v)], 'inclusive');
    expect(mirrored?.upperFence).toBe(-fence);
    expect(mirrored?.outliers).toEqual([]);
  });

  it('keeps derived statistics finite for extreme finite observations', () => {
    const stats = computeBoxWhiskerStats(
      [-Number.MAX_VALUE, -Number.MAX_VALUE, Number.MAX_VALUE, Number.MAX_VALUE],
      'exclusive',
    );
    expect(stats).not.toBeNull();
    expect(Object.values(stats as object).flat().filter(value => typeof value === 'number')
      .every(Number.isFinite)).toBe(true);
  });
});

describe('boxWhiskerGeometry', () => {
  it('uses stable equal series slots and a fixed 6% local gutter', () => {
    const first = boxWhiskerGeometry(20, 400, 1, 2, 0, 0, 33);
    const second = boxWhiskerGeometry(20, 400, 1, 2, 0, 1, 33);
    expect(first).not.toBeNull();
    expect(second).not.toBeNull();
    expect(first?.boxWidth).toBeCloseTo(second?.boxWidth ?? 0, 10);
    // gapWidth is expressed as a percentage of one box/series slot, not as a
    // percentage of the complete multi-series group.  One category interval
    // therefore contains `seriesCount + gapWidth / 100` slot units.
    const seriesSlotWidth = 400 / (2 + 0.33);
    expect(first?.boxWidth).toBeCloseTo(seriesSlotWidth * (1 - BOX_WHISKER_SLOT_GUTTER_FRACTION), 10);
    expect((second?.centerX ?? 0) - (first?.centerX ?? 0)).toBeCloseTo(seriesSlotWidth, 10);
    expect((second?.boxX ?? 0) - ((first?.boxX ?? 0) + (first?.boxWidth ?? 0)))
      .toBeCloseTo(seriesSlotWidth * BOX_WHISKER_SLOT_GUTTER_FRACTION, 10);
  });

  it('does not multiply the authored gap around a many-series category group', () => {
    const first = boxWhiskerGeometry(0, 600, 1, 6, 0, 0, 33);
    expect(first?.boxWidth).toBeCloseTo(
      (600 / (6 + 0.33)) * (1 - BOX_WHISKER_SLOT_GUTTER_FRACTION),
      10,
    );
  });

  it('keeps a series position independent of whether peer series are populated', () => {
    const firstCategory = boxWhiskerGeometry(0, 600, 2, 3, 0, 2, 33);
    const secondCategory = boxWhiskerGeometry(0, 600, 2, 3, 1, 2, 33);
    expect((secondCategory?.centerX ?? 0) - (firstCategory?.centerX ?? 0)).toBeCloseTo(300, 10);
  });

  it('rejects invalid geometry inputs and bounds aggregate point counting', () => {
    expect(boxWhiskerGeometry(0, 100, 0, 1, 0, 0, 33)).toBeNull();
    expect(boxWhiskerPointCount([[[1, 2]], [[3]]], 10)).toBe(3);
    expect(boxWhiskerPointCount([[[...new Array(6)]], [[...new Array(5)]]], 10)).toBe(11);
  });
});

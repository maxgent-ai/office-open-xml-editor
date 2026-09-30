/** Fixed empty space inside each equal per-series slot. */
export const BOX_WHISKER_SLOT_GUTTER_FRACTION = 0.06;

export interface BoxWhiskerStats {
  q1: number;
  median: number;
  q3: number;
  lowerFence: number;
  upperFence: number;
  whiskerLo: number;
  whiskerHi: number;
  mean: number;
  outliers: number[];
  /** Sorted non-outlier observations, including repeated values. */
  inner: number[];
}

export interface BoxWhiskerGeometry {
  boxX: number;
  boxWidth: number;
  centerX: number;
}

function median(sorted: readonly number[]): number {
  const middle = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 1) return sorted[middle];
  // Halving before addition avoids overflow for two large finite values.
  return sorted[middle - 1] / 2 + sorted[middle] / 2;
}

function finiteMean(values: readonly number[]): number {
  let scale = 0;
  for (const value of values) scale = Math.max(scale, Math.abs(value));
  if (scale === 0) return 0;
  let normalized = 0;
  for (const value of values) normalized += value / scale;
  return (normalized / values.length) * scale;
}

/**
 * Tukey fence `edge ± 1.5 × IQR`, where `edge` is the quartile the fence
 * extends from and `other` the opposite quartile. When an intermediate step
 * overflows, evaluate the same expression, with the same rounding steps, on
 * quarter-scaled quartiles and multiply by four; saturate only when that
 * product overflows. Office's own arithmetic at this range is not measured.
 */
function tukeyFence(edge: number, other: number): number {
  const direct = edge + (edge - other) * 1.5;
  if (Number.isFinite(direct)) return direct;
  const quarterEdge = edge / 4;
  const scaled = quarterEdge + (quarterEdge - other / 4) * 1.5;
  const result = scaled * 4;
  if (Number.isFinite(result)) return result;
  return scaled < 0 ? -Number.MAX_VALUE : Number.MAX_VALUE;
}

/**
 * Linear interpolation at a 1-based order-statistic position, clamped to the
 * sample. Weighting both neighbours (rather than adding a scaled difference)
 * keeps the result finite for opposite-signed extreme values; the final clamp
 * keeps rounding (e.g. of subnormal ties) from leaving the neighbour range.
 */
function quantileAt(sorted: readonly number[], position: number): number {
  if (position <= 1) return sorted[0];
  if (position >= sorted.length) return sorted[sorted.length - 1];
  const index = Math.floor(position);
  const fraction = position - index;
  const below = sorted[index - 1];
  const above = sorted[index];
  if (fraction === 0 || below === above) return below;
  const value = below * (1 - fraction) + above * fraction;
  return Math.min(above, Math.max(below, value));
}

/**
 * Compute box statistics from finite observations.
 * Missing/non-finite observations are discarded and repeats are retained.
 *
 * [MS-ODRAWXML] ST_QuartileMethod (2.24.4.17) describes the two methods only
 * as including or excluding the median; the interpolation is observed
 * PowerPoint 16.113 behavior (vector PDF, electronic-distribution engine) on
 * synthetic controls with the same irregular data under both methods,
 * n = 1–25, 30 and 31 (every n mod 4): every box edge within 0.41pt
 * (0.23 axis units) of these values, while each other Hyndman–Fan
 * definition misses at least one measured group by more than 30pt.
 * - `exclusive`: position (n + 1)p, the QUARTILE.EXC convention, clamped to
 *   the smallest/largest observation (for n = 2 the box spans the sample).
 * - `inclusive`: position (n − 1)p + 1, the QUARTILE.INC convention.
 * The previous split-halves medians matched only 14/26 (exclusive) and
 * 13/26 (inclusive) of the measurable groups. Excel's own export (print
 * engine) of a 40-point inclusive box chart is consistent as well.
 */
export function computeBoxWhiskerStats(
  values: readonly (number | null | undefined)[],
  method: string,
): BoxWhiskerStats | null {
  const sorted = values
    .filter((value): value is number => typeof value === 'number' && Number.isFinite(value))
    .sort((a, b) => a - b);
  if (sorted.length === 0) return null;

  const n = sorted.length;
  const center = median(sorted);
  const position = method === 'inclusive'
    ? (p: number) => (n - 1) * p + 1
    : (p: number) => (n + 1) * p;
  const q1 = quantileAt(sorted, position(0.25));
  const q3 = quantileAt(sorted, position(0.75));
  const lowerFence = tukeyFence(q1, q3);
  const upperFence = tukeyFence(q3, q1);
  const inner: number[] = [];
  const outliers: number[] = [];
  for (const value of sorted) {
    // The fences are strict: equality remains a whisker candidate.
    if (value < lowerFence || value > upperFence) outliers.push(value);
    else inner.push(value);
  }

  return {
    q1,
    median: center,
    q3,
    lowerFence,
    upperFence,
    whiskerLo: inner[0] ?? sorted[0],
    whiskerHi: inner[inner.length - 1] ?? sorted[sorted.length - 1],
    mean: finiteMean(sorted),
    outliers,
    inner,
  };
}

/**
 * Minimum value-axis distance, in points, between two painted observation
 * dots of one box-and-whisker series. It equals the fixed 3pt dot diameter.
 */
export const BOX_WHISKER_DOT_SPACING_PT = 3;

/**
 * Select the observation dots PowerPoint paints for one series, category by
 * category in axis order. `positionOf` maps a value to its device coordinate
 * on the value axis; `minSpacing` is BOX_WHISKER_DOT_SPACING_PT in the same
 * device units. Each result lists its category's dot values in ascending order.
 *
 * Neither ECMA-376 nor [MS-ODRAWXML] defines which observations
 * CT_SeriesElementVisibilities@nonoutliers/@outliers paint. This is observed
 * PowerPoint 16.113 behavior (vector PDF, electronic-distribution engine) on
 * synthetic three-series controls: 225 series/category groups and 995 dots
 * under both quartile methods, every dot count and position (within 1pt)
 * reproduced with the quartiles of computeBoxWhiskerStats.
 *
 * - One instance of the lowest and of the highest non-outlier value is the
 *   whisker end and gets no dot, even when it lies inside the box and so has
 *   no whisker cap either. Further copies of an end value, and values near
 *   it, remain candidates; the omitted end blocks nothing.
 * - Shown non-outliers and outliers form one ascending pass. A candidate is
 *   painted only when it lies at least 3pt from the last dot painted, so
 *   duplicates and near-coincident values collapse onto the lower dot. The
 *   boundary lies between 2.983pt (collapses) and 3.027pt (both paint) over
 *   80 value pairs, each drawn at two value scales fourfold apart with the
 *   same point gaps (identical results, so not axis values), and again
 *   between 2.994pt and 3.033pt in a plot shrunk to 0.71 of the height (so
 *   not a fraction of the plot). Which side 3pt itself falls on is not
 *   measured. Chains compare with the last painted dot: dots spaced 2pt
 *   apart paint every second one.
 * - The last painted dot carries across categories and restarts for each
 *   series. In an outliers-only series, the same outlier in each category
 *   paints only in the first category; the next series paints it again.
 * - Hidden non-outliers up to 2.65pt below an outlier do not suppress it.
 *   That hidden outliers likewise take no part, and that a category without
 *   observations keeps the carried dot, are inferred, not measured.
 */
export function boxWhiskerObservationDots(
  statsByCategory: readonly (BoxWhiskerStats | null)[],
  showNonoutliers: boolean,
  showOutliers: boolean,
  positionOf: (value: number) => number,
  minSpacing: number,
): number[][] {
  let lastPosition: number | null = null;
  return statsByCategory.map(stats => {
    if (!stats || (!showNonoutliers && !showOutliers)) return [];
    // `inner` and `outliers` are each sorted, and every outlier lies outside
    // the inner range, so the pass is low outliers, inner values, high ones.
    const inner = showNonoutliers ? stats.inner.slice(1, -1) : [];
    const low = showOutliers ? stats.outliers.filter(value => value < stats.lowerFence) : [];
    const high = showOutliers ? stats.outliers.filter(value => value > stats.upperFence) : [];
    const dots: number[] = [];
    for (const value of [...low, ...inner, ...high]) {
      const position = positionOf(value);
      if (lastPosition !== null && Math.abs(position - lastPosition) < minSpacing) continue;
      dots.push(value);
      lastPosition = position;
    }
    return dots;
  });
}

/**
 * Upper bound on the dots boxWhiskerObservationDots can select for one
 * category, without layout: every shown observation except the two whisker
 * ends, which never get a dot. The spacing rule can only lower the count.
 */
export function boxWhiskerDotCandidateCount(
  stats: BoxWhiskerStats,
  showNonoutliers: boolean,
  showOutliers: boolean,
): number {
  return (showNonoutliers ? Math.max(0, stats.inner.length - 2) : 0)
    + (showOutliers ? stats.outliers.length : 0);
}

/** Count raw observations with overflow-safe early termination. */
export function boxWhiskerPointCount(
  groups: readonly (readonly (readonly unknown[])[])[],
  limit: number,
): number {
  let count = 0;
  for (const series of groups) {
    for (const observations of series) {
      count += observations.length;
      if (!Number.isSafeInteger(count) || count > limit) return limit + 1;
    }
  }
  return count;
}

/**
 * Place a series in a stable equal slot inside a ChartEx category group.
 * Empty peers keep their slots, so the same series never shifts horizontally
 * between categories. The 6% gutter is local to each series slot.
 */
export function boxWhiskerGeometry(
  plotX: number,
  plotWidth: number,
  categoryCount: number,
  seriesCount: number,
  categoryIndex: number,
  seriesIndex: number,
  gapWidthPercent: number,
): BoxWhiskerGeometry | null {
  if (
    !Number.isFinite(plotX)
    || !Number.isFinite(plotWidth)
    || plotWidth <= 0
    || !Number.isInteger(categoryCount)
    || categoryCount <= 0
    || !Number.isInteger(seriesCount)
    || seriesCount <= 0
    || !Number.isInteger(categoryIndex)
    || categoryIndex < 0
    || categoryIndex >= categoryCount
    || !Number.isInteger(seriesIndex)
    || seriesIndex < 0
    || seriesIndex >= seriesCount
    || !Number.isFinite(gapWidthPercent)
    || gapWidthPercent < 0
  ) return null;

  // Excel divides the plot into `categoryCount` full category intervals.  The
  // first and last category centres therefore sit half an interval from the
  // plot edges.  Using `categoryCount + 1` incorrectly compresses a
  // formula-only (one category, many series) box chart into the middle half of
  // the plot.
  const categoryInterval = plotWidth / categoryCount;
  // `gapWidth` is a percentage of one data-point slot.  A category interval
  // contains all series slots plus one gap slot, so applying it to the whole
  // group (`interval / (1 + gap)`) makes a six-series formula chart about 20%
  // too narrow.  This is the same unit relation used by clustered columns.
  const seriesSlotWidth = categoryInterval / (seriesCount + gapWidthPercent / 100);
  const groupWidth = seriesSlotWidth * seriesCount;
  const gutter = seriesSlotWidth * BOX_WHISKER_SLOT_GUTTER_FRACTION;
  const boxWidth = seriesSlotWidth - gutter;
  const groupLeft = plotX + categoryInterval * (categoryIndex + 0.5) - groupWidth / 2;
  const boxX = groupLeft + seriesIndex * seriesSlotWidth + gutter / 2;
  return { boxX, boxWidth, centerX: boxX + boxWidth / 2 };
}

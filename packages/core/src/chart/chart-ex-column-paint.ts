import type { ChartModel, ChartSeries } from '../types/chart.js';
import type { DateCategoryAxisPlan } from './date-axis.js';
import { chartSeriesVariesByPoint } from './effective-style.js';
import { resolveCategoryGapWidthPercent } from './category-spacing.js';
import { catAxisReversed, planValueAxis, type ValueAxisPlan } from './shared/axis.js';
import { chartExSeriesFormatIndex } from './shared/palette.js';
import { computeSecondaryAxis, forEachErrorBarEndpoint } from './shared/secondary-axis.js';

/** Whether the delegated bar painter uses the simple ChartEx column geometry
 * shared with preflight. Other bar layouts retain the general painter path,
 * whose grouping, orientation and axis ownership rules are broader. */
export function chartExColumnSharedPlanApplies(
  chart: ChartModel,
  series: readonly ChartSeries[],
): boolean {
  const barSeries = new Set(series);
  if (chart.series.some(owner => !barSeries.has(owner))) return false;
  const defaultHorizontal = chart.chartType === 'clusteredBarH'
    || chart.chartType === 'stackedBarH'
    || chart.chartType === 'stackedBarHPct';
  const defaultGrouping = chart.chartType.startsWith('stacked')
    ? chart.chartType.endsWith('Pct') ? 'percentStacked' : 'stacked'
    : 'clustered';
  const groupKeys = new Set<string>();
  for (const owner of series) {
    if (owner.useSecondaryAxis === true) return false;
    const horizontal = owner.barGroupDirection != null
      ? owner.barGroupDirection === 'bar' : defaultHorizontal;
    const grouping = owner.barGroupGrouping ?? defaultGrouping;
    if (horizontal || grouping === 'stacked' || grouping === 'percentStacked') return false;
    groupKeys.add(owner.barGroupIndex != null ? `group-${owner.barGroupIndex}` : 'primary-default');
    if (groupKeys.size > 1) return false;
  }
  return true;
}

/** Primary value-axis plan for the unstacked vertical delegate used by every
 * ChartEx column family. The painter supplies its measured axis length; the
 * resource preflight omits it because authored clipping bounds are size-free. */
export function planChartExColumnValueAxis(
  chart: ChartModel,
  series: readonly ChartSeries[],
  categoryCount: number,
  axisLenPt?: number,
): ValueAxisPlan {
  let dataMin = 0;
  let dataMax = 0;
  const include = (value: number): void => {
    dataMin = Math.min(dataMin, value);
    dataMax = Math.max(dataMax, value);
  };
  for (const owner of series) {
    for (let index = 0; index < Math.min(owner.values.length, categoryCount); index++) {
      const value = owner.values[index];
      if (value == null || !Number.isFinite(value)) continue;
      include(value);
    }
    forEachErrorBarEndpoint(
      owner,
      'y',
      index => index < categoryCount ? owner.values[index] ?? null : null,
      include,
    );
  }
  if (chart.valMax != null) dataMax = chart.valMax;
  if (chart.valMin != null) dataMin = chart.valMin;
  if (dataMin === 0 && dataMax === 0) dataMax = 1;
  return planValueAxis(chart, dataMin, dataMax, axisLenPt, false, 'vertical');
}

/** The exact non-zero clipped span predicate used by the bar painter. */
export function clippedBarBodyHasVisibleSpan(
  first: number,
  second: number,
  clipMin: number,
  clipMax: number,
): boolean {
  const clamp = (value: number): number => Math.max(clipMin, Math.min(clipMax, value));
  const start = clamp(Math.min(first, second));
  const end = clamp(Math.max(first, second));
  return Number.isFinite(start) && Number.isFinite(end) && end > start;
}

export interface ChartExBarBodyVisibilityInput {
  categoryStart: number;
  categoryEnd: number;
  valueStart: number;
  valueEnd: number;
  clipCategory: boolean;
  categoryClipStart?: number;
  categoryClipEnd?: number;
  valueClipStart?: number;
  valueClipEnd?: number;
}

/** The final category-intersection and value-span decision used immediately
 * before a ChartEx bar body is painted. Coordinates are normalized or pixels;
 * only a common 0..1 / plot-rectangle interval is required. */
export function chartExBarBodyIsVisible(
  geometry: ChartExBarBodyVisibilityInput,
): boolean {
  const categoryVisible = !geometry.clipCategory || clippedBarBodyHasVisibleSpan(
    geometry.categoryStart, geometry.categoryEnd,
    geometry.categoryClipStart ?? 0, geometry.categoryClipEnd ?? 1,
  );
  return categoryVisible && clippedBarBodyHasVisibleSpan(
    geometry.valueStart, geometry.valueEnd,
    geometry.valueClipStart ?? 0, geometry.valueClipEnd ?? 1,
  );
}

/** Generic bar-layout visibility for ChartEx delegates that cannot use the
 * compact vertical clustered-column plan. It mirrors the bar painter's group,
 * stack, date-coordinate and primary value-axis calculations in normalized
 * plot coordinates, then calls the same final body predicate as paint. */
export function planChartExGenericBarBodyVisibility(
  chart: ChartModel,
  series: readonly ChartSeries[],
  categoryCount: number,
  dateAxis: DateCategoryAxisPlan | null,
): readonly ReadonlySet<number>[] {
  const defaultHorizontal = chart.chartType === 'clusteredBarH'
    || chart.chartType === 'stackedBarH'
    || chart.chartType === 'stackedBarHPct';
  const defaultStacked = chart.chartType.startsWith('stacked');
  const defaultPercent = chart.chartType === 'stackedBarPct'
    || chart.chartType === 'stackedBarHPct';
  const horizontal = (owner: ChartSeries): boolean => owner.barGroupDirection != null
    ? owner.barGroupDirection === 'bar' : defaultHorizontal;
  const grouping = (owner: ChartSeries): string => owner.barGroupGrouping
    ?? (defaultPercent ? 'percentStacked' : defaultStacked ? 'stacked' : 'clustered');
  const stacked = (owner: ChartSeries): boolean => {
    const value = grouping(owner);
    return value === 'stacked' || value === 'percentStacked';
  };
  const percent = (owner: ChartSeries): boolean => grouping(owner) === 'percentStacked';
  const groupKey = (owner: ChartSeries): string => owner.barGroupIndex != null
    ? `group-${owner.barGroupIndex}`
    : owner.useSecondaryAxis === true ? 'secondary-default' : 'primary-default';
  const groups = new Map<string, ChartSeries[]>();
  const seriesSet = new Set(series);
  for (const owner of series) {
    const key = groupKey(owner);
    const members = groups.get(key);
    if (members) members.push(owner); else groups.set(key, [owner]);
  }
  const hasSecondary = !defaultHorizontal && chart.secondaryValAxis != null
    && chart.series.some(owner => owner.useSecondaryAxis === true);
  const plotGroupBySeries = new Map<ChartSeries, NonNullable<ChartModel['plotGroups']>[number]>();
  const axisPercentState = new Map<string, { count: number; percentCount: number }>();
  for (const plotGroup of chart.plotGroups ?? []) {
    if (plotGroup.seriesCount > 0) {
      const state = axisPercentState.get(plotGroup.valueAxis)
        ?? { count: 0, percentCount: 0 };
      state.count++;
      if (plotGroup.grouping === 'percentStacked') state.percentCount++;
      axisPercentState.set(plotGroup.valueAxis, state);
    }
    for (let index = plotGroup.seriesStart;
      index < plotGroup.seriesStart + plotGroup.seriesCount; index++) {
      const owner = chart.series[index];
      if (owner) plotGroupBySeries.set(owner, plotGroup);
    }
  }
  const percentMultiplier = (owner: ChartSeries): number => {
    if (!percent(owner)) return 1;
    const plotGroup = plotGroupBySeries.get(owner);
    if (!plotGroup) return 100;
    const state = axisPercentState.get(plotGroup.valueAxis);
    return state != null && state.count === state.percentCount ? 100 : 1;
  };
  const primarySeries = hasSecondary
    ? series.filter(owner => owner.useSecondaryAxis !== true) : [...series];
  const primaryGroups = new Map<string, ChartSeries[]>();
  for (const owner of primarySeries) {
    const key = groupKey(owner);
    const members = primaryGroups.get(key);
    if (members) members.push(owner); else primaryGroups.set(key, [owner]);
  }
  const primaryPlotGroups = (chart.plotGroups ?? []).filter(plotGroup =>
    plotGroup.seriesCount > 0 && plotGroup.valueAxis !== 'secondary'
  );
  const primaryPercent = primaryPlotGroups.length > 0
    ? primaryPlotGroups.every(plotGroup => plotGroup.grouping === 'percentStacked')
    : primarySeries.some(percent);
  let dataMin = 0;
  let dataMax = 0;
  const include = (value: number): void => {
    if (!Number.isFinite(value)) return;
    dataMin = Math.min(dataMin, value);
    dataMax = Math.max(dataMax, value);
  };
  for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
    for (const members of primaryGroups.values()) {
      const owner = members[0];
      if (!owner) continue;
      const isStacked = stacked(owner);
      const isPercent = percent(owner);
      const denominator = isPercent
        ? members.reduce((sum, member) =>
          sum + Math.abs(member.values[categoryIndex] ?? 0), 0) || 1
        : 1;
      let positive = 0;
      let negative = 0;
      for (const member of members) {
        const raw = member.values[categoryIndex] ?? 0;
        const value = isPercent ? raw / denominator * percentMultiplier(member) : raw;
        if (!isStacked) include(value);
        else if (value < 0) negative += value;
        else positive += value;
      }
      if (isStacked) {
        include(negative);
        include(positive);
      }
    }
  }
  for (const owner of chart.series) {
    if (seriesSet.has(owner)
      || owner.seriesType !== 'line' && owner.seriesType !== 'area'
      || hasSecondary && owner.useSecondaryAxis === true) continue;
    for (let index = 0; index < Math.min(owner.values.length, categoryCount); index++) {
      const value = owner.values[index];
      if (value != null) include(value);
    }
  }
  if (primaryPercent) {
    if (primarySeries.some(owner => owner.values.some(value => value != null && value > 0))) {
      dataMax = Math.max(dataMax, 100);
    }
    if (primarySeries.some(owner => owner.values.some(value => value != null && value < 0))) {
      dataMin = Math.min(dataMin, -100);
    }
  }
  if (chart.valMax != null) dataMax = primaryPercent ? chart.valMax * 100 : chart.valMax;
  if (chart.valMin != null) dataMin = primaryPercent ? chart.valMin * 100 : chart.valMin;
  if (dataMin === 0 && dataMax === 0) dataMax = 1;
  const valueAxis = planValueAxis(
    chart, dataMin, dataMax, undefined, primaryPercent,
    defaultHorizontal ? 'horizontal' : 'vertical',
  );
  const plottedBarValue = (owner: ChartSeries, categoryIndex: number): number => {
    const raw = owner.values[categoryIndex] ?? 0;
    if (!stacked(owner)) return raw;
    const members = groups.get(groupKey(owner)) ?? [owner];
    const isPercent = percent(owner);
    const denominator = isPercent
      ? members.reduce((sum, member) =>
        sum + Math.abs(member.values[categoryIndex] ?? 0), 0) || 1
      : 1;
    const value = isPercent ? raw / denominator * percentMultiplier(owner) : raw;
    let cumulative = 0;
    for (let index = 0; index <= members.indexOf(owner); index++) {
      const candidateRaw = members[index]?.values[categoryIndex] ?? 0;
      const candidate = isPercent
        ? candidateRaw / denominator * percentMultiplier(owner) : candidateRaw;
      if ((candidate < 0) === (value < 0)) cumulative += candidate;
    }
    return cumulative;
  };
  const secondaryBarSeries = hasSecondary
    ? series.filter(owner => owner.useSecondaryAxis === true) : [];
  const secondaryPlotGroups = (chart.plotGroups ?? []).filter(plotGroup =>
    plotGroup.seriesCount > 0 && plotGroup.valueAxis === 'secondary'
  );
  const secondaryPercent = secondaryPlotGroups.length > 0
    ? secondaryPlotGroups.every(plotGroup => plotGroup.grouping === 'percentStacked')
    : secondaryBarSeries.some(percent);
  const secondaryScaleSeries = chart.series.map(owner => {
    if (owner.useSecondaryAxis !== true) return owner;
    if (seriesSet.has(owner)) {
      return {
        ...owner,
        values: owner.values.map((value, index) => value == null
          ? value : plottedBarValue(owner, index)),
      };
    }
    if (!secondaryPercent) return owner;
    return {
      ...owner,
      values: owner.values.map(value => value == null ? value : value * 100),
    };
  });
  const secondaryScale = chart.secondaryValAxis && hasSecondary
    ? computeSecondaryAxis(
      chart.secondaryValAxis,
      secondaryScaleSeries,
      1,
      defaultHorizontal ? 'x' : 'y',
      secondaryPercent,
      secondaryBarSeries.length > 0,
    ) : null;
  const secondaryBase = secondaryScale
    ? chart.secondaryCatAxis?.crossesAt != null
      && Number.isFinite(chart.secondaryCatAxis.crossesAt)
      ? Math.max(secondaryScale.min, Math.min(
        secondaryScale.max, chart.secondaryCatAxis.crossesAt,
      ))
      : chart.secondaryCatAxis?.crosses === 'max'
        ? secondaryScale.max
        : chart.secondaryCatAxis?.crosses === 'min'
          ? secondaryScale.min
          : Math.max(secondaryScale.min, Math.min(secondaryScale.max, 0))
    : 0;
  const secondaryPosition = secondaryScale?.makeToY(0, 1);
  const catReversed = catAxisReversed(chart);
  const categorySlot = (index: number, isHorizontal: boolean): number => isHorizontal
    ? (catReversed ? index : categoryCount - 1 - index)
    : (catReversed ? categoryCount - 1 - index : index);
  const categorySize = (index: number): number => dateAxis
    ? dateAxis.categoryBandFractions[index] ?? 0
    : 1 / Math.max(1, categoryCount);
  const categoryStart = (index: number, isHorizontal: boolean): number => dateAxis
    ? (dateAxis.positions[index] ?? Number.NaN) - categorySize(index) / 2
    : categorySlot(index, isHorizontal) / Math.max(1, categoryCount);
  const result = series.map(() => new Set<number>());
  for (let categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++) {
    const positiveOffsets = new Map<string, number>();
    const negativeOffsets = new Map<string, number>();
    for (let seriesIndex = 0; seriesIndex < series.length; seriesIndex++) {
      const owner = series[seriesIndex]!;
      const members = groups.get(groupKey(owner)) ?? [owner];
      const isHorizontal = horizontal(owner);
      const isStacked = stacked(owner);
      const isPercent = percent(owner);
      const denominator = isPercent
        ? members.reduce((sum, member) =>
          sum + Math.abs(member.values[categoryIndex] ?? 0), 0) || 1
        : 1;
      const raw = owner.values[categoryIndex] ?? 0;
      const value = isPercent ? raw / denominator * percentMultiplier(owner) : raw;
      const negative = value < 0;
      const key = groupKey(owner);
      const positiveOffset = positiveOffsets.get(key) ?? 0;
      const negativeOffset = negativeOffsets.get(key) ?? 0;
      const memberIndex = Math.max(0, members.indexOf(owner));
      const size = categorySize(categoryIndex);
      const effectiveCount = isStacked ? 1 : Math.max(1, members.length);
      const rawOverlap = members[0]?.barGroupOverlap ?? chart.barOverlap ?? 0;
      const overlap = isStacked || !Number.isFinite(rawOverlap)
        ? 0 : Math.max(-100, Math.min(100, rawOverlap));
      const gap = resolveCategoryGapWidthPercent(
        members[0]?.barGroupGapWidth ?? chart.barGapWidth, 'chartex',
      );
      const width = size / (
        1 + (effectiveCount - 1) * (1 - overlap / 100) + gap / 100
      );
      const pitch = isStacked ? 0 : width * (1 - overlap / 100);
      const clusterWidth = width + (effectiveCount - 1) * pitch;
      const start = categoryStart(categoryIndex, isHorizontal)
        + (size - clusterWidth) / 2 + (isStacked ? 0 : memberIndex * pitch);
      // The generic painter's horizontal branch always maps through valX,
      // even when the series is bound to the secondary axis. Keep preflight
      // on that same primary value axis so resource reachability matches paint.
      const secondary = !isHorizontal && hasSecondary && owner.useSecondaryAxis === true;
      const axisStartValue = isStacked
        ? (negative ? negativeOffset : positiveOffset)
        : secondary ? secondaryBase : 0;
      const axisEndValue = isStacked ? axisStartValue + value : value;
      const valueStart = secondary && secondaryPosition
        ? secondaryPosition(axisStartValue) : valueAxis.frac(axisStartValue);
      const valueEnd = secondary && secondaryPosition
        ? secondaryPosition(axisEndValue) : valueAxis.frac(axisEndValue);
      if (chartExBarBodyIsVisible({
        categoryStart: start,
        categoryEnd: start + width,
        valueStart,
        valueEnd,
        clipCategory: !isHorizontal,
      })) result[seriesIndex]!.add(categoryIndex);
      if (isStacked) {
        if (negative) negativeOffsets.set(key, negativeOffset + value);
        else positiveOffsets.set(key, positiveOffset + value);
      }
    }
  }
  return result;
}

/** Formatting-index selection shared with the ChartEx bar paint path. */
export function chartExColumnPointStyleIndex(
  chart: ChartModel,
  series: ChartSeries,
  seriesIndex: number,
  pointIndex: number,
): number {
  return chartSeriesVariesByPoint(chart, seriesIndex)
    ? pointIndex : chartExSeriesFormatIndex(series, seriesIndex);
}

/** Series/category bodies retained by the ChartEx clustered-column date-axis
 * delegate. Geometry is normalized to the plot width, so preflight and paint
 * share the same authored date range without needing a canvas rectangle. */
export interface ChartExColumnDateCategoryVisibilityPlan {
  positions: readonly number[];
  categoryBandFractions: readonly number[];
  seriesCount: number;
  overlapPct: number;
  gapWidthPct: number;
}

export function planChartExColumnDateCategoryVisibility(
  chart: ChartModel,
  series: readonly ChartSeries[],
  dateAxis: DateCategoryAxisPlan | null,
): ChartExColumnDateCategoryVisibilityPlan | null {
  if (!dateAxis) return null;
  const owner = series[0];
  const rawOverlap = owner?.barGroupOverlap ?? chart.barOverlap ?? 0;
  const overlapPct = Number.isFinite(rawOverlap)
    ? Math.max(-100, Math.min(100, rawOverlap)) : 0;
  const gapWidthPct = resolveCategoryGapWidthPercent(
    owner?.barGroupGapWidth ?? chart.barGapWidth,
    'chartex',
  );
  return {
    positions: dateAxis.positions,
    categoryBandFractions: dateAxis.categoryBandFractions,
    seriesCount: Math.max(1, series.length),
    overlapPct,
    gapWidthPct,
  };
}

export function chartExColumnDateCategoryIsVisible(
  plan: ChartExColumnDateCategoryVisibilityPlan | null,
  seriesIndex: number,
  categoryIndex: number,
): boolean {
  if (!plan) return true;
  const position = plan.positions[categoryIndex];
  const categorySize = plan.categoryBandFractions[categoryIndex];
  if (position == null || categorySize == null) return false;
  const pitchFactor = 1 - plan.overlapPct / 100;
  const denom = 1 + (plan.seriesCount - 1) * pitchFactor + plan.gapWidthPct / 100;
  const barWidth = categorySize / denom;
  const clusterGap = barWidth * pitchFactor;
  const clusterWidth = barWidth + (plan.seriesCount - 1) * clusterGap;
  const categoryStart = position - categorySize / 2;
  const barStart = categoryStart + (categorySize - clusterWidth) / 2
    + seriesIndex * clusterGap;
  return barStart + barWidth > 0 && barStart < 1;
}

import type { ChartDataPointOverride, ChartModel, ChartSeries } from '../types/chart.js';
import { chartCategories } from './category-spacing.js';
import { planHistogramBins } from './histogram-binning.js';
import { planParetoLayout } from './pareto-layout.js';
import { MAX_CANVAS_CHART_POINTS } from './resource-limits.js';
import { chartExSeriesFormatIndex } from './shared/palette.js';
import type { ValueAxisPlan } from './shared/axis.js';
import { chartExDelegateModel } from './effective-style.js';
import {
  chartExColumnDateCategoryIsVisible,
  chartExColumnPointStyleIndex,
  chartExColumnSharedPlanApplies,
  planChartExGenericBarBodyVisibility,
  type ChartExColumnDateCategoryVisibilityPlan,
  clippedBarBodyHasVisibleSpan,
  planChartExColumnDateCategoryVisibility,
  planChartExColumnValueAxis,
} from './chart-ex-column-paint.js';
import {
  deletedLegendEntryIndices,
  legendEntryIsVisible,
  legendEntryRanges,
} from './legend-entry-plan.js';
import { emptyChartExLegendSeries } from './chart-ex-legend-series.js';
import { chartDateAxisPlan } from './shared/secondary-axis.js';

export type ChartExColumnPaintPlan =
  | {
    kind: 'columns';
    series: ChartSeries[];
    categoryCount: number;
    delegateChart: ChartModel;
    valueAxis: ValueAxisPlan;
    dateCategoryVisibility: ChartExColumnDateCategoryVisibilityPlan | null;
    genericBodyVisibility: readonly ReadonlySet<number>[] | null;
    sharedGeometry: boolean;
  }
  | { kind: 'tooManyInputPoints' };

function columnPlan(
  chart: ChartModel,
  series: ChartSeries[],
  categories: string[],
  auxiliarySeries: ChartSeries[] = [],
): Extract<ChartExColumnPaintPlan, { kind: 'columns' }> {
  const delegateChart = chartExDelegateModel(chart, {
    chartType: 'clusteredBar', categories, series: [...series, ...auxiliarySeries],
  });
  const categoryCount = chartCategories(delegateChart).length;
  const sharedGeometry = chartExColumnSharedPlanApplies(delegateChart, series);
  const dateAxis = chartDateAxisPlan(delegateChart, chartCategories(delegateChart));
  return {
    kind: 'columns', series, categoryCount, delegateChart,
    valueAxis: planChartExColumnValueAxis(delegateChart, series, categoryCount),
    dateCategoryVisibility: planChartExColumnDateCategoryVisibility(
      delegateChart, series, sharedGeometry ? dateAxis : null,
    ),
    genericBodyVisibility: sharedGeometry ? null : planChartExGenericBarBodyVisibility(
      delegateChart, series, categoryCount, dateAxis,
    ),
    sharedGeometry,
  };
}

/** Data-mark domain for image preflight and paint-work accounting. Reuse the
 * histogram/Pareto planners used by renderHistogramChart/renderParetoChart;
 * indexed point formatting must follow the displayed bins/reordered bars.
 * MS-ODRAWXML §2.24.3.77 CT_Series supplies formatIdx, ownerIdx and dataPt.
 * Aggregation and hidden-owner policy are already materialized by the parser;
 * do not aggregate again or infer hidden layout from paint formatting here.
 * prepareColumn supplies the painter's optional outline-owner carrier before
 * sorting (its fill is noFill, but an authored point fill still owns its atom).
 */
export function planChartExColumnPaint(
  chart: ChartModel,
  prepareColumn: (series: ChartSeries, index: number, count: number) => ChartSeries
    = series => series,
): ChartExColumnPaintPlan | null {
  if (!['clusteredColumn', 'histogram', 'pareto'].includes(chart.chartType)) return null;
  const empty = (): ChartExColumnPaintPlan => columnPlan(chart, [], []);
  if (chart.chartexSuppressGeometry) return empty();
  if (chart.chartType === 'histogram') {
    const source = chart.series[0];
    if (!source) return empty();
    const bins = planHistogramBins(
      source.values, chart.chartexHistogramBinning ?? {}, source.valFormatCode,
      chart.date1904 === true,
    );
    if (bins.kind === 'tooManyInputPoints') return bins;
    return columnPlan(chart, [
      { ...source, categories: undefined, values: bins.counts, valFormatCode: null },
    ], bins.categories);
  }
  const raw = chart.chartType === 'pareto'
    ? chart.series.filter(series => series.seriesType !== 'line') : chart.series;
  let sourcePoints = 0;
  for (const series of raw) {
    sourcePoints += series.values.length;
    if (sourcePoints > MAX_CANVAS_CHART_POINTS) return { kind: 'tooManyInputPoints' };
  }
  const columns = raw.map((series, index) => prepareColumn(series, index, raw.length));
  const layouts = chart.chartType === 'pareto' || chart.chartexParetoSortDescending
    ? columns.map(series => planParetoLayout(series, chart.categories, {
      sortDescending: chart.chartexParetoSortDescending ?? true,
      keepUnvaluedCategories: true,
    })) : null;
  const displayed = layouts ? layouts.map(layout => layout.orderedSeries) : columns;
  let categories = displayed[0]?.categories ?? chart.categories;
  if (chart.chartType === 'pareto') {
    const owner = layouts?.[chart.chartexParetoOwnerIndex ?? 0];
    const first = layouts?.[0];
    if (!owner || !first || owner.points.length === 0 || first.points.length === 0) return empty();
    // renderParetoChart appends an empty category for its flat cumulative-line
    // endpoint. A longer later owner can still paint a valued bar in that slot.
    if (chart.chartexParetoFlatEndpoint === true
      && first.categories.length === first.points.length) categories = [...categories, ''];
  }
  // Match the actual bar delegate, including first-series category precedence
  // and chartCategories' ordinal fallback when no category text exists.
  const series = chart.chartType === 'pareto' ? displayed : displayed.filter(series =>
    series.seriesType !== 'line' && series.seriesType !== 'scatter' && series.seriesType !== 'area');
  const authoredAuxiliary = chart.chartType === 'clusteredColumn'
    ? displayed.filter(series =>
      series.seriesType === 'line' || series.seriesType === 'scatter'
      || series.seriesType === 'area')
    : [];
  const generatedAuxiliary = chart.chartType === 'pareto' && series[0]
    ? [{ ...series[0], values: [], seriesType: 'line' as const, useSecondaryAxis: true }]
    : chart.chartexShowUnpairedPercentageAxis && series[0]
      ? [{ ...series[0], values: [], seriesType: 'line' as const, useSecondaryAxis: true }]
      : [];
  return columnPlan(chart, series, categories, [...authoredAuxiliary, ...generatedAuxiliary]);
}

/** Non-empty column bodies only; empty bins/categories keep ticks but consume
 * no fill paint. Point overrides are indexed in the plan's displayed domain. */
export function visitChartExColumnPaintSites(
  plan: Extract<ChartExColumnPaintPlan, { kind: 'columns' }>,
  visit: (series: ChartSeries, point: ChartDataPointOverride | undefined,
    styleIndex: number, count: number) => void | false,
): void {
  for (let seriesIndex = 0; seriesIndex < plan.series.length; seriesIndex++) {
    const series = plan.series[seriesIndex]!;
    const overrides = new Map((series.dataPointOverrides ?? []).map(point => [point.idx, point]));
    for (let index = 0; index < Math.min(series.values.length, plan.categoryCount); index++) {
      const value = series.values[index];
      if (value == null || !Number.isFinite(value) || value === 0) continue;
      if (plan.sharedGeometry) {
        if (!chartExColumnDateCategoryIsVisible(
          plan.dateCategoryVisibility, seriesIndex, index,
        )) continue;
        if (!clippedBarBodyHasVisibleSpan(
          plan.valueAxis.frac(0), plan.valueAxis.frac(value), 0, 1,
        )) continue;
      } else if (plan.genericBodyVisibility?.[seriesIndex]?.has(index) !== true) continue;
      const styleIndex = chartExColumnPointStyleIndex(
        plan.delegateChart, series, seriesIndex, index,
      );
      if (visit(series, overrides.get(index), styleIndex, plan.series.length) === false) return;
    }
  }
}

/** Legend-key fill sites produced by the delegated bar painter. Reachability
 * is computed from the same synthetic legend series used by paint; notably it
 * has no values, so a one-series varyColors chart exposes no point keys. */
export function visitChartExColumnLegendSites(
  plan: Extract<ChartExColumnPaintPlan, { kind: 'columns' }>,
  visit: (series: ChartSeries, point: ChartDataPointOverride | undefined,
    styleIndex: number, count: number) => void | false,
): void {
  if (!plan.delegateChart.showLegend || plan.series.length === 0) return;
  const legendChart = {
    ...plan.delegateChart,
    series: plan.series.map(series => emptyChartExLegendSeries(series.name, series)),
  };
  const ranges = legendEntryRanges(legendChart, true);
  const deleted = deletedLegendEntryIndices(legendChart);
  for (let seriesIndex = 0; seriesIndex < plan.series.length; seriesIndex++) {
    if (!legendEntryIsVisible(ranges, deleted, seriesIndex)) continue;
    const series = plan.series[seriesIndex]!;
    if (visit(
      series, undefined, chartExSeriesFormatIndex(series, seriesIndex), plan.series.length,
    ) === false) return;
  }
}

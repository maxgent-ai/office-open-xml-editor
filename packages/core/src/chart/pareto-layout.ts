import type {
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartSeries,
} from '../types/chart';

export interface ParetoPoint {
  sourceIndex: number;
  category: string;
  value: number;
  cumulativeFraction: number;
}

export interface ParetoLayout {
  points: ParetoPoint[];
  /** Source series reordered for the frequency bars. */
  orderedSeries: ChartSeries;
  /** Reordered series whose values are the cumulative 0..1 fractions. */
  series: ChartSeries;
  categories: string[];
}

export interface ParetoLayoutOptions {
  /** Owner-backed Pareto bars sort frequencies; a standalone paretoLine does not. */
  sortDescending?: boolean;
  /** ChartEx aggregation preserves named categories lacking a numeric point. */
  keepUnvaluedCategories?: boolean;
}

function remapIndexed<T extends { idx: number }>(
  values: readonly T[] | null | undefined,
  newIndexBySource: ReadonlyMap<number, number>,
): T[] | null | undefined {
  if (values == null) return values;
  return values.flatMap(value => {
    const idx = newIndexBySource.get(value.idx);
    return idx == null ? [] : [{ ...value, idx }];
  });
}

function reorderNullable<T>(
  values: readonly T[] | null | undefined,
  sourceIndices: readonly number[],
): Array<T | null> | null | undefined {
  if (values == null) return values;
  return sourceIndices.map(index => values[index] ?? null);
}

/**
 * Derive a deterministic Pareto order without mutating the authored model.
 *
 * Finite non-negative values participate. Ties retain source order, missing or
 * invalid values are omitted from frequency points, and a zero-total series
 * produces finite zero cumulative values. ChartEx can retain named category
 * ticks with no value: after the sorted points, or at their source slot when
 * authored order is kept. Indexed point/label properties follow their source
 * point through the reorder.
 */
export function planParetoLayout(
  series: ChartSeries,
  chartCategories: readonly string[],
  options: ParetoLayoutOptions = {},
): ParetoLayout {
  const retained = series.values
    .map((value, sourceIndex) => ({ value, sourceIndex }))
    .filter((entry): entry is { value: number; sourceIndex: number } =>
      entry.value != null && Number.isFinite(entry.value) && entry.value >= 0
    );
  if (options.sortDescending !== false) {
    retained.sort((a, b) => b.value - a.value || a.sourceIndex - b.sourceIndex);
  }

  // Normalize before summing so finite values near Number.MAX_VALUE cannot
  // overflow the denominator to Infinity and collapse early fractions to 0.
  // The scale is the maximum retained value: with authored (unsorted) order the
  // first point need not be the largest, and a leading zero would zero every
  // fraction.
  const scale = retained.reduce((max, entry) => Math.max(max, entry.value), 0);
  const normalizedTotal = scale > 0
    ? retained.reduce((sum, entry) => sum + entry.value / scale, 0)
    : 0;
  let running = 0;
  const points = retained.map((entry): ParetoPoint => {
    if (scale > 0) running += entry.value / scale;
    return {
      sourceIndex: entry.sourceIndex,
      category: series.categories?.[entry.sourceIndex]
        ?? chartCategories[entry.sourceIndex]
        ?? String(entry.sourceIndex + 1),
      value: entry.value,
      cumulativeFraction: normalizedTotal > 0
        ? (running >= normalizedTotal ? 1 : running / normalizedTotal)
        : 0,
    };
  });
  const sourceIndices = points.map(point => point.sourceIndex);
  if (options.keepUnvaluedCategories) {
    for (let index = 0; index < chartCategories.length; index++) {
      if (series.values[index] == null && chartCategories[index] !== '') {
        sourceIndices.push(index);
      }
    }
    // [MS-ODRAWXML] CT_NumericValue/CT_StringValue@idx identify the source
    // point. Without the frequency sort, bars keep authored (source-index)
    // order, so an unvalued named category keeps its own slot instead of
    // moving after every valued point (which also mislabelled other series'
    // bars drawn against the first series' categories). With sorting it
    // still follows the sorted points. Evidence limit: no Office control
    // places an interior unvalued slot; the renderers' flat-endpoint rule is
    // unchanged.
    if (options.sortDescending === false) sourceIndices.sort((a, b) => a - b);
  }
  const newIndexBySource = new Map(
    sourceIndices.map((sourceIndex, newIndex) => [sourceIndex, newIndex]),
  );
  const categories = sourceIndices.map(index => series.categories?.[index]
    ?? chartCategories[index] ?? String(index + 1));
  const values = sourceIndices.map(index => {
    const value = series.values[index];
    return value != null && Number.isFinite(value) && value >= 0 ? value : null;
  });
  // A slot without a value adds nothing, so it carries the cumulative share
  // of the slots before it (0 before the first valued point).
  const fractionBySource = new Map(
    points.map(point => [point.sourceIndex, point.cumulativeFraction]),
  );
  let carried = 0;
  const cumulative = sourceIndices.map(index => {
    carried = fractionBySource.get(index) ?? carried;
    return carried;
  });
  const reordered = {
    ...series,
    categories,
    catFormatCodes: reorderNullable(series.catFormatCodes, sourceIndices),
    dataPointColors: reorderNullable(series.dataPointColors, sourceIndices),
    dataLabelColors: reorderNullable(series.dataLabelColors, sourceIndices),
    dataPointOverrides: remapIndexed<ChartDataPointOverride>(
      series.dataPointOverrides,
      newIndexBySource,
    ),
    dataLabelOverrides: remapIndexed<ChartDataLabelOverride>(
      series.dataLabelOverrides,
      newIndexBySource,
    ),
  };

  return {
    points,
    categories,
    orderedSeries: {
      ...reordered,
      values,
    },
    series: {
      ...reordered,
      values: cumulative,
    },
  };
}

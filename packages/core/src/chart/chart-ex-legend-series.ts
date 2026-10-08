import type { ChartSeries } from '../types/chart.js';

/** The synthetic series painted by ChartEx legends. The empty value domain is
 * observable: varyColors cannot create point keys from the plotted source
 * series after the painter replaces it. Trendlines remain separate keys. */
export function emptyChartExLegendSeries(
  name: string,
  source?: { trendLines?: ChartSeries['trendLines'] },
): ChartSeries {
  return { name, values: [], color: null, trendLines: source?.trendLines };
}

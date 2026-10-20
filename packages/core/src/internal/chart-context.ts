import { safeUtf16Prefix } from './utf16-prefix.js';

export const MAX_CHART_CONTEXT_TEXT_CHARACTERS = 65_536;
export const DEFAULT_CHART_CONTEXT_TEXT_CHARACTERS = 16_384;

export interface BoundedChartContextText {
  readonly text: string;
  readonly truncated: boolean;
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

/** The bounded summary only needs this immutable subset of a chart model. */
export interface ChartContextModel {
  readonly chartType: string;
  readonly title: string | null;
  readonly categories: readonly string[];
  readonly series: readonly Readonly<{
    readonly name: string;
    readonly values: readonly (number | null)[];
  }>[];
}

function normalizeLimit(requested: number | undefined): number {
  const value = requested ?? DEFAULT_CHART_CONTEXT_TEXT_CHARACTERS;
  if (!Number.isFinite(value) || value < 0) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  return Math.min(MAX_CHART_CONTEXT_TEXT_CHARACTERS, Math.floor(value));
}

/**
 * Build the same compact chart summary for every OOXML format without first
 * materializing an unbounded string. The model is immutable and parser-bounded;
 * iteration stops as soon as the requested text budget is exhausted.
 */
export function boundedChartContextText(
  chart: ChartContextModel,
  requestedMaxTextCharacters?: number,
): BoundedChartContextText {
  const maxTextCharacters = normalizeLimit(requestedMaxTextCharacters);
  const chunks: string[] = [];
  let textCharacters = 0;
  let truncated = false;
  let hasPart = false;

  const append = (value: string, beginsPart: boolean): boolean => {
    if (beginsPart && hasPart) {
      if (textCharacters >= maxTextCharacters) {
        truncated = true;
        return false;
      }
      chunks.push('\n');
      textCharacters++;
    }
    hasPart = true;
    const allowed = Math.max(0, maxTextCharacters - textCharacters);
    const chunk = safeUtf16Prefix(value, allowed);
    chunks.push(chunk);
    textCharacters += chunk.length;
    if (chunk.length < value.length) {
      truncated = true;
      return false;
    }
    return true;
  };

  if (!append(`Chart type: ${chart.chartType}`, true)) {
    return { text: chunks.join(''), truncated, textCharacters, maxTextCharacters };
  }
  if (chart.title !== null && !append(`Title: ${chart.title}`, true)) {
    return { text: chunks.join(''), truncated, textCharacters, maxTextCharacters };
  }
  if (chart.categories.length > 0) {
    if (!append('Categories: ', true)) {
      return { text: chunks.join(''), truncated, textCharacters, maxTextCharacters };
    }
    for (const [index, category] of chart.categories.entries()) {
      if (!append(`${index === 0 ? '' : ', '}${category}`, false)) {
        return { text: chunks.join(''), truncated, textCharacters, maxTextCharacters };
      }
    }
  }
  for (const series of chart.series) {
    if (!append(`Series ${series.name || '(unnamed)'}: `, true)) break;
    for (const [index, value] of series.values.entries()) {
      if (!append(`${index === 0 ? '' : ', '}${value === null ? '' : value}`, false)) {
        return { text: chunks.join(''), truncated, textCharacters, maxTextCharacters };
      }
    }
    if (truncated) break;
  }
  return { text: chunks.join(''), truncated, textCharacters, maxTextCharacters };
}

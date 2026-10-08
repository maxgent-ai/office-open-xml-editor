import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import { normalizeFontMetricFamily } from '@silurus/ooxml-core';
import type { FontResolution } from '../layout/font-service.js';
import { referenceFontLineMetrics } from '../reference-font-line-metrics.js';

/** Vertical admission shared by ordinary text and numbering markers. Resource
 * identity, tuple ambiguity and glyph coverage are resolved before this gate.
 * Canvas sides remain the fallback; exact spacing suppresses design geometry. */
export function selectedFontLineMetric(
  selected: FontResolution | undefined,
  localFont: ResolvedFontMetric | undefined,
  naturalMetricAllowed: boolean,
) {
  const resourceMetric = (naturalMetricAllowed || localFont?.designAscentRatio == null)
    && (localFont?.lineHeightRatio != null || localFont?.designAscentRatio != null
      || localFont?.eastAsianLineHeightRatio != null) ? localFont : undefined;
  const referenceMetric = naturalMetricAllowed && !resourceMetric
    && mayUseAuthoredReferenceVerticalMetric(selected)
    ? referenceFontLineMetrics(selected.requestedFamily, selected.weight, selected.style)
    : undefined;
  return { resourceMetric, referenceMetric, lineMetric: resourceMetric ?? referenceMetric };
}

/** A subset face lends design metrics only to spans whose every scalar it owns.
 * A FontFace tuple can win CSS selection while Canvas silently paints missing
 * glyphs from another face. Ranges are normalized once at the service boundary. */
export function metricCoversText(metric: ResolvedFontMetric, text: string): boolean {
  const ranges = metric.unicodeRanges;
  if (ranges === undefined || text.length === 0) return true;
  for (const scalar of text) {
    const codePoint = scalar.codePointAt(0)!;
    let low = 0;
    let high = ranges.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (ranges[middle][1] < codePoint) low = middle + 1;
      else high = middle;
    }
    if (low >= ranges.length || ranges[low][0] > codePoint) return false;
  }
  return true;
}


/** An authored-family key alone cannot prove the selected route. Match the
 * chosen face tuple and acquisition source before admitting its metric;
 * a substitute may preserve the request while painting another face. */
export function metricMatchesSelectedFont(
  metric: ResolvedFontMetric,
  selected: FontResolution,
  text: string,
): boolean {
  if (normalizeFontMetricFamily(metric.family)
    !== normalizeFontMetricFamily(selected.resolvedFamily)
    || (metric.weight ?? 400) !== selected.weight
    || (metric.style ?? 'normal') !== selected.style
    || !metricCoversText(metric, text)) return false;
  const identity = metric.sourceIdentity;
  if (selected.source === 'embedded') {
    // The production loader routes Canvas to a resource-derived private family.
    // Its metric must come from that same registration, including when another
    // document embeds different bytes under the same authored family/style.
    return selected.resourceIdentity === undefined
      ? identity?.startsWith('embedded:') === true
      : identity === selected.resourceIdentity;
  }
  if (selected.source === 'local' || selected.source === 'substitute') {
    // The same CSS tuple may be supplied by an unrelated caller or another
    // loaded face. Only the resource that created this resolver route may lend
    // vertical geometry; a matching family name is insufficient.
    return identity !== undefined && identity === selected.resourceIdentity;
  }
  return false;
}


/** Multiple resources may be registered under one CSS family/style. The
 * authored alias does not control Canvas's choice among overlapping faces;
 * only geometry shared by every covering resource is safe to admit. */
export function sameResourceGeometry(a: ResolvedFontMetric, b: ResolvedFontMetric): boolean {
  return a.lineHeightRatio === b.lineHeightRatio
    && a.designAscentRatio === b.designAscentRatio
    && a.designDescentRatio === b.designDescentRatio
    && a.lineGapRatio === b.lineGapRatio
    && a.eastAsianLineHeightRatio === b.eastAsianLineHeightRatio
    && a.fontBoxRatio === b.fontBoxRatio;
}


export type MetricTupleIndex = ReadonlyMap<string, readonly ResolvedFontMetric[]>;

export const metricTupleIndexes = new WeakMap<object, MetricTupleIndex>();


export function selectedMetricKey(
  family: string,
  weight: number,
  style: string,
): string {
  return `${normalizeFontMetricFamily(family)}\0${weight}\0${style}`;
}


/** A frozen service snapshot is immutable for the document lifetime, so one
 * tuple/source index can serve every paragraph. A custom mutable service gets
 * one temporary index per synchronous buildSegments call, then can change its
 * entries before the next call. */
export function indexedFontMetrics(
  metrics: Readonly<Record<string, ResolvedFontMetric>>,
): MetricTupleIndex {
  const frozenRoot = Object.isFrozen(metrics);
  if (frozenRoot) {
    const retained = metricTupleIndexes.get(metrics);
    if (retained) return retained;
  }
  const values = Object.values(metrics);
  const cacheable = frozenRoot && values.every((metric) => Object.isFrozen(metric));
  const mutable = new Map<string, ResolvedFontMetric[]>();
  for (const metric of values) {
    const key = selectedMetricKey(metric.family, metric.weight ?? 400, metric.style ?? 'normal');
    const candidates = mutable.get(key);
    if (candidates) candidates.push(metric);
    else mutable.set(key, [metric]);
  }
  for (const candidates of mutable.values()) Object.freeze(candidates);
  if (cacheable) metricTupleIndexes.set(metrics, mutable);
  return mutable;
}


export function selectResourceMetric(
  index: MetricTupleIndex,
  selected: FontResolution,
  probeText: string,
): ResolvedFontMetric | undefined {
  if (selected.source !== 'embedded' && selected.source !== 'local'
    && selected.source !== 'substitute') return undefined;
  const candidates = index.get(selectedMetricKey(
    selected.resolvedFamily, selected.weight, selected.style,
  )) ?? [];
  let first: ResolvedFontMetric | undefined;
  for (const metric of candidates) {
    if (!metricMatchesSelectedFont(metric, selected, probeText)) continue;
    first ??= metric;
  }
  if (!first) return undefined;
  // Source identity proves which resource was requested, but not which of the
  // overlapping CSS faces Canvas actually paints. Compare every covering
  // resource in the tuple, including resources acquired from other sources.
  for (const metric of candidates) {
    if (metricCoversText(metric, probeText)
      && !sameResourceGeometry(first, metric)) return undefined;
  }
  return first;
}


/** Width inference has a stricter, independent projection. Conflicting xAvg
 * metadata cannot revoke a valid vertical metric for the same selected face. */
export function selectResourceAverageWidthRatio(
  index: MetricTupleIndex,
  selected: FontResolution,
  probeText: string,
): number | undefined {
  if (selected.source !== 'embedded' && selected.source !== 'local'
    && selected.source !== 'substitute') return undefined;
  const candidates = index.get(selectedMetricKey(
    selected.resolvedFamily, selected.weight, selected.style,
  )) ?? [];
  const withSpace = `${probeText} `;
  const first = candidates.find((metric) => metricMatchesSelectedFont(metric, selected, withSpace));
  const ratio = first?.averageCharWidthRatio;
  if (ratio == null || ratio <= 0) return undefined;
  return candidates.every((metric) => !metricCoversText(metric, withSpace)
    || metric.averageCharWidthRatio === ratio) ? ratio : undefined;
}


export function mayUseExactLocalReferenceWidthMetric(
  selected: FontResolution | undefined,
): selected is FontResolution {
  // The width floor requires an exact selected face. A native CSS family list
  // cannot establish which fallback Canvas painted or measured the glyph.
  return selected?.source === 'local'
    && selected.resourceIdentity?.startsWith('office-local:') === true;
}


export function mayUseAuthoredReferenceVerticalMetric(
  selected: FontResolution | undefined,
): selected is FontResolution {
  // Library pagination policy for an unavailable authored face: keep the
  // document's known OpenType line box while Canvas paints/measures a serif or
  // sans fallback. Word's §17.3.1.33 automatic/atLeast line advances follow
  // the authored font; a controlled missing-Calibri case matched Word's five
  // pages only with this vertical projection. This says nothing about glyph
  // coverage or advances, which remain selected-resource/Canvas measurements.
  // A loaded local() tuple may also use the pinned reference when no parsed
  // resource metric exists. A registered substitute is a different selected
  // face: if its resource metric is unavailable, use its Canvas line box in
  // both axes instead of importing the authored font's vertical geometry.
  return selected?.source === 'native'
    || mayUseExactLocalReferenceWidthMetric(selected);
}

/**
 * Font sources (`LoadOptions.fonts`) for DrawingML charts, shared by DOCX,
 * XLSX and PPTX. Library resource policy, not OOXML behaviour.
 *
 * - Demand: the families a chart paints with, from one generic walk over the
 *   canonical {@link ChartModel}, so a face named only by a chart is demanded
 *   too.
 * - Paint: charts paint the substitutions a font source actually loaded
 *   through {@link ChartModel.fontSubstitutions}. Chart labels are not split
 *   by script, so script-scoped substitutes are left out.
 */
import type { ChartModel } from '../types/chart.js';
import type { FontSubstituteScript } from './substitute-script.js';

/** Keys of a chart model whose string value names a text face: every
 * `…FontFace` / `fontFace` (title, axes, axis titles, data labels, legend,
 * legend entries, rich runs, secondary axes) and the theme faces charts
 * default to (ECMA-376 §21.2.2.216 txPr, §20.1.4.2 font scheme). */
const CHART_FACE_KEY = /^(?:fontFace|[A-Za-z]+FontFace|themeMajorFontLatin|themeMinorFontLatin)$/;

/** The families a chart paints with, from one generic walk over its model.
 * Theme references (`+mj-*` / `+mn-*`, §20.1.4.1.16) resolve to the chart's
 * theme faces. The library fallback faces of chart stacks (Calibri, Arial)
 * are not demanded: they are fallbacks, painted with a source face only when
 * one loaded for them. */
export function chartFontFamilies(chart: unknown): string[] {
  const families = new Set<string>();
  const model = chart as { themeMajorFontLatin?: string | null; themeMinorFontLatin?: string | null } | null;
  const pending: unknown[] = [chart];
  const seen = new Set<object>();
  while (pending.length) {
    const value = pending.pop();
    if (value === null || typeof value !== 'object' || seen.has(value) || ArrayBuffer.isView(value)) continue;
    seen.add(value);
    for (const [key, child] of Object.entries(value)) {
      if (typeof child === 'string') {
        if (!CHART_FACE_KEY.test(key)) continue;
        const face = child.startsWith('+mj') ? model?.themeMajorFontLatin
          : child.startsWith('+mn') ? model?.themeMinorFontLatin : child;
        const name = face?.trim();
        if (name && !name.startsWith('+')) families.add(name);
      } else if (child !== null && typeof child === 'object') {
        pending.push(child);
      }
    }
  }
  return [...families];
}

/** `LoadedFontSources.substitutions`: lower-cased authored family → served family. */
export type LoadedChartFontSubstitutions = ReadonlyMap<string, { readonly family: string; readonly script?: FontSubstituteScript }>;

const unscopedBySnapshot = new WeakMap<LoadedChartFontSubstitutions, Readonly<Record<string, string>> | null>();

/** The unscoped substitutions of one loaded snapshot as the
 * {@link ChartModel.fontSubstitutions} record, or `undefined` when there are
 * none. The record has a null prototype (keys are document font names, e.g.
 * `__proto__`) and is memoised per snapshot (snapshots are immutable). */
export function chartFontSubstitutions(
  substitutions: LoadedChartFontSubstitutions | undefined,
): Readonly<Record<string, string>> | undefined {
  if (!substitutions?.size) return undefined;
  let unscoped = unscopedBySnapshot.get(substitutions);
  if (unscoped === undefined) {
    const record = Object.create(null) as Record<string, string>;
    for (const [family, substitution] of substitutions) if (!substitution.script) record[family] = substitution.family;
    unscoped = Object.keys(record).length ? Object.freeze(record) : null;
    unscopedBySnapshot.set(substitutions, unscoped);
  }
  return unscoped ?? undefined;
}

/** The chart model to paint: `chart` with `fontSubstitutions`, or `chart`
 * unchanged without any (no font source, nothing loaded). */
export function withChartFontSubstitutions<T extends ChartModel>(
  chart: T,
  substitutions: Readonly<Record<string, string>> | undefined,
): T {
  return substitutions ? { ...chart, fontSubstitutions: substitutions } : chart;
}

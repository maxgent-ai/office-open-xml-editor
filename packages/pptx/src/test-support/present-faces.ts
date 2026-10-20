import { normalizeLocalFontMetricFamily, type OfficeFontFallbackRoute } from '@silurus/ooxml-core';
import { renderTextBody as renderTextBodyRaw } from '../renderer.js';
import type { TextBody } from '../types.js';

type RenderTextBodyArgs = Parameters<typeof renderTextBodyRaw>;

/**
 * Line metrics admit a face's PowerPoint catalog geometry only when that face
 * is present (an exact local() office route was loaded). These controls model
 * PowerPoint with its authored faces installed, so declare every family the
 * body names as present, in each weight/style tuple, painting as itself.
 */
export function presentFacesContext(body: TextBody): NonNullable<RenderTextBodyArgs[13]> {
  const families = new Set<string>();
  for (const paragraph of body.paragraphs ?? []) {
    if (paragraph.defFontFamily) families.add(paragraph.defFontFamily);
    for (const run of paragraph.runs ?? []) {
      if (run.type !== 'text') continue;
      for (const family of [run.fontFamily, run.fontFamilyEa, run.fontFamilyCs]) {
        if (family) families.add(family);
      }
    }
  }
  const officeFontRoutes: Record<string, OfficeFontFallbackRoute> = {};
  for (const family of families) {
    for (const weight of [400, 700]) {
      for (const style of ['normal', 'italic'] as const) {
        const key = normalizeLocalFontMetricFamily(family);
        const routeKey = weight === 400 && style === 'normal' ? key : `${key}:${weight}:${style}`;
        officeFontRoutes[routeKey] = {
          requestedFamily: family, family, source: 'local',
          resourceIdentity: `office-local:local("${family}")`, weight, style,
          metric: { family, weight, style },
        } as OfficeFontFallbackRoute;
      }
    }
  }
  return { themeMajorFont: null, themeMinorFont: null, dpr: 1, officeFontRoutes };
}

/** renderTextBody with the body's faces declared present unless a context is given. */
export function renderTextBodyWithPresentFaces(...args: RenderTextBodyArgs): ReturnType<typeof renderTextBodyRaw> {
  const next = [...args] as RenderTextBodyArgs;
  for (let index = next.length; index < 13; index++) (next as unknown[])[index] = undefined;
  if (next[13] === undefined) next[13] = presentFacesContext(next[1]);
  return renderTextBodyRaw(...next);
}

import { findReferenceFontMetrics } from '@silurus/ooxml-core';
import { excelDrawingMlLineRatios } from '@silurus/ooxml-core/internal/office-auto-line';
import type { OfficeFontFallbackRequest, OfficeFontFallbackRoute, SpaceLine } from '@silurus/ooxml-core';
import type { ShapeText, ShapeTextRun } from './types.js';

type TextRun = Extract<ShapeTextRun, { type: 'text' }>;

/** Excel's natural line box of one shape-text run, as em ratios. */
export interface ShapeRunLineRatios {
  ascentRatio: number;
  descentRatio: number;
}

/** The single face that owns a run's line box, or undefined when the run
 * names no face or routes East Asian / complex-script text to another face.
 * The shape paint path selects a:latin; a distinct a:ea or a:cs face would
 * need script-run routing before its metrics could own the line. */
function soleFace(run: TextRun): string | undefined {
  const face = run.fontFace?.trim();
  if (!face) return undefined;
  const same = (other: string | undefined) => !other?.trim()
    || other.trim().toLocaleLowerCase('en-US') === face.toLocaleLowerCase('en-US');
  return same(run.fontFaceEa) && same(run.fontFaceCs) ? face : undefined;
}

/** Text runs of a shape body whose face could own Excel's line box. These are
 * the tuples the workbook preflights as exact local faces. */
export function shapeLineFontRuns(text: ShapeText): TextRun[] {
  const runs: TextRun[] = [];
  for (const paragraph of text.paragraphs) for (const run of paragraph.runs) {
    if (run.type === 'text' && soleFace(run)) runs.push(run);
  }
  return runs;
}

/**
 * A spacing value (`a:lnSpc`, `a:spcBef`, `a:spcAft`) as Excel lays shape
 * text out with it. Excel rounds each value on its own to a whole unit,
 * halves rounding up, before any line box or gap is built: `a:spcPts` to whole
 * points, `a:spcPct` to whole percent. The percentage's result in points is
 * not rounded further, and spcAft and spcBef on one gap are rounded one by one,
 * not as a sum.
 *
 * Observed Excel for Mac 16.113.2 behaviour (#1604 boundary and rounding
 * controls). Neighbouring values painted identically within each whole unit
 * and differently across units:
 * - lnSpc spcPts: Meiryo 14 pt 36-40 pt, Yu Gothic 14 pt 26.5-30.5 pt and
 *   Meiryo 24 pt 62.5-66.5 pt in 0.25 pt steps plus 0.01-0.05 pt steps.
 *   36.25 painted like 36, 36.5-37.35 like 37, 37.5-38.25 like 38.
 * - lnSpc spcPct: the same faces across their 4d boundary (Meiryo 132-140 %,
 *   Yu Gothic 114-122 %) in 0.5 % steps plus 0.1 % steps. 136.1-136.4 %
 *   painted like 136 %, 136.5-136.7 % like 137 %, 132.5 % like 133 %.
 *   Meiryo 24 pt at 136.4 % kept its natural descent. Unrounded, its H
 *   (63.835 pt) would pass 4d (63.834 pt) and step by 4.3 pt. H itself is not
 *   rounded: Yu Gothic 14 pt at 118 % (H 27.64 pt) kept its descent, which it
 *   would not if H rounded to 28 pt.
 * - spcBef/spcAft spcPts 3.0-4.0 pt in 0.1 pt steps: 3.0-3.4 painted like 3,
 *   3.5-4.0 like 4, 10.5 like 11. Both 3.3 pt on one gap gave 6 pt, not 7.
 * - spcBef/spcAft spcPct 10-15 % of the natural line in 0.5 % steps plus
 *   12.2/12.8 %: 10.5 and 11 % painted alike, as did 11.5-12.2 %, 12.5-13 %,
 *   13.5-14 % and 14.5-15 %.
 */
export function excelShapeSpacing(spacing: SpaceLine | null | undefined): SpaceLine | null | undefined {
  if (!spacing) return spacing;
  if (spacing.type === 'pts') return { type: 'pts', val: Math.floor(spacing.val + 0.5) };
  return { type: 'pct', val: Math.floor(spacing.val / 1000 + 0.5) * 1000 };
}

/** One key rule for workbook, worker, and synchronous shape paint. */
export function officeRequestKey(request: OfficeFontFallbackRequest): string {
  const family = request.family.trim().toLowerCase();
  const weight = request.weight ?? 400;
  const style = request.style ?? 'normal';
  return weight === 400 && style === 'normal' ? family : `${family}:${weight}:${style}`;
}

export function shapeOfficeRouteKey(run: TextRun): string {
  return officeRequestKey({ family: run.fontFace!, weight: run.bold ? 700 : 400,
    style: run.italic ? 'italic' : 'normal' });
}

/** A loaded face's browser-observable font box, as em ratios. */
export interface ShapeFontBox {
  ascent: number;
  descent: number;
}

/** Reads the font box of a route's loaded face. Returns undefined when the
 * engine does not expose it. */
export type ShapeFontBoxProbe = (route: OfficeFontFallbackRoute) => ShapeFontBox | undefined;

// Probe size and tolerance for matching a loaded face to a catalog profile.
// Blink rounds a face's ascent and descent to whole pixels on most platforms,
// so at 1000 px the observation carries at most 0.0005 em of rounding. Two
// pixels (0.002 em, about 4 design units at 2048 upm) covers that with margin
// and stays well below the smallest gap between disagreeing catalog profiles
// whose boxes differ at all (Helvetica Neue Bold, 0.014 em). Copies whose boxes
// are identical, such as Times New Roman's, match together and are declined.
// A wider tolerance can only make more profiles match, and a match with
// disagreeing profiles is declined, so it never admits a wrong profile.
const PROBE_PX = 1000;
const PROBE_TOLERANCE_EM = 2 / PROBE_PX;

const probedBoxes = new WeakMap<OfficeFontFallbackRoute, ShapeFontBox | null>();

/** A probe that measures the route's face on `ctx`, once per route object.
 * TextMetrics.fontBoundingBoxAscent/Descent report the face's own ascent and
 * descent (hhea on CoreText, usWin or typo on other rasterizers), not the
 * requested family's name, so they identify which copy local() loaded. */
export function canvasShapeFontBoxProbe(
  ctx: Pick<CanvasRenderingContext2D, 'font' | 'measureText'>,
): ShapeFontBoxProbe {
  return (route) => {
    const cached = probedBoxes.get(route);
    if (cached !== undefined) return cached ?? undefined;
    const saved = ctx.font;
    let box: ShapeFontBox | null = null;
    try {
      ctx.font = `${route.style} ${route.weight} ${PROBE_PX}px "${route.family}"`;
      const measured = ctx.measureText('H');
      const ascent = measured.fontBoundingBoxAscent;
      const descent = measured.fontBoundingBoxDescent;
      if (Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0) {
        box = { ascent: ascent / PROBE_PX, descent: descent / PROBE_PX };
      }
    } finally {
      ctx.font = saved;
    }
    probedBoxes.set(route, box);
    return box ?? undefined;
  };
}

type Profile = ReturnType<typeof findReferenceFontMetrics>[number];

const isSystemProfile = (profile: Profile) =>
  profile.source === 'macos-system' || profile.source === 'macos-supplemental';

function projectProfile(profile: Profile): ShapeRunLineRatios | undefined {
  if (profile.source === 'published-open-font') return undefined;
  if (profile.farEastCodePage == null || !profile.win) return undefined;
  return excelDrawingMlLineRatios({
    faceSource: isSystemProfile(profile) ? 'system' : 'office-bundle',
    unitsPerEm: profile.unitsPerEm,
    hhea: profile.hhea,
    win: profile.win,
    typoMetrics: profile.typoMetrics,
    farEastCodePage: profile.farEastCodePage,
  }) ?? undefined;
}

/** One ratio pair when every profile projects to it, else undefined. */
function agreed(profiles: readonly Profile[]): ShapeRunLineRatios | undefined {
  let ratios: ShapeRunLineRatios | undefined;
  for (const profile of profiles) {
    const projected = projectProfile(profile);
    if (!projected || (ratios && (projected.ascentRatio !== ratios.ascentRatio
      || projected.descentRatio !== ratios.descentRatio))) return undefined;
    ratios = projected;
  }
  return ratios;
}

/** Does the observed font box equal one of the boxes a rasterizer can report
 * for this profile (hhea, usWin, or typo when USE_TYPO_METRICS is set)? */
function matchesProfile(box: ShapeFontBox, profile: Profile): boolean {
  const upm = profile.unitsPerEm;
  const pairs: Array<readonly [number, number]> = [[profile.hhea[0], -profile.hhea[1]]];
  if (profile.win) pairs.push(profile.win);
  if (profile.typoMetrics) pairs.push([profile.typoMetrics[0], -profile.typoMetrics[1]]);
  return pairs.some(([ascent, descent]) =>
    Math.abs(box.ascent - ascent / upm) <= PROBE_TOLERANCE_EM
    && Math.abs(box.descent - descent / upm) <= PROBE_TOLERANCE_EM);
}

/**
 * Excel's natural line box for one run (see `excelDrawingMlLineRatios`).
 *
 * The static catalog is reference geometry, not proof of the bytes behind
 * local(). Admit it only after the exact style has loaded. When every catalog
 * profile of the tuple projects to the same box, that box is used.
 *
 * Otherwise the loaded copy's provenance decides, and it must be proven. Excel
 * for Mac uses a macOS system copy of a family over the copy bundled in Office
 * (#1604: its Times New Roman had the system hhea lineGap 87, 1.150 em, where
 * the Office copy gives 1.107 em). A local() route does not say which copy it
 * loaded, so `probe` compares the loaded face's font box with each profile.
 * The run resolves only when the matching profiles all project to one box,
 * and, if the family has a system profile, only when they are all system
 * profiles: an Office copy on the viewer's machine does not prove the system
 * copy that Excel would prefer is absent. Times New Roman's two copies share
 * their ascent and descent (they differ only in lineGap, which Canvas does not
 * expose), so it stays unresolved. Anything unresolved returns undefined and
 * the caller keeps the ordinary Canvas line box: no probe, no match, an
 * ambiguous match, a published open font, missing OS/2 data, or a macOS Far
 * East face (not measured).
 */
export function shapeRunLineRatios(
  run: TextRun,
  route: OfficeFontFallbackRoute | undefined,
  probe?: ShapeFontBoxProbe,
): ShapeRunLineRatios | undefined {
  const family = soleFace(run);
  if (!family || !route || route.source !== 'local' || route.metric.synthesized
    || !route.resourceIdentity.startsWith('office-local:')) return undefined;
  if (route.requestedFamily.toLocaleLowerCase('en-US') !== family.toLocaleLowerCase('en-US')
    || route.weight !== (run.bold ? 700 : 400)
    || route.style !== (run.italic ? 'italic' : 'normal')) return undefined;
  const profiles = findReferenceFontMetrics(family, { weight: route.weight, style: route.style });
  if (profiles.length === 0) return undefined;
  const unanimous = agreed(profiles);
  if (unanimous || profiles.length === 1) return unanimous;
  const box = probe?.(route);
  if (!box) return undefined;
  const matching = profiles.filter((profile) => matchesProfile(box, profile));
  if (matching.length === 0) return undefined;
  if (profiles.some(isSystemProfile) && !matching.every(isSystemProfile)) return undefined;
  return agreed(matching);
}

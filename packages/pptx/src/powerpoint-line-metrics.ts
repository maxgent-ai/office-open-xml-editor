import { findReferenceFontMetrics } from '@silurus/ooxml-core';

/**
 * PowerPoint's split of a text line at its baseline (observed behaviour; the
 * DrawingML text model in ECMA-376 §21.1.2 does not say where the baseline
 * sits inside a line).
 *
 * PowerPoint for Mac 16.113.2 PDF controls (#1610). Two decks with 432
 * boxes, every face verified from the embedded PDF fonts:
 * - The line box is 1.2 × the largest run size for every face.
 * - A run's ascent share of that box is usWinAscent / (usWinAscent +
 *   usWinDescent). A face that sets fsSelection USE_TYPO_METRICS uses
 *   (sTypoAscender + sTypoLineGap) / (sTypoAscender − sTypoDescender +
 *   sTypoLineGap) instead.
 *   - Faces that separate the tables: Yu Gothic (usWin 0.765, hhea 0.798),
 *     Baskerville Old Face, Stencil, and Gabriola (typo 0.814, usWin 0.647).
 *   - Unlike Excel, there is no Far East 1.3× leading and no dependence on
 *     the face's hhea table.
 * - Font resolution follows the face PowerPoint really used, not the name:
 *   - Office-bundled faces and the macOS Supplemental faces (Times New
 *     Roman, Georgia, Courier New) are used as named. A face present in both
 *     resolves to the Supplemental copy (Times New Roman's hhea lineGap 87 is
 *     the one embedded).
 *   - PowerPoint substitutes faces that only macOS /System/Library/Fonts
 *     provides:
 *     - Palatino → Palatino Linotype;
 *     - Helvetica and Helvetica Neue → Arial (embedded under the Helvetica
 *       name with Arial's glyphs and usWin metrics);
 *     - Avenir, Menlo and Hiragino Sans → the deck's theme font or a Far
 *       East fallback.
 *   - Only the first two mappings are pinned. Every other system-only face
 *     returns undefined, and the caller keeps its ordinary line model.
 */
const OBSERVED_SUBSTITUTES: Readonly<Record<string, string>> = {
  palatino: 'Palatino Linotype',
  helvetica: 'Arial',
  'helvetica neue': 'Arial',
};

/**
 * Resolved shares keyed by face/weight/style. Font names come from document
 * content, so the cache is a bounded LRU: at most SHARE_CACHE_LIMIT entries,
 * the least recently used evicted first. A miss only re-reads the static
 * reference table, so the cap trades a little lookup work for bounded memory.
 */
export const SHARE_CACHE_LIMIT = 256;
const shareCache = new Map<string, number | null>();

/** @internal Test hook: current number of cached face keys. */
export function powerPointShareCacheSize(): number {
  return shareCache.size;
}

export function powerPointAscentShare(
  family: string,
  bold: boolean,
  italic: boolean,
): number | undefined {
  const key = `${family.trim().toLocaleLowerCase('en-US')}|${bold ? 700 : 400}|${italic ? 'i' : 'n'}`;
  const cached = shareCache.get(key);
  if (cached !== undefined) {
    shareCache.delete(key);
    shareCache.set(key, cached);
    return cached ?? undefined;
  }
  const share = resolveShare(family, bold, italic);
  shareCache.set(key, share ?? null);
  if (shareCache.size > SHARE_CACHE_LIMIT) {
    const oldest = shareCache.keys().next().value;
    if (oldest !== undefined) shareCache.delete(oldest);
  }
  return share;
}

function resolveShare(family: string, bold: boolean, italic: boolean): number | undefined {
  const trimmed = family.trim();
  if (!trimmed) return undefined;
  const name = OBSERVED_SUBSTITUTES[trimmed.toLocaleLowerCase('en-US')] ?? trimmed;
  const profiles = findReferenceFontMetrics(name, {
    weight: bold ? 700 : 400, style: italic ? 'italic' : 'normal',
  });
  const supplemental = profiles.filter((p) => p.source === 'macos-supplemental');
  const chosen = supplemental.length > 0
    ? supplemental : profiles.filter((p) => p.source === 'office-mac');
  if (chosen.length === 0) return undefined;
  let share: number | undefined;
  for (const profile of chosen) {
    let next: number | undefined;
    if (profile.typoMetrics) {
      const [ascender, descender, lineGap] = profile.typoMetrics;
      const above = ascender + Math.max(0, lineGap);
      next = above / (above - descender);
    } else if (profile.win) {
      const [ascent, descent] = profile.win;
      next = ascent / (ascent + descent);
    }
    if (next === undefined || !Number.isFinite(next) || next <= 0 || next >= 1) return undefined;
    if (share !== undefined && Math.abs(share - next) > 1e-12) return undefined;
    share = next;
  }
  return share;
}

/** One run's contribution to a line: its authored size and ascent share. */
export interface PowerPointLineRun {
  sizePx: number;
  share: number;
}

/**
 * The natural line box of one line. Each run claims 1.2 × its own size,
 * split by its share. The line unions the runs' ascent and descent parts and
 * rescales them into 1.2 × the largest size.
 *
 * Measured on mixed-face lines (Arial+Meiryo in both orders, Arial+MS Gothic,
 * Calibri+Yu Gothic, Arial+Gabriola) and mixed sizes (40+100 pt): all exact.
 * Taking the largest descent instead is off by up to 5 px (1/100 in), and the
 * largest ascent by 13 px.
 */
export function powerPointNaturalLine(runs: readonly PowerPointLineRun[]): { ascent: number; descent: number } {
  let maxSize = 0;
  let ascent = 0;
  let descent = 0;
  for (const run of runs) {
    maxSize = Math.max(maxSize, run.sizePx);
    ascent = Math.max(ascent, 1.2 * run.sizePx * run.share);
    descent = Math.max(descent, 1.2 * run.sizePx * (1 - run.share));
  }
  const height = 1.2 * maxSize;
  if (!(ascent + descent > 0)) return { ascent: height * 0.8, descent: height * 0.2 };
  const a = height * ascent / (ascent + descent);
  return { ascent: a, descent: height - a };
}

/**
 * PowerPoint rounds `spcPts` to whole points before laying out the line
 * (#1610 controls: 40.5 → 41, 45.25 → 45, 48.33 → 48, 49.25 → 49 and
 * 50.75 → 51 pt, each over eight lines). Font sizes stay exact: 10.5, 11.5,
 * 13.33, 40.5 and 55.5 pt keep a 1.2 × size line.
 */
export function powerPointExactLinePoints(points: number): number {
  return Math.floor(points + 0.5);
}

/*
 * PDF exports quantize: each baseline lands on a whole device unit counted
 * from the anchored text top (1/100 in in the #1610 controls; 1/150 in in a
 * corpus export whose 12 pt lines alternate 13.92/14.88 pt). The unit belongs
 * to the export device, not to the slide, so layout keeps continuous
 * positions; the controls match to within half a unit.
 */

/**
 * Office-observed automatic single-line allocation from one static OpenType
 * face. Historical Word for Mac synthetic controls isolated OS/2 code-page
 * bits 17–20 and found 1.3× hhea for that class. Windows Word 16.0.20430 instead
 * matches 1.3× usWin: Yu Gothic/Mincho have identical vertical tables across the tested
 * Mac/Windows faces but different hhea/usWin boxes; MS Mincho, MS Gothic and
 * Meiryo are equal-box counterexamples. Controls cover 8/12/20pt grid/snap-off
 * and 10/30pt atLeast lines. Library policy selects that Windows result from
 * each face's own data, independent of the viewer OS. Other faces retain
 * signed hhea lineGap above the baseline. ECMA-376 §17.3.1.33 defines the
 * spacing multiplier, not this table choice. Missing usWin data declines the
 * Far East projection: hhea or Canvas ink cannot establish the Windows box.
 */
export const OFFICE_FAR_EAST_SINGLE_LINE_FACTOR = 1.3;

function farEastWindowsLineBox(winAscent?: number, winDescent?: number) {
  if (winAscent == null || winDescent == null
    || !Number.isFinite(winAscent) || !Number.isFinite(winDescent)
    || winAscent < 0 || winDescent < 0 || winAscent + winDescent <= 0) return null;
  const half = ((OFFICE_FAR_EAST_SINGLE_LINE_FACTOR - 1) / 2) * (winAscent + winDescent);
  return { ascent: winAscent + half, descent: winDescent + half };
}

export function officeOpenTypeAutoLineRatios(metrics: Readonly<{
  unitsPerEm: number;
  hheaAscent: number;
  hheaDescent: number;
  hheaLineGap: number;
  winAscent?: number;
  winDescent?: number;
  farEastCodePage: boolean;
}>): Readonly<{
  lineHeightRatio: number;
  designAscentRatio: number;
  designDescentRatio: number;
}> | null {
  const { unitsPerEm, hheaAscent, hheaDescent, hheaLineGap, farEastCodePage } = metrics;
  if (!(Number.isFinite(unitsPerEm) && unitsPerEm > 0
    && Number.isFinite(hheaAscent) && hheaAscent >= 0
    && Number.isFinite(hheaDescent) && hheaDescent <= 0
    && Number.isFinite(hheaLineGap))) return null;
  const glyphBox = hheaAscent - hheaDescent;
  if (!(glyphBox > 0)) return null;
  const box = farEastCodePage
    ? farEastWindowsLineBox(metrics.winAscent, metrics.winDescent)
    : { ascent: hheaAscent + hheaLineGap, descent: -hheaDescent };
  if (!box) return null;
  const { ascent, descent } = box;
  if (!(ascent >= 0 && descent >= 0 && ascent + descent > 0)) return null;
  return Object.freeze({
    lineHeightRatio: (ascent + descent) / unitsPerEm,
    designAscentRatio: ascent / unitsPerEm,
    designDescentRatio: descent / unitsPerEm,
  });
}

/**
 * Excel's natural single line for DrawingML shape text, from one static
 * OpenType face (observed behaviour; ECMA-376 does not select font tables).
 *
 * Excel for Mac 16.113.2 PDF controls (#1604): 177 shape bodies over Calibri,
 * Arial, Times New Roman, Yu Gothic, Meiryo, Meiryo UI, MS Gothic, Baskerville
 * Old Face, Stencil, Gabriola, Palatino and Helvetica at 8–72 pt, each face
 * verified as embedded. The metric table depends on where the face comes from:
 *
 * - `office-bundle` (fonts shipped inside Office.app) follow the Windows
 *   metrics. The glyph box is the OS/2 usWin extent. Baskerville Old Face and
 *   Stencil measured 1.141 and 1.185 em (their hhea box is 1.0 em).
 *   - A face in the Far East code-page class (OS/2 ulCodePageRange1 bits
 *     17–20) gets 1.3 × the usWin box. Half of the added leading goes above
 *     the ascent and half below the descent: Yu Gothic 1.673 em (its hhea box
 *     would give 1.433 em), Meiryo 1.95 em, MS Gothic 1.3 em. The class
 *     follows the face, not `lang` or the text: Meiryo with lang en-US and
 *     Latin-only text keeps 1.95 em, and Arial with lang ja-JP keeps 1.15 em.
 *   - A face with fsSelection USE_TYPO_METRICS uses sTypoAscender + sTypoLineGap
 *     above and -sTypoDescender below (Gabriola 1.700 em; usWin would give
 *     1.841 em).
 *   - Other faces add the Windows TEXTMETRIC external leading
 *     max(0, hhea.lineGap − (usWin box − hhea box)) above the ascent (Arial
 *     1.150 em, Calibri 1.221 em).
 * - `system` (fonts macOS itself provides) use hhea: ascender + lineGap above,
 *   -descender below (Palatino 1.100 em where usWin gives 1.656 em, Helvetica
 *   1.000 em). When both copies of a family exist, Excel used the system face:
 *   its Times New Roman carries the macOS hhea lineGap 87, giving 1.150 em.
 *   No system Far East face was measured, so that class returns null.
 *
 * Callers must gate this to Excel shape text and a known font-metric source.
 */
export function excelDrawingMlLineRatios(metrics: Readonly<{
  faceSource: 'office-bundle' | 'system';
  unitsPerEm: number;
  hhea: readonly [ascender: number, descender: number, lineGap: number];
  win: readonly [ascent: number, descent: number];
  /** OS/2 typo metrics when USE_TYPO_METRICS is set; otherwise absent. */
  typoMetrics?: readonly [ascender: number, descender: number, lineGap: number];
  farEastCodePage: boolean;
}>): Readonly<{ ascentRatio: number; descentRatio: number }> | null {
  const { faceSource, unitsPerEm, hhea, win, typoMetrics, farEastCodePage } = metrics;
  const values = [unitsPerEm, ...hhea, ...win, ...(typoMetrics ?? [])];
  if (!(unitsPerEm > 0) || !values.every(Number.isFinite)) return null;
  const [hheaAscent, hheaDescent, hheaLineGap] = hhea;
  const [winAscent, winDescent] = win;
  let ascent: number;
  let descent: number;
  if (faceSource === 'system') {
    if (farEastCodePage) return null;
    ascent = hheaAscent + Math.max(0, hheaLineGap);
    descent = -hheaDescent;
  } else if (farEastCodePage) {
    const box = farEastWindowsLineBox(winAscent, winDescent);
    if (!box) return null;
    ({ ascent, descent } = box);
  } else if (typoMetrics) {
    ascent = typoMetrics[0] + Math.max(0, typoMetrics[2]);
    descent = -typoMetrics[1];
  } else {
    const box = winAscent + winDescent;
    const externalLeading = Math.max(0, hheaLineGap - (box - (hheaAscent - hheaDescent)));
    ascent = winAscent + externalLeading;
    descent = winDescent;
  }
  if (!(ascent >= 0 && descent >= 0 && ascent + descent > 0)) return null;
  return Object.freeze({ ascentRatio: ascent / unitsPerEm, descentRatio: descent / unitsPerEm });
}

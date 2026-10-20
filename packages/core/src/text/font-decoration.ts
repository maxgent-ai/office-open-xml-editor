import { findReferenceFontMetrics } from '../fonts/reference-font-metrics.js';

/** A face's own decoration geometry, as fractions of the font size measured
 *  from the alphabetic baseline (positive = below). */
export interface FontDecorationGeometry {
  /** Top edge of the single underline. */
  readonly underlineTopEm: number;
  readonly underlineThicknessEm: number;
  /** Top edge of the single strikethrough (negative: above the baseline). */
  readonly strikeTopEm: number;
  readonly strikeThicknessEm: number;
}

/**
 * The catalogued face's post underline and OS/2 strikeout metrics. Windows
 * Word and PowerPoint 16.0.20430 draw single underline and strikethrough from
 * them (#1896 controls, eight faces at 48pt, within the PDF's 0.12pt grid).
 * Null when the face is not catalogued or records no usable values; callers
 * then keep their existing geometry.
 */
export function referenceFontDecoration(
  family: string,
  weight: number,
  style: 'normal' | 'italic',
): FontDecorationGeometry | null {
  for (const profile of findReferenceFontMetrics(family, { weight, style })) {
    const decoration = profile.decoration;
    if (!decoration || !(profile.unitsPerEm > 0)) continue;
    const [underlinePosition, underlineThickness, strikeoutPosition, strikeoutSize] = decoration;
    if (!(underlineThickness > 0) || !(strikeoutSize > 0) || !(strikeoutPosition > 0)) continue;
    const em = profile.unitsPerEm;
    return Object.freeze({
      underlineTopEm: -underlinePosition / em,
      underlineThicknessEm: underlineThickness / em,
      strikeTopEm: -strikeoutPosition / em,
      strikeThicknessEm: strikeoutSize / em,
    });
  }
  return null;
}

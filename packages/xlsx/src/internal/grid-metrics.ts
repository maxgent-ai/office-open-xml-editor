import { PT_TO_PX } from '@silurus/ooxml-core';

/** ECMA-376 §18.3.1.13 fallback when the default font's MDW is unavailable. */
export const MDW_FALLBACK = 8;

/** Convert an OOXML column width (maximum-digit-width units) to CSS pixels. */
export function colWidthToPx(width: number, mdw: number = MDW_FALLBACK): number {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** ECMA-376 §18.3.1.13: encode a UI width in characters (padding excluded) as a
 *  stored width, using the normative 5 px padding at the current MDW. Then
 *  decode it with the stored-width formula. Both steps truncate. */
export function uiColWidthToPx(chars: number, mdw: number = MDW_FALLBACK): number {
  const stored = Math.trunc(((chars * mdw + 5) / mdw) * 256) / 256;
  return colWidthToPx(stored, mdw);
}

/** ECMA-376 §18.3.1.81: baseColWidth is a digit count, excluding padding.
 *  At 96 DPI the normative padding is 5 px (two 2 px margins and gridline).
 *
 *  Observed Windows Excel behaviour (not in the specification): the derived
 *  width is rounded UP to the next multiple of 8 px. Excel 16.0 build 20430
 *  with Calibri 11 (MDW 7) shows base 8/10/12 as 64/80/96 px (not 61/75/89),
 *  and its PDF export keeps the same ratios. At 200 % display scaling base 13
 *  reads back as 78 pt (104 CSS px) rather than the normative 76.875 pt, which
 *  is the same rule in CSS px (13 × 7.5 + 5 = 102.5 -> 104). The rounding is
 *  applied here in CSS px at 100 %, before zoom and device-pixel scaling; MDW
 *  is the measured integer digit width. An explicit defaultColWidth is a
 *  stored width and is not rounded.
 *
 *  Layout is independent of the viewer OS; Windows Excel is the reference. */
export function baseColWidthToPx(baseWidth: number, mdw: number): number {
  return Math.ceil(Math.round(baseWidth * mdw + 5) / 8) * 8;
}

/** In-memory inverse used by the viewer's resize interaction. */
export function pxToColWidth(px: number, mdw: number = MDW_FALLBACK): number {
  return px / mdw;
}

/** Convert OOXML row height points to CSS pixels at 96 DPI. */
export function rowHeightToPx(heightPt: number): number {
  return Math.round(heightPt * PT_TO_PX);
}

/** In-memory inverse used by the viewer's resize interaction. */
export function pxToRowHeight(px: number): number {
  return px / PT_TO_PX;
}

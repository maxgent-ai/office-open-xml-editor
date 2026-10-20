/**
 * Em-box centre above the alphabetic baseline, as a fraction of the font size,
 * used only when the Canvas engine does not report
 * `TextMetrics.fontBoundingBoxAscent/Descent`.
 *
 * Library fallback, not a spec value: ECMA-376 (§17.6.20 / §21.1.2.1.1
 * vertical text) does not define where a sideways glyph sits on the column, and
 * the font's own metrics are preferred whenever the engine exposes them. The
 * ratio matches the em-box centre measured for the sideways-digit offset in
 * issue #771 (about 0.38 em with headless Skia's default CJK/Latin faces); it
 * is not a per-font constant and is never consulted on engines that report the
 * font bounding box.
 */
export const EM_BOX_CENTER_FALLBACK_EM = 0.38;

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Cross-axis offset from the alphabetic baseline to the font's em-box centre:
 * `(fontBoundingBoxAscent − fontBoundingBoxDescent) / 2`.
 *
 * In vertical text the upright cells (drawn with a `middle` baseline) centre
 * their em box on the column line, while a sideways Latin/digit glyph is drawn
 * on its alphabetic baseline. Shifting the sideways draw by this font metric
 * lets mixed columns share one centreline. `size` is the font size in the
 * caller's unit; the result is in that unit.
 */
export function emBoxCenterAboveBaseline(ctx: Ctx2D, sample: string, size: number): number {
  const prevBaseline = ctx.textBaseline;
  ctx.textBaseline = 'alphabetic';
  const m = ctx.measureText(sample);
  ctx.textBaseline = prevBaseline;
  const asc = m.fontBoundingBoxAscent;
  const desc = m.fontBoundingBoxDescent;
  if (typeof asc === 'number' && typeof desc === 'number' && (asc !== 0 || desc !== 0)) {
    return (asc - desc) / 2;
  }
  return EM_BOX_CENTER_FALLBACK_EM * size;
}

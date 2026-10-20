// Canvas metric shared by the DOCX tbRl and PPTX eaVert rotate-layout
// vertical painters. Both draw a vertical column by rotating the page +90° and
// running the horizontal layout. The cross-axis em-box centre used for the
// sideways cells lives in `em-box-center.ts`.

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Along-column offset, in px, from a glyph's own cell centre to its INK centre
 * when the glyph is drawn UPRIGHT — i.e. `(actualBoundingBoxAscent −
 * actualBoundingBoxDescent)/2` measured with a `middle` textBaseline.
 *
 * For an upright-drawn glyph the page transform maps the glyph's VERTICAL extent
 * onto the along-column axis, so a glyph whose ink is not vertically centred in
 * its em box (most visibly a substituted vertical bracket form ︵ ︶ ﹁ ﹂,
 * whose ink hugs one end of the cell) lands off the cell centre by this amount.
 * The renderer shifts the draw by `+this` so the ink re-centres — a per-GLYPH
 * measured metric (`actualBoundingBox*` is the tight ink box), NOT a constant.
 * For an ordinary ideograph/kana this is ≈0, so upright CJK cells are unaffected.
 *
 * Returns 0 when the Canvas does not report `actualBoundingBox*` (older engines):
 * the glyph then draws at the cell centre exactly as before this metric existed.
 */
export function inkCenterAboveMiddlePx(ctx: Ctx2D, drawStr: string): number {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const m = ctx.measureText(drawStr);
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  const asc = m.actualBoundingBoxAscent;
  const desc = m.actualBoundingBoxDescent;
  if (typeof asc === 'number' && typeof desc === 'number') {
    return (asc - desc) / 2;
  }
  return 0;
}

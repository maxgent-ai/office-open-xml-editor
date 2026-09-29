import { crispOffset } from '../canvas/crisp.js';
import { pptxUnderlineDashArray } from '../draw/dash.js';

/**
 * Draw a text underline at the given baseline. The style vocabulary is
 * DrawingML ST_TextUnderlineType (ECMA-376 §20.1.10.82): `sng` / `dbl` /
 * `heavy` / `dotted` / `dottedHeavy` / `dash` / `dashHeavy` / `dashLong` /
 * `dashLongHeavy` / `dotDash` / `dotDashHeavy` / `dotDotDash` /
 * `dotDotDashHeavy` / `wavy` / `wavyHeavy` / `wavyDbl`. We map each value to a
 * Canvas dash pattern + line-weight pair; "wavy*" is approximated by a sine
 * curve traced as a polyline so the glyph stays legibly distinct from "dotted".
 *
 * This is the single source of truth for run-underline geometry across the
 * pptx / docx renderers (hoisted from the pptx renderer verbatim). The word-only
 * WordprocessingML ST_Underline vocabulary (§17.18.99: single/wave/wavyDouble/…)
 * is normalized to this DrawingML vocabulary by the docx renderer before it
 * calls in, so both formats share one dispatch. `undefined` style renders the
 * default single (`sng`) rule.
 *
 * @param dpr device-pixel ratio; horizontal strokes snap onto the nearest crisp
 *   device row via {@link crispOffset} so an odd-device-width rule stays crisp
 *   on a DPR=1 display. The wavy variant is not axis-aligned per pixel and
 *   deliberately omits the offset.
 */
export function drawUnderline(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number,
  baseline: number,
  width: number,
  sizePx: number,
  color: string | CanvasGradient | CanvasPattern,
  style: string | undefined,
  dpr = 1,
  authoredLineWidthPx?: number,
): void {
  const baseLineW = authoredLineWidthPx != null
    ? Math.max(0, authoredLineWidthPx) : Math.max(1, sizePx * 0.05);
  // ST_TextUnderlineType (§20.1.10.82) has two distinct "heavy" spellings:
  // the bare `heavy` value itself, and the `*Heavy` suffix on the dotted/dash/
  // wavy families (`dottedHeavy`, `dashHeavy`, …). Both denote the same ~1.8×
  // weight class, so both must be recognized here.
  const heavy = style === 'heavy' || (style?.endsWith('Heavy') ?? false);
  const lineW = heavy ? baseLineW * 1.8 : baseLineW;
  const y = baseline + Math.max(2, lineW);
  // Crispness nudge (see crispOffset): a horizontal underline whose device-pixel
  // width is odd straddles two device rows on a DPR=1 display (blurry). Snapping
  // the line's y onto the nearest crisp device position centers an odd-width
  // stroke on one device row → crisp. Applied to the straight / dbl / dashed
  // branches only; the wavy variant is not axis-aligned per pixel, so it
  // deliberately omits the offset.
  const crispY = crispOffset(y, lineW, dpr);
  ctx.strokeStyle = color;
  ctx.lineWidth = lineW;
  ctx.setLineDash([]);

  // Dash patterns scaled by lineW so they stay proportional at any font size,
  // computed via core's shared pptxUnderlineDashArray (§20.1.10.82
  // ST_TextUnderlineType — the run-underline enum, distinct from the
  // §20.1.10.49 preset line dash though it reuses a few shape names) at the
  // call site below.

  if (style && style.startsWith('wavy')) {
    // Sine wave with amplitude ≈ lineW and wavelength ≈ 6×lineW.
    const amp = lineW;
    const wavelength = lineW * 6;
    ctx.beginPath();
    ctx.moveTo(x, y);
    const step = Math.max(1, lineW * 0.5);
    for (let dx = 0; dx <= width; dx += step) {
      const yy = y + Math.sin((dx / wavelength) * Math.PI * 2) * amp;
      ctx.lineTo(x + dx, yy);
    }
    ctx.stroke();
    if (style === 'wavyDbl') {
      // Second wave below, offset by 2.5×amp.
      ctx.beginPath();
      ctx.moveTo(x, y + amp * 2.5);
      for (let dx = 0; dx <= width; dx += step) {
        const yy = y + amp * 2.5 + Math.sin((dx / wavelength) * Math.PI * 2) * amp;
        ctx.lineTo(x + dx, yy);
      }
      ctx.stroke();
    }
    return;
  }

  if (style === 'dbl') {
    const offset = lineW * 1.4;
    // Two parallel rules straddling y; snap each onto its own crisp device row.
    const y1 = y - offset / 2;
    const y2 = y + offset / 2;
    ctx.beginPath();
    ctx.moveTo(x, y1 + crispOffset(y1, lineW, dpr));
    ctx.lineTo(x + width, y1 + crispOffset(y1, lineW, dpr));
    ctx.moveTo(x, y2 + crispOffset(y2, lineW, dpr));
    ctx.lineTo(x + width, y2 + crispOffset(y2, lineW, dpr));
    ctx.stroke();
    return;
  }

  ctx.setLineDash(pptxUnderlineDashArray(style ?? 'sng', lineW));
  ctx.beginPath();
  ctx.moveTo(x, y + crispY);
  ctx.lineTo(x + width, y + crispY);
  ctx.stroke();
  ctx.setLineDash([]);
}

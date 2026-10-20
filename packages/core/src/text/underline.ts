import { crispOffset } from '../canvas/crisp.js';
import { pptxUnderlineDashArray } from '../draw/dash.js';

// ── Decoration geometry ─────────────────────────────────────────────────────
// ECMA-376 names the underline / strike style (§20.1.10.82 ST_TextUnderlineType,
// §20.1.10.79 ST_TextStrikeType, §17.18.99 ST_Underline) but not its position or
// thickness. A catalogued face supplies its own rule (OpenType post
// underlinePosition / underlineThickness, OS/2 yStrikeoutPosition /
// yStrikeoutSize; see `font-decoration.ts`), which Windows Word and PowerPoint
// 16.0.20430 draw (#1896 controls). The em fractions below are library policy
// for faces the reference table does not catalogue (the Canvas API exposes no
// font tables).

/** Decoration stroke width as a fraction of the font size (minimum 1 px). */
export const DECORATION_LINE_WIDTH_EM = 0.05;
/** Strike-through rise above the alphabetic baseline, as a fraction of the font size. */
export const STRIKE_RISE_EM = 0.32;
/** `heavy` / `*Heavy` underline weight relative to the single rule. */
const HEAVY_WEIGHT_FACTOR = 1.8;
/** `dbl` underline: distance between the two rules, in line widths. */
const DOUBLE_RULE_GAP_LINE_WIDTHS = 1.4;
/** `wavy*` underline: sine wavelength, in line widths (amplitude is one line width). */
const WAVE_LENGTH_LINE_WIDTHS = 6;
/** `wavyDbl` underline: offset of the second wave, in amplitudes. */
const DOUBLE_WAVE_OFFSET_AMPLITUDES = 2.5;

/**
 * Draw a text underline at the given baseline. The style vocabulary is
 * DrawingML ST_TextUnderlineType (ECMA-376 §20.1.10.82): `sng` / `dbl` /
 * `heavy` / `dotted` / `dottedHeavy` / `dash` / `dashHeavy` / `dashLong` /
 * `dashLongHeavy` / `dotDash` / `dotDashHeavy` / `dotDotDash` /
 * `dotDotDashHeavy` / `wavy` / `wavyHeavy` / `wavyDbl`. We map each value to a
 * Canvas dash pattern + line-weight pair; "wavy*" is approximated by a sine
 * curve traced as a polyline so the glyph stays legibly distinct from "dotted".
 *
 * This is the shared run-underline dispatch for DrawingML text (hoisted from
 * the pptx renderer verbatim). `undefined` style renders the default single
 * (`sng`) rule.
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
  fontRule?: Readonly<{ topPx: number; thicknessPx: number }>,
): void {
  const baseLineW = authoredLineWidthPx != null
    ? Math.max(0, authoredLineWidthPx)
    : fontRule ? fontRule.thicknessPx : Math.max(1, sizePx * DECORATION_LINE_WIDTH_EM);
  // ST_TextUnderlineType (§20.1.10.82) has two distinct "heavy" spellings:
  // the bare `heavy` value itself, and the `*Heavy` suffix on the dotted/dash/
  // wavy families (`dottedHeavy`, `dashHeavy`, …). Both denote the same ~1.8×
  // weight class ({@link HEAVY_WEIGHT_FACTOR}), so both must be recognized here.
  const heavy = style === 'heavy' || (style?.endsWith('Heavy') ?? false);
  // With the face's own rule (post underlinePosition / underlineThickness),
  // Windows PowerPoint 16.0.20430 draws the single rule there and a heavy
  // rule 1.5x as thick with its bottom at the single rule's centre (#1896
  // controls: Arial, Times New Roman, Calibri at 48pt). Other styles centre on
  // the single rule. Without it, the library geometry below applies.
  const lineW = heavy ? baseLineW * (fontRule ? 1.5 : HEAVY_WEIGHT_FACTOR) : baseLineW;
  const singleCentre = fontRule ? baseline + fontRule.topPx + baseLineW / 2 : undefined;
  const y = singleCentre === undefined
    ? baseline + Math.max(2, lineW)
    : heavy ? singleCentre - lineW / 2 : singleCentre;
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
    // Sine wave with amplitude = lineW and wavelength = WAVE_LENGTH_LINE_WIDTHS × lineW.
    const amp = lineW;
    const wavelength = lineW * WAVE_LENGTH_LINE_WIDTHS;
    ctx.beginPath();
    ctx.moveTo(x, y);
    const step = Math.max(1, lineW * 0.5);
    for (let dx = 0; dx <= width; dx += step) {
      const yy = y + Math.sin((dx / wavelength) * Math.PI * 2) * amp;
      ctx.lineTo(x + dx, yy);
    }
    ctx.stroke();
    if (style === 'wavyDbl') {
      // Second wave below, offset by DOUBLE_WAVE_OFFSET_AMPLITUDES × amp.
      ctx.beginPath();
      ctx.moveTo(x, y + amp * DOUBLE_WAVE_OFFSET_AMPLITUDES);
      for (let dx = 0; dx <= width; dx += step) {
        const yy = y + amp * DOUBLE_WAVE_OFFSET_AMPLITUDES + Math.sin((dx / wavelength) * Math.PI * 2) * amp;
        ctx.lineTo(x + dx, yy);
      }
      ctx.stroke();
    }
    return;
  }

  if (style === 'dbl') {
    const offset = lineW * DOUBLE_RULE_GAP_LINE_WIDTHS;
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

// Authored line breaks in chart label text. A category cell such as
// "FY2024\n1Q営業利益", a series name or a number format with a literal line
// feed reaches the renderer as one string; PowerPoint and Excel start a new line
// at every authored break. `CanvasRenderingContext2D.fillText` instead replaces
// each break with a space, so every label path that draws a whole string goes
// through these helpers. Text without a break takes the exact single-line path
// it took before (same `fillText`/`measureText` arguments).

import { elideToWidth } from './text-elide.js';

const BREAK = /\r\n|\r|\n/;
const OUTER_BREAKS = /^(?:\r\n|\r|\n)+|(?:\r\n|\r|\n)+$/g;

/** Line pitch for stacked axis tick-label lines: the pitch the wrapped column
 *  and ChartEx category labels already use (font size + 2 px). */
export function axisLabelLineHeightPx(fontPx: number): number {
  return fontPx + 2;
}

/** The authored lines of `text`. Text without a break is returned unchanged as
 *  one line. Leading/trailing breaks are trimmed (as `wrapMeasuredText` trims
 *  outer whitespace); interior blank lines are kept. */
export function authoredTextLines(text: string): string[] {
  if (!BREAK.test(text)) return [text];
  return text.replace(OUTER_BREAKS, '').split(BREAK);
}

/** Number of authored lines in `text` (1 for text without a break). */
export function authoredLineCount(text: string): number {
  return authoredTextLines(text).length;
}

/** Largest authored line count over `texts` (1 when empty). */
export function maxAuthoredLineCount(texts: Iterable<string>): number {
  let count = 1;
  for (const text of texts) count = Math.max(count, authoredLineCount(text));
  return count;
}

/** Extra block height the additional authored lines of the tallest label add
 *  beyond a single line. A rotated label stacks its lines perpendicular to the
 *  text direction, so only the vertical component (|cos θ|) of that stack
 *  grows a horizontal band. Exactly 0 when no label contains a break. */
export function authoredExtraLinesHeight(
  texts: Iterable<string>,
  lineHeightPx: number,
  rotationRad = 0,
): number {
  const extraLines = maxAuthoredLineCount(texts) - 1;
  if (extraLines <= 0) return 0;
  const run = extraLines * lineHeightPx;
  return rotationRad === 0 ? run : run * Math.abs(Math.cos(rotationRad));
}

/** Width of the widest authored line with the current `ctx.font`. */
export function measureAuthoredTextWidth(ctx: CanvasRenderingContext2D, text: string): number {
  const lines = authoredTextLines(text);
  if (lines.length === 1) return ctx.measureText(lines[0]).width;
  let width = 0;
  for (const line of lines) width = Math.max(width, ctx.measureText(line).width);
  return width;
}

/** Offset of the first line so a block of `lineCount` lines keeps the anchor
 *  semantics of the current `textBaseline`: `top`/`hanging` grow downward,
 *  `middle` is centred on the anchor, and bottom/alphabetic baselines grow
 *  upward so the last line stays on the anchor. */
export function authoredBlockStartOffset(
  ctx: CanvasRenderingContext2D,
  lineCount: number,
  lineHeightPx: number,
): number {
  const span = (lineCount - 1) * lineHeightPx;
  const baseline = ctx.textBaseline;
  if (baseline === 'top' || baseline === 'hanging') return 0;
  if (baseline === 'middle') return -span / 2;
  return -span;
}

/** Draw `text` at `(x, y)` with the caller's align/baseline, one `fillText`
 *  per authored line spaced `lineHeightPx` apart (see
 *  {@link authoredBlockStartOffset}). With `maxWidthPx`, each line is elided
 *  separately. Text without a break is drawn by the single call it used
 *  before: `fillText(text, x, y)` (or the elided text). */
export function fillAuthoredText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  lineHeightPx: number,
  maxWidthPx?: number,
): void {
  fillAuthoredLines(ctx, text, x, y, lineHeightPx, maxWidthPx, false);
}

/** Like {@link fillAuthoredText}, but the first line stays where the
 *  single-line label is drawn and further lines grow downward whatever the
 *  baseline — for labels hanging below a horizontal axis. */
export function fillAuthoredTextDownward(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  lineHeightPx: number,
  maxWidthPx?: number,
): void {
  fillAuthoredLines(ctx, text, x, y, lineHeightPx, maxWidthPx, true);
}

function fillAuthoredLines(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  lineHeightPx: number,
  maxWidthPx: number | undefined,
  downward: boolean,
): void {
  const lines = authoredTextLines(text);
  const fit = (line: string): string => maxWidthPx == null ? line : elideToWidth(ctx, line, maxWidthPx);
  if (lines.length === 1) {
    ctx.fillText(fit(lines[0]), x, y);
    return;
  }
  const startY = downward ? y : y + authoredBlockStartOffset(ctx, lines.length, lineHeightPx);
  lines.forEach((line, index) => {
    const drawn = fit(line);
    if (drawn) ctx.fillText(drawn, x, startY + index * lineHeightPx);
  });
}

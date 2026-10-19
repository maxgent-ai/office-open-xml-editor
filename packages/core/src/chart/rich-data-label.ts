import type { ChartTextRun } from '../types/chart.js';
import { chartTextFontSizePx } from './layout.js';
import { chartFaceWithFallback } from './shared/fonts.js';

export interface RichDataLabelOptions {
  runs: readonly ChartTextRun[];
  ptToPx: number;
  fontFamily: string;
  fallbackBold: boolean;
  fallbackItalic?: boolean;
  /** Normalized baseline shift inherited from the label's default run. */
  fallbackBaseline?: number;
  /** Effective label-default text paint is absent/unresolved. */
  fallbackColorHidden?: boolean;
  fontFamilyForFace?: (face: string) => string;
}

export interface RichDataLabelPaintRun {
  text: string;
  font: string;
  fillStyle: string | null;
  width: number;
  fontSizePx: number;
  baselineShiftPx: number;
}

export interface RichDataLabelBlock {
  lines: RichDataLabelPaintRun[][];
  /** Per-line alignment authored on the owning DrawingML paragraph. */
  lineAligns: Array<string | null>;
  lineHeights: number[];
  lineWidths: number[];
  width: number;
  height: number;
}

const MAX_SCALARS = 4096;
const MAX_LINES = 4;
const MAX_RUNS = MAX_SCALARS;

export function resolveRichDataLabelBlock(
  ctx: CanvasRenderingContext2D,
  options: RichDataLabelOptions,
  fallbackFontSizePx: number,
  fallbackColor: string,
): RichDataLabelBlock | null {
  const lines: RichDataLabelPaintRun[][] = [[]];
  const lineAligns: Array<string | null> = [null];
  let scalarCount = 0;
  const cleanColor = (value: string | null | undefined): string | null => {
    if (!value) return null;
    const hex = value.startsWith('#') ? value.slice(1) : value;
    return /^[0-9A-Fa-f]{6}([0-9A-Fa-f]{2})?$/.test(hex) ? `#${hex}` : null;
  };
  outer: for (let runIndex = 0;
    runIndex < options.runs.length && runIndex < MAX_RUNS;
    runIndex++) {
    const source = options.runs[runIndex];
    if (source.text !== '\n' && lineAligns[lines.length - 1] == null) {
      lineAligns[lines.length - 1] = source.paragraphAlign ?? null;
    }
    const fontSizePx = chartTextFontSizePx(source.fontSizeHpt, options.ptToPx)
      ?? fallbackFontSizePx;
    const authoredFace = source.fontFace?.trim().replaceAll('"', '');
    const safeFace = authoredFace && !authoredFace.startsWith('+') ? authoredFace : null;
    const fontFamily = authoredFace && options.fontFamilyForFace
      ? options.fontFamilyForFace(authoredFace)
      : safeFace ? chartFaceWithFallback(safeFace) : options.fontFamily;
    const italic = source.italic ?? options.fallbackItalic ?? false;
    const font = `${italic ? 'italic ' : ''}${(source.bold ?? options.fallbackBold) ? 'bold ' : ''}${fontSizePx}px ${fontFamily}`;
    const baselineShiftPx = (source.baseline ?? options.fallbackBaseline ?? 0) * fontSizePx;
    const sourcePaintAuthored = source.colorPaintAuthored === true
      || source.color != null || source.colorHidden === true;
    const fillStyle = sourcePaintAuthored
      ? (source.colorHidden === true ? null : cleanColor(source.color))
      : options.fallbackColorHidden === true ? null : fallbackColor;
    let chunk = '';
    const flush = (): void => {
      if (!chunk) return;
      ctx.font = font;
      lines[lines.length - 1].push({
        text: chunk, font, fillStyle,
        width: ctx.measureText(chunk).width,
        fontSizePx,
        baselineShiftPx,
      });
      chunk = '';
    };
    let skipLfAfterCr = false;
    for (const sourceCharacter of source.text) {
      if (skipLfAfterCr && sourceCharacter === '\n') {
        skipLfAfterCr = false;
        continue;
      }
      skipLfAfterCr = false;
      const character = sourceCharacter === '\r' ? '\n' : sourceCharacter;
      if (sourceCharacter === '\r') skipLfAfterCr = true;
      if (scalarCount >= MAX_SCALARS) {
        flush();
        break outer;
      }
      scalarCount++;
      if (character === '\n') {
        flush();
        if (lines.length >= MAX_LINES) break outer;
        lines.push([]);
        // A newline embedded in one DrawingML run stays in the same paragraph,
        // so the paragraph alignment continues on the following visual line.
        // The synthetic paragraph-boundary run has no alignment; the next real
        // run will set the new paragraph's value as before.
        lineAligns.push(source.paragraphAlign ?? null);
      } else chunk += character;
    }
    flush();
  }
  if (!lines.some(line => line.length > 0)) return null;
  const lineHeights = lines.map(line => Math.max(
    fallbackFontSizePx * 1.15,
    ...line.map(run => run.fontSizePx * 1.15 + Math.abs(run.baselineShiftPx)),
  ));
  const lineWidths = lines.map(line => line.reduce((sum, run) => sum + run.width, 0));
  return {
    lines,
    lineAligns,
    lineHeights,
    lineWidths,
    width: Math.max(0, ...lineWidths),
    height: lineHeights.reduce((sum, height) => sum + height, 0),
  };
}

export function paintRichDataLabelBlock(
  ctx: CanvasRenderingContext2D,
  block: RichDataLabelBlock,
  x: number,
  y: number,
  textAlign: CanvasTextAlign = 'center',
  textBaseline: CanvasTextBaseline = 'middle',
  containerWidth = block.width,
): void {
  const top = textBaseline === 'top'
    ? y : textBaseline === 'bottom' ? y - block.height : y - block.height / 2;
  let lineTop = top;
  for (let lineIndex = 0; lineIndex < block.lines.length; lineIndex++) {
    const line = block.lines[lineIndex];
    const lineWidth = block.lineWidths[lineIndex];
    const paragraphAlign = block.lineAligns[lineIndex];
    const effectiveAlign: CanvasTextAlign = paragraphAlign === 'l' ? 'left'
      : paragraphAlign === 'r' ? 'right'
        : paragraphAlign === 'ctr' ? 'center' : textAlign;
    // `just` and `dist` need word/glyph expansion. Retain the authored token
    // while avoiding invented distribution geometry.
    const containerLeft = textAlign === 'left' ? x
      : textAlign === 'right' ? x - containerWidth : x - containerWidth / 2;
    const containerRight = containerLeft + containerWidth;
    let cursorX = effectiveAlign === 'left' ? containerLeft
      : effectiveAlign === 'right' ? containerRight - lineWidth
        : containerLeft + (containerWidth - lineWidth) / 2;
    const paintY = textBaseline === 'top'
      ? lineTop : textBaseline === 'bottom'
        ? lineTop + block.lineHeights[lineIndex]
        : lineTop + block.lineHeights[lineIndex] / 2;
    for (const run of line) {
      ctx.font = run.font;
      if (run.fillStyle == null) {
        cursorX += run.width;
        continue;
      }
      ctx.fillStyle = run.fillStyle;
      ctx.textAlign = 'left';
      ctx.textBaseline = textBaseline;
      ctx.fillText(run.text, cursorX, paintY - run.baselineShiftPx);
      cursorX += run.width;
    }
    lineTop += block.lineHeights[lineIndex];
  }
}

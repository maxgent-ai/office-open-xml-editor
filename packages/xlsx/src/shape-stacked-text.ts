import {
  layoutStackedText, paintStackedRotated, paintStackedVerticalForm, STACKED_CELL_FACTOR, stackedClusters,
  stackedFaceBox, stackedUprightBaseline, stackedVerticalGlyphText,
  type DrawingMlInputRun, type DrawingMlTextRect, type StackedGlyph, type StackedParagraph,
} from '@silurus/ooxml-core/internal/drawingml-text';
import { excelShapeSpacing } from './shape-office-line.js';
import type { ShapeText, ShapeTextRun } from './types.js';

type TextRun = Extract<ShapeTextRun, { type: 'text' }>;
export interface StackedEquation {
  width: number;
  height: number;
  draw(x: number, y: number): void;
}

type Style = { equation?: StackedEquation; font: string; px: number; face?: string; color: string; spacing: number; bold: boolean; italic: boolean };
type Glyph = StackedGlyph<Style> & { cell: number; ascent: number; descent: number; nominal?: true };

/**
 * ECMA-376 §20.1.10.83: letters stacked top-to-bottom; Rtl reverses columns.
 * Library policy: reuse core's face-cell layout, not Excel's glyph-ink pitch.
 * Issue #1668 Excel controls measured unequal Arial 24 pt baseline gaps
 * 26.16, 25.92, 22.08, 20.88, 26.16 pt, and Yu Gothic 24 pt gaps
 * 33.84, 34.08, 32.88, 34.08, 30.96, 33.12 pt. The shared 7/6 face-cell
 * factor comes from PowerPoint observations; the face box always comes from
 * font tables (see core stacked-faces.ts). This is not an Excel fidelity claim.
 * Glyph metrics, vertical forms and shaping follow the available font.
 * Excel's PDF clips at the worksheet print region, not at each shape: keep
 * overflow for the caller's worksheet/viewport clip (no added body clip).
 * OMML uses an atomic upright box from the existing equation renderer; its
 * height advances along the column and its width sizes the column (library
 * policy, not a claim about Office equation positioning).
 */
export function drawShapeStackedText(
  ctx: CanvasRenderingContext2D,
  body: ShapeText,
  direction: 'wordArtVert' | 'wordArtVertRtl',
  rect: DrawingMlTextRect,
  pxPerPt: number,
  resolveFont: (run: TextRun, text: string) => { font: string; px: number; face?: string },
  equation: (run: Extract<ShapeTextRun, { type: 'math' }>, precedingSizePt: number) => StackedEquation | undefined,
): void {
  const measure = (text: string, style: Style): Glyph[] => {
    ctx.font = style.font;
    const box = stackedFaceBox(style.face ?? '', style.bold, style.italic);
    const probe = box ? undefined : ctx.measureText('M');
    // Unknown faces use the drawing engine's box; one em is the portable
    // fallback when that engine does not expose font-box metrics.
    const ascent = box ? box.ascent * style.px : probe?.fontBoundingBoxAscent ?? style.px;
    const descent = box ? box.descent * style.px : probe?.fontBoundingBoxDescent ?? 0;
    const cell = STACKED_CELL_FACTOR * (ascent + descent);
    const face = { family: style.face ?? '', bold: style.bold, italic: style.italic };
    return stackedClusters(text, face, {
      upright: cell, verticalGlyph: style.px, rotated: (ch) => ctx.measureText(ch).width,
    }).map(({ text: ch, kind, advance, nominal }) => ({ text: ch, style, kind, advance: advance + style.spacing,
      thickness: cell, space: ch === ' ', cell, ascent, descent, ...(nominal ? { nominal } : {}) }));
  };
  const sameStyle = (a: Style, b: Style) => !a.equation && !b.equation && a.font === b.font && a.color === b.color && a.spacing === b.spacing;
  const paragraphs: StackedParagraph<Style>[] = body.paragraphs.map((para) => {
    const runs: DrawingMlInputRun<Style>[] = [];
    let mark: Style | undefined;
    let previousFaceText = '';
    let precedingSizePt = 11;
    for (const run of para.runs) {
      if (run.type === 'break') { runs.push({ type: 'break' }); continue; }
      if (run.type === 'math') {
        const object = equation(run, precedingSizePt);
        if (!object) continue; // Same skip contract as horizontal shape text.
        const style: Style = { font: '', px: 0, color: run.color ?? '#000000',
          spacing: 0, bold: false, italic: false, equation: object };
        runs.push({ type: 'object', width: object.height, style, display: run.display });
        previousFaceText = '';
        continue;
      }
      precedingSizePt = run.size > 0 ? run.size : 11;
      // Split only at face changes. Equivalent adjacent run styles can still
      // join in the shared breaker, preserving combining/ZWJ clusters.
      let piece = '';
      let style: Style | undefined;
      const flush = () => { if (style && piece) runs.push({ type: 'text', text: piece, style }); };
      for (const ch of run.text) {
        // UAX #24: marks/joiners keep the base face, including split runs.
        const faceText = /^[\p{M}\p{Cf}]$/u.test(ch) && previousFaceText ? previousFaceText : ch;
        const resolved = resolveFont(run, faceText);
        previousFaceText = faceText;
        const next: Style = { ...resolved, color: run.color ?? '#000000',
          spacing: (run.spacing ?? 0) * pxPerPt, bold: run.bold, italic: run.italic };
        if (style && !sameStyle(style, next)) { flush(); piece = ''; }
        style = next;
        piece += ch;
      }
      style ??= { ...resolveFont(run, 'M'), color: run.color ?? '#000000',
        spacing: (run.spacing ?? 0) * pxPerPt, bold: run.bold, italic: run.italic };
      mark = style;
      flush();
      if (run.text === '') runs.push({ type: 'text', text: '', style });
    }
    const defaultMark: Style = { font: `${11 * pxPerPt}px sans-serif`, px: 11 * pxPerPt,
      color: '#000000', spacing: 0, bold: false, italic: false };
    return { runs, alignment: para.align, lineSpacing: excelShapeSpacing(para.spaceLine),
      spaceBefore: excelShapeSpacing(para.spaceBefore), spaceAfter: excelShapeSpacing(para.spaceAfter),
      emptyThickness: measure('M', mark ?? defaultMark)[0].thickness };
  });
  const layout = layoutStackedText<Style, Glyph>(paragraphs, {
    direction, rect, anchor: body.anchor, anchorCtr: body.anchorCtr,
    wrap: body.wrap !== 'none', spcFirstLastPara: body.spcFirstLastPara,
    lnSpcReduction: body.autoFit === 'norm' ? body.lnSpcReduction ?? 0 : 0,
    pxPerPt, glyphs: measure, sameStyle,
    objectGlyph: ({ style }) => {
      const object = style.equation!;
      return { text: '', style, kind: 'upright', advance: object.height,
        thickness: object.width, space: false, cell: object.height, ascent: 0, descent: 0 };
    },
  });
  ctx.save();
  for (const glyph of layout.glyphs) {
    if (glyph.style.equation) {
      glyph.style.equation.draw(glyph.axisX - glyph.thickness / 2, glyph.cellTop);
      continue;
    }
    if (glyph.space) continue;
    ctx.font = glyph.style.font;
    ctx.fillStyle = glyph.style.color;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'alphabetic';
    // The face box is held in px here: paint geometry takes size 1.
    if (glyph.kind === 'upright') {
      ctx.fillText(glyph.text, glyph.axisX, stackedUprightBaseline(glyph.cellTop, glyph.cell, glyph, 1));
    } else {
      // Library policy: a vertical glyph without a presentation form turns
      // like a rotated one from the cell top (PPTX centres it in its 1 em
      // cell instead; the PowerPoint controls do not cover Excel).
      const form = glyph.kind === 'verticalGlyph' ? stackedVerticalGlyphText(glyph) : null;
      if (form !== null) paintStackedVerticalForm(ctx, form, glyph.axisX, glyph.cellTop, glyph, 1);
      else paintStackedRotated(ctx, glyph.text, glyph.axisX, glyph.cellTop, glyph, 1);
    }
  }
  ctx.restore();
}

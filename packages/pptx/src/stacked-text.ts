// Stacked vertical text, `a:bodyPr@vert` `wordArtVert` / `wordArtVertRtl`
// (ECMA-376 §20.1.10.83). Layout lives in core (`layoutStackedText`, with the
// face and character data in `stacked-faces.ts`); this module turns PowerPoint
// segments into measured glyphs and paints the placed glyphs.
//
// Glyph classes (issue #1626 controls, PowerPoint reference PDF export):
//   upright        — every character outside the sideways set: stands upright,
//                    centred on the column axis, in a cell of 7/6 × face box ×
//                    size; its baseline sits usWinDescent × size (the face
//                    box descent) above the cell bottom.
//   verticalGlyph  — a sideways character the drawing face has a vertical
//                    glyph for: a 1 em cell, baseline face-box ascent below
//                    the cell top, the face box centred on the axis.
//   rotated        — a sideways character without one: turned 90° clockwise,
//                    advancing by its horizontal advance, face box centred on
//                    the axis (ascent to the page right).
// Canvas cannot select a face's `vert` glyph portably, so a vertical glyph is
// drawn as its Unicode vertical presentation form where one exists (、。 and
// the fullwidth brackets) and otherwise as the character turned 90° clockwise
// inside its 1 em cell, which is how those faces design the remaining forms
// (ー ～ ＝ ＿ ￣ ｜ … ‥ ‐ arrows ― ∥).
//
// Decorations (underline, strike, highlight, glyph effects) of a stacked body
// are not painted: no control measured them.

import {
  layoutStackedText,
  STACKED_CELL_FACTOR,
  stackedCellBoxOverride,
  stackedFaceBox,
  stackedSidewaysCharacter,
  stackedVerticalGlyph,
  type DrawingMlInputRun,
  type DrawingMlLineSpacing,
  type DrawingMlTextRect,
  type StackedGlyph,
  type StackedParagraph,
} from '@silurus/ooxml-core/internal/drawingml-text';
import { graphemeClusterOffsets, verticalBracketFormSubstitute, verticalFormSubstitute } from '@silurus/ooxml-core';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

/** The segment fields stacked layout and paint read (renderer's LayoutSegment). */
export interface StackedSegmentStyle {
  math?: { width: number; ascent: number; descent: number };
  font: string;
  sizePx: number;
  color: string;
  noFill?: boolean;
  letterSpacingPx?: number;
  faceFamily?: string;
  faceFamilyLatin?: string;
  /** Face box (em) of the drawing face: usWin, or typo under USE_TYPO_METRICS. */
  lineMetric?: { readonly glyph: { ascent: number; descent: number } | undefined };
  /** Face box of the run's latin face. */
  lineMetricLatin?: { readonly glyph: { ascent: number; descent: number } | undefined } | null;
  /** The run's resolved hyperlink, handed to the selection overlay per glyph. */
  hyperlink?: HyperlinkTarget;
}

export interface StackedParagraphInput<T extends StackedSegmentStyle> {
  runs: readonly DrawingMlInputRun<T>[];
  alignment?: string;
  lineSpacing?: DrawingMlLineSpacing;
  spaceBefore?: DrawingMlLineSpacing;
  spaceAfter?: DrawingMlLineSpacing;
  /** Style of the paragraph mark (sizes an empty column). */
  markStyle: T;
  eastAsianLineBreak?: boolean;
}

export interface StackedBodyInput<T extends StackedSegmentStyle> {
  vert: 'wordArtVert' | 'wordArtVertRtl';
  rect: DrawingMlTextRect;
  anchor?: string;
  anchorCtr?: boolean;
  wrap: boolean;
  spcFirstLastPara?: boolean;
  lnSpcReduction?: number;
  pxPerPt: number;
  paragraphs: readonly StackedParagraphInput<T>[];
  /** Existing equation painter, supplied by the renderer's math-resource owner. */
  drawObject?(style: T, x: number, y: number): void;
  sameStyle?(a: T, b: T): boolean;
}

export interface StackedGlyphRun {
  text: string;
  font: string;
  fontSize: number;
  /** Glyph cell in body coordinates. */
  x: number;
  y: number;
  w: number;
  h: number;
  hyperlink?: HyperlinkTarget;
}

type Ctx2D = CanvasRenderingContext2D;

/**
 * Face box of a segment in em: the #1610 glyph box of its face, else the
 * catalogued box of that family at any weight, else the resolved Canvas font
 * box (an unknown face).
 */
function faceBox(ctx: Ctx2D, style: StackedSegmentStyle): { ascent: number; descent: number } {
  const known = style.lineMetric?.glyph
    ?? (style.faceFamily
      ? stackedFaceBox(style.faceFamily, /\bbold\b/.test(style.font), /\bitalic\b/.test(style.font)) : undefined);
  if (known) return known;
  ctx.font = style.font;
  const m = ctx.measureText('M');
  const size = style.sizePx || 1;
  const ascent = m.fontBoundingBoxAscent;
  const descent = m.fontBoundingBoxDescent;
  if (Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0) {
    return { ascent: ascent / size, descent: descent / size };
  }
  return { ascent: 0.9, descent: 0.25 };
}

interface Measured<T> extends StackedGlyph<T> {
  box: { ascent: number; descent: number };
  cell: number;
}

function makeMeasurer<T extends StackedSegmentStyle>(ctx: Ctx2D) {
  const boxes = new Map<T, { ascent: number; descent: number }>();
  return (text: string, style: T): Measured<T>[] => {
    let box = boxes.get(style);
    if (!box) {
      box = faceBox(ctx, style);
      boxes.set(style, box);
    }
    const family = style.faceFamily ?? '';
    const cellEm = stackedCellBoxOverride(family) ?? box.ascent + box.descent;
    const cell = STACKED_CELL_FACTOR * cellEm * style.sizePx;
    // The run's latin face also sizes the column (MS Mincho glyphs of a run
    // whose latin face is Arial sit in an Arial-thick column, wordartvert
    // slides 10-11), as it sizes a horizontal line (#1610).
    const latinFamily = style.faceFamilyLatin ?? family;
    const latinBox = style.lineMetricLatin?.glyph
      ?? stackedFaceBox(latinFamily, /\bbold\b/.test(style.font), /\bitalic\b/.test(style.font));
    const latinEm = latinFamily === family ? cellEm
      : stackedCellBoxOverride(latinFamily) ?? (latinBox ? latinBox.ascent + latinBox.descent : cellEm);
    const thickness = Math.max(cell, STACKED_CELL_FACTOR * latinEm * style.sizePx);
    const spc = style.letterSpacingPx ?? 0;
    const out: Measured<T>[] = [];
    // One stacked cell per grapheme cluster: a base with its combining marks
    // or variation selectors (é, か + ゙) stands in one cell, measured and
    // painted together. The cluster takes the class of its base character:
    // every measured sideways character is a single code point, and a mark
    // attached to it (for example a variation selector) does not change how
    // PowerPoint turns the base.
    let start = 0;
    const ends = graphemeClusterOffsets(text);
    ends.push(text.length);
    for (const end of ends) {
      const ch = text.slice(start, end);
      start = end;
      if (!ch) continue;
      const cp = ch.codePointAt(0) ?? 0;
      let kind: Measured<T>['kind'] = 'upright';
      let advance = cell;
      if (stackedSidewaysCharacter(cp)) {
        if (stackedVerticalGlyph(family, cp)) {
          kind = 'verticalGlyph';
          advance = style.sizePx;
        } else {
          kind = 'rotated';
          ctx.font = style.font;
          advance = ctx.measureText(ch).width;
        }
      }
      out.push({ text: ch, style, kind, advance: advance + spc, thickness, space: ch === ' ', box, cell });
    }
    return out;
  };
}

/**
 * Lay out and paint a stacked body. `ctx` is in slide coordinates; the body
 * rectangle has already been placed by the caller (shape rotation and flips
 * are applied outside, as for horizontal text). Returns one run per glyph for
 * the selection overlay.
 */
export function renderStackedText<T extends StackedSegmentStyle>(
  ctx: Ctx2D,
  body: StackedBodyInput<T>,
): StackedGlyphRun[] {
  const measure = makeMeasurer<T>(ctx);
  const paragraphs: StackedParagraph<T>[] = body.paragraphs.map((p) => ({
    // Upright OMML's height is the advance along a stacked column; the
    // horizontal equation width remains its cross-column extent.
    runs: p.runs.map((run) => run.type === 'object' && run.style.math
      ? { ...run, width: run.style.math.ascent + run.style.math.descent } : run),
    alignment: p.alignment,
    lineSpacing: p.lineSpacing,
    spaceBefore: p.spaceBefore,
    spaceAfter: p.spaceAfter,
    emptyThickness: measure('M', p.markStyle)[0].thickness,
    eastAsianLineBreak: p.eastAsianLineBreak,
  }));
  const layout = layoutStackedText<T, Measured<T>>(paragraphs, {
    direction: body.vert,
    rect: body.rect,
    anchor: body.anchor,
    anchorCtr: body.anchorCtr,
    wrap: body.wrap,
    spcFirstLastPara: body.spcFirstLastPara,
    lnSpcReduction: body.lnSpcReduction,
    pxPerPt: body.pxPerPt,
    glyphs: measure,
    sameStyle: body.sameStyle,
    objectGlyph: ({ style }) => {
      if (!style.math || !body.drawObject) throw new Error('Missing stacked equation adapter');
      const height = style.math.ascent + style.math.descent;
      return { text: '', style, kind: 'upright', advance: height, thickness: style.math.width,
        space: false, box: { ascent: 0, descent: 0 }, cell: height };
    },
  });
  const runs: StackedGlyphRun[] = [];
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  for (const placed of layout.glyphs) {
    const g = placed;
    const { style, box } = g;
    if (style.math) {
      body.drawObject!(style, placed.axisX - g.thickness / 2, placed.cellTop);
      continue;
    }
    const size = style.sizePx;
    runs.push({
      text: g.text, font: style.font, fontSize: size,
      x: placed.axisX - g.thickness / 2, y: placed.cellTop, w: g.thickness, h: g.advance,
      ...(style.hyperlink ? { hyperlink: style.hyperlink } : {}),
    });
    if (g.space || style.noFill) continue;
    ctx.font = style.font;
    ctx.fillStyle = style.color;
    const cp = g.text.codePointAt(0) ?? 0;
    if (g.kind === 'upright') {
      ctx.textAlign = 'center';
      ctx.textBaseline = 'alphabetic';
      ctx.fillText(g.text, placed.axisX, placed.cellTop + g.cell - box.descent * size);
    } else if (g.kind === 'verticalGlyph') {
      // A presentation form replaces a lone character only; a cluster with a
      // mark keeps its own glyphs and takes the turned fallback.
      const single = String.fromCodePoint(cp) === g.text;
      const form = single ? verticalFormSubstitute(cp) ?? verticalBracketFormSubstitute(cp) : null;
      const left = placed.axisX - (box.ascent + box.descent) * size / 2;
      if (form !== null) {
        ctx.textAlign = 'left';
        ctx.textBaseline = 'alphabetic';
        ctx.fillText(String.fromCodePoint(form), left, placed.cellTop + box.ascent * size);
      } else {
        ctx.save();
        ctx.translate(left + size / 2, placed.cellTop + size / 2);
        ctx.rotate(Math.PI / 2);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(g.text, 0, 0);
        ctx.restore();
      }
    } else {
      ctx.save();
      ctx.translate(placed.axisX - (box.ascent - box.descent) * size / 2, placed.cellTop);
      ctx.rotate(Math.PI / 2);
      ctx.textAlign = 'left';
      ctx.textBaseline = 'alphabetic';
      ctx.fillText(g.text, 0, 0);
      ctx.restore();
    }
  }
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  return runs;
}

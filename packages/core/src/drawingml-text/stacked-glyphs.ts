import { graphemeClusterOffsets } from '../text/sea-break.js';
import { verticalBracketFormSubstitute, verticalFormSubstitute } from '../text/vertical-orientation.js';
import type { StackedGlyphKind } from './stacked.js';
import { stackedSidewaysCharacter, stackedVerticalGlyphSource } from './stacked-faces.js';

/**
 * Glyph primitives of stacked vertical text (`a:bodyPr@vert` `wordArtVert` /
 * `wordArtVertRtl`, ECMA-376 §20.1.10.83) shared by the PPTX and XLSX hosts.
 * The hosts keep their own face-box resolution, column thickness, decoration
 * and selection; these helpers fix the per-glyph class, advance and paint
 * frame (issue #1626 PowerPoint controls; see `stacked.ts` and
 * `stacked-faces.ts`).
 *
 * Paint geometry takes the face box in `box` units and `size` (px per unit):
 * a host holding the box in em passes the font size, a host holding it in px
 * passes 1.
 */

export interface StackedFace {
  family: string;
  bold: boolean;
  italic: boolean;
}

export interface StackedCluster {
  text: string;
  kind: StackedGlyphKind;
  /** Advance down the column, before rPr@spc. */
  advance: number;
  /** A vertical glyph drawn as its unturned horizontal glyph (vrt2-only
   * coverage outside Basic Latin; see `stackedVerticalGlyphSource`). */
  nominal?: true;
}

export interface StackedClusterAdvances {
  /** Cell of an upright glyph (7/6 × face box × size). */
  upright: number;
  /** Cell of a vertical glyph: one em. */
  verticalGlyph: number;
  /** Horizontal advance of a rotated glyph. */
  rotated(text: string): number;
}

/**
 * One stacked cell per grapheme cluster: a base with its combining marks or
 * variation selectors (é, か + ゙) stands in one cell. The cluster takes the
 * class of its base character: every measured sideways character is a single
 * code point, and a mark attached to it does not change how PowerPoint turns
 * the base. `upright` forces every cluster upright (paragraph markers).
 */
export function stackedClusters(
  text: string,
  face: StackedFace,
  advances: StackedClusterAdvances,
  upright = false,
): StackedCluster[] {
  const out: StackedCluster[] = [];
  const ends = graphemeClusterOffsets(text);
  ends.push(text.length);
  let start = 0;
  for (const end of ends) {
    const ch = text.slice(start, end);
    start = end;
    if (!ch) continue;
    const cp = ch.codePointAt(0) ?? 0;
    const source = upright ? undefined : stackedVerticalGlyphSource(face.family, cp, face.bold, face.italic);
    if (upright || !stackedSidewaysCharacter(cp)) {
      out.push({ text: ch, kind: 'upright', advance: advances.upright });
    } else if (source) {
      out.push({ text: ch, kind: 'verticalGlyph', advance: advances.verticalGlyph,
        ...(source === 'nominal' ? { nominal: true } : {}) });
    } else {
      out.push({ text: ch, kind: 'rotated', advance: advances.rotated(ch) });
    }
  }
  return out;
}

/**
 * The text a vertical glyph paints in the vertical-form frame
 * ({@link paintStackedVerticalForm}), or null when the host turns it instead.
 * A `nominal` glyph paints its own text unturned. Canvas cannot select a
 * face's `vert` glyph portably, so a `vert` glyph is drawn as its Unicode
 * vertical presentation form where one exists (、。 and the fullwidth
 * brackets). A form replaces a lone character only; a cluster with a mark
 * keeps its own glyphs.
 */
export function stackedVerticalGlyphText(glyph: { text: string; nominal?: true }): string | null {
  return glyph.nominal ? glyph.text : stackedVerticalForm(glyph.text);
}

/** The Unicode vertical presentation form of a lone character, or null. */
export function stackedVerticalForm(text: string): string | null {
  const cp = text.codePointAt(0) ?? 0;
  if (String.fromCodePoint(cp) !== text) return null;
  const form = verticalFormSubstitute(cp) ?? verticalBracketFormSubstitute(cp);
  return form === null ? null : String.fromCodePoint(form);
}

type FaceBox = { readonly ascent: number; readonly descent: number };

/** Alphabetic baseline of an upright glyph: face-box descent above the cell bottom. */
export function stackedUprightBaseline(cellTop: number, cell: number, box: FaceBox, size: number): number {
  return cellTop + cell - box.descent * size;
}

/**
 * Paint a vertical presentation form: the face box centred on the axis, its
 * baseline face-box ascent below the cell top. `beforeFill` sees the left
 * edge and baseline (host decorations).
 */
export function paintStackedVerticalForm(
  ctx: CanvasRenderingContext2D,
  form: string,
  axisX: number,
  cellTop: number,
  box: FaceBox,
  size: number,
  beforeFill?: (left: number, baseline: number) => void,
): void {
  const left = axisX - (box.ascent + box.descent) * size / 2;
  const baseline = cellTop + box.ascent * size;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  beforeFill?.(left, baseline);
  ctx.fillText(form, left, baseline);
}

/**
 * Paint a glyph turned 90° clockwise from the cell top, the face box centred
 * on the axis (ascent to the page right). `beforeFill` runs in the rotated
 * frame with the baseline at y = 0 (host decorations).
 */
export function paintStackedRotated(
  ctx: CanvasRenderingContext2D,
  text: string,
  axisX: number,
  cellTop: number,
  box: FaceBox,
  size: number,
  beforeFill?: () => void,
): void {
  ctx.save();
  ctx.translate(axisX - (box.ascent - box.descent) * size / 2, cellTop);
  ctx.rotate(Math.PI / 2);
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  beforeFill?.();
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

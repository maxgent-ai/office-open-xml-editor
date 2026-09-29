// Per-line bidi ordering for the xlsx renderer (rich-text cell runs).
//
// We reorder a cell line's rich-text runs at SEGMENT granularity (1:1 with the
// runs — every per-run font/colour property is preserved) using the shared
// UAX#9 engine (rule L2), and let Canvas shape each run internally when it is
// drawn with `ctx.direction` set to the run's resolved direction. The whole run
// string is drawn in one fillText, so Canvas resolves any residual
// intra-run bidi.

import {
  resolveBaseDirection,
  hasStrongRtl,
} from '@silurus/ooxml-core';

/**
 * Resolve a cell's base direction from its xf @readingOrder
 * (ECMA-376 §18.8.1: 1 = LTR, 2 = RTL, 0/absent = Context → UAX#9 first-strong).
 */
export function cellBaseRtl(readingOrder: number | undefined, text: string): boolean {
  if (readingOrder === 2) return true;
  if (readingOrder === 1) return false;
  return resolveBaseDirection(undefined, text) === 'rtl';
}

import { computeSegmentLineVisualOrder, segmentLineHasRtl, type SegmentLineVisualOrder } from '@silurus/ooxml-core/internal/bidi-line';

/**
 * Resolve whether one cell line / paragraph needs the bidi pass and its base
 * direction, from the xf @readingOrder (§18.8.1) and the line/paragraph text.
 * Gated so pure-LTR text (no strong-RTL char and not explicitly RTL) keeps the
 * exact pre-bidi path: `needBidi` false ⇒ no UAX#9 reorder. `baseRtl` is only
 * resolved when the pass will run. The single source of truth for the gate +
 * base-direction rule shared by the non-wrap ({@link cellBaseRtl} per LF line)
 * and wrap (per LF paragraph) rich-text paths — keeping the two in lockstep.
 */
export function resolveCellBidi(
  readingOrder: number | undefined,
  text: string,
): { needBidi: boolean; baseRtl: boolean } {
  const needBidi = readingOrder === 2 || hasStrongRtl(text);
  return { needBidi, baseRtl: needBidi && cellBaseRtl(readingOrder, text) };
}

export type LineVisualOrder = SegmentLineVisualOrder;
export const segmentsHaveRtl = (segments: readonly unknown[]): boolean =>
  segmentLineHasRtl(segments);
export const computeLineVisualOrder = (
  segments: readonly unknown[], baseRtl: boolean,
): LineVisualOrder => computeSegmentLineVisualOrder(segments, baseRtl);

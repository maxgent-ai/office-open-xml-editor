import { EMU_PER_PX } from '@silurus/ooxml-core';
import type { DrawingAnchorTag } from '../types.js';

/** Library compatibility policy for UNTAGGED models only (no `anchorTag`):
 * whether an anchor saved as `editAs="oneCell"` has the complete positive
 * native extent required to ignore its `to` marker. Both dimensions form one
 * DrawingML size; a partial/malformed extent falls back as a unit so culling
 * and paint always derive the same rectangle. Tagged models never consult
 * this; see {@link resolveCellAnchorRect}. */
export function usesNativeOneCellExtent(anchor: {
  readonly editAs?: string;
  readonly nativeExtCx?: number;
  readonly nativeExtCy?: number;
}): boolean {
  return anchor.editAs === 'oneCell'
    && (anchor.nativeExtCx ?? 0) > 0
    && (anchor.nativeExtCy ?? 0) > 0;
}

/** Anchor facts read by {@link resolveCellAnchorRect}. Markers, `editAs`,
 * `anchorTag` and every extent are normative ECMA-376 acquisition facts; how
 * they combine into a display rectangle is library policy (documented on the
 * resolver). */
export interface CellAnchorSizeFacts {
  readonly fromCol: number;
  readonly fromColOff: number;
  readonly fromRow: number;
  readonly fromRowOff: number;
  readonly toCol: number;
  readonly toColOff: number;
  readonly toRow: number;
  readonly toRowOff: number;
  readonly editAs?: string;
  readonly nativeExtCx?: number;
  readonly nativeExtCy?: number;
  readonly anchorTag?: DrawingAnchorTag;
  readonly anchorExtCx?: number;
  readonly anchorExtCy?: number;
}

/** The single axis operation the resolver needs: scaled sheet-space px of the
 * leading edge of a 1-based column/row band. */
export interface CellAnchorAxis {
  offsetOf(index1: number): number;
}

/** EMU display size retained from an anchor's prepared initial display. Pure
 * input: this module holds no state. Internal viewer/worker projections bind
 * a compact reference (internal/initial-anchor-sizes.ts) and every consumer
 * (paint, hit/outline, culling/decode) looks it up for the same anchor. A
 * supplied value is used exactly, including a non-positive (non-renderable)
 * size; it is never replaced by the current, edited `to` marker. */
export interface RetainedAnchorExtent {
  readonly cx: number;
  readonly cy: number;
}

/** Sheet-space rectangle in scaled px. Width/height may be <= 0 (or NaN for
 * malformed numeric input); callers treat a non-positive size as not
 * renderable, exactly as before. */
export interface SheetAnchorRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/**
 * Resolve an anchor's sheet-space display rectangle. Picture paint, shape-group
 * paint, element hit/outline projection, and render culling/decode sizing all
 * call this one function so they always agree on the same rectangle.
 *
 * The top-left is always the `from` marker. The size is chosen as follows:
 * - Untagged (`anchorTag` absent: old models, VML/OLE previews, legacy binary
 *   conversion). Library compatibility policy, unchanged: use the complete
 *   positive native extent when {@link usesNativeOneCellExtent}; otherwise use
 *   the `to` marker, including when only part of the native extent is present.
 * - `twoCellAnchor` (normative, ECMA-376 Part 1 §20.5.2.33). `from`/`to` is
 *   the initial display rectangle for every `editAs`. `editAs` only governs
 *   how later band edits move or resize the object (§20.5.3.2).
 *   Library policy: for `editAs="oneCell"`, a retained prepared-initial
 *   extent supplied by the caller replaces `to` exactly; a non-positive one
 *   stays non-renderable (no fallback). Without one (the stateless public
 *   renderer has no edit history) `from`/`to` is used, which is correct for
 *   the INITIAL display only; live band-edit retention is promised solely on
 *   the internal viewer path that binds a prepared reference. The child/group
 *   xfrm ext is never used for tagged anchors.
 * - `oneCellAnchor` (normative, §20.5.2.24). Use the anchor-level `<xdr:ext>`.
 *   Library validity policy for this new metadata: a missing or zero component
 *   yields a non-positive, non-renderable rectangle. No fallback is guessed.
 * - Any other tag value uses `from`/`to`. No kind is inferred from geometry.
 */
export function resolveCellAnchorRect(
  anchor: CellAnchorSizeFacts,
  colAxis: CellAnchorAxis,
  rowAxis: CellAnchorAxis,
  scale: number,
  retainedInitialExtent?: RetainedAnchorExtent,
): SheetAnchorRect {
  const x = colAxis.offsetOf(anchor.fromCol + 1) + (anchor.fromColOff * scale) / EMU_PER_PX;
  const y = rowAxis.offsetOf(anchor.fromRow + 1) + (anchor.fromRowOff * scale) / EMU_PER_PX;
  const fixed = fixedExtentEmu(anchor, retainedInitialExtent);
  if (fixed) {
    return {
      x,
      y,
      width: (fixed.cx * scale) / EMU_PER_PX,
      height: (fixed.cy * scale) / EMU_PER_PX,
    };
  }
  const x2 = colAxis.offsetOf(anchor.toCol + 1) + (anchor.toColOff * scale) / EMU_PER_PX;
  const y2 = rowAxis.offsetOf(anchor.toRow + 1) + (anchor.toRowOff * scale) / EMU_PER_PX;
  return { x, y, width: x2 - x, height: y2 - y };
}

/** EMU size that replaces the `to` marker, or null when `to` is used. */
function fixedExtentEmu(
  anchor: CellAnchorSizeFacts,
  retained: RetainedAnchorExtent | undefined,
): { cx: number; cy: number } | null {
  if (anchor.anchorTag === undefined) {
    return usesNativeOneCellExtent(anchor)
      ? { cx: anchor.nativeExtCx as number, cy: anchor.nativeExtCy as number }
      : null;
  }
  if (anchor.anchorTag === 'oneCellAnchor') {
    return { cx: anchor.anchorExtCx ?? 0, cy: anchor.anchorExtCy ?? 0 };
  }
  if (
    anchor.anchorTag === 'twoCellAnchor'
    && anchor.editAs === 'oneCell'
    && retained !== undefined
  ) {
    return { cx: retained.cx, cy: retained.cy };
  }
  return null;
}

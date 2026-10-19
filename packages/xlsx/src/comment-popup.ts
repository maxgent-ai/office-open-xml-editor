/**
 * Pure geometry for the cell-comment hover popup. Kept free of DOM so it can be
 * unit-tested in isolation; the viewer owns the actual element and feeds it
 * measured rects.
 *
 * Excel shows a comment/note as a small box anchored at the commented cell's
 * top-right corner, opening toward the sheet body. This calculator reproduces
 * that placement and adds viewport-aware flipping/clamping so the popup is
 * always fully visible.
 *
 * All coordinates are in canvasArea CSS-pixel space — the same space
 * `getCellRect` and the selection overlay use. The cell rect is therefore
 * already RTL-mirrored by the caller (via `screenX`); the `rtl` flag here only
 * picks the *default* open direction (left for RTL so the popup grows into the
 * mirrored sheet body rather than off the right edge).
 */

/** Gap (CSS px) between the cell edge and the popup. */
const COMMENT_POPUP_GAP = 8;

export interface CommentPopupGeometry {
  /** The commented cell's on-screen rect (canvasArea space). */
  cell: { x: number; y: number; w: number; h: number };
  /** The popup's measured size. */
  popup: { w: number; h: number };
  /** The visible viewport (canvasArea client size). */
  viewport: { w: number; h: number };
  /** True when the current sheet is laid out right-to-left. */
  rtl: boolean;
}

/**
 * Compute the top-left position of the popup, in canvasArea CSS pixels.
 *
 * Horizontal: prefer the side that opens into the sheet body (right for LTR,
 * left for RTL). If that side overflows the viewport, flip to the opposite
 * side. If neither side fits (a popup wider than the room on both sides), clamp
 * the preferred side into the viewport.
 *
 * Vertical: anchor at the cell's top edge, then clamp so the popup stays fully
 * within [0, viewport.h].
 */
export function computeCommentPopupPosition(geo: CommentPopupGeometry): {
  left: number;
  top: number;
} {
  const { cell, popup, viewport, rtl } = geo;
  const gap = COMMENT_POPUP_GAP;

  const rightOf = cell.x + cell.w + gap; // popup's left when opening to the right
  const leftOf = cell.x - gap - popup.w; // popup's left when opening to the left

  const fitsRight = rightOf + popup.w <= viewport.w;
  const fitsLeft = leftOf >= 0;

  let left: number;
  if (!rtl) {
    // LTR: prefer right, flip to left, else clamp.
    if (fitsRight) left = rightOf;
    else if (fitsLeft) left = leftOf;
    else left = rightOf;
  } else {
    // RTL: prefer left, flip to right, else clamp.
    if (fitsLeft) left = leftOf;
    else if (fitsRight) left = rightOf;
    else left = leftOf;
  }

  // Final clamp so a popup wider than the available room on the chosen side is
  // still pulled fully into view.
  left = Math.max(0, Math.min(left, viewport.w - popup.w));

  let top = cell.y;
  top = Math.max(0, Math.min(top, viewport.h - popup.h));

  return { left, top };
}

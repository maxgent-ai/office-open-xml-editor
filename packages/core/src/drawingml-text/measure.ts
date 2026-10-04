import { FONT_TRACKING_SENTINEL } from '../internal/font-measurement-sentinels.js';
/**
 * Canvas run advance including DrawingML `a:rPr@spc` (§21.1.2.3.9).
 * A native Canvas tracking measurement includes one trailing spacing unit;
 * OOXML spacing belongs only between characters, so remove that unit. The
 * fallback uses code points and is shared by measure and paint adapters.
 */
const nativeSpacing = new WeakMap<CanvasRenderingContext2D, boolean>();

export function drawingMlCodePointCount(text: string): number {
  let n = 0;
  for (const _ of text) n++;
  return n;
}

function hasNativeSpacing(ctx: CanvasRenderingContext2D): boolean {
  const cached = nativeSpacing.get(ctx);
  if (cached !== undefined) return cached;
  const tracking = ctx as CanvasRenderingContext2D & { letterSpacing?: string };
  const previous = tracking.letterSpacing;
  if (typeof previous !== 'string') {
    nativeSpacing.set(ctx, false);
    return false;
  }
  let supported = false;
  try {
    tracking.letterSpacing = '0px';
    const natural = ctx.measureText(FONT_TRACKING_SENTINEL).width;
    tracking.letterSpacing = '1px';
    const tracked = ctx.measureText(FONT_TRACKING_SENTINEL).width;
    supported = Number.isFinite(natural) && Number.isFinite(tracked) && tracked !== natural;
  } catch {
    supported = false;
  } finally {
    try { tracking.letterSpacing = previous; } catch { /* inert context */ }
  }
  nativeSpacing.set(ctx, supported);
  return supported;
}

export function measureDrawingMlAdvance(
  ctx: CanvasRenderingContext2D,
  text: string,
  letterSpacingPx = 0,
): number {
  const tracking = ctx as CanvasRenderingContext2D & { letterSpacing?: string };
  const previous = tracking.letterSpacing;
  if (letterSpacingPx !== 0 && hasNativeSpacing(ctx)) {
    try {
      tracking.letterSpacing = `${letterSpacingPx}px`;
      const tracked = ctx.measureText(text).width;
      if (Number.isFinite(tracked)) return text.length > 0 ? tracked - letterSpacingPx : tracked;
    } finally {
      try { tracking.letterSpacing = previous; } catch { /* inert context */ }
    }
  }
  return ctx.measureText(text).width
    + letterSpacingPx * Math.max(0, drawingMlCodePointCount(text) - 1);
}

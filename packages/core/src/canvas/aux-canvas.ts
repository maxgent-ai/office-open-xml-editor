/**
 * Auxiliary (offscreen) canvas allocation shared across the renderers and the
 * metafile rasterizers.
 *
 * Lives under `canvas/` (a leaf utility layer) so both `shape/` (effects,
 * scene3d) and `image/` (wmf/emf/dib metafile players) can depend on it without
 * an unnatural `image → shape` edge. `shape/effects.ts` re-exports it for its
 * historical import path.
 */

/** A canvas usable off-DOM: an OffscreenCanvas in a worker, else a detached
 *  `<canvas>`. */
export type AuxCanvas = HTMLCanvasElement | OffscreenCanvas;
/** The 2D context type of an {@link AuxCanvas}. */
export type AuxContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
type SourceContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

function dimensions(w: number, h: number): [number, number] {
  return [Math.max(1, Math.ceil(w)), Math.max(1, Math.ceil(h))];
}

/**
 * Allocate an auxiliary canvas. Prefers OffscreenCanvas (no DOM pollution, works
 * inside a worker); falls back to a detached <canvas>; returns null when neither
 * is available (e.g. a headless unit-test environment) so callers can no-op.
 * Dimensions are clamped to >=1 to avoid zero-size canvas errors.
 */
export function createAuxCanvas(w: number, h: number): AuxCanvas | null {
  const [cw, ch] = dimensions(w, h);
  if (typeof OffscreenCanvas !== 'undefined') {
    return new OffscreenCanvas(cw, ch);
  }
  if (typeof document !== 'undefined') {
    const c = document.createElement('canvas');
    c.width = cw;
    c.height = ch;
    return c;
  }
  return null;
}

/**
 * Allocate a scratch canvas for a specific live context. The final constructor
 * fallback supports Canvas-compatible node implementations such as skia-canvas;
 * it intentionally runs after browser/worker strategies because constructing an
 * HTMLCanvasElement directly throws. Every strategy is isolated so draw callers
 * can safely degrade when allocation is unavailable.
 */
export function createAuxCanvasForContext(
  ctx: SourceContext,
  w: number,
  h: number,
): AuxCanvas | null {
  const [cw, ch] = dimensions(w, h);
  if (typeof OffscreenCanvas !== 'undefined') {
    try {
      return new OffscreenCanvas(cw, ch);
    } catch {
      // Continue to the DOM canvas strategy.
    }
  }
  if (typeof document !== 'undefined') {
    try {
      const canvas = document.createElement('canvas');
      canvas.width = cw;
      canvas.height = ch;
      return canvas;
    } catch {
      // Continue to the context-compatible constructor strategy.
    }
  }
  try {
    // This canvas comes from the caller's rendering context. Its constructor
    // preserves the host canvas class when document.createElement is unavailable.
    const constructor = ctx.canvas?.constructor;
    if (typeof constructor !== 'function') return null;
    const CanvasConstructor = constructor as new (width: number, height: number) => AuxCanvas;
    return new CanvasConstructor(cw, ch);
  } catch {
    return null;
  }
}

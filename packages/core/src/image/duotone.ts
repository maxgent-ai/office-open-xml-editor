import { createAuxCanvas } from '../canvas/aux-canvas.js';
import { interpolateColorChannel } from '../shape/color-space.js';
import { setImageEffectFrame, type EffectBuffer } from './image-effect-frame.js';
// Shared DrawingML duotone image effect (ECMA-376 §20.1.8.23 CT_DuotoneEffect)
// for the docx, pptx and xlsx renderers. Spec text (§20.1.8.23):
//
//   "For each pixel, combines clr1 and clr2 through a linear interpolation to
//    determine the new color for that pixel."
//
// `clr1` is the first `EG_ColorChoice` child, `clr2` the second. The
// interpolation factor is the pixel's luminance (0 = darkest → `clr1`, 1 =
// lightest → `clr2`), so a near-white photo under a `black`↔`light-pink` duotone
// is remapped along that ramp to pink (the sample-9.xlsx "Gift budget" picture).
// This matches PowerPoint / Word / Excel, which recolor by luminance.
//
// The pixel transform lives here as a pure function over an ImageData-shaped
// buffer so it is unit-testable without a canvas; the thin canvas wrapper
// (getImageData → transform → putImageData → ImageBitmap) is provided for the
// renderers and works on both OffscreenCanvas (browser/worker) and node-canvas
// (skia) via the shared 2D context surface.

import { decodedBitmapTargetResizeOptions } from './raster-target.js';

/** A duotone effect resolved to its two endpoint colours. Both are 6- or 8-char
 *  uppercase hex WITHOUT a leading `#` (the form the Rust parsers emit). `clr1`
 *  is the dark endpoint (luminance 0), `clr2` the light endpoint (luminance 1),
 *  matching the child order of `<a:duotone>` in §20.1.8.23. Any per-colour
 *  transforms (lumMod/lumOff/tint/satMod/…) are already baked into these hexes
 *  by the parser's colour-resolution machinery. */
export interface Duotone {
  /** First `EG_ColorChoice` child — dark endpoint, RRGGBB or RRGGBBAA. */
  clr1: string;
  /** Second `EG_ColorChoice` child — light endpoint, RRGGBB or RRGGBBAA. */
  clr2: string;
}

/** An ImageData-shaped RGBA buffer: `data` is `RGBA…` bytes row-major, length
 *  `width*height*4`. A real {@link ImageData} satisfies this; tests pass a plain
 *  object so the transform is exercised without a canvas. */
export interface RgbaBuffer {
  data: Uint8ClampedArray | Uint8Array;
  width: number;
  height: number;
}

/** Parse a 6-char hex (no `#`) to `[r, g, b]` (0–255). Returns `null` when the
 *  string is not exactly 6 hex digits, so a malformed colour disables the
 *  effect rather than corrupting pixels. */
export function hex6ToRgb(hex: string): [number, number, number] | null {
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return null;
  return [
    parseInt(hex.slice(0, 2), 16),
    parseInt(hex.slice(2, 4), 16),
    parseInt(hex.slice(4, 6), 16),
  ];
}

/** Rec. 601 relative luminance (0–1) of an sRGB byte triple. Rec. 601 is the
 *  classic grayscale weighting used by DrawingML raster effects (the same
 *  coefficients Office uses for its "washout"/recolor luminance ramp); it is a
 *  gamma-space approximation, matching how the effect is applied to stored
 *  (non-linear) pixel values rather than linear light. */
export function luminance601(r: number, g: number, b: number): number {
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255;
}

/** In-place duotone remap of an RGBA buffer along the `clr1`→`clr2` ramp by
 *  per-pixel Rec. 601 luminance. Endpoint opacity is interpolated and modulates source alpha; opaque
 *  endpoints therefore preserve alpha. Transparent RGB is transformed too: a later ordered
 *  alphaRepl/alphaInv can make those colors visible. Returns the same buffer. */
export function duotoneImageData(buf: RgbaBuffer, dark: string, light: string): RgbaBuffer {
  const c1 = /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(dark) ? hex6ToRgb(dark.slice(0, 6)) : null;
  const c2 = /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(light) ? hex6ToRgb(light.slice(0, 6)) : null;
  if (!c1 || !c2) return buf;
  const da = dark.length === 8 ? parseInt(dark.slice(6), 16) / 255 : 1;
  const la = light.length === 8 ? parseInt(light.slice(6), 16) / 255 : 1;
  const [dr, dg, db] = c1;
  const [lr, lg, lb] = c2;
  const d = buf.data;
  for (let i = 0; i < d.length; i += 4) {
    const t = luminance601(d[i], d[i + 1], d[i + 2]);
    d[i] = interpolateColorChannel(dr, lr, t);
    d[i + 1] = interpolateColorChannel(dg, lg, t);
    d[i + 2] = interpolateColorChannel(db, lb, t);
    // §20.1.8.23 interpolates the complete endpoint color; its opacity
    // modulates source coverage, preserving transparent source pixels.
    d[i + 3] = Math.round(d[i + 3] * (da + (la - da) * t));
  }
  return buf;
}

/** Minimal 2D context surface the canvas wrapper needs. Both
 *  `OffscreenCanvasRenderingContext2D` and node-canvas's context satisfy it. */
interface Ctx2DLike {
  drawImage(img: CanvasImageSource, dx: number, dy: number): void;
  getImageData(sx: number, sy: number, sw: number, sh: number): ImageData;
  putImageData(data: ImageData, dx: number, dy: number): void;
  createImageData?(width: number, height: number): ImageData;
}

/** An offscreen drawing surface {@link applyDuotone} draws onto and reads back.
 *  It must both expose a 2D context (for the pixel round-trip) AND be a valid
 *  `ImageBitmapSource` (so the recoloured surface can be baked into a new
 *  `ImageBitmap`). A real `OffscreenCanvas` satisfies both; the intersection lets
 *  the factory return one without any cast, and a node-canvas / test mock is
 *  cast (once) at the factory boundary where the shape is known. */
export type OffscreenSurface = ImageBitmapSource & {
  readonly width: number;
  readonly height: number;
  getContext(id: '2d'): Ctx2DLike | null;
};

/** Factory for an offscreen drawing surface. Injected so the renderers pass an
 *  `OffscreenCanvas` (browser/worker) or a node-canvas (skia) constructor; the
 *  transform code stays environment-agnostic. Returns `null` when no surface can
 *  be made, in which case {@link applyDuotone} returns the source untransformed. */
export type OffscreenFactory = (w: number, h: number) => OffscreenSurface | null;

/** Shared Canvas factory: OffscreenCanvas in workers/modern browsers, or a
 * main-thread HTML canvas. A Window that can paint images does not need an
 * OffscreenCanvas merely to run its authored pixel corrections. Node hosts
 * supply the existing factory/shim; no environment-specific image algorithm. */
export const defaultOffscreenFactory: OffscreenFactory = (w, h) =>
  createAuxCanvas(w, h) as OffscreenSurface | null;

/** Apply a duotone effect to a decoded image, returning a NEW `ImageBitmap`
 *  (via `createImageBitmap`) recoloured along the `clr1`→`clr2` luminance ramp.
 *  The source is drawn onto an offscreen surface at its native pixel size, its
 *  pixels are remapped by {@link duotoneImageData}, and the result is baked back
 *  into an `ImageBitmap` so the render path is unchanged (still a
 *  `CanvasImageSource` drawn by `drawImageCropped`). Alpha is preserved.
 *
 *  Returns the ORIGINAL source unchanged when the surface/ImageData pipeline is
 *  unavailable (e.g. `createImageBitmap`/`OffscreenCanvas` missing), so a picture
 *  never vanishes — it just renders without the recolor. Callers cache the
 *  result keyed by (image path + both colours) so the transform runs once. */
export async function applyDuotone(
  img: CanvasImageSource,
  duotone: Duotone,
  opts: ImageTransformOptions,
): Promise<CanvasImageSource> {
  return applyImageDataTransform(
    img,
    (data) => {
      duotoneImageData(data, duotone.clr1, duotone.clr2);
    },
    opts,
  );
}

export interface ImageTransformOptions {
  width: number;
  height: number;
  offscreenFactory?: OffscreenFactory;
  /** Optional final display target. The source is always transformed at the
   * authored pixel grid; resizing occurs only while baking the result. */
  targetWidthPx?: number;
  targetHeightPx?: number;
  /** Preserve readback failures for the strict chain error boundary. */
  strict?: boolean;
  /** Release readback backing before a spatial/container effect allocates scratch. */
  releaseReadbackSurface?: boolean;
}

/** Run an in-place ImageData transform over a decoded image at its native
 *  pixel grid and bake the result into a new `ImageBitmap`. Returns the
 *  ORIGINAL source when the surface/readback pipeline is unavailable. */
export async function applyImageDataTransform(
  img: CanvasImageSource,
  transform: (data: ImageData) => void | EffectBuffer,
  opts: ImageTransformOptions,
): Promise<CanvasImageSource> {
  const { width, height } = opts;
  if (width <= 0 || height <= 0) return img;
  if (typeof createImageBitmap === 'undefined') return img;
  const factory = opts.offscreenFactory ?? defaultOffscreenFactory;
  const surface = factory(width, height);
  if (!surface) return img;
  let completed = false;
  try {
    const ctx = surface.getContext('2d');
    if (!ctx) return img;
    ctx.drawImage(img, 0, 0);
    let data: ImageData;
    try {
      data = ctx.getImageData(0, 0, width, height);
    } catch (error) {
      if (opts.strict) throw error;
      // Tainted canvas / unsupported readback — leave the picture unrecoloured.
      return img;
    }
    const mutable = surface as { width: number; height: number };
    if (opts.releaseReadbackSurface) mutable.width = 0;
    let result = transform(data);
    const output = { width: result?.width ?? width, height: result?.height ?? height };
    const frame = result?.frame;
    if (result && result !== data) {
      if (!opts.releaseReadbackSurface) mutable.width = 0;
      if (typeof ImageData !== 'undefined') {
        data = new ImageData(result.data as Uint8ClampedArray<ArrayBuffer>, result.width, result.height);
      } else {
        // Canvas implementations without the DOM constructor supply the same
        // generic allocation through their 2D context. Allocate/copy while the
        // readback backing store is released, then drop the kernel output before
        // restoring the canvas: at most four full RGBA surfaces coexist.
        if (!ctx.createImageData) throw new Error('ImageData allocation is unavailable');
        const written = ctx.createImageData(result.width, result.height);
        written.data.set(result.data); data = written;
      }
      result = undefined;
    }
    if (opts.releaseReadbackSurface || output.width !== width || output.height !== height || mutable.width === 0) {
      mutable.width = output.width; mutable.height = output.height;
    }
    ctx.putImageData(data, 0, 0);
    const resizeOptions = decodedBitmapTargetResizeOptions(
      output.width,
      output.height,
      opts.targetWidthPx,
      opts.targetHeightPx,
    );
    const bitmap = resizeOptions
      ? await createImageBitmap(surface, resizeOptions)
      : await createImageBitmap(surface);
    if (frame) {
      const sx = bitmap.width / output.width, sy = bitmap.height / output.height;
      setImageEffectFrame(bitmap, { x: frame.x * sx, y: frame.y * sy,
        width: frame.width * sx, height: frame.height * sy });
    }
    completed = true;
    return bitmap;
  } finally {
    // Failed readback, evaluation or native bake owns no result. Release its
    // backing store immediately; do not retain a discarded full-size surface
    // until GC. Successful Node shims may return the surface itself as bitmap.
    if (!completed) (surface as { width: number }).width = 0;
  }
}

import { OoxmlError } from '../errors/ooxml-error.js';
import { imageEffectFrame } from './image-effect-frame.js';
// Shared `<a:srcRect>` source-rectangle crop (ECMA-376 §20.1.8.55) for the docx,
// pptx and xlsx renderers. The crop is a fraction of the image's NATIVE pixel
// grid. Browser rasters may use a display-sized grid because the same fractional
// mapping is preserved under axis-wise resampling; metafiles must rasterize the
// whole picture FRAME (see `metafileRasterSize`) because their player maps that
// authored frame onto the output before the crop. Centralised here so all three
// renderers crop identically (previously triplicated, and metafiles diverged).

import { isMetafileMime } from './wmf';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** A `<a:srcRect>` crop: signed fractional insets measured from each edge.
 *  The visible region is `[l, t, 1−r, 1−b]` of the source. `ST_Percentage` is
 *  not range-limited, so negative and greater-than-one fractions are retained. */
export interface SrcRect {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** Whether the logical source rectangle intersects the bitmap at positive
 * area. Shared by decode preflight and paint so a fully cropped blip never
 * allocates an oversized fallback raster. */
export function srcRectHasVisibleArea(srcRect: SrcRect | null | undefined): boolean {
  if (!srcRect) return true;
  const values = [srcRect.l, srcRect.t, srcRect.r, srcRect.b];
  if (!values.every(Number.isFinite)) return false;
  const x0 = srcRect.l;
  const y0 = srcRect.t;
  const x1 = 1 - srcRect.r;
  const y1 = 1 - srcRect.b;
  return x1 > x0 && y1 > y0
    && Math.min(1, x1) > Math.max(0, x0)
    && Math.min(1, y1) > Math.max(0, y0);
}

/** Native pixel size of a decoded image (ImageBitmap exposes `width`/`height`;
 *  an `<img>` element exposes `naturalWidth`/`naturalHeight`). */
export function imageNaturalSize(img: CanvasImageSource): { w: number; h: number } {
  const el = img as {
    naturalWidth?: number;
    naturalHeight?: number;
    width?: number;
    height?: number;
  };
  const w = el.naturalWidth || (typeof el.width === 'number' ? el.width : 0) || 0;
  const h = el.naturalHeight || (typeof el.height === 'number' ? el.height : 0) || 0;
  return { w, h };
}

/** Immutable normalized source intersection and destination projection. */
export interface ImageCropProjection {
  readonly sourceX: number; readonly sourceY: number;
  readonly sourceWidth: number; readonly sourceHeight: number;
  readonly dxFraction: number; readonly dyFraction: number;
  readonly dwFraction: number; readonly dhFraction: number;
}

interface CropMapping {
  sx: number; sy: number; sw: number; sh: number;
  dxFraction: number; dyFraction: number; dwFraction: number; dhFraction: number;
}

/** One §20.1.8.55 calculation for every image paint. Non-finite fractions
 * mean no crop; a crop with no positive logical extent paints nothing. */
function imageCropProjection(
  srcRect: SrcRect | null | undefined,
): Readonly<ImageCropProjection> | null {
  if (!srcRect) return null;
  const values = [srcRect.l, srcRect.t, srcRect.r, srcRect.b];
  if (!values.every(Number.isFinite)) return null;
  if (!(srcRect.l || srcRect.t || srcRect.r || srcRect.b)) return null;
  const logicalX0 = srcRect.l, logicalY0 = srcRect.t;
  const logicalX1 = 1 - srcRect.r, logicalY1 = 1 - srcRect.b;
  const logicalW = logicalX1 - logicalX0, logicalH = logicalY1 - logicalY0;
  if (!(logicalW > 0) || !(logicalH > 0)) return Object.freeze({
    sourceX: 0, sourceY: 0, sourceWidth: 0, sourceHeight: 0,
    dxFraction: 0, dyFraction: 0, dwFraction: 0, dhFraction: 0,
  });
  const sourceX0 = Math.max(0, logicalX0), sourceY0 = Math.max(0, logicalY0);
  const sourceX1 = Math.min(1, logicalX1), sourceY1 = Math.min(1, logicalY1);
  const sourceW = Math.max(0, sourceX1 - sourceX0), sourceH = Math.max(0, sourceY1 - sourceY0);
  const plan = Object.freeze({
    sourceX: sourceX0, sourceY: sourceY0, sourceWidth: sourceW, sourceHeight: sourceH,
    dxFraction: (sourceX0 - logicalX0) / logicalW,
    dyFraction: (sourceY0 - logicalY0) / logicalH,
    dwFraction: sourceW / logicalW, dhFraction: sourceH / logicalH,
  });
  return plan;
}

/** {@link imageCropProjection} for effect and projected-surface frames, where
 * an unrepresentable crop must fail closed instead of silently meaning "no
 * crop". Rejects non-finite fractions and intermediate overflow (for example
 * l = r = -1e308 yields an infinite logical width behind finite fractions). */
export function checkedImageCropProjection(
  srcRect: SrcRect | null | undefined,
): Readonly<ImageCropProjection> | null {
  if (srcRect) {
    const { l, t, r, b } = srcRect;
    if (![l, t, r, b, 1 - r - l, 1 - b - t].every(Number.isFinite)) {
      throw new RangeError('Image crop is not representable');
    }
  }
  const plan = imageCropProjection(srcRect);
  if (plan && !Object.values(plan).every(Number.isFinite)) throw new RangeError('Image crop projection overflow');
  return plan;
}

/** The decoded grid supplies pixels only, never retained layout geometry. */
export function cropSourceMapping(
  img: CanvasImageSource, srcRect: SrcRect | null | undefined,
  projection?: Readonly<ImageCropProjection> | null,
): CropMapping | null {
  const plan = projection === undefined ? imageCropProjection(srcRect) : projection;
  if (!plan) return null;
  const { w, h } = imageNaturalSize(img);
  if (w <= 0 || h <= 0) return null;
  return {
    sx: plan.sourceX * w, sy: plan.sourceY * h,
    sw: plan.sourceWidth * w, sh: plan.sourceHeight * h,
    dxFraction: plan.dxFraction, dyFraction: plan.dyFraction,
    dwFraction: plan.dwFraction, dhFraction: plan.dhFraction,
  };
}

export interface ImagePaintTransform {
  readonly rotation?: number; readonly flipH?: boolean; readonly flipV?: boolean; readonly alpha?: number;
}
export interface ImagePaintProjection {
  readonly width: number; readonly height: number;
  readonly transformed: boolean;
  readonly centerX: number; readonly centerY: number; readonly rotationRad: number;
  readonly scaleX: 1 | -1; readonly scaleY: 1 | -1; readonly alpha: number;
  readonly destination: Readonly<{ x: number; y: number; width: number; height: number }>;
  readonly crop: Readonly<ImageCropProjection> | null;
}
export interface ImagePaintProjectionOptions {
  readonly transform?: Readonly<ImagePaintTransform>;
}

/** Canonical local-coordinate crop/transform projection. This is a pure
 * drawing primitive, not a layout or contour algorithm. */
export function imagePaintProjection(
  width: number, height: number, srcRect: SrcRect | null | undefined,
  transform: Readonly<ImagePaintTransform> = {},
): Readonly<ImagePaintProjection> {
  const rotation = transform.rotation ?? 0;
  const rotationProduct = rotation * Math.PI;
  const rotationRad = rotationProduct / 180;
  const transformed = rotation !== 0 || Boolean(transform.flipH) || Boolean(transform.flipV);
  const centerX = width / 2, centerY = height / 2;
  const crop = imageCropProjection(srcRect);
  const destination = Object.freeze({ x: transformed ? -centerX : 0, y: transformed ? -centerY : 0, width, height });
  return Object.freeze({ width, height, transformed, centerX, centerY, rotationRad,
    scaleX: transform.flipH ? -1 : 1, scaleY: transform.flipV ? -1 : 1,
    alpha: transform.alpha ?? 1, destination, crop });
}

/** Apply the projection's drawing transforms only. No recorder, bounds or
 * measurement result is produced. */
export function withImagePaintProjection(
  ctx: AnyCtx, x: number, y: number, plan: Readonly<ImagePaintProjection>,
  draw: (x: number, y: number, width: number, height: number) => void,
): void {
  const hasAlpha = plan.alpha < 1;
  if (hasAlpha) { ctx.save(); ctx.globalAlpha *= plan.alpha; }
  try {
    if (!plan.transformed) {
      draw(x + plan.destination.x, y + plan.destination.y, plan.width, plan.height);
    } else {
      ctx.save();
      try {
        ctx.translate(x + plan.centerX, y + plan.centerY);
        ctx.rotate(plan.rotationRad); ctx.scale(plan.scaleX, plan.scaleY);
        draw(plan.destination.x, plan.destination.y, plan.width, plan.height);
      } finally { ctx.restore(); }
    }
  } finally { if (hasAlpha) ctx.restore(); }
}

/** Draw `img` into the destination box `[dx, dy, dw, dh]`, honoring an optional
 *  `<a:srcRect>` crop. The destination box is unchanged — the visible slice is
 *  stretched to fill it (the 9-arg `drawImage` behavior). Crop applies to raster
 *  blips AND metafiles alike: a cropped metafile must have been rasterized at its
 *  full frame via {@link metafileRasterSize}, so its bitmap is the full source. */
export function drawImageCropped(
  ctx: AnyCtx, img: CanvasImageSource, srcRect: SrcRect | null | undefined,
  dx: number, dy: number, dw: number, dh: number,
  options?: Readonly<ImagePaintProjectionOptions>,
): void {
  const paint = (x: number, y: number, width: number, height: number, projection?: Readonly<ImageCropProjection> | null) => {
    const frame = imageEffectFrame(img);
    if (frame) {
      // §20.1.8.15 grow changes effect coverage, never the authored picture
      // frame. A source crop still selects that frame; only an uncropped
      // picture paints the Gaussian's transparent outset beyond its bounds.
      const { w, h } = imageNaturalSize(img);
      let crop: Readonly<ImageCropProjection> | null;
      try { crop = projection === undefined ? checkedImageCropProjection(srcRect) : projection; }
      catch { throw new OoxmlError('unsupported-image-effect', 'Unrepresentable image effect crop'); }
      const draw = (...coordinates: [number, number, number, number] | [number, number, number, number, number, number, number, number]) => {
        // Canvas may silently ignore non-finite arguments. The effect frame
        // must not turn a malformed crop/overflow into a missing picture.
        if (!coordinates.every(Number.isFinite)) throw new OoxmlError('unsupported-image-effect', 'Unrepresentable image effect paint projection');
        if (coordinates.length === 4) ctx.drawImage(img, ...coordinates);
        else ctx.drawImage(img, ...coordinates);
      };
      if (!crop) {
        draw(x - frame.x / frame.width * width, y - frame.y / frame.height * height,
          w / frame.width * width, h / frame.height * height);
      } else if (crop.sourceWidth > 0 && crop.sourceHeight > 0) {
        draw(frame.x + crop.sourceX * frame.width, frame.y + crop.sourceY * frame.height,
          crop.sourceWidth * frame.width, crop.sourceHeight * frame.height,
          x + crop.dxFraction * width, y + crop.dyFraction * height,
          crop.dwFraction * width, crop.dhFraction * height);
      }
      return;
    }
    const c = cropSourceMapping(img, srcRect, projection);
    if (!c) ctx.drawImage(img, x, y, width, height);
    else if (c.sw > 0 && c.sh > 0 && c.dwFraction > 0 && c.dhFraction > 0) {
      ctx.drawImage(img, c.sx, c.sy, c.sw, c.sh,
        x + c.dxFraction * width, y + c.dyFraction * height, c.dwFraction * width, c.dhFraction * height);
    }
  };
  // Keep all existing seven-argument DOCX/XLSX/PPTX calls exactly on their
  // crop-only path; an authored transform is an explicit eighth argument.
  if (!options) { paint(dx, dy, dw, dh); return; }
  const plan = imagePaintProjection(dw, dh, srcRect, options.transform);
  withImagePaintProjection(ctx, dx, dy, plan, (x, y, width, height) => paint(x, y, width, height, plan.crop));
}

/**
 * Required full-source raster resolution for a destination measured in device
 * pixels. A positive `<a:srcRect>` crop magnifies a source slice to fill the
 * destination, so the full decoded source needs proportionally more pixels;
 * negative insets create transparent outsets and therefore need fewer.
 *
 * The result is a decode request, not an allocation promise. The decoder maps
 * the full source grid to both requested axes (the same resampling the final
 * destination draw would otherwise perform) and applies the shared ceiling.
 */
export function sourceRasterTargetSize(
  destinationWidthPx: number,
  destinationHeightPx: number,
  srcRect?: SrcRect | null,
): { width: number; height: number } | null {
  if (!Number.isFinite(destinationWidthPx) || !Number.isFinite(destinationHeightPx)) return null;
  if (!(destinationWidthPx > 0) || !(destinationHeightPx > 0)) return null;
  const logicalWidth = srcRect ? 1 - srcRect.l - srcRect.r : 1;
  const logicalHeight = srcRect ? 1 - srcRect.t - srcRect.b : 1;
  if (!Number.isFinite(logicalWidth) || !Number.isFinite(logicalHeight)) return null;
  if (!(logicalWidth > 0) || !(logicalHeight > 0) || !srcRectHasVisibleArea(srcRect)) return null;
  const width = Math.ceil(destinationWidthPx / logicalWidth);
  const height = Math.ceil(destinationHeightPx / logicalHeight);
  return Number.isFinite(width) && Number.isFinite(height)
    ? { width, height }
    : null;
}

/** Raster target size (pt) for decoding an embedded image. A browser raster's
 *  retained pixel target is handled separately, so its point box passes through. A metafile
 *  (WMF/EMF) with an `<a:srcRect>` crop must be rasterized at its FULL picture
 *  frame, not the visible sub-rectangle — the player maps the frame onto the
 *  raster (see `playEmf`), and the crop is relative to that frame. Scale the box
 *  up by `1/(1−l−r)` and `1/(1−t−b)` so the rasterised frame and the fractional
 *  crop align (e.g. one composite EMF cropped into subfigures). Uncropped
 *  metafiles and all raster blips pass the box through unchanged. */
export function metafileRasterSize(
  mimeType: string,
  srcRect: SrcRect | null | undefined,
  widthPt: number,
  heightPt: number,
): { widthPt: number; heightPt: number } | null {
  if (!srcRect || !isMetafileMime(mimeType)) return { widthPt, heightPt };
  if (!srcRectHasVisibleArea(srcRect)) return null;
  const fracW = 1 - srcRect.l - srcRect.r;
  const fracH = 1 - srcRect.t - srcRect.b;
  return { widthPt: widthPt / fracW, heightPt: heightPt / fracH };
}

/** Physical extent of the whole (uncropped) source of a stretched blip fill
 * (§20.1.8.55/.56): the fillRect insets shrink the destination, and the
 * srcRect window is stretched across it, so the full source spans the
 * destination divided by the visible source fraction. CT_Blip effects that
 * are measured in EMUs (§20.1.8.15) are evaluated against this frame, the
 * same mapping core uses for nested stretched fills. */
export function stretchedSourceFrame(
  width: number, height: number,
  srcRect: SrcRect | null | undefined,
  fillRect: Readonly<{ l?: number; t?: number; r?: number; b?: number }> | null | undefined,
): { width: number; height: number } {
  const destW = width * (1 - (fillRect?.l ?? 0) - (fillRect?.r ?? 0));
  const destH = height * (1 - (fillRect?.t ?? 0) - (fillRect?.b ?? 0));
  const cropW = 1 - (srcRect?.l ?? 0) - (srcRect?.r ?? 0);
  const cropH = 1 - (srcRect?.t ?? 0) - (srcRect?.b ?? 0);
  return {
    width: Math.max(0, cropW > 0 ? destW / cropW : destW),
    height: Math.max(0, cropH > 0 ? destH / cropH : destH),
  };
}

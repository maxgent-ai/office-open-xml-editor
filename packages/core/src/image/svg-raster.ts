// Shared SVG rasterization primitive; worker transport lives in svg-decode-bridge.
import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import { closeImageBitmapIfSupported } from './image-bitmap-lifecycle.js';

export interface SvgDecodeTarget {
  /** Minimum retained source-grid width in device pixels. */
  readonly targetWidthPx?: number;
  /** Minimum retained source-grid height in device pixels. */
  readonly targetHeightPx?: number;
}

export type SvgBlobDecoder = (
  blob: Blob,
  target?: SvgDecodeTarget,
) => Promise<ImageBitmap>;

function finitePositive(value: number | undefined): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? Math.ceil(value)
    : undefined;
}

/** Bound a requested SVG raster surface to the same per-image hard quota used
 * by ordinary decoded OOXML images. Both axes scale together, preserving the
 * authored vector aspect ratio when a host asks for an oversized target. */
export function boundedSvgRasterSize(width: number, height: number): { width: number; height: number } {
  if (!Number.isFinite(width) || !Number.isFinite(height) || !(width > 0) || !(height > 0)) {
    throw new Error(`invalid SVG raster size: ${width}x${height}`);
  }
  // Divide the square roots separately: `width * height` can overflow to
  // Infinity for finite hostile inputs, which would collapse the scale to zero
  // and produce NaN dimensions after multiplication.
  const pixelScale = Math.sqrt(MAX_RASTER_PIXELS) / Math.sqrt(width) / Math.sqrt(height);
  const scale = Math.min(
    1,
    MAX_RASTER_DIMENSION / width,
    MAX_RASTER_DIMENSION / height,
    pixelScale,
  );
  return {
    width: Math.min(MAX_RASTER_DIMENSION, Math.max(1, Math.floor(width * scale))),
    height: Math.min(MAX_RASTER_DIMENSION, Math.max(1, Math.floor(height * scale))),
  };
}

/** Smallest bounded source grid that covers the requested axes without changing
 * the SVG's intrinsic aspect ratio. Target axes are coverage requirements, as
 * they are for ordinary raster decode; they are not a replacement aspect ratio.
 * Consequently the grid is not the target box: a square SVG requested at 96×87
 * yields 96×96. Retained-surface ceilings for an SVG raster must be derived from
 * this grid (see `assertSvgRasterCoverage`), never from `width × height` of the
 * request. */
function svgRasterCoverageSize(
  sourceWidth: number,
  sourceHeight: number,
  targetWidth: number | undefined,
  targetHeight: number | undefined,
): { width: number; height: number } {
  const widthScale = targetWidth === undefined ? 0 : targetWidth / sourceWidth;
  const heightScale = targetHeight === undefined ? 0 : targetHeight / sourceHeight;
  const requiredScale = targetWidth === undefined && targetHeight === undefined
    ? 1
    : Math.max(widthScale, heightScale);
  const quotaScale = Math.min(
    MAX_RASTER_DIMENSION / sourceWidth,
    MAX_RASTER_DIMENSION / sourceHeight,
    Math.sqrt(MAX_RASTER_PIXELS) / Math.sqrt(sourceWidth) / Math.sqrt(sourceHeight),
  );
  if (!Number.isFinite(requiredScale) || !(requiredScale > 0)
    || !Number.isFinite(quotaScale) || !(quotaScale > 0)) {
    throw new Error(`invalid SVG raster scale: ${Math.min(requiredScale, quotaScale)}`);
  }
  if (quotaScale < requiredScale) {
    // The quota, rather than a caller axis, dominates. Round inward so the
    // retained surface cannot cross that quota.
    return boundedSvgRasterSize(
      Math.max(1, Math.floor(sourceWidth * quotaScale)),
      Math.max(1, Math.floor(sourceHeight * quotaScale)),
    );
  }
  // Assign the scale-dominating requested axis directly before deriving the
  // other. This avoids floating roundoff growing a 400px request to 401px.
  if (targetWidth !== undefined && widthScale >= heightScale) {
    return boundedSvgRasterSize(
      targetWidth,
      Math.max(1, Math.ceil(sourceHeight * targetWidth / sourceWidth)),
    );
  }
  if (targetHeight !== undefined) {
    return boundedSvgRasterSize(
      Math.max(1, Math.ceil(sourceWidth * targetHeight / sourceHeight)),
      targetHeight,
    );
  }
  return boundedSvgRasterSize(Math.ceil(sourceWidth), Math.ceil(sourceHeight));
}

/** Fail closed when a decoded SVG raster is larger than the coverage grid its
 * request admits. The admitted grid is `svgRasterCoverageSize` of the
 * surface's own aspect ratio, which a conforming decoder already returned:
 * one requested axis matches its target and the other follows the authored
 * aspect, or both shrink under the per-image quota. A surface that exceeds
 * every requested axis was over-rasterized (for example by an injected
 * decoder) and is rejected here instead of being retained. The hard per-image
 * dimension/pixel quota is enforced separately by the caller. */
export function assertSvgRasterCoverage(
  width: number,
  height: number,
  target: SvgDecodeTarget,
): void {
  const targetWidth = finitePositive(target.targetWidthPx);
  const targetHeight = finitePositive(target.targetHeightPx);
  if (targetWidth === undefined && targetHeight === undefined) return;
  if (!(width > 0) || !(height > 0)) return;
  const exceedsWidth = targetWidth === undefined || width > targetWidth;
  const exceedsHeight = targetHeight === undefined || height > targetHeight;
  if (!exceedsWidth || !exceedsHeight) return;
  const admitted = svgRasterCoverageSize(width, height, targetWidth, targetHeight);
  throw new OoxmlDecodedImageLimitError(
    'image-pixels',
    admitted.width * admitted.height,
    width * height,
  );
}

/** Decode SVG with the Window's standards-compliant image pipeline, then
 * transfer a bounded, display-sized ImageBitmap back to the render worker.
 * Dedicated workers cannot rely on `createImageBitmap(svgBlob)` in Chromium. */
export async function decodeSvgBlobOnMainThread(
  blob: Blob,
  target: SvgDecodeTarget = {},
): Promise<ImageBitmap> {
  if (typeof Image === 'undefined') throw new Error('SVG host decode requires HTMLImageElement');
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image();
    const targetWidth = finitePositive(target.targetWidthPx);
    const targetHeight = finitePositive(target.targetHeightPx);
    await new Promise<void>((resolve, reject) => {
      image.onload = () => {
        if (typeof image.decode === 'function') image.decode().then(resolve).catch(resolve);
        else resolve();
      };
      image.onerror = () => reject(new Error('SVG host decode failed'));
      image.src = url;
    });
    // A viewBox-only SVG receives the HTML default object size (300×150), with
    // its viewBox ratio reflected in naturalWidth/naturalHeight. Truly
    // dimensionless sources can still report zero; use that same default box so
    // the explicit Canvas target below remains drawable and deterministic.
    const hasNaturalSize = Number.isFinite(image.naturalWidth) && image.naturalWidth > 0
      && Number.isFinite(image.naturalHeight) && image.naturalHeight > 0;
    const sourceWidth = hasNaturalSize ? image.naturalWidth : 300;
    const sourceHeight = hasNaturalSize ? image.naturalHeight : 150;
    const size = svgRasterCoverageSize(sourceWidth, sourceHeight, targetWidth, targetHeight);
    image.width = size.width;
    image.height = size.height;
    // Rasterize explicitly instead of createImageBitmap(image): Chromium can
    // return a transparent bitmap for a viewBox-only SVG even after the image
    // loaded successfully. Worker render mode already requires OffscreenCanvas;
    // transferToImageBitmap moves that backing store without retaining a second
    // full-size surface on Window.
    let bitmap: ImageBitmap;
    if (typeof OffscreenCanvas !== 'undefined') {
      const canvas = new OffscreenCanvas(size.width, size.height);
      const context = canvas.getContext('2d');
      if (!context) throw new Error('SVG host raster target has no 2-D context');
      context.drawImage(image, 0, 0, size.width, size.height);
      bitmap = canvas.transferToImageBitmap();
    } else {
      const canvas = document.createElement('canvas');
      canvas.width = size.width;
      canvas.height = size.height;
      const context = canvas.getContext('2d');
      if (!context) throw new Error('SVG host raster target has no 2-D context');
      context.drawImage(image, 0, 0, size.width, size.height);
      bitmap = await createImageBitmap(canvas);
    }
    const pixels = bitmap.width * bitmap.height;
    if (
      bitmap.width > MAX_RASTER_DIMENSION
      || bitmap.height > MAX_RASTER_DIMENSION
      || pixels > MAX_RASTER_PIXELS
    ) {
      closeImageBitmapIfSupported(bitmap);
      throw new OoxmlDecodedImageLimitError('image-pixels', MAX_RASTER_PIXELS, pixels);
    }
    return bitmap;
  } finally {
    URL.revokeObjectURL(url);
  }
}

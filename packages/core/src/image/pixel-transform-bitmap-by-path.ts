// DrawingML `<a:duotone>` (§20.1.8.23) adapter shared by all three formats.
// Base and recoloured surfaces live in one per-document weighted decoded owner,
// so the expensive transform is memoized without a parallel lifecycle or byte-
// accounting system.

import {
  captureDecodedBitmapCacheEpoch,
  dropCachedDerivedBitmapNamespace,
  getCachedBitmapByPath,
  getCachedDerivedBitmap,
  resolvedCachedBitmapVariantKey,
  type CachedBitmapOptions,
} from './bitmap-image-by-path';
import {
  type OffscreenFactory,
} from './duotone';
import { imageEffectError } from './image-effect-error.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import { imageNaturalSize } from './crop';
import { MAX_IMAGE_EFFECT_BASE_PIXELS, MAX_RASTER_PIXELS } from './pixel-budget.js';
import { decodedBitmapTargetResizeOptions } from './raster-target.js';
import { carryIncompleteMetafileReport } from './raster-or-metafile.js';

type FetchImage = (path: string, mime: string) => Promise<Blob>;

/** Shared ownership/cache machinery, independent of the pixel algorithm. */
export interface BitmapPixelTransform {
  readonly key: string;
  readonly strict: boolean;
  preflight(pixels: number, width?: number, height?: number): void;
  prepare?(base: ImageBitmap): Promise<void>;
  apply(base: ImageBitmap, options: import('./duotone.js').ImageTransformOptions): Promise<CanvasImageSource>;
}

const DUOTONE_CACHE_NAMESPACE = 'duotone';

/**
 * Decode a raster, metafile or SVG blip at `imagePath` and, when `duotone` is set,
 * recolour it along the `clr1`→`clr2` luminance ramp (§20.1.8.23), returning a
 * drawable source cached per document then by `duotoneCacheKey`.
 *
 * With NO duotone this is a thin pass-through to {@link getCachedBitmapByPath}
 * (the shared base cache) — no second-layer entry is created, so a non-duotone
 * picture behaves byte-for-byte as before. With a duotone the base bitmap is
 * decoded (and cached) once, then the recolour runs once per colour pair and is
 * memoized here.
 *
 * The recolour needs a readable pixel grid, so it only applies to a decoded
 * bitmap (SVG requires `opts.svgDecoder`); a `null` base (an unsupported
 * metafile — true EMF / geometry-less WMF) propagates as `null` and the draw
 * site skips it. When the
 * offscreen pixel pipeline is unavailable, {@link applyDuotone} returns the base
 * unchanged, so the picture still draws (just without the recolour).
 *
 * @param opts.offscreenFactory optional surface factory for environments without
 *   a global `OffscreenCanvas` (node); forwarded to {@link applyDuotone}.
 */
export async function getCachedPixelTransformBitmapByPath(
  imagePath: string,
  mimeType: string,
  duotone: BitmapPixelTransform | null | undefined,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions & {
    offscreenFactory?: OffscreenFactory;
    /** When true, an authored duotone that cannot be applied returns null
     * instead of silently drawing the original pixels. Chart picture markers
     * use this fail-closed mode; established shape consumers retain their
     * legacy fallback unless they opt in. */
    failClosedOnDuotoneFailure?: boolean;
  } = {},
): Promise<ImageBitmap | null> {
  // A complete chain owns its entire acquisition lifecycle: preflight,
  // source/dependency decode (including SVG), preparation, fill rasterization,
  // and bitmap bake. Host broken-image fallbacks may swallow native errors;
  // normalize here before any such fallback can turn failure into an omission.
  try {
    const { offscreenFactory, failClosedOnDuotoneFailure = false, ...requestedBitmapOpts } = opts;
    const bitmapOpts = duotone
      ? {
          ...requestedBitmapOpts,
          // Peak pipeline: source ImageBitmap + offscreen backing + ImageData +
          // result ImageBitmap. Bound the base to one quarter of the document
          // byte ceiling so transient pixel work cannot silently double it.
          maxRetainedPixels: Math.min(
            requestedBitmapOpts.maxRetainedPixels ?? MAX_RASTER_PIXELS,
            MAX_IMAGE_EFFECT_BASE_PIXELS,
          ),
        }
      : requestedBitmapOpts;
    // For DrawingML pixel effects (§20.1.8.23), keep authored raster pixels
    // until display resampling. SVG has no authored pixel grid: retain the
    // full-source display/crop coverage target so its decoder can preserve the
    // intrinsic aspect ratio. The resolved base grid keys each effect separately.
    // This sampling choice and the effect-base ceiling are library policy;
    // §20.1.8.23 defines the pixel transform, not vector sampling resolution.
    // Both paths reject an over-budget effect base rather than approximating it.
    const sourceBitmapOpts = duotone && mimeType !== 'image/svg+xml'
      ? { ...bitmapOpts, targetWidthPx: undefined, targetHeightPx: undefined }
      : bitmapOpts;
    // An over-long effect list is rejected before anything is fetched or
    // decoded; the cumulative pixel work is checked once the grid is known.
    duotone?.preflight(0);
    const epoch = duotone
      ? captureDecodedBitmapCacheEpoch(fetchImage, DUOTONE_CACHE_NAMESPACE)
      : undefined;
    // Base, colour-free bitmap from the shared path-keyed cache.
    const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, sourceBitmapOpts);
    // No duotone → return the base directly (no second-layer entry). A `null`
    // (unsupported metafile) propagates unchanged.
    if (!duotone || !base) return base;
    duotone.preflight(base.width * base.height, base.width, base.height);
    // Strict and compatibility callers must not share a derived cache entry: a
    // compatibility pass-through must never make a later strict lookup succeed.
    const resolvedBaseKey = await resolvedCachedBitmapVariantKey(
      imagePath,
      mimeType,
      fetchImage,
      sourceBitmapOpts,
      epoch,
      base,
    );
    const resizeOptions = decodedBitmapTargetResizeOptions(
      Number(base.width),
      Number(base.height),
      requestedBitmapOpts.targetWidthPx,
      requestedBitmapOpts.targetHeightPx,
    );
    const resizeKey = resizeOptions
      ? `|resize:${resizeOptions.resizeWidth}x${resizeOptions.resizeHeight}`
      : '';
    // Tuple boundaries remain unambiguous even when paths/effect recipes
    // contain punctuation used by the individual cache namespaces.
    const key = JSON.stringify([resolvedBaseKey, duotone.key, resizeKey, failClosedOnDuotoneFailure]);
    return await getCachedDerivedBitmap(
      DUOTONE_CACHE_NAMESPACE,
      key,
      fetchImage,
      async () => {
        const { w, h } = imageNaturalSize(base);
        if (w <= 0 || h <= 0) {
          if (duotone.strict) throw new OoxmlError('unsupported-image-effect', 'Invalid image effect source dimensions');
          return { bitmap: failClosedOnDuotoneFailure ? null : base, owned: false };
        }
        // Before the transform allocates its surfaces: passes × pixels must fit
        // the shared effect-work ceiling (a quota error, never a partial list).
        const transformOptions = {
          width: w,
          height: h,
          offscreenFactory,
          targetWidthPx: requestedBitmapOpts.targetWidthPx,
          targetHeightPx: requestedBitmapOpts.targetHeightPx,
          releaseReadbackSurface: duotone.strict,
          strict: duotone.strict,
        };
        // CT_Blip pixel effects (grayscl/biLevel/clrChange, with the duotone at
        // its document position) share the duotone's decode, cache and budget.
        const recoloured = await duotone.apply(base, transformOptions);
        // `applyDuotone` returns a CanvasImageSource; when the pixel pipeline ran
        // it is a fresh ImageBitmap, otherwise it is the unchanged current
        // source. Strict callers fail closed. Compatibility callers still bake a
        // display-sized copy so an unavailable effect surface cannot silently
        // defeat the caller's bounded-resolution request.
        if (recoloured === base) {
          // A complete ordered chain must never degrade to uncorrected pixels
          // when Canvas/readback is unavailable. Keep only the legacy duotone
          // caller's explicit compatibility fallback.
          if (duotone.strict) {
            throw new OoxmlError('unsupported-image-effect', 'Image effects unavailable');
          }
          if (failClosedOnDuotoneFailure) return { bitmap: null, owned: false };
          if (!resizeOptions) return { bitmap: base, owned: false };
          if (typeof createImageBitmap === 'undefined') {
            throw new Error('createImageBitmap is unavailable for duotone fallback resampling');
          }
          const resized = await createImageBitmap(base, resizeOptions);
          carryIncompleteMetafileReport(base, resized);
          return { bitmap: resized, owned: resized !== base };
        }
        const bitmap = recoloured as ImageBitmap;
        // A partially drawn metafile stays partial after an effect or a
        // resample; the derived entry (and every later cache hit on it) keeps
        // the base picture's report.
        carryIncompleteMetafileReport(base, bitmap);
        return { bitmap, owned: bitmap !== base };
      },
      epoch,
      duotone.prepare ? () => (duotone.prepare as (base: ImageBitmap) => Promise<void>)(base) : undefined,
    );
  } catch (error) {
    if (duotone?.strict) throw imageEffectError(error);
    throw error;
  }
}

/** Drop only the duotone namespace. Full document teardown calls the shared
 * decoded-owner drop once. */
export function dropDuotoneBitmapCache(fetchImage: FetchImage): void {
  dropCachedDerivedBitmapNamespace(fetchImage, DUOTONE_CACHE_NAMESPACE);
}

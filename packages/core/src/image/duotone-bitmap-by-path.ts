import { imageEffectError } from './image-effect-error.js';
import { ImageEffectWorkLedger } from './image-effect-work.js';
import { preferVectorBlip } from './blip-gate.js';
import { acquireCurrentBitmapCacheLease } from './bitmap-image-by-path.js';
import { MAX_IMAGE_EFFECT_PIXEL_WORK, OoxmlDecodedImageLimitError } from './pixel-budget.js';
// Ordered DrawingML blip effects for complete CT_Blip chains (DOCX, PPTX and
// XLSX); a lone legacy duotone keeps the algorithm-free duotone-only owner.
import { getCachedPixelTransformBitmapByPath } from './pixel-transform-bitmap-by-path.js';
import { getCachedDuotoneOnlyBitmapByPath, duotoneOnlyCacheKey } from './duotone-only-bitmap-by-path.js';
import { applyImageDataTransform, type Duotone } from './duotone.js';
import { applyBlipPixelEffects, assertBlipPixelEffectsBudget, blipPixelEffectsKey,
  isBlipPixelEffects, blipEffectImageFills, blipPixelEffectWork, type BlipPixelEffects } from './blip-effects.js';
import { OoxmlError } from '../errors/ooxml-error.js';
export { dropDuotoneBitmapCache } from './pixel-transform-bitmap-by-path.js';

export function duotoneCacheKey(path: string, value?: Duotone | BlipPixelEffects | null): string {
  return isBlipPixelEffects(value) ? `${path}|fx:${blipPixelEffectsKey(value)}` : duotoneOnlyCacheKey(path, value);
}
export async function getCachedDuotoneBitmapByPath(
  imagePath: string, mimeType: string, value: Duotone | BlipPixelEffects | null | undefined,
  fetchImage: Parameters<typeof getCachedPixelTransformBitmapByPath>[3],
  opts: Parameters<typeof getCachedPixelTransformBitmapByPath>[4] = {},
): Promise<ImageBitmap | null> {
  if (!isBlipPixelEffects(value)) return getCachedDuotoneOnlyBitmapByPath(imagePath, mimeType, value, fetchImage, opts);
  try {
    assertBlipPixelEffectsBudget(value, 0);
    const release = blipEffectImageFills(value.effects).length ? acquireCurrentBitmapCacheLease(fetchImage) : undefined;
    return await loadEffects(imagePath, mimeType, value, fetchImage, opts, { work: 0, execution: new ImageEffectWorkLedger() }).finally(() => release?.());
  } catch (error) {
    throw imageEffectError(error);
  }
}
// Validation failures must reject the promise so the outer source lease is released.
async function loadEffects(
  imagePath: string, mimeType: string, value: BlipPixelEffects,
  fetchImage: Parameters<typeof getCachedPixelTransformBitmapByPath>[3],
  opts: NonNullable<Parameters<typeof getCachedPixelTransformBitmapByPath>[4]>,
  ledger: { work: number; execution: ImageEffectWorkLedger },
): Promise<ImageBitmap | null> {
  // JSON tuples encode valid finite physical dimensions without decimal
  // ambiguity. Reject invalid numbers before JSON can collapse them to null.
  if (![opts.widthPt ?? 0, opts.heightPt ?? 0].every(v => Number.isFinite(v) && v >= 0)) {
    throw new OoxmlError('unsupported-image-effect', 'Invalid physical image effect dimensions');
  }
  const physical = (width?: number, height?: number): BlipPixelEffects => {
    const effects = { ...value,
      emuPerPixelX: opts.widthPt && width ? opts.widthPt * 12700 / width : value.emuPerPixelX,
      emuPerPixelY: opts.heightPt && height ? opts.heightPt * 12700 / height : value.emuPerPixelY,
    };
    if (![effects.emuPerPixelX ?? 9525, effects.emuPerPixelY ?? 9525].every(v => Number.isFinite(v) && v > 0)) {
      throw new OoxmlError('unsupported-image-effect', 'Unrepresentable physical image effect sampling');
    }
    return effects;
  };
  const images = new Map<object, CanvasImageSource>();
  return getCachedPixelTransformBitmapByPath(imagePath, mimeType, {
    key: `fx:${blipPixelEffectsKey(value)}|physical:${JSON.stringify([opts.widthPt ?? 0, opts.heightPt ?? 0])}`,
    strict: true,
    preflight: (pixels, width, height) => {
      assertBlipPixelEffectsBudget(physical(width, height), pixels, width && height ? { width, height } : undefined);
    },
    prepare: async base => {
      const effects = physical(base.width, base.height);
      for (const fill of blipEffectImageFills(effects.effects)) {
        const vector = !fill.blipEffects?.length && preferVectorBlip(fill);
        const path = vector ? fill.svgImagePath as string : fill.imagePath;
        const mime = vector ? 'image/svg+xml' : fill.mimeType;
        // A nested fill owns its own physical image frame. Tile sampling is
        // set by DPI (§20.1.8.14/58), not by the containing picture's extent.
        // Stretched sources map their crop to fillRect within that extent.
        // Keep nested outputs native-sized; outer display resampling must not
        // change a tile period or a later effect's pixel/physical conversion.
        const nested = { ...effects, effects: fill.blipEffects ?? [], duotone: fill.duotone,
          hostRotation: fill.rotWithShape === false ? 0 : effects.hostRotation };
        const nestedOpts = { ...opts, targetWidthPx: undefined, targetHeightPx: undefined,
          widthPt: 0, heightPt: 0 };
        if (fill.tile) {
          nested.emuPerPixelX = nested.emuPerPixelY = 914400 / (fill.dpi || 96);
        } else {
          const crop = fill.srcRect, rect = fill.fillRect;
          const cw = 1 - (crop?.l ?? 0) - (crop?.r ?? 0), ch = 1 - (crop?.t ?? 0) - (crop?.b ?? 0);
          // An empty crop paints no source. Keep finite sampling for its
          // bounded decode; do not divide by a zero/negative visible extent.
          nestedOpts.widthPt = base.width * (effects.emuPerPixelX ?? 9525) / 12700
            * (1 - (rect?.l ?? 0) - (rect?.r ?? 0)) / (cw > 0 ? cw : 1);
          nestedOpts.heightPt = base.height * (effects.emuPerPixelY ?? 9525) / 12700
            * (1 - (rect?.t ?? 0) - (rect?.b ?? 0)) / (ch > 0 ? ch : 1);
        }
        // A nested authored duotone belongs to this strict chain too; its
        // readback/bake failure cannot use the standalone compatibility fallback.
        const image = fill.blipEffects?.length || fill.duotone
          ? await loadEffects(path, mime, nested, fetchImage, nestedOpts, ledger)
          : await getCachedDuotoneOnlyBitmapByPath(path, mime, fill.duotone, fetchImage, nestedOpts);
        if (!image) throw new OoxmlError('unsupported-image-effect', 'Image fill decode unavailable');
        images.set(fill, image);
      }
      let additionalPixels = 0;
      for (const image of new Set(images.values())) {
        if (image === base) continue;
        const size = image as { width: number; height: number };
        additionalPixels += size.width * size.height;
      }
      assertBlipPixelEffectsBudget(effects, base.width * base.height, base, additionalPixels);
      // Nested transforms charge this same ledger before they paint. Reserve
      // the complete parent plan only once its decoded tile sources are known;
      // this includes every recursive mask, blur tap and fill draw iteration.
      ledger.work += blipPixelEffectWork(effects, base.width, base.height, { images });
      if (ledger.work > MAX_IMAGE_EFFECT_PIXEL_WORK) throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK, ledger.work);
    },
    apply: (base, options) => applyImageDataTransform(base,
      data => applyBlipPixelEffects(data, physical(base.width, base.height), { offscreenFactory: opts.offscreenFactory, images, baseImage: base, workLedger: ledger.execution }), options),
  }, fetchImage, opts);
}

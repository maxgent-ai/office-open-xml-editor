// The lone-duotone path keeps the established duotone cache key and does not
// need the ordered effect interpreter. Both adapters inject their algorithm
// into the same document-owned cache/decode primitive; no second cache or
// error policy.
import { getCachedPixelTransformBitmapByPath } from './pixel-transform-bitmap-by-path.js';
import { applyDuotone, type Duotone } from './duotone.js';

export function duotoneOnlyCacheKey(path: string, duotone?: Duotone | null): string {
  return duotone ? `${path}|duo:${duotone.clr1}:${duotone.clr2}` : path;
}
export function getCachedDuotoneOnlyBitmapByPath(
  imagePath: string, mimeType: string, duotone: Duotone | null | undefined,
  fetchImage: Parameters<typeof getCachedPixelTransformBitmapByPath>[3],
  opts: Parameters<typeof getCachedPixelTransformBitmapByPath>[4] = {},
): Promise<ImageBitmap | null> {
  return getCachedPixelTransformBitmapByPath(imagePath, mimeType, duotone ? {
    key: `duo:${duotone.clr1}:${duotone.clr2}`,
    strict: false,
    preflight() {},
    apply: (base, options) => applyDuotone(base, duotone, options),
  } : undefined, fetchImage, opts);
}

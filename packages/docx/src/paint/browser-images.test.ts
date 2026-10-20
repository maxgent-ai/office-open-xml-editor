import { afterEach, describe, expect, it, vi } from 'vitest';
import { getCachedBitmapByPath, OoxmlDecodedImageLimitError } from '@silurus/ooxml-core';
import { decodeRaster, imageKey, preloadPaintImages } from './browser-images.js';

vi.mock('@silurus/ooxml-core', async importOriginal => ({
  ...await importOriginal<typeof import('@silurus/ooxml-core')>(),
  getCachedBitmapByPath: vi.fn(async () => ({ width: 1, height: 1 })),
  resolvedCachedBitmapVariantKey: vi.fn(async (path: string) => path),
}));

afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });

class PixelCanvas {
  width = 1;
  height = 1;
  pixels = new Uint8ClampedArray([255, 255, 255, 128]);
  getContext() {
    return {
      drawImage() {},
      getImageData: () => ({ data: this.pixels, width: 1, height: 1 }),
      putImageData: (image: ImageData) => { this.pixels = new Uint8ClampedArray(image.data); },
    };
  }
}
const duotone = { clr1: '112233', clr2: 'AABBCC' };
function bitmapFactory() {
  vi.stubGlobal('createImageBitmap', vi.fn(async (canvas: PixelCanvas) => ({
    width: 1, height: 1, pixels: [...canvas.pixels], close() {},
  })));
}

describe('authored DOCX image colour effects', () => {
  it('uses a DOM canvas when OffscreenCanvas is unavailable and preserves alpha', async () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', { createElement: () => new PixelCanvas() });
    bitmapFactory();
    const bitmap = await decodeRaster('picture.png', 'image/png', undefined,
      async () => new Blob(), 1, 1, duotone);
    expect(bitmap).toMatchObject({ pixels: [170, 187, 204, 128] });
  });

  it('reports an unavailable pixel path by rejection instead of returning untransformed pixels', async () => {
    vi.stubGlobal('OffscreenCanvas', class { getContext() { return null; } });
    await expect(decodeRaster('picture.emf', 'image/x-emf', undefined,
      async () => new Blob(), 1, 1, duotone)).rejects.toThrow('2D canvas');
  });

  it('retains duotone when a failed preferred SVG falls back to its metafile source', async () => {
    vi.stubGlobal('OffscreenCanvas', PixelCanvas);
    bitmapFactory();
    vi.mocked(getCachedBitmapByPath).mockRejectedValueOnce(new Error('bad SVG'));
    const decoded = await preloadPaintImages([{
      kind: 'image', resourceKey: 'fallback', partPath: 'picture.emf', mimeType: 'image/x-emf',
      svgImagePath: 'preferred.svg', duotone, intrinsicSize: { widthPt: 1, heightPt: 1 },
    }], [{ resourceKind: 'image', resourceKey: 'fallback', widthPt: 1, heightPt: 1 }],
    async () => new Blob());
    expect(decoded.get(imageKey('picture.emf', undefined, duotone)))
      .toMatchObject({ pixels: [170, 187, 204, 128] });
    expect(getCachedBitmapByPath).toHaveBeenCalledTimes(2);
  });

  it('keeps a preferred SVG quota failure terminal instead of retrying another representation', async () => {
    const failure = new OoxmlDecodedImageLimitError('image-pixels', 1, 2);
    vi.mocked(getCachedBitmapByPath).mockRejectedValueOnce(failure);
    await expect(preloadPaintImages([{
      kind: 'image', resourceKey: 'quota', partPath: 'picture.emf', mimeType: 'image/x-emf',
      svgImagePath: 'preferred.svg', duotone, intrinsicSize: { widthPt: 1, heightPt: 1 },
    }], [{ resourceKind: 'image', resourceKey: 'quota', widthPt: 1, heightPt: 1 }],
    async () => new Blob())).rejects.toBe(failure);
    expect(getCachedBitmapByPath).toHaveBeenCalledTimes(1);
  });

  it.each([
    { partPath: 'picture.svg', mimeType: 'image/svg+xml' },
    { partPath: 'picture.emf', mimeType: 'image/x-emf', svgImagePath: 'preferred.svg' },
    { partPath: 'picture.wmf', mimeType: 'image/x-wmf' },
  ])('applies duotone to $mimeType through the same pixel pipeline', async descriptor => {
    vi.stubGlobal('OffscreenCanvas', PixelCanvas);
    bitmapFactory();
    const fetch = async () => new Blob();
    const svgDecoder = vi.fn();
    const decoded = await preloadPaintImages([{
      ...descriptor, kind: 'image', resourceKey: 'image', duotone,
      intrinsicSize: { widthPt: 1, heightPt: 1 },
    }], [{ resourceKind: 'image', resourceKey: 'image', widthPt: 1, heightPt: 1 }],
    fetch, undefined, 1, svgDecoder);
    expect(decoded.get(imageKey(descriptor.partPath, undefined, duotone)))
      .toMatchObject({ pixels: [170, 187, 204, 128] });
    expect(getCachedBitmapByPath).toHaveBeenCalledWith(
      descriptor.svgImagePath ?? descriptor.partPath,
      descriptor.svgImagePath ? 'image/svg+xml' : descriptor.mimeType,
      fetch, expect.objectContaining({ svgDecoder }),
    );
  });
});

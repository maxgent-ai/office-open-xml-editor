import { afterEach, expect, it, vi } from 'vitest';
import { MAX_IMAGE_EFFECT_BASE_PIXELS } from '@silurus/ooxml-core';
import { decodeRaster, preloadPaintImages } from './browser-images.js';

// Exercise the production preload/decode/cache/effect path in both realms.
// Blob decoding deliberately rejects SVG, as Chromium's createImageBitmap does.
afterEach(() => vi.unstubAllGlobals());
it.each(['main', 'worker'] as const)('rasterizes SVG through the supported %s decoder before lum', async (realm) => {
  const written: number[][] = [];
  const bitmap = () => ({ width: 2, height: 1, close() {} }) as ImageBitmap;
  class Surface {
    constructor(readonly width: number, readonly height: number) {}
    transferToImageBitmap() { return bitmap(); }
    getContext() {
      return {
        drawImage() {},
        getImageData: () => ({ width: 2, height: 1, data: new Uint8ClampedArray([0, 0, 0, 255, 0, 0, 0, 255]) }),
        putImageData: (data: ImageData) => written.push([...data.data]),
      };
    }
  }
  vi.stubGlobal('OffscreenCanvas', Surface);
  vi.stubGlobal('createImageBitmap', vi.fn(async (source: unknown) => {
    if (source instanceof Blob) throw new Error('SVG Blob decode unsupported');
    return bitmap();
  }));
  const workerDecoder = vi.fn(async () => bitmap());
  if (realm === 'main') {
    vi.stubGlobal('Image', class {
      naturalWidth = 2; naturalHeight = 1;
      onload?: () => void;
      set src(_value: string) { this.onload?.(); }
    });
  } else {
    vi.stubGlobal('Image', undefined);
  }
  const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
  const images = await preloadPaintImages([{
    kind: 'image', resourceKey: 'picture', partPath: 'word/media/picture.svg',
    mimeType: 'image/svg+xml', intrinsicSize: { widthPt: 2, heightPt: 1 },
    blipEffects: [{ type: 'luminance', bright: .9, contrast: 0 }],
  }], [{
    resourceKind: 'image', resourceKey: 'picture', widthPt: 2, heightPt: 1,
  }], fetchImage, undefined, 1, realm === 'worker' ? workerDecoder : undefined);
  expect(images.size).toBe(1);
  expect(written).toEqual([[230, 230, 230, 255, 230, 230, 230, 255]]);
  if (realm === 'worker') expect(workerDecoder).toHaveBeenCalledWith(expect.any(Blob), {
    targetWidthPx: 2, targetHeightPx: 1,
  });
});

it('closes an over-budget SVG base before allocating any effect surfaces', async () => {
  const close = vi.fn();
  const svgDecoder = vi.fn(async () => ({ width: MAX_IMAGE_EFFECT_BASE_PIXELS + 1, height: 1, close }) as ImageBitmap);
  const surface = vi.fn();
  vi.stubGlobal('OffscreenCanvas', surface);
  await expect(decodeRaster('word/media/too-large.svg', 'image/svg+xml', undefined,
    async () => new Blob(['<svg/>']), 1, 1, undefined, undefined,
    undefined, undefined, svgDecoder, undefined, [{ type: 'luminance', bright: .9, contrast: 0 }],
  )).rejects.toMatchObject({ code: 'ooxml-decoded-image-limit' });
  expect(close).toHaveBeenCalledOnce();
  expect(surface).not.toHaveBeenCalled();
});

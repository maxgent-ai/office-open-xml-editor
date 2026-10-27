import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  MAX_IMAGE_EFFECT_BASE_PIXELS,
  decodeSvgBlobOnMainThread,
  type SvgBlobDecoder,
  type SvgDecodeTarget,
} from '@silurus/ooxml-core';
import { decodeRaster, preloadPaintImages, type DocxFetchImage } from './browser-images.js';

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

// #1950: an `asvg:svgBlip` picture whose box aspect differs from the SVG's.
// The host rasterizer covers the box without distorting the SVG, so a square
// SVG in a 1.00 x 0.90 in box retains a 96x96 grid at 96 DPI. Worker mode must
// admit exactly that grid (as main mode draws the SVG), not the 96x87 box.
describe('aspect-mismatched SVG pictures (#1950)', () => {
  const SQUARE_SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 96 96"/>';
  class HostImage {
    naturalWidth = 96; naturalHeight = 96; width = 0; height = 0;
    onload?: () => void;
    set src(_value: string) { queueMicrotask(() => this.onload?.()); }
  }
  // Production Window rasterizer; only its DOM image is stubbed.
  const hostDecoder = vi.fn(async (blob: Blob, target?: SvgDecodeTarget) => {
    vi.stubGlobal('Image', HostImage);
    try {
      return await decodeSvgBlobOnMainThread(blob, target);
    } finally {
      vi.stubGlobal('Image', undefined);
    }
  });

  beforeEach(() => {
    hostDecoder.mockClear();
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(readonly width: number, readonly height: number) {}
      getContext() { return { drawImage() {} }; }
      transferToImageBitmap() {
        return { width: this.width, height: this.height, close() {} } as unknown as ImageBitmap;
      }
    });
  });

  const preload = (heightIn: number, fetchImage: DocxFetchImage, svgDecoder?: SvgBlobDecoder) =>
    preloadPaintImages([{
      kind: 'image', resourceKey: 'picture', partPath: 'word/media/image1.png',
      mimeType: 'image/png', svgImagePath: 'word/media/image2.svg',
      intrinsicSize: { widthPt: 72, heightPt: 72 * heightIn },
    }], [{
      resourceKind: 'image', resourceKey: 'picture', widthPt: 72, heightPt: 72 * heightIn,
    }], fetchImage, undefined, 96 / 72, svgDecoder);

  const fetchSvgOnly = () => vi.fn(async (path: string) => {
    if (path !== 'word/media/image2.svg') throw new Error(`unexpected fallback fetch: ${path}`);
    return new Blob([SQUARE_SVG], { type: 'image/svg+xml' });
  });

  it.each([0.9, 0.98, 0.99, 1])('worker retains the 96x96 coverage grid for a 1.00 x %s in box', async (heightIn) => {
    vi.stubGlobal('Image', undefined);
    const fetchImage = fetchSvgOnly();
    const images = await preload(heightIn, fetchImage, hostDecoder);
    const [image] = [...images.values()] as ImageBitmap[];
    expect({ width: image.width, height: image.height }).toEqual({ width: 96, height: 96 });
    expect(hostDecoder).toHaveBeenCalledWith(expect.any(Blob), {
      targetWidthPx: 96, targetHeightPx: Math.ceil(96 * heightIn),
    });
  });

  // Main mode draws the SVG element itself; no raster target is involved.
  it('main draws the SVG element for a 1.00 x 0.90 in box', async () => {
    vi.stubGlobal('Image', HostImage);
    const images = await preload(0.9, fetchSvgOnly());
    expect([...images.values()][0]).toBeInstanceOf(HostImage);
  });

  it('still rejects a worker SVG raster beyond the coverage grid', async () => {
    vi.stubGlobal('Image', undefined);
    const close = vi.fn();
    const oversized = vi.fn(async () => ({ width: 960, height: 960, close }) as unknown as ImageBitmap);
    await expect(preload(0.9, fetchSvgOnly(), oversized))
      .rejects.toMatchObject({ code: 'ooxml-decoded-image-limit', metric: 'image-pixels' });
    expect(close).toHaveBeenCalledTimes(1);
  });
});

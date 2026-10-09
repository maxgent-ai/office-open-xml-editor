import { beforeEach, describe, expect, it, vi } from 'vitest';
import { OptionalImageCodecUnavailableError, TiffDecodeError } from '@silurus/ooxml-core';
import type { Slide } from './types.js';

const coreMocks = vi.hoisted(() => ({
  bitmap: { width: 1920, height: 720, close() {} } as unknown as ImageBitmap,
  decode: vi.fn(),
}));

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  coreMocks.decode.mockResolvedValue(coreMocks.bitmap);
  return {
    ...actual,
    getCachedDuotoneBitmapByPath: coreMocks.decode,
  };
});

import { renderSlide } from './renderer.js';

function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function canvas(
  drawImage?: ReturnType<typeof vi.fn>,
  fillText?: ReturnType<typeof vi.fn>,
): HTMLCanvasElement {
  const state: Record<string, unknown> = {
    fillStyle: '', strokeStyle: '', globalAlpha: 1, lineWidth: 1,
  };
  const context = new Proxy(state, {
    get(target, property: string) {
      if (property in target) return target[property];
      if (property === 'drawImage' && drawImage) return drawImage;
      if (property === 'fillText' && fillText) return fillText;
      return () => undefined;
    },
    set(target, property: string, value) {
      target[property] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return {
    width: 0, height: 0, style: {}, offsetWidth: 960,
    getContext: () => context,
  } as unknown as HTMLCanvasElement;
}

describe('PPTX display-sized picture decode', () => {
  beforeEach(() => {
    coreMocks.decode.mockReset();
    coreMocks.decode.mockResolvedValue(coreMocks.bitmap);
  });

  it('labels an opaque legacy PICT picture without invoking a decoder, and keeps other pictures', async () => {
    const fillText = vi.fn();
    const frame = {
      type: 'picture', x: 0, y: 0, width: 4_572_000, height: 3_429_000,
      rotation: 0, flipH: false, flipV: false,
    };
    const slide = { index: 0, slideNumber: 1, background: null, elements: [
      { ...frame, imagePath: 'media/image-a', mimeType: 'image/x-pict' },
      { ...frame, imagePath: 'media/image-b', mimeType: 'image/png' },
    ] } as Slide;
    await renderSlide(canvas(undefined, fillText), slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob([pngHeader(2, 1) as BlobPart], { type: 'image/png' })),
    });
    expect(fillText).toHaveBeenCalledWith('PICT image unsupported', 240, 180, 480);
    expect(coreMocks.decode.mock.calls.map(([path]) => path)).toEqual(['media/image-b']);
  });

  it('propagates a recognized TIFF codec failure instead of silently omitting the picture', async () => {
    const error = new TiffDecodeError('Unsupported TIFF compression');
    coreMocks.decode.mockRejectedValue(error);
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'picture',
        x: 0,
        y: 0,
        width: 4_572_000,
        height: 3_429_000,
        rotation: 0,
        flipH: false,
        flipV: false,
        imagePath: 'ppt/media/unsupported.tiff',
        mimeType: 'image/tiff',
      }],
    } as Slide;

    await expect(renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['tiff'], { type: 'image/tiff' })),
    })).rejects.toBe(error);
  });

  it('paints a placeholder and continues when only the optional TIFF codec is absent', async () => {
    coreMocks.decode.mockRejectedValue(new OptionalImageCodecUnavailableError('tiff'));
    const fillText = vi.fn();
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'picture',
        x: 0,
        y: 0,
        width: 4_572_000,
        height: 3_429_000,
        rotation: 0,
        flipH: false,
        flipV: false,
        imagePath: 'ppt/media/optional.tiff',
        mimeType: 'image/tiff',
      }],
    } as Slide;

    await expect(renderSlide(canvas(undefined, fillText), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['tiff'], { type: 'image/tiff' })),
    })).resolves.toBeDefined();

    expect(fillText).toHaveBeenCalledWith('TIFF image unavailable', 240, 180, 480);
  });

  it('contains a missing optional TIFF codec in a slide background', async () => {
    coreMocks.decode.mockRejectedValue(new OptionalImageCodecUnavailableError('tiff'));
    const fillText = vi.fn();
    const slide = {
      index: 0,
      slideNumber: 1,
      background: {
        fillType: 'image',
        imagePath: 'ppt/media/background.tiff',
        mimeType: 'image/tiff',
      },
      elements: [],
    } as Slide;

    await expect(renderSlide(canvas(undefined, fillText), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['tiff'], { type: 'image/tiff' })),
    })).resolves.toBeDefined();

    expect(fillText).toHaveBeenCalledWith('TIFF image unavailable', 480, 360, 960);
  });

  it('derives full-source device pixels from the frame, effective DPR, and srcRect crop', async () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [{
        type: 'picture',
        x: 0,
        y: 0,
        width: 4_572_000, // half of the 10-inch slide => 480 CSS px
        height: 3_429_000, // half slide height => 360 CSS px
        rotation: 0,
        flipH: false,
        flipV: false,
        imagePath: 'ppt/media/108mp-poster.png',
        mimeType: 'image/png',
        srcRect: { l: 0.25, t: 0, r: 0.25, b: 0 },
      }],
    } as Slide;
    const fetchImage = vi.fn(async () =>
      new Blob([pngHeader(12_090, 9_063) as BlobPart], { type: 'image/png' }));

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
    });

    expect(coreMocks.decode).toHaveBeenCalled();
    for (const call of coreMocks.decode.mock.calls) {
      expect(call[4]).toMatchObject({
        // 480 CSS px × DPR 2, then / 50% visible source width, retained at
        // 2× that display grid while the geometry share has room.
        targetWidthPx: 3840,
        targetHeightPx: 1440,
      });
    }
  });

  it('applies background fillRect to both decode sizing and cropped destination paint', async () => {
    const drawImage = vi.fn();
    const slide = {
      index: 0,
      slideNumber: 1,
      background: {
        fillType: 'image',
        imagePath: 'ppt/media/background.png',
        mimeType: 'image/png',
        srcRect: { l: 0.25, t: 0, r: 0.1, b: 0 },
        fillRect: { l: 0.1, t: -0.1, r: 0.2, b: 0.05 },
      },
      elements: [],
    } as Slide;

    await renderSlide(canvas(drawImage), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage: vi.fn(async () =>
        new Blob([pngHeader(12_090, 8_838) as BlobPart], { type: 'image/png' })),
    });

    expect(coreMocks.decode).toHaveBeenCalledWith(
      'ppt/media/background.png',
      'image/png',
      undefined,
      expect.any(Function),
      expect.objectContaining({ targetWidthPx: 4136, targetHeightPx: 3024 }),
    );
    expect(drawImage).toHaveBeenCalledWith(
      coreMocks.bitmap,
      480, 0, 1248, 720,
      96, -72, 672, 756,
    );
  });

  it('keeps DrawingML pixel effects and metafiles on their authored source grid', async () => {
    const slide = {
      index: 0,
      slideNumber: 1,
      background: {
        fillType: 'image',
        imagePath: 'ppt/media/background.png',
        mimeType: 'image/png',
        duotone: { clr1: '000000', clr2: 'FFFFFF' },
      },
      elements: [{
        type: 'picture',
        x: 0,
        y: 0,
        width: 598_125,
        height: 689_458,
        rotation: 0,
        flipH: false,
        flipV: false,
        imagePath: 'ppt/media/icon.wmf',
        mimeType: 'image/wmf',
        duotone: { clr1: '112233', clr2: 'FFFFFF' },
      }],
    } as Slide;
    const fetchImage = vi.fn(async (path: string) => path.endsWith('.png')
      ? new Blob([pngHeader(1_600, 1_200) as BlobPart], { type: 'image/png' })
      : new Blob([new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a]) as BlobPart], { type: 'image/wmf' }));

    await renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 2,
      fetchImage,
    });

    expect(coreMocks.decode).toHaveBeenCalled();
    for (const call of coreMocks.decode.mock.calls) {
      expect(call[4]).not.toHaveProperty('targetWidthPx');
      expect(call[4]).not.toHaveProperty('targetHeightPx');
    }
  });
});

describe('PPTX slide paint atomicity', () => {
  beforeEach(() => {
    coreMocks.decode.mockReset();
  });

  // Chrome rasterizes a canvas recording in chunks split at task boundaries,
  // and antialiased-path coverage can depend on that chunking (issue #1580).
  // A slide must therefore not be cleared or drawn until every decode it needs
  // has settled, and must then be painted without yielding to another task.
  it('paints nothing until all decodes settle, then paints in one task', async () => {
    const pending = new Map<string, (bitmap: ImageBitmap) => void>();
    coreMocks.decode.mockImplementation((path: string) => new Promise((resolve) => {
      pending.set(path, resolve);
    }));
    const events: string[] = [];
    let taskBoundarySinceFirstDraw = false;
    const record = (name: string) => {
      if (events.length === 0) setTimeout(() => { taskBoundarySinceFirstDraw = true; }, 0);
      else if (taskBoundarySinceFirstDraw) events.push('<task boundary>');
      events.push(name);
    };
    const context = new Proxy({} as Record<string, unknown>, {
      get(target, property: string) {
        if (property in target) return target[property];
        if (['fillRect', 'drawImage', 'fill', 'stroke'].includes(property)) {
          return () => record(property);
        }
        if (property === 'getTransform') {
          return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
        }
        return () => undefined;
      },
      set(target, property: string, value) {
        target[property] = value;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;
    const target = {
      set width(_value: number) { record('clear'); },
      height: 0,
      style: {},
      offsetWidth: 960,
      getContext: () => context,
    } as unknown as HTMLCanvasElement;
    const picture = (imagePath: string) => ({
      type: 'picture',
      x: 0,
      y: 0,
      width: 4_572_000,
      height: 3_429_000,
      rotation: 0,
      flipH: false,
      flipV: false,
      imagePath,
      mimeType: 'image/png',
    });
    const slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [picture('ppt/media/a.png'), picture('ppt/media/b.png')],
    } as Slide;

    const render = renderSlide(target, slide, 9_144_000, 6_858_000, {
      width: 960,
      dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['png'], { type: 'image/png' })),
    });
    await vi.waitFor(() => expect(pending.size).toBe(2));
    pending.get('ppt/media/a.png')!(coreMocks.bitmap);
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(events).toEqual([]);

    pending.get('ppt/media/b.png')!(coreMocks.bitmap);
    await render;
    expect(events).toEqual(['clear', 'fillRect', 'drawImage', 'drawImage']);
  });

  const picturesSlide = (...paths: string[]) => ({
    index: 0,
    slideNumber: 1,
    background: null,
    elements: paths.map((imagePath) => ({
      type: 'picture',
      x: 0,
      y: 0,
      width: 4_572_000,
      height: 3_429_000,
      rotation: 0,
      flipH: false,
      flipV: false,
      imagePath,
      mimeType: imagePath.endsWith('.tiff') ? 'image/tiff' : 'image/png',
    })),
  }) as Slide;
  const renderPictures = (slide: Slide) => renderSlide(canvas(), slide, 9_144_000, 6_858_000, {
    width: 960,
    dpr: 1,
    fetchImage: vi.fn(async () => new Blob(['png'], { type: 'image/png' })),
  });

  it('rejects with a fatal decode error without waiting for later pictures', async () => {
    const fatal = new TiffDecodeError('Unsupported TIFF compression');
    coreMocks.decode.mockImplementation((path: string) => path.endsWith('a.tiff')
      ? Promise.reject(fatal)
      : new Promise(() => {}));

    await expect(renderPictures(picturesSlide('ppt/media/a.tiff', 'ppt/media/b.png')))
      .rejects.toBe(fatal);
  });

  it('reports the earliest fatal decode error in paint order', async () => {
    const first = new TiffDecodeError('first picture');
    const second = new TiffDecodeError('second picture');
    coreMocks.decode.mockImplementation((path: string) => path.endsWith('a.tiff')
      ? new Promise((_, reject) => setTimeout(() => reject(first), 10))
      : Promise.reject(second));

    await expect(renderPictures(picturesSlide('ppt/media/a.tiff', 'ppt/media/b.tiff')))
      .rejects.toBe(first);
  });
});

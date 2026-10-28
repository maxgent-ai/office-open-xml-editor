import { describe, it, expect, afterEach, vi } from 'vitest';
import { getCachedSvgImageByPath, dropSvgImageCache } from './svg-image-by-path';
import { getCachedBitmapByPath } from './bitmap-image-by-path';
import { decodeSvgBlobOnMainThread } from './svg-raster';

describe('getCachedSvgImageByPath', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('fetches bytes, makes an object URL, loads an <img>, dedupes by path', async () => {
    let created = 0;
    let revoked = 0;
    vi.stubGlobal('URL', { createObjectURL: () => { created++; return `blob:${created}`; },
                           revokeObjectURL: () => { revoked++; } });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    const b = await getCachedSvgImageByPath('word/media/i.svg', fetchImage);
    expect(a).toBe(b);
    expect(fetchImage).toHaveBeenCalledTimes(1);
    expect(created).toBe(1);
    expect(revoked).toBe(1); // raw Blob handle is not retained with the image
  });

  it('self-evicts on failure (no poisoned cache) and revokes the failed object URL', async () => {
    let created = 0;
    let revoked = 0;
    vi.stubGlobal('URL', {
      createObjectURL: () => { created++; return `blob:${created}`; },
      revokeObjectURL: () => { revoked++; },
    });
    // <img> that always fails to load.
    class FailImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onerror && this.onerror()); } }
    vi.stubGlobal('Image', FailImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    await expect(getCachedSvgImageByPath('p.svg', fetchImage)).rejects.toThrow();
    expect(revoked).toBe(1); // failed URL revoked, not leaked
    // Second call must RETRY (cache self-evicted), not return a cached rejection.
    await expect(getCachedSvgImageByPath('p.svg', fetchImage)).rejects.toThrow();
    expect(fetchImage).toHaveBeenCalledTimes(2);
  });

  it('namespaces the cache by fetchImage — same zip path from two documents does not cross-contaminate', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    // Two different documents reference the SAME internal zip path. Their byte
    // sources (fetchImage closures) differ — the cache must not serve doc A's
    // decoded SVG for doc B's request.
    const fetchA = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const fetchB = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    const a = await getCachedSvgImageByPath('word/media/image1.svg', fetchA);
    const b = await getCachedSvgImageByPath('word/media/image1.svg', fetchB);
    expect(fetchA).toHaveBeenCalledTimes(1);
    expect(fetchB).toHaveBeenCalledTimes(1); // B must consult its OWN source, not hit A's cache
    expect(a).not.toBe(b); // distinct decoded images
    // Within one document, the path still dedupes.
    const a2 = await getCachedSvgImageByPath('word/media/image1.svg', fetchA);
    expect(a2).toBe(a);
    expect(fetchA).toHaveBeenCalledTimes(1);
  });

  it('dropSvgImageCache forgets decoded images and lets them re-decode', async () => {
    let revoked = 0;
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => { revoked++; } });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', FakeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    await getCachedSvgImageByPath('a.svg', fetchImage);
    await getCachedSvgImageByPath('b.svg', fetchImage);
    expect(revoked).toBe(2); // each URL was already released after decode
    dropSvgImageCache(fetchImage); // e.g. on Document.destroy()
    expect(revoked).toBe(2);
    await getCachedSvgImageByPath('a.svg', fetchImage);
    expect(fetchImage).toHaveBeenCalledTimes(3); // cache cleared → fresh decode
  });

  it('awaits img.decode() before resolving when available', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    let decoded = false;
    class DecodeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      decode() { return Promise.resolve().then(() => { decoded = true; }); }
      set src(_v: string) { queueMicrotask(() => this.onload && this.onload()); } }
    vi.stubGlobal('Image', DecodeImg);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
    await getCachedSvgImageByPath('d.svg', fetchImage);
    expect(decoded).toBe(true); // resolved only after decode() completed
  });

  it('uses the shared decoded owner when HTMLImageElement is unavailable', async () => {
    vi.stubGlobal('Image', undefined);
    const bitmap = { width: 20, height: 10, close() {} } as unknown as ImageBitmap;
    const createImageBitmap = vi.fn(async () => bitmap);
    vi.stubGlobal('createImageBitmap', createImageBitmap);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    const result = await getCachedSvgImageByPath('worker.svg', fetchImage);
    expect(result).toBe(bitmap);
    expect(createImageBitmap).toHaveBeenCalledTimes(1);
    expect(fetchImage).toHaveBeenCalledWith('worker.svg', 'image/svg+xml');
  });

  it('uses the Window decode bridge at the requested display size in a worker', async () => {
    vi.stubGlobal('Image', undefined);
    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      throw new Error('Chromium workers cannot decode this SVG Blob');
    }));
    const bitmap = { width: 640, height: 360, close() {} } as unknown as ImageBitmap;
    const workerDecoder = vi.fn(async () => bitmap);
    const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));

    const result = await getCachedSvgImageByPath('worker-bridged.svg', fetchImage, {
      targetWidthPx: 640,
      targetHeightPx: 360,
      workerDecoder,
    });

    expect(result).toBe(bitmap);
    expect(workerDecoder).toHaveBeenCalledWith(expect.any(Blob), {
      targetWidthPx: 640,
      targetHeightPx: 360,
    });
  });

  it('admits at most two concurrent SVG fetch/decode operations per document', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload?.()); } }
    vi.stubGlobal('Image', FakeImg);
    let active = 0;
    let maximum = 0;
    const releases: Array<() => void> = [];
    const fetchImage = vi.fn(async () => {
      active++;
      maximum = Math.max(maximum, active);
      await new Promise<void>((resolve) => releases.push(resolve));
      active--;
      return new Blob(['<svg/>'], { type: 'image/svg+xml' });
    });

    const pending = [0, 1, 2].map((i) => getCachedSvgImageByPath(`${i}.svg`, fetchImage));
    await Promise.resolve();
    await Promise.resolve();
    expect(fetchImage).toHaveBeenCalledTimes(2);
    releases.shift()?.();
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(fetchImage).toHaveBeenCalledTimes(3);
    while (releases.length > 0) releases.shift()?.();
    await Promise.all(pending);
    expect(maximum).toBe(2);
  });

  it('shares the same two decode slots across SVG and raster work', async () => {
    vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
    class FakeImg { onload: (() => void) | null = null; onerror: (() => void) | null = null;
      set src(_v: string) { queueMicrotask(() => this.onload?.()); } }
    vi.stubGlobal('Image', FakeImg);
    vi.stubGlobal('createImageBitmap', vi.fn(async () => (
      { width: 1, height: 1, close() {} } as unknown as ImageBitmap
    )));
    let active = 0;
    let maximum = 0;
    const releases: Array<() => void> = [];
    const fetchImage = vi.fn(async (_path: string, mime: string) => {
      active++;
      maximum = Math.max(maximum, active);
      await new Promise<void>((resolve) => releases.push(resolve));
      active--;
      return new Blob([mime === 'image/svg+xml' ? '<svg/>' : new Uint8Array([0])], { type: mime });
    });

    const pending = [
      getCachedSvgImageByPath('a.svg', fetchImage),
      getCachedBitmapByPath('a.png', 'image/png', fetchImage),
      getCachedSvgImageByPath('b.svg', fetchImage),
      getCachedBitmapByPath('b.png', 'image/png', fetchImage),
    ];
    await vi.waitFor(() => expect(active).toBe(2));
    expect(fetchImage).toHaveBeenCalledTimes(2);
    while (releases.length > 0) {
      releases.shift()?.();
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
    await Promise.all(pending);
    expect(maximum).toBe(2);
  });

  describe('worker SVG coverage admission (#1950)', () => {
    /** The production Window rasterizer, with only the DOM image stubbed: it
     * reports a 96x96 intrinsic size (a square viewBox) and rasterizes into an
     * OffscreenCanvas whose transferred bitmap has the canvas grid. */
    const hostDecoder = (naturalWidth: number, naturalHeight: number) =>
      async (blob: Blob, target?: { targetWidthPx?: number; targetHeightPx?: number }) => {
        class HostImage {
          naturalWidth = naturalWidth; naturalHeight = naturalHeight;
          width = 0; height = 0;
          onload: (() => void) | null = null;
          set src(_v: string) { queueMicrotask(() => this.onload?.()); }
        }
        vi.stubGlobal('Image', HostImage);
        try {
          return await decodeSvgBlobOnMainThread(blob, target);
        } finally {
          vi.stubGlobal('Image', undefined);
        }
      };

    function stubWorkerRealm() {
      vi.stubGlobal('Image', undefined);
      vi.stubGlobal('URL', { createObjectURL: () => 'blob:x', revokeObjectURL: () => {} });
      vi.stubGlobal('OffscreenCanvas', class {
        constructor(readonly width: number, readonly height: number) {}
        getContext() { return { drawImage() {} }; }
        transferToImageBitmap() {
          return { width: this.width, height: this.height, close() {} } as unknown as ImageBitmap;
        }
      });
    }

    it.each([0.9, 0.98, 0.99, 1])(
      'admits the aspect-preserving 96x96 grid of a square SVG in a 96 x (96*%s) box',
      async (ratio) => {
        stubWorkerRealm();
        const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
        const targetHeightPx = Math.ceil(96 * ratio);
        const result = await getCachedSvgImageByPath('square.svg', fetchImage, {
          targetWidthPx: 96,
          targetHeightPx,
          workerDecoder: hostDecoder(96, 96),
        }) as ImageBitmap;
        expect({ width: result.width, height: result.height }).toEqual({ width: 96, height: 96 });
      },
    );

    it('rejects and closes a worker raster larger than the coverage grid it was asked for', async () => {
      stubWorkerRealm();
      const close = vi.fn();
      // Ten times the request on both axes: no aspect ratio needs that grid.
      const workerDecoder = vi.fn(async () => (
        { width: 960, height: 870, close } as unknown as ImageBitmap
      ));
      const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
      await expect(getCachedSvgImageByPath('oversized.svg', fetchImage, {
        targetWidthPx: 96,
        targetHeightPx: 87,
        workerDecoder,
      })).rejects.toMatchObject({
        code: 'ooxml-decoded-image-limit',
        metric: 'image-pixels',
        limit: 96 * 87,
        observed: 960 * 870,
      });
      expect(close).toHaveBeenCalledTimes(1);
    });

    it('keeps non-finite worker raster dimensions a terminal quota failure', async () => {
      stubWorkerRealm();
      const close = vi.fn();
      const workerDecoder = vi.fn(async () => (
        { width: Infinity, height: Infinity, close } as unknown as ImageBitmap
      ));
      const fetchImage = vi.fn(async () => new Blob(['<svg/>'], { type: 'image/svg+xml' }));
      await expect(getCachedSvgImageByPath('non-finite.svg', fetchImage, {
        targetWidthPx: 96,
        targetHeightPx: 87,
        workerDecoder,
      })).rejects.toMatchObject({ code: 'ooxml-decoded-image-limit', metric: 'image-dimension' });
      expect(close).toHaveBeenCalledTimes(1);
    });
  });
});

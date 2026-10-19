import { afterEach, expect, it, vi } from 'vitest';
import { OoxmlError, OoxmlResourceLimitError } from '../errors/ooxml-error.js';
import { OoxmlDecodedImageLimitError } from './pixel-budget.js';
import { TiffDecodeError } from './tiff-contract.js';
import { OptionalImageCodecUnavailableError } from './optional-image-fallback.js';
import { OoxmlIncompleteMetafileError } from './raster-or-metafile.js';
import { applyBlipPixelEffects } from './blip-effects.js';
import { getCachedDuotoneBitmapByPath } from './duotone-bitmap-by-path.js';
import { dropBitmapCacheByPath } from './bitmap-image-by-path.js';
import type { OffscreenFactory } from './duotone.js';

const errors = [
  new OoxmlError('unsupported-image-effect', 'Invalid authored chain'),
  new OoxmlDecodedImageLimitError('image-effect-work', 10, 11),
  new OoxmlResourceLimitError('Archive quota', { stage: 'parsing', violation: {
    format: 'docx', operation: 'open', resource: 'archive', metric: 'entry-count', limit: 10, observed: 11,
    configurable: false, usage: { archiveEntryCount: 11, declaredInflatedBytes: 0,
      operationInflatedBytes: 0, distinctInflatedBytes: 0 },
  } }),
  new TiffDecodeError('Codec failed'),
  new OptionalImageCodecUnavailableError('tiff'),
  new OoxmlIncompleteMetafileError('emf', ['unsupported record']),
];
afterEach(() => vi.unstubAllGlobals());

it.each(errors.flatMap(error => ['dependency', 'readback', 'bake'].map(stage => ({ error, stage }))))(
  'preserves $error.name identity during strict $stage failures', async ({ error, stage }) => {
    const base = { width: 1, height: 1, close() {} } as ImageBitmap;
    vi.stubGlobal('createImageBitmap', async (source: unknown) => {
      if (!(source instanceof Blob) && stage === 'bake') throw error;
      return base;
    });
    const fetchImage = async (path: string) => {
      if (path === 'nested.png' && stage === 'dependency') throw error;
      return new Blob([new Uint8Array([1])]);
    };
    const offscreenFactory = (() => ({ width: 1, height: 1, getContext: () => ({
      drawImage() {},
      getImageData() {
        if (stage === 'readback') throw error;
        return { width: 1, height: 1, data: new Uint8ClampedArray(4) };
      },
      putImageData() {},
    }) })) as unknown as OffscreenFactory;
    const effects = stage === 'dependency'
      ? [{ type: 'fillOverlay' as const, blend: 'over' as const, fill: {
        fillType: 'image' as const, imagePath: 'nested.png', mimeType: 'image/png', stretch: true,
      } }] : [{ type: 'grayscale' as const }];
    try {
      await expect(getCachedDuotoneBitmapByPath('base.png', 'image/png', { effects }, fetchImage, { offscreenFactory }))
        .rejects.toBe(error);
    } finally { dropBitmapCacheByPath(fetchImage); }
  });


it.each(errors)('preserves %s during evaluated fill rasterization', error => {
  expect(() => applyBlipPixelEffects({ width: 1, height: 1, data: new Uint8ClampedArray(4) }, {
    effects: [{ type: 'fillOverlay', blend: 'over', fill: {
      fillType: 'image', imagePath: 'nested.png', mimeType: 'image/png', stretch: true,
    } }],
  }, { offscreenFactory: () => { throw error; } })).toThrow(error);
});

import { describe, expect, it, vi } from 'vitest';
import { effectFillPixels, effectFillWork } from './blip-fill';
import type { BlipEffectFill } from './blip-effects';
import { setImageEffectFrame } from './image-effect-frame';
import type { OffscreenFactory } from './duotone';

function paint(fill: BlipEffectFill, source = { width: 1, height: 1 } as CanvasImageSource) {
  const ctx = {
    translate: vi.fn(), scale: vi.fn(), save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
    // Abort the old unbounded loop on its first draw, so a regression cannot hang CI.
    drawImage: vi.fn(() => { if (fill.fillType === 'image' && fill.tile?.tx === 27273042316900) throw new Error('unsafe tile reached paint'); }),
    getImageData: () => ({ data: new Uint8ClampedArray(4), width: 1, height: 1 }),
  };
  const surface = { width: 1, height: 1, getContext: () => ctx };
  const run = () => effectFillPixels(fill, { data: new Uint8ClampedArray(4), width: 1, height: 1 },
    12700, 12700, { offscreenFactory: (() => surface) as unknown as OffscreenFactory,
      images: new Map([[fill, source]]) });
  return { ctx, surface, run };
}
const image = (tile: { sx?: number; sy?: number; tx?: number; ty?: number; flip?: string }, dpi = 72): BlipEffectFill =>
  ({ fillType: 'image', imagePath: 'tile.png', mimeType: 'image/png', tile, dpi });

describe('effect fill tile geometry', () => {
  it('rejects the reviewer’s unrepresentable tile index with a typed resource error', () => {
    const p = paint(image({ sx: 10 / 100000, sy: 1000000000 / 100000, tx: 27273042316900 }, 65535));
    expect(p.run).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
    expect(p.ctx.drawImage).not.toHaveBeenCalled();
    expect(p.surface.width).toBe(0);
  });

  it.each([{ sx: 0 }, { sy: -1 }, { sx: Infinity }, { tx: NaN }, { ty: Infinity }])(
    'rejects invalid geometry before drawing: %j', tile => {
      const p = paint(image(tile));
      expect(p.run).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
      expect(p.ctx.drawImage).not.toHaveBeenCalled();
    });

  it('rejects a non-finite DPI rather than substituting a default', () => {
    const p = paint(image({}, NaN));
    expect(p.run).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
    expect(p.ctx.drawImage).not.toHaveBeenCalled();
  });

  it.each([3.25, -2.75])('normalizes phase and retains alternate-tile flips for offset %s', tx => {
    const p = paint(image({ tx: tx * 12700, flip: 'x' }));
    p.run();
    expect(p.ctx.drawImage).toHaveBeenCalledTimes(9);
    expect(p.ctx.translate.mock.calls[1]).toEqual([-.75, 0]);
    expect(p.ctx.scale.mock.calls[1]).toEqual([1, 1]);
    expect(p.ctx.scale.mock.calls[2]).toEqual([-1, 1]);
  });
});

it('rejects an unavailable recursive pattern instead of substituting its foreground color', () => {
  const p = paint({ fillType: 'pattern', preset: 'unknown', fg: 'FF0000', bg: '0000FF' });
  expect(p.run).toThrowError(expect.objectContaining({ code: 'unsupported-image-effect' }));
  expect(p.surface.width).toBe(0);
});

it('tiles a grown source by its authored frame and budgets all overlapping outsets', () => {
  const source = { width: 6, height: 6 } as CanvasImageSource;
  setImageEffectFrame(source, { x: 2, y: 2, width: 2, height: 2 });
  const fill = image({});
  const p = paint(fill, source);
  const work = effectFillWork(fill, { width: 1, height: 1 }, 12700, 12700, { images: new Map([[fill, source]]) });
  p.run();
  expect(p.ctx.drawImage).toHaveBeenCalledTimes(25);
  expect(work).toBe(43); // 25 draws + 16 bounded overdraw samples + rasterization/readback
  expect(p.ctx.translate.mock.calls[2][0] - p.ctx.translate.mock.calls[1][0]).toBe(2);
});

it('repeats the srcRect window, not the whole source, as the tile period (Annex L.4.8.4.3)', () => {
  const source = { width: 4, height: 4 } as CanvasImageSource;
  const fill = { ...image({}), srcRect: { l: 0.25, t: 0.25, r: 0.25, b: 0.25 } } as BlipEffectFill;
  const p = paint(fill, source);
  p.run();
  // 4 px at 72 dpi cropped to its central half: a 2 pt period.
  expect(p.ctx.translate.mock.calls[2][0] - p.ctx.translate.mock.calls[1][0]).toBe(2);
});

it('tiles nothing for a crop with no positive extent', () => {
  const fill = { ...image({}), srcRect: { l: 0.6, t: 0, r: 0.6, b: 0 } } as BlipEffectFill;
  const p = paint(fill, { width: 4, height: 4 } as CanvasImageSource);
  p.run();
  expect(p.ctx.drawImage).not.toHaveBeenCalled();
});

it('covers the inverse-rotated box of a counter-rotated (rotWithShape=0) tiled brush', () => {
  const draws = (rotation: number) => {
    const fill = { ...image({}), rotWithShape: false } as BlipEffectFill;
    const ctx = {
      translate: vi.fn(), rotate: vi.fn(), scale: vi.fn(), save() {}, restore() {}, beginPath() {}, rect() {}, clip() {},
      drawImage: vi.fn(),
      getImageData: () => ({ data: new Uint8ClampedArray(8 * 2 * 4), width: 8, height: 2 }),
    };
    const surface = { width: 8, height: 2, getContext: () => ctx };
    effectFillPixels(fill, { data: new Uint8ClampedArray(8 * 2 * 4), width: 8, height: 2 }, 12700, 12700,
      { offscreenFactory: (() => surface) as unknown as OffscreenFactory, images: new Map([[fill, { width: 1, height: 1 } as CanvasImageSource]]) },
      rotation);
    return ctx.drawImage.mock.calls.length;
  };
  // An 8×2 box turned 90° spans 2×8 in brush space: more rows must be tiled.
  expect(draws(90)).toBeGreaterThan(draws(0));
});

import { Canvas } from 'skia-canvas';
import { afterEach, expect, it, vi } from 'vitest';
import { paintVmlImageMask } from './vml-mask.js';
import type { CanvasPaintContext } from './types.js';

afterEach(() => vi.unstubAllGlobals());
it('leaves the page intact when an optional image resource has no drawable', () => {
  vi.stubGlobal('OffscreenCanvas', Canvas);
  const canvas = new Canvas(20, 20);
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = 'red'; ctx.fillRect(0,0,20,20);
  const context = { ctx, scale: 1, dpr: 1, resources: { paint() {} } } as unknown as CanvasPaintContext;
  paintVmlImageMask('unavailable', {xPt:0,yPt:0,widthPt:20,heightPt:20}, context,
    mask => { mask.fillStyle='white'; mask.fillRect(0,0,20,20); });
  expect([...ctx.getImageData(10,10,1,1).data]).toEqual([255,0,0,255]);
});

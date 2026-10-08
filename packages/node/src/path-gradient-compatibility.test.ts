import { describe, expect, it } from 'vitest';
import { resolveFill, type GradientFill } from '@silurus/ooxml-core';
import { buildPresetGeometryFillPath } from '../../core/src/shape/preset-geometry';
import type { FillOutline } from '../../core/src/shape/path-gradient';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const stops = [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }];
const area = { l: .2, r: .2, t: .2, b: .2 };
const segment = { l: .2, r: .2, t: .5, b: .5 };
const preset = (name: string): FillOutline => (ctx, x, y, w, h) => {
  buildPresetGeometryFillPath(ctx, name, x, y, w, h, []);
};
const lShape: FillOutline = (ctx, x, y, w, h) => {
  [[0, 0], [1, 0], [1, .3], [.3, .3], [.3, 1], [0, 1]]
    .forEach(([u, v], i) => ctx[i ? 'lineTo' : 'moveTo'](x + u * w, y + v * h));
  ctx.closePath();
};

describe.skipIf(!skia)('path shade native compatibility boundaries', () => {
  function expectNative(recipe: GradientFill, outline: FillOutline, w = 120, h = 120) {
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const actual = new Canvas(260, 200).getContext('2d') as unknown as CanvasRenderingContext2D;
    const expected = new Canvas(260, 200).getContext('2d') as unknown as CanvasRenderingContext2D;
    for (const ctx of [actual, expected]) {
      ctx.translate(180, 30); ctx.rotate(.3); ctx.scale(-1, 1);
    }
    actual.fillStyle = resolveFill(recipe, actual, 0, 0, w, h, 0, undefined, undefined, outline) as CanvasGradient;
    // Independent previous-renderer oracle: authored midpoint, max-axis rect
    // radius or diagonal shape radius, with native Canvas stop interpolation.
    const f = recipe.fillToRect;
    const cx = w * (f?.l ?? 0) + w * (1 - (f?.l ?? 0) - (f?.r ?? 0)) / 2;
    const cy = h * (f?.t ?? 0) + h * (1 - (f?.t ?? 0) - (f?.b ?? 0)) / 2;
    const rx = Math.max(Math.abs(cx), Math.abs(w - cx));
    const ry = Math.max(Math.abs(cy), Math.abs(h - cy));
    const gradient = expected.createRadialGradient(cx, cy, 0, cx, cy,
      recipe.path === 'rect' ? Math.max(rx, ry) : Math.sqrt(rx * rx + ry * ry));
    for (const stop of recipe.stops) gradient.addColorStop(stop.position, `#${stop.color}`);
    expected.fillStyle = gradient;
    for (const ctx of [actual, expected]) ctx.fillRect(0, 0, w, h);
    expect(Buffer.from(actual.getImageData(0, 0, 260, 200).data)
      .equals(Buffer.from(expected.getImageData(0, 0, 260, 200).data))).toBe(true);
  }
  const recipe = (path: 'rect' | 'shape', fillToRect?: GradientFill['fillToRect']): GradientFill =>
    ({ fillType: 'gradient', gradType: 'radial', angle: 0, path, fillToRect, stops });

  it('retains native rect point foci on supported concave outlines and translucent ellipses', () => {
    for (const name of ['plus', 'star4', 'star5']) for (const focus of [
      undefined, { l: 0, t: 0, r: 1, b: 1 }, { l: .25, t: .25, r: .75, b: .75 },
      { l: 1, t: 1, r: 0, b: 0 },
    ]) expectNative(recipe('rect', focus), preset(name));
    expectNative({ ...recipe('rect'), stops: [{ position: 0, color: 'FF000000' },
      { position: 1, color: 'FF0000FF' }] }, preset('ellipse'));
  });

  it('retains native rect area paint for host outlines outside the center kernel', () => {
    for (const outline of [preset('rightArrow'), preset('chevron'), preset('donut'), lShape]) {
      expectNative(recipe('rect', area), outline);
    }
    expectNative(recipe('rect', { l: 0, r: 0, t: 0, b: 0 }), preset('rightArrow'), 200, 20);
  });

  it('keeps area shades unchanged by duplicate joins and forward collinear subdivisions', () => {
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const outlines: FillOutline[] = [
      (ctx, x, y, w, h) => ctx.rect(x, y, w, h),
      (ctx, x, y, w, h) => {
        // Includes a redundant vertex across the implicit closing edge.
        [[0, 0], [.5, 0], [1, 0], [1, 0], [1, .5], [1, 1], [.5, 1], [0, 1], [0, .5]]
          .forEach(([u, v], i) => ctx[i ? 'lineTo' : 'moveTo'](x + u * w, y + v * h));
        ctx.closePath();
      },
    ];
    const images = outlines.map(outline => {
      const ctx = new Canvas(120, 120).getContext('2d') as unknown as CanvasRenderingContext2D;
      ctx.fillStyle = resolveFill(recipe('shape', area), ctx, 0, 0, 120, 120,
        0, undefined, undefined, outline) as CanvasPattern;
      ctx.fillRect(0, 0, 120, 120);
      // A native midpoint fallback would fail this interior focus contract.
      expect([...ctx.getImageData(30, 30, 1, 1).data]).toEqual([0, 0, 0, 255]);
      return Buffer.from(ctx.getImageData(0, 0, 120, 120).data);
    });
    expect(images[1].equals(images[0])).toBe(true);
  });

  it('retains native midpoint paint for horizontal and vertical segment foci', () => {
    for (const name of ['rect', 'roundRect', 'plus', 'star4', 'star5']) {
      expectNative(recipe(name === 'rect' || name === 'roundRect' ? 'shape' : 'rect', segment), preset(name));
    }
    expectNative(recipe('shape', { l: .5, r: .5, t: .2, b: .2 }), preset('rect'));
  });
});

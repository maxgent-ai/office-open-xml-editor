import { describe, expect, it } from 'vitest';
import { flattenFillOutline, isStrictlyStarShaped, pathShadeFocusRect } from './path-gradient';
import type { GradientFill } from '../types/common';

const box = { x: 0, y: 0, w: 100, h: 100 };
const fill = (fillToRect?: GradientFill['fillToRect']): GradientFill =>
  ({ fillType: 'gradient', gradType: 'radial', path: 'shape', angle: 0, stops: [], fillToRect });
const stub = {} as CanvasRenderingContext2D;

describe('path shade focus rectangle', () => {
  it('retains the authored inset area and degenerate axes', () => {
    expect(pathShadeFocusRect(fill())).toEqual({ origin: [.5, .5], size: [0, 0] });
    expect(pathShadeFocusRect(fill({ l: 0, t: 0, r: 0, b: 0 }))).toEqual({ origin: [0, 0], size: [1, 1] });
    const area = pathShadeFocusRect(fill({ l: .2, t: .5, r: .2, b: .5 }));
    expect(area.origin).toEqual([.2, .5]);
    expect(area.size[0]).toBeCloseTo(.6); expect(area.size[1]).toBe(0);
    expect(pathShadeFocusRect(fill({ l: .25, t: .25, r: .75, b: .75 }))).toEqual({ origin: [.25, .25], size: [0, 0] });
  });
});

describe('star-shaped outline classification', () => {
  const polygons = (draw: (ctx: CanvasRenderingContext2D) => void) =>
    flattenFillOutline(ctx => draw(ctx), stub, box);
  it('accepts convex, curved and star polygons about the center', () => {
    const ellipse = polygons(ctx => { ctx.moveTo(100, 50); ctx.ellipse(50, 50, 50, 30, 0, 0, Math.PI * 2); });
    const star = polygons(ctx => {
      for (let i = 0; i < 10; i++) {
        const a = i * Math.PI / 5; const r = i % 2 ? 20 : 50;
        ctx.lineTo(50 + r * Math.cos(a), 50 + r * Math.sin(a));
      }
      ctx.closePath();
    });
    expect(isStrictlyStarShaped(ellipse, [50, 50], box)).toBe(true);
    expect(isStrictlyStarShaped(star, [50, 50], box)).toBe(true);
  });
  it('rejects holes, centers on the outline and concavities hiding edges', () => {
    const donut = polygons(ctx => {
      ctx.moveTo(100, 50); ctx.arc(50, 50, 50, 0, Math.PI * 2); ctx.closePath();
      ctx.moveTo(75, 50); ctx.arc(50, 50, 25, Math.PI * 2, 0, true); ctx.closePath();
    });
    const chevron = polygons(ctx => { ctx.moveTo(0, 0); ctx.lineTo(50, 0); ctx.lineTo(100, 50); ctx.lineTo(50, 100); ctx.lineTo(0, 100); ctx.lineTo(50, 50); ctx.closePath(); });
    const arrow = polygons(ctx => {
      [[0, 30], [60, 30], [60, 0], [100, 50], [60, 100], [60, 70], [0, 70]]
        .forEach(([x, y], i) => (i ? ctx.lineTo(x, y) : ctx.moveTo(x, y)));
      ctx.closePath();
    });
    expect(isStrictlyStarShaped(donut, [50, 50], box)).toBe(false);
    expect(isStrictlyStarShaped(chevron, [50, 50], box)).toBe(false);
    expect(isStrictlyStarShaped(arrow, [50, 50], box)).toBe(false);
  });
  it('flattens through the outline\'s own transforms without a real context', () => {
    const [square] = polygons(ctx => { ctx.save(); ctx.translate(10, 0); ctx.rect(0, 0, 20, 20); ctx.restore(); });
    expect(square.slice(0, 4)).toEqual([[10, 0], [30, 0], [30, 20], [10, 20]]);
  });
});

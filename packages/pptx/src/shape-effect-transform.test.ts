import { afterEach, describe, expect, it, vi } from 'vitest';
import { cacheDevicePaint, projectedEffectGeometry, transformedEffectBounds } from './renderer';

describe('shape effect bounds', () => {
  afterEach(() => vi.unstubAllGlobals());
  it('includes the active group transform translation and scale', () => {
    expect(transformedEffectBounds(
      { a: 1.5, b: 0, c: 0, d: 1.5, e: 72, f: 45 },
      464,
      120,
      146,
      200,
    )).toEqual({ x: 768, y: 225, w: 219, h: 300 });
  });

  it('returns the axis-aligned device bounds after rotation', () => {
    expect(transformedEffectBounds(
      { a: 0, b: 2, c: -2, d: 0, e: 500, f: 30 },
      10,
      20,
      40,
      30,
    )).toEqual({ x: 400, y: 50, w: 60, h: 80 });
  });

  it('keeps the authored shadow anchor distinct from the projected AABB corner', () => {
    const geometry = projectedEffectGeometry(
      { prst: 'orthographicFront' },
      { a: 0, b: 1, c: -1, d: 0, e: 75, f: -25 },
      0,
      0,
      100,
      50,
      0,
      'tr',
    );
    expect(geometry.bbox).toEqual({ x: 25, y: -25, w: 50, h: 100 });
    expect(geometry.anchor).toEqual([75, 75]);
    expect(geometry.anchor).not.toEqual([geometry.bbox.x + geometry.bbox.w, geometry.bbox.y]);
  });

  it('includes raster padding rounded up to device pixels under camera zoom', () => {
    const geometry = projectedEffectGeometry(
      { prst: 'orthographicFront', zoom: 4 },
      { a: 2, b: 0, c: 0, d: 2, e: 0, f: 0 },
      0, 0, 100, 50, 1.1, 'tr',
    );
    // 1.1 CSS px rounds to 3 device px. Camera zoom scales that local margin.
    expect(geometry.bbox.x).toBeCloseTo(-312, 10);
    expect(geometry.bbox.w).toBeCloseTo(824, 10);
  });

  it('bounds and anchors the exposed back surface at the extrusion base', () => {
    const geometry = projectedEffectGeometry(
      { prst: 'orthographicFront', rot: { lat: 0, lon: 150, rev: 0 } },
      { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      0, 0, 100, 50, 0, 'tr', 40,
    );
    // Parallel projection: width=100*cos(30), centre=50+40*sin(150).
    // The back-facing top-right corner is the left end of the projected box.
    expect(geometry.bbox.x).toBeCloseTo(70 - 50 * Math.cos(Math.PI / 6), 10);
    expect(geometry.bbox.w).toBeCloseTo(100 * Math.cos(Math.PI / 6), 10);
    expect(geometry.anchor[0]).toBeCloseTo(geometry.bbox.x, 10);
    expect(geometry.anchor[1]).toBe(0);
  });

  it('keeps a camera-crossing rear face and its effect anchor in the same finite scene', () => {
    const units = 96, w = 8 * units, h = units, depth = 8 * units;
    const extent = 2 * Math.sqrt(3) * units;
    const distance = extent + 4 * units + .02 * (2 * extent + 4 * units);
    const q = [-1, 1].flatMap(sx => [-1, 1].map(sy => {
      const z = sx * extent + 4 * units;
      return [w / 2 + (-sx * 2 + 4 * Math.sqrt(3)) * units / (1 - z / distance),
        h / 2 + sy * h / 2 / (1 - z / distance)];
    }));
    const geometry = projectedEffectGeometry(
      { prst: 'perspectiveFront', fov: 90, rot: { lat: 0, lon: 120, rev: 0 } },
      { a: 2, b: 0, c: 0, d: 2, e: 5, f: 7 },
      0, 0, w, h, 0, 'tr', depth,
    );
    expect(geometry.bbox.x).toBeCloseTo(5 + 2 * Math.min(...q.map(p => p[0])), 7);
    expect(geometry.bbox.y).toBeCloseTo(7 + 2 * Math.min(...q.map(p => p[1])), 7);
    expect(geometry.bbox.w).toBeCloseTo(2 * (Math.max(...q.map(p => p[0])) - Math.min(...q.map(p => p[0]))), 7);
    expect(geometry.bbox.h).toBeCloseTo(2 * (Math.max(...q.map(p => p[1])) - Math.min(...q.map(p => p[1]))), 7);
    expect(geometry.anchor[0]).toBeCloseTo(5 + 2 * q[2][0], 7);
    expect(geometry.anchor[1]).toBeCloseTo(7 + 2 * q[2][1], 7);
  });

  it('reuses one device-space projection across multiple effect passes', () => {
    const cacheContext = { setTransform: vi.fn() };
    class TestOffscreenCanvas {
      constructor(public width: number, public height: number) {}
      getContext(): typeof cacheContext { return cacheContext; }
    }
    vi.stubGlobal('OffscreenCanvas', TestOffscreenCanvas);
    const expensiveProjection = vi.fn();
    const replayContext = {
      save: vi.fn(),
      restore: vi.fn(),
      setTransform: vi.fn(),
      drawImage: vi.fn(),
    } as unknown as CanvasRenderingContext2D;
    const replay = cacheDevicePaint(
      expensiveProjection,
      { a: 2, b: 0, c: 0, d: 2, e: 20, f: 30 },
      { x: 10, y: 15, w: 100, h: 60 },
    );
    replay(replayContext);
    replay(replayContext);
    expect(expensiveProjection).toHaveBeenCalledTimes(1);
    expect(replayContext.drawImage).toHaveBeenCalledTimes(2);
  });

  it('declines an oversized projection cache and safely replays the source', () => {
    const constructor = vi.fn();
    vi.stubGlobal('OffscreenCanvas', constructor);
    const projection = vi.fn();
    const replay = cacheDevicePaint(
      projection,
      { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      { x: 0, y: 0, w: 100_000, h: 100_000 },
      { w: 960, h: 540 },
    );
    replay({} as CanvasRenderingContext2D);
    replay({} as CanvasRenderingContext2D);
    expect(constructor).not.toHaveBeenCalled();
    expect(projection).toHaveBeenCalledTimes(2);
  });

  it('still caches a bounded projection that crosses a viewport edge', () => {
    const cacheContext = { setTransform: vi.fn() };
    let created = 0;
    class CrossingCanvas {
      constructor() { created += 1; }
      getContext(): typeof cacheContext { return cacheContext; }
    }
    vi.stubGlobal('OffscreenCanvas', CrossingCanvas);
    const projection = vi.fn();
    const replay = cacheDevicePaint(
      projection,
      { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 },
      { x: -2, y: 20, w: 100, h: 60 },
      { w: 960, h: 540 },
    );
    const target = {
      save: vi.fn(), restore: vi.fn(), setTransform: vi.fn(), drawImage: vi.fn(),
    } as unknown as CanvasRenderingContext2D;
    replay(target);
    replay(target);
    expect(created).toBe(1);
    expect(projection).toHaveBeenCalledTimes(1);
  });
});

import { describe, expect, it } from 'vitest';
import type { DrawingMLShapePaintPlan } from './drawingml-shape';
import { drawingMLGeometryBounds, resolveDrawingMLGeometry, requireResolvedDrawingMLGeometry } from './drawingml-geometry';
import { GeometryWorkBudgetError } from './path-data';

const rectangle: DrawingMLShapePaintPlan = {
  rect: { x: 10, y: 20, w: 30, h: 40 },
  geometry: { kind: 'preset', name: 'rect', adjustments: [] },
  fill: { fillType: 'solid', color: 'FF0000' }, stroke: { color: '000000', width: 2 },
  transform: { rotationDeg: 0, flipH: false, flipV: false },
};
function retained(plan: DrawingMLShapePaintPlan) {
  return { ...plan, resolvedGeometry: resolveDrawingMLGeometry(plan, 1) };
}
describe('shared retained DrawingML geometry', () => {
  it('excludes unfilled/unstroked paths from enclosure without discarding their source order', () => {
    const plan = retained({ ...rectangle, stroke: null,
      geometry: { kind: 'custom', subpaths: [
        [{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1, y: 0 }, { cmd: 'lineTo', x: 0, y: 1 }, { cmd: 'close' }],
        [{ cmd: 'moveTo', x: -100, y: -100 }, { cmd: 'lineTo', x: 100, y: 100 }],
      ], paint: [{}, { fill: 'none', stroke: false }] },
    });
    expect(plan.resolvedGeometry.paths).toHaveLength(2);
    expect(drawingMLGeometryBounds(plan)).toEqual({ x: 10, y: 20, w: 30, h: 40 });
    expect(plan.resolvedGeometry.fillSilhouette).toHaveLength(4);
  });
  it('retains the retracted leader and full arrow body rather than the anchor rectangle', () => {
    const plan = retained({ ...rectangle, rect: { x: 10, y: 20, w: 30, h: 0 }, fill: null,
      geometry: { kind: 'preset', name: 'line', adjustments: [] },
      stroke: { color: '000000', width: 2, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
    });
    const leader = plan.resolvedGeometry.paths.find(p => p.stroke !== null)!;
    expect(leader.path).toEqual([{ op: 'moveTo', args: [10, 20] }, { op: 'lineTo', args: [31, 20] }]);
    expect(drawingMLGeometryBounds(plan)).toEqual({ x: 10, y: 15, w: 30, h: 10 });
  });
  it('trims a decorated custom path for its stroke but keeps the untrimmed fill', () => {
    // 2 pt line, lg/lg triangle: e = 2 pt, the line stops 10 − 1 = 9 pt short.
    const plan = retained({ ...rectangle, rect: { x: 0, y: 0, w: 40, h: 40 },
      geometry: { kind: 'custom', subpaths: [[{ cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'cubicBezTo', x1: 1, y1: 0, x2: 1, y2: 0, x: 1, y: 1 }]] },
      stroke: { color: '000000', width: 2, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
    });
    const [fill, line] = plan.resolvedGeometry.paths;
    expect(fill.stroke).toBeNull();
    expect(fill.path.at(-1)!.args.slice(4)).toEqual([40, 40]);
    expect(line.fill).toBe('none');
    const [ex, ey] = line.path.at(-1)!.args.slice(4);
    expect(Math.hypot(ex - 40, ey - 40)).toBeCloseTo(9, 6);
    // The split paths stay within the acquired work the painter validates.
    expect(requireResolvedDrawingMLGeometry(plan, 1)).toBe(plan.resolvedGeometry);
  });
  it('retracts a bent leader across a terminal leg shorter than the retraction', () => {
    // bentConnector2 ends with a 1 pt leg; lg/lg triangle on a 2 pt line stops the line 9 pt short.
    const plan = retained({ ...rectangle, rect: { x: 0, y: 0, w: 100, h: 1 }, fill: null,
      geometry: { kind: 'preset', name: 'bentConnector2', adjustments: [] },
      stroke: { color: '000000', width: 2, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
    });
    const leader = plan.resolvedGeometry.paths.find(p => p.stroke !== null)!;
    const [ex, ey] = leader.path.at(-1)!.args;
    expect(Math.hypot(ex - 100, ey - 1)).toBeCloseTo(9, 6);
    expect(ey).toBeCloseTo(0, 6);
  });
  it('keeps origin translation compatible with a cloned immutable local plan', () => {
    const original = retained(rectangle), copy = structuredClone(original);
    const moved = { ...copy, rect: { ...copy.rect, x: 110, y: 220 } };
    expect(requireResolvedDrawingMLGeometry(moved, 1)).toEqual(original.resolvedGeometry);
    expect(drawingMLGeometryBounds(moved)).toEqual({ x: 109, y: 219, w: 32, h: 42 });
    expect(drawingMLGeometryBounds(original, { a: 0, b: 1, c: -1, d: 0, e: 100, f: 200 })).toEqual({ x: 39, y: 209, w: 42, h: 32 });
    expect(() => requireResolvedDrawingMLGeometry({ ...copy, rect: { ...copy.rect, w: 31 } }, 1)).toThrow(/stale/);
  });
  it('fails geometry work atomically and refuses nonfinite resolved commands', () => {
    const before = JSON.stringify(rectangle);
    expect(() => resolveDrawingMLGeometry(rectangle, 1, 1)).toThrow(GeometryWorkBudgetError);
    expect(JSON.stringify(rectangle)).toBe(before);
    const acquired = retained(rectangle);
    expect(() => requireResolvedDrawingMLGeometry({ ...acquired, resolvedGeometry: { ...acquired.resolvedGeometry, workUnits: 1 } }, 1)).toThrow(GeometryWorkBudgetError);
    expect(() => requireResolvedDrawingMLGeometry({ ...acquired, rect: { ...acquired.rect, x: NaN } }, 1)).toThrow(/frame/);
    expect(() => resolveDrawingMLGeometry({ ...rectangle, geometry: { kind: 'custom', subpaths: [[{ cmd: 'moveTo', x: NaN, y: 0 }]] } }, 1)).toThrow(/Non-finite/);
  });
});

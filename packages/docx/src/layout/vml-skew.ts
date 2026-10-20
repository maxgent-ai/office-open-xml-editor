import { composeAffine, inverseMapAffinePoint } from './affine.js';
import { GeometryWorkBudgetError, resolveDrawingMLGeometry } from '@silurus/ooxml-core';
import type { DrawingMLShapePaintPlan, GeometryPath } from '@silurus/ooxml-core';
import type { DeepReadonly, VmlShapeAcquisitionFacts, Matrix2DData } from './types.js';
import { projectPath } from './vml-project-path.js';

type Point = readonly [number, number];
type AcquiredShapePlan = DrawingMLShapePaintPlan & {
  resolvedGeometry: NonNullable<DrawingMLShapePaintPlan['resolvedGeometry']>;
};

/** Part 4 §19.2.2.31: transform acquired vectors, including separately
 * acquired arrow contours. Keep the original image/brush frame and physical
 * pen. Re-resolving an arrow from a transformed/retracted shaft would change
 * its tip and discard its skew. No image data is transformed here. */
export function acquireVmlAffineGeometry(
  plan: DrawingMLShapePaintPlan, matrix: Matrix2DData, maximumWork = Infinity,
): AcquiredShapePlan {
  if (matrix.a === 1 && matrix.b === 0 && matrix.c === 0 && matrix.d === 1 && matrix.e === 0 && matrix.f === 0)
    return { ...plan, resolvedGeometry: plan.resolvedGeometry ?? resolveDrawingMLGeometry(plan, 1, maximumWork) };
  return acquireVmlProjectedGeometry(plan, ([x, y]) => [matrix.a * x + matrix.c * y + matrix.e, matrix.b * x + matrix.d * y + matrix.f], maximumWork);
}
/** Project every retained vector through a single acquired point map. */
export function acquireVmlProjectedGeometry(
  plan: DrawingMLShapePaintPlan, effect: (p: Point) => Point, maximumWork = Infinity,
): AcquiredShapePlan {
  maximumWork = Math.min(maximumWork, 1_048_576);
  const geometry = plan.resolvedGeometry ?? resolveDrawingMLGeometry(plan, 1, maximumWork);
  let work = geometry.workUnits;
  const charge = (n = 1) => {
    work += n;
    if (work > maximumWork)
      throw new GeometryWorkBudgetError('VML affine geometry work budget exceeded');
  };
  const project = ([x, y]: Point): Point => {
    charge();
    const dx = plan.rect.x - geometry.originX, dy = plan.rect.y - geometry.originY;
    const q = effect([x + dx, y + dy]);
    return [q[0] - dx, q[1] - dy];
  };
  const path = (source: GeometryPath, map = project): GeometryPath => {
    return projectPath(source, map, charge).map(command => {
      // projectPath supplies only its acquired polygon, in output points.
      if (command.cmd === 'close') {
        charge();
        return { op: 'closePath' as const, args: [] };
      }
      if (command.cmd !== 'moveTo' && command.cmd !== 'lineTo')
        throw new Error('VML projection returned an unacquired curve');
      charge(3);
      return { op: command.cmd, args: [command.x, command.y] };
    });
  };
  const paths = geometry.paths.map(p => ({ ...p, path: path(p.path) }));
  const fillSilhouette = path(geometry.fillSilhouette);
  const arrows = geometry.arrows.map(arrow => ({ ...arrow, tipX: 0, tipY: 0, angle: 0,
    geometry: { ...arrow.geometry, path: path(arrow.geometry.path, ([x, y]) => project([
        arrow.tipX + Math.cos(arrow.angle) * x - Math.sin(arrow.angle) * y,
        arrow.tipY + Math.sin(arrow.angle) * x + Math.cos(arrow.angle) * y,
      ])) },
  }));
  return { ...plan, resolvedGeometry: { ...geometry, paths, fillSilhouette, arrows, workUnits: work } };
}
export function acquireVmlSkew(plan: DrawingMLShapePaintPlan, skew: DeepReadonly<NonNullable<VmlShapeAcquisitionFacts['vmlSkew']>>, maximumWork = Infinity): ReturnType<typeof acquireVmlAffineGeometry> {
  return acquireVmlAffineGeometry(plan, acquireVmlSkewMatrix(plan, skew), maximumWork);
}
/** Conjugate the post-host skew into the retained host's local point frame.
 * Keeping the original host transform preserves image rotation/reflection;
 * §19.2.2.31 excludes only skew from image data, not the host transform. */
export function acquireVmlSkewMatrix(plan: Pick<DrawingMLShapePaintPlan, 'rect' | 'transform'>, skew: DeepReadonly<NonNullable<VmlShapeAcquisitionFacts['vmlSkew']>>): Matrix2DData {
  const { x, y, w, h } = plan.rect;
  const [sxx, sxy, syx, syy, px, py] = skew.matrix;
  if (px !== 0 || py !== 0)
    throw new Error('Projective VML skew requires resolved reciprocal-factor semantics');
  const cx = x + w / 2, cy = y + h / 2;
  const ox = cx + w * skew.origin[0], oy = cy + h * skew.origin[1];
  const dx = skew.offset.xPt + skew.offset.xFraction * w;
  const dy = skew.offset.yPt + skew.offset.yFraction * h;
  const theta = plan.transform.rotationDeg * Math.PI / 180;
  const cos = Math.cos(theta), sin = Math.sin(theta);
  const fx = plan.transform.flipH ? -1 : 1, fy = plan.transform.flipV ? -1 : 1;
  const project = (x: number, y: number) => {
    const vx = (x - cx) * fx, vy = (y - cy) * fy;
    const u = cx + cos * vx - sin * vy - ox, v = cy + sin * vx + cos * vy - oy;
    const sx = ox + sxx * u + syx * v + dx - cx;
    const sy = oy + sxy * u + syy * v + dy - cy;
    return { x: cx + fx * (cos * sx + sin * sy), y: cy + fy * (-sin * sx + cos * sy) };
  };
  const p = project(0, 0), a = project(1, 0), b = project(0, 1);
  return { a: a.x - p.x, b: a.y - p.y, c: b.x - p.x, d: b.y - p.y, e: p.x, f: p.y };
}
/** Express an ancestor's page-space vector effect in a retained host frame.
 * Host rotation/reflection remains the painter's single outer transform. */
export function conjugateVmlHostMatrix(plan: Pick<DrawingMLShapePaintPlan, 'rect' | 'transform'>, matrix: Matrix2DData): Matrix2DData {
  const { x, y, w, h } = plan.rect, theta = plan.transform.rotationDeg * Math.PI / 180;
  const fx = plan.transform.flipH ? -1 : 1, fy = plan.transform.flipV ? -1 : 1;
  const a = Math.cos(theta) * fx, b = Math.sin(theta) * fx, c = -Math.sin(theta) * fy, d = Math.cos(theta) * fy;
  const cx = x + w / 2, cy = y + h / 2;
  const host = { a, b, c, d, e: cx - a * cx - c * cy, f: cy - b * cx - d * cy };
  const project = (xPt: number, yPt: number) => {
    const p = inverseMapAffinePoint(host, { xPt, yPt });
    if (!p)
      throw new RangeError('Invalid VML host orientation');
    return p;
  };
  const o = project(0, 0), u = project(1, 0), v = project(0, 1);
  return composeAffine({ a: u.xPt - o.xPt, b: u.yPt - o.yPt, c: v.xPt - o.xPt, d: v.yPt - o.yPt, e: o.xPt, f: o.yPt }, composeAffine(matrix, host));
}

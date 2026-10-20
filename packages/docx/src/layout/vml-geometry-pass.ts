import { GeometryWorkBudgetError, resolveDrawingMLGeometry } from '@silurus/ooxml-core';
import type { DrawingMLShapePaintPlan } from '@silurus/ooxml-core';
/** A fill/pen/shadow pass changes paint, never already acquired vectors.
 * Resolve the new scheduling/style binding, retaining source contours and
 * arrow tips. This also avoids applying endpoint retraction a second time. */
export function resolveVmlGeometryPass(plan: DrawingMLShapePaintPlan, maximumWork = Infinity): NonNullable<DrawingMLShapePaintPlan['resolvedGeometry']> {
  const source = plan.resolvedGeometry;
  const resolved = resolveDrawingMLGeometry(plan, 1, maximumWork);
  if (!source)
    return resolved;
  const workUnits = resolved.workUnits + source.workUnits;
  if (workUnits > maximumWork)
    throw new GeometryWorkBudgetError('VML geometry pass work budget exceeded');
  if (!plan.stroke)
    return { ...resolved, originX: source.originX, originY: source.originY,
      paths: source.paths.map(p => ({ ...p, stroke: null })), fillSilhouette: source.fillSilhouette,
      arrows: [], workUnits };
  if (resolved.paths.length !== source.paths.length)
    throw new Error('VML geometry pass changed path scheduling');
  return { ...resolved, originX: source.originX, originY: source.originY,
    paths: resolved.paths.map((p, i) => ({ ...p, path: source.paths[i]!.path })),
    fillSilhouette: source.fillSilhouette,
    arrows: resolved.arrows.map((arrow, i) => {
      const original = source.arrows[i];
      if (!original)
        throw new Error('VML geometry pass changed arrow scheduling');
      return { ...arrow, tipX: original.tipX, tipY: original.tipY, angle: original.angle,
        geometry: { ...arrow.geometry, path: original.geometry.path } };
    }), workUnits };
}

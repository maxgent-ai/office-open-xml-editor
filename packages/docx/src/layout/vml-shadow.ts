import { acquireVmlProjectedGeometry } from './vml-skew.js';
import { resolveVmlGeometryPass } from './vml-geometry-pass.js';
import { drawingMLGeometryBounds, GeometryWorkBudgetError } from '@silurus/ooxml-core';
import type { DrawingMLShapePaintPlan } from '@silurus/ooxml-core';
import type { ShapeRun } from '../types.js';
import type { DeepReadonly, DrawingLayout, DrawingPaintCommand, LayoutRect } from './types.js';

type ShapeCommand = Extract<DrawingPaintCommand, { kind: 'drawingml-shape' }>;
type Point = readonly [number, number];

/** Effect ink can leave the source allocation box. Keep flow/wrap ownership
 * unchanged and enclose the retained shadow commands in the drawing's ink box,
 * including upright story transforms. No painter or text service is involved. */
export function shadowInkBounds(commands: readonly DrawingPaintCommand[], base: LayoutRect,
  owner?: DrawingLayout['transform']): LayoutRect {
  let left = base.xPt, top = base.yPt, right = left + base.widthPt, bottom = top + base.heightPt;
  const include = (x: number, y: number) => {
    left = Math.min(left, x); top = Math.min(top, y); right = Math.max(right, x); bottom = Math.max(bottom, y);
  };
  for (const command of commands) {
    if (command.kind === 'drawingml-shape') {
      const box = drawingMLGeometryBounds(command.plan, owner);
      if (box) { include(box.x, box.y); include(box.x + box.w, box.y + box.h); }
    } else if (command.kind === 'watermark-text') {
      const r = command.rect, cx = r.xPt + r.widthPt / 2, cy = r.yPt + r.heightPt / 2;
      const w = command.fitShape ? r.widthPt : command.sourceBounds.widthPt;
      const h = command.fitShape ? r.heightPt : command.sourceBounds.heightPt;
      const angle = command.rotationDeg * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle);
      const ink = command.maskBounds ?? {xPt:r.xPt,yPt:r.yPt,widthPt:w,heightPt:h};
      for (const [dx, dy] of [[0, 0], [ink.widthPt, 0], [ink.widthPt, ink.heightPt], [0, ink.heightPt]]) {
        const vx = (ink.xPt + dx - cx)*(command.flipH ? -1 : 1),
          vy = (ink.yPt + dy - cy)*(command.flipV ? -1 : 1);
        const x = cx + vx * cos - vy * sin, y = cy + vx * sin + vy * cos;
        include(owner ? owner.a * x + owner.c * y + owner.e : x,
          owner ? owner.b * x + owner.d * y + owner.f : y);
      }
    }
  }
  return { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
}

/** Part 4 §19.1.2.18/§19.1.3.6: a single translated silhouette, two for
 * double/emboss, or a homogeneous perspective matrix about the shadow origin.
 * Acquisition owns all geometry; paint only replays ordinary shape commands.
 * Absolute offsets are points, fractions follow the final shape extent. */
export function planVmlShadows(
  shape: DeepReadonly<ShapeRun>, plan: DrawingMLShapePaintPlan,
  maximumWork = Infinity,
  effectRect = plan.rect,
): readonly ShapeCommand[] {
  const shadow = shape.vmlShadow;
  if (!shadow || shadow.opacity === 0) return [];
  // Library resource policy: perspective subdivision can amplify a tiny
  // authored path. Bound acquired operations independently of source length.
  maximumWork = Math.min(maximumWork, 1_048_576);
  const { x, y, w, h } = plan.rect;
  const effect = effectRect;
  let work = 0;
  const charge = (n = 1) => {
    work += n;
    if (work > maximumWork) throw new GeometryWorkBudgetError('VML shadow geometry allowance exceeded');
  };
  const retained: ShapeCommand[] = [];
  const offsets = shadow.kind === 'double' || shadow.kind === 'emboss'
    ? [[shadow.offset, shadow.color], [shadow.offset2, shadow.color2]] as const
    : [[shadow.offset, shadow.color]] as const;
  for (const [offset, color] of offsets) {
    const dx = offset.xPt + offset.xFraction * effect.w, dy = offset.yPt + offset.yFraction * effect.h;
    // The shadow uses its own solid pen. Omit the source brush entirely;
    // retained layout data cannot carry an explicit undefined property.
    const shadowStroke = plan.stroke ? { ...plan.stroke, color } : null;
    if (shadowStroke) delete shadowStroke.fill;
    let shadowPlan: DrawingMLShapePaintPlan = {
      ...plan, rect: { x: x + dx, y: y + dy, w, h },
      fill: !shadow.obscured || shape.fill || shape.vmlPicture ? { fillType: 'solid', color } : null,
      stroke: shadowStroke,
      resolvedGeometry: plan.resolvedGeometry,
    };
    if (shadow.kind === 'perspective') {
      const [sxx, sxy, syx, syy, px, py] = shadow.matrix;
      const ox = effect.x + effect.w * (0.5 + shadow.origin[0]), oy = effect.y + effect.h * (0.5 + shadow.origin[1]);
      const theta = plan.transform.rotationDeg * Math.PI / 180;
      const cos = Math.cos(theta), sin = Math.sin(theta);
      const project = (point: Point): Point => {
        charge();
        // MS-OI29500 §2.1.1807(c) places rotation/reflection before
        // the shadow matrix. Its row-vector convention maps x with sxx/syx
        // and y with sxy/syy. Matrix p values are reciprocal EMU (or reciprocal
        // extent fractions); use a homogeneous divisor in that unit domain.
        const vx0 = (point[0] - (x + w / 2)) * (plan.transform.flipH ? -1 : 1);
        const vy0 = (point[1] - (y + h / 2)) * (plan.transform.flipV ? -1 : 1);
        const vx = x + w / 2 + cos * vx0 - sin * vy0 - ox;
        const vy = y + h / 2 + sin * vx0 + cos * vy0 - oy;
        const z = 1 + (shadow.absolutePerspective ? 12700 * (px * vx + py * vy)
          : px * vx / (w || 1) + py * vy / (h || 1));
        if (z === 0 || !Number.isFinite(z)) throw new RangeError('VML shadow crosses the perspective horizon');
        return [ox + (sxx * vx + syx * vy + dx) / z, oy + (sxy * vx + syy * vy + dy) / z];
      };
      const geometry = plan.resolvedGeometry;
      if (!geometry) throw new Error('VML shadow requires acquired source geometry');
      // The same acquired geometry owns arrow contours and shafts. Project
      // both, rather than regenerating heads at the projected endpoints.
      const styled = {...shadowPlan, rect: plan.rect};
      const sourceGeometry = resolveVmlGeometryPass(styled, maximumWork-work);
      charge(sourceGeometry.workUnits);
      const projected = acquireVmlProjectedGeometry(
        {...styled, resolvedGeometry: sourceGeometry}, project, maximumWork-work,
      );
      charge(projected.resolvedGeometry.workUnits);
      retained.push({kind: 'drawingml-shape', shadowPass: true,
        plan: {...projected, transform: {rotationDeg: 0, flipH: false, flipV: false}},
        opacity: shadow.opacity});
      continue;
    }
    const resolvedGeometry = resolveVmlGeometryPass(shadowPlan, maximumWork - work);
    charge(resolvedGeometry.workUnits);
    retained.push({ kind: 'drawingml-shape', shadowPass: true, plan: { ...shadowPlan, resolvedGeometry }, opacity: shadow.opacity });
  }
  return retained;
}

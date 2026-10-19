import { drawingMLGeometryBounds } from '@silurus/ooxml-core';
import { composeAffine, mapAffinePoint } from './affine.js';
import { snapshotPlainData } from './plain-data.js';
import type { ShapeDrawingPlanResult } from './shape-drawing-plan.js';
import type { DeepReadonly, DrawingPaintCommand, LayoutRect, VmlShapeAcquisitionFacts } from './types.js';
import { acquireVmlAffineGeometry, conjugateVmlHostMatrix } from './vml-skew.js';
type ShapeCommand = Extract<DrawingPaintCommand,
  { kind: 'noop' | 'drawingml-shape' | 'drawingml-image-fill' | 'watermark-text' }>;

/** CT_Group/Part 4 §19.2.2.31: transform acquired descendant vectors,
 * retaining each image's original brush frame. The parser composes ancestors
 * once relative to the leaf; bind that effect to this occurrence's placement. */
export function acquireVmlGroupVectors(
  result: ShapeDrawingPlanResult,
  effect: DeepReadonly<NonNullable<VmlShapeAcquisitionFacts['vmlGroupTransform']>>,
  bounds: LayoutRect,
  maximumWork = Infinity,
): ShapeDrawingPlanResult {
  if (result.status !== 'planned')
    return result;
  const [a, b, c, d, e, f] = effect;
  const matrix = { a, b, c, d,
    e: bounds.xPt + e - a * bounds.xPt - c * bounds.yPt,
    f: bounds.yPt + f - b * bounds.xPt - d * bounds.yPt,
  };
  let remaining = maximumWork;
  const transformCommand = (command: ShapeCommand): ShapeCommand => {
    if (command.kind === 'noop')
      return command;
    if (command.kind === 'watermark-text') {
      const r = command.rect;
      const local = conjugateVmlHostMatrix({
        rect: { x: r.xPt, y: r.yPt, w: r.widthPt, h: r.heightPt },
        transform: { rotationDeg: command.rotationDeg,
          flipH: command.flipH ?? false, flipV: command.flipV ?? false },
      }, matrix);
      const mask = command.maskBounds;
      const points = mask ? [
        { xPt: mask.xPt, yPt: mask.yPt },
        { xPt: mask.xPt + mask.widthPt, yPt: mask.yPt },
        { xPt: mask.xPt + mask.widthPt, yPt: mask.yPt + mask.heightPt },
        { xPt: mask.xPt, yPt: mask.yPt + mask.heightPt },
      ].map(p => mapAffinePoint(local, p)) : [];
      const left = Math.min(...points.map(p => p.xPt));
      const top = Math.min(...points.map(p => p.yPt));
      return { ...command,
        glyphTransform: command.glyphTransform
          ? composeAffine(local, command.glyphTransform) : local,
        ...(mask ? { maskBounds: {
            xPt: left, yPt: top,
            widthPt: Math.max(...points.map(p => p.xPt)) - left,
            heightPt: Math.max(...points.map(p => p.yPt)) - top,
          } } : {}),
      };
    }
    const local = conjugateVmlHostMatrix(command.plan, matrix);
    const plan = acquireVmlAffineGeometry(command.plan, local, remaining);
    remaining -= plan.resolvedGeometry.workUnits;
    const ink = command.kind === 'drawingml-image-fill' && command.strokeMask
      ? drawingMLGeometryBounds({ ...plan, transform: {
          rotationDeg: 0, flipH: false, flipV: false,
        } }) : null;
    return { ...command, plan, ...(ink ? { maskBounds: {
          xPt: ink.x, yPt: ink.y, widthPt: ink.w, heightPt: ink.h,
        } } : {}) };
  };
  const before = result.before?.map(transformCommand)
    .filter((c): c is Exclude<ShapeCommand, { kind: 'noop' }> => c.kind !== 'noop');
  return snapshotPlainData({ ...result, ...(before ? { before } : {}),
    command: transformCommand(result.command),
  }, 'VML group vectors');
}

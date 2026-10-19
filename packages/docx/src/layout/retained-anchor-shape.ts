import type { AnchorAcquisitionInput } from './anchor-input.js';
import type { ParagraphShapeRun, TextLayoutService } from './text.js';
import type { LayoutRect } from './types.js';
import { planShapeDrawing, type ShapeDrawingPlanResult } from './shape-drawing-plan.js';

/** Same planner as ordinary anchored groups. Resolved child rotation/reflection
 * belongs to the child, not the outer group. Do not apply the raw transform
 * chain again: native acquisition has already resolved it into this frame. */
export function planRetainedAnchorShape(
  run: ParagraphShapeRun,
  acquisition: AnchorAcquisitionInput,
  rect: LayoutRect,
  text?: TextLayoutService,
  imageFillResourceKey?: string,
  maximumGeometryWork?: number,
  vmlPictureResourceKey?: string,
  vmlStrokeResourceKey?: string,
  vmlEdgeResourcePrefix?: string,
): ShapeDrawingPlanResult {
  const child = acquisition.group?.resolvedChildFrame;
  return planShapeDrawing(child ? {
    ...run, rotation: child.rotationDeg, flipH: child.flipH, flipV: child.flipV,
  } : run, rect, text, run.vmlTextPathInput, imageFillResourceKey, maximumGeometryWork, vmlPictureResourceKey, vmlStrokeResourceKey, vmlEdgeResourcePrefix);
}

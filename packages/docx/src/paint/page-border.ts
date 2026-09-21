import type { PageBorderLayout } from '../layout/types.js';
import { composeAffine, scaleAffine } from './affine.js';
import { paintStrokeSegment } from './canvas-border.js';
import { applyCanvasTransform } from './canvas-transform.js';
import { canvasPaintFrame } from './deferred-paint-frame.js';
import type { CanvasPaintContext } from './types.js';

export function paintPageBorderLayout(
  pageBorder: PageBorderLayout,
  context: CanvasPaintContext,
): void {
  const pointToCss = composeAffine(
    context.pointToCss ?? scaleAffine(context.scale),
    pageBorder.logicalToPhysical,
  );
  const borderContext: CanvasPaintContext = {
    ...context,
    pointToCss,
  };
  const frame = canvasPaintFrame(context.ctx, () => {
    applyCanvasTransform(context.ctx, pageBorder.logicalToPhysical);
  });
  frame(() => {
    for (const segment of pageBorder.segments) {
      // Preserve the historical half-CSS-pixel minimum used by DOCX page
      // borders without changing the retained authored point width.
      paintStrokeSegment(segment, borderContext, 0.5);
    }
  })();
}

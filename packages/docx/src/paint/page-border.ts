import type { PageBorderLayout } from '../layout/types.js';
import { composeAffine, scaleAffine } from './affine.js';
import { oneDevicePixelCssWidth, paintStrokeSegment } from './canvas-border.js';
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
      // Keep authored page-border layout geometry while sharing the raster
      // floor used by the other retained WordprocessingML borders.
      paintStrokeSegment(segment, borderContext, oneDevicePixelCssWidth(context));
    }
  })();
}

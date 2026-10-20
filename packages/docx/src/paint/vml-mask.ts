import { createAuxCanvas } from '@silurus/ooxml-core';
import type { LayoutRect } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';

/** Bounded temporary surface for VML compositing. No resources are cached. */
export function paintVmlRasterPass(
  bounds: LayoutRect,
  context: CanvasPaintContext,
  paint: (ctx: CanvasRenderingContext2D) => void,
): void {
  if (!(bounds.widthPt > 0 && bounds.heightPt > 0)) return;
  const scale = Math.min(
    context.scale * context.dpr,
    2048 / Math.max(bounds.widthPt, bounds.heightPt),
  );
  const canvas = createAuxCanvas(bounds.widthPt * scale, bounds.heightPt * scale);
  if (!canvas) throw new Error('Canvas cannot allocate a VML brush surface');
  try {
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | null;
    if (!ctx) throw new Error('Canvas cannot paint a VML brush surface');
    ctx.scale(scale, scale);
    ctx.translate(-bounds.xPt, -bounds.yPt);
    paint(ctx);
    context.ctx.drawImage(canvas, bounds.xPt, bounds.yPt, bounds.widthPt, bounds.heightPt);
  } finally {
    canvas.width = 1;
    canvas.height = 1;
  }
}

/** Replay an acquired text/pen mask with a retained image brush. The scratch
 * two surfaces are each bounded to 2048² pixels (32 MiB together) and
 * released after this pass; document size cannot amplify allocation. Resource
 * ink is painted first, so an unavailable optional resource stays transparent. */
export function paintVmlImageMask(
  resourceKey: string,
  bounds: LayoutRect,
  context: CanvasPaintContext,
  mask: (ctx: CanvasRenderingContext2D) => void,
  imageBounds: LayoutRect = bounds,
  imageTransform?: Readonly<{ rotationDeg: number; flipH?: boolean; flipV?: boolean }>,
): void {
  paintVmlRasterPass(bounds, context, (ctx) => {
    ctx.save();
    if (imageTransform) {
      const x = imageBounds.xPt + imageBounds.widthPt / 2,
        y = imageBounds.yPt + imageBounds.heightPt / 2;
      ctx.translate(x, y);
      ctx.scale(imageTransform.flipH ? -1 : 1, imageTransform.flipV ? -1 : 1);
      ctx.rotate((imageTransform.rotationDeg * Math.PI) / 180);
      ctx.translate(-x, -y);
    }
    context.resources.paint(resourceKey, 'image', imageBounds, ctx, bounds);
    ctx.restore();
    ctx.globalCompositeOperation = 'destination-in';
    paintVmlRasterPass(bounds, { ...context, ctx }, mask);
  });
}

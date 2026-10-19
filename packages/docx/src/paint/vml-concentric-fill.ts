import { paintVmlRasterPass } from './vml-mask.js';
import { withDrawingMLShapeTransform } from '@silurus/ooxml-core';
import type { DrawingPaintCommand } from '../layout/types.js';
import type { CanvasPaintContext } from './types.js';

/** Replay already acquired path operations and affine ring frames. */
export function paintVmlConcentricFill(
  command: Extract<DrawingPaintCommand, { kind: 'drawingml-shape' }>,
  context: CanvasPaintContext,
): void {
  const fill = command.concentricFill;
  if (!fill) return;
  const { x, y, w, h } = command.plan.rect;
  const geometry = command.plan.resolvedGeometry;
  withDrawingMLShapeTransform(context.ctx as CanvasRenderingContext2D, command.plan, () => {
    paintVmlRasterPass(
      fill.bounds ?? { xPt: x, yPt: y, widthPt: w, heightPt: h },
      context,
      (ctx) => {
        ctx.globalAlpha *= command.opacity ?? 1;
        const append = (path: (typeof geometry.paths)[number]['path'], amount: number) => {
          const sx = fill.focus.width + (1 - fill.focus.width) * amount;
          const sy = fill.focus.height + (1 - fill.focus.height) * amount;
          ctx.save();
          ctx.translate(x + fill.focus.l * w * (1 - amount), y + fill.focus.t * h * (1 - amount));
          ctx.scale(sx, sy);
          ctx.translate(-geometry.originX, -geometry.originY);
          for (const operation of path) {
            // These operations were validated/sealed by geometry acquisition.
            const args = operation.args;
            if (operation.op === 'roundRect') {
              ctx.roundRect(
                args[0]!,
                args[1]!,
                args[2]!,
                args[3]!,
                args[4] === 0
                  ? args[5]!
                  : Array.from({ length: (args.length - 5) / 2 }, (_, i) => ({
                      x: args[5 + i * 2]!,
                      y: args[6 + i * 2]!,
                    })),
              );
            } else {
              const method = ctx[operation.op] as (...args: number[]) => void;
              method.apply(ctx, [...args]);
            }
          }
          ctx.restore();
        };
        for (const path of geometry.paths) {
          if (path.fill === 'none') continue;
          for (const ring of fill.rings) {
            ctx.save();
            ctx.beginPath();
            append(path.path, ring.outside);
            ctx.clip(path.fillRule);
            // Copy inside this ring's outer contour, preserving already painted
            // outer rings. The final surface composites once, so translucent stops
            // are never multiplied by the number of concentric copies.
            ctx.globalCompositeOperation = 'copy';
            ctx.fillStyle = ring.color;
            ctx.fillRect(
              fill.bounds?.xPt ?? x,
              fill.bounds?.yPt ?? y,
              fill.bounds?.widthPt ?? w,
              fill.bounds?.heightPt ?? h,
            );
            ctx.restore();
          }
          if (fill.focus.width > 0 && fill.focus.height > 0) {
            ctx.save();
            ctx.beginPath();
            append(path.path, 0);
            ctx.clip(path.fillRule);
            ctx.globalCompositeOperation = 'copy';
            ctx.fillStyle = fill.centerColor;
            ctx.fillRect(
              fill.bounds?.xPt ?? x,
              fill.bounds?.yPt ?? y,
              fill.bounds?.widthPt ?? w,
              fill.bounds?.heightPt ?? h,
            );
            ctx.restore();
          }
        }
      },
    );
  });
}

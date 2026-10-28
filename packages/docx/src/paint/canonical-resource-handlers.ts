import { paintOptionalImagePlaceholder, renderChart } from '@silurus/ooxml-core';
import { withChartFontSubstitutions } from '@silurus/ooxml-core/internal/chart-font-sources';
import type {
  ChartImageLookup,
  ChartThreeDRenderer,
  ChartRegionMapRenderer,
  ChartExRenderer,
} from '@silurus/ooxml-core';
import type {
  LayoutRect,
} from '../layout/types.js';
import {
  isUnavailablePaintResourceHandle,
  type ResolvedPaintResource,
} from './resource-session.js';
import { drawableHandle, paintImageResource } from './image-resource.js';
export { paintImageResource } from './image-resource.js';
import type {
  CanvasPaintResourceHandlers,
  PaintCanvas2D,
} from './types.js';


function paintDrawableResource(
  resource: ResolvedPaintResource<'math'>,
  bounds: LayoutRect,
  ctx: PaintCanvas2D,
): void {
  const drawable = drawableHandle(resource);
  if (!drawable) {
    if (isUnavailablePaintResourceHandle(resource.handle)
      && resource.handle.placeholder === 'tiff') {
      paintOptionalImagePlaceholder(ctx as CanvasRenderingContext2D, 'tiff', {
        x: bounds.xPt,
        y: bounds.yPt,
        width: bounds.widthPt,
        height: bounds.heightPt,
      });
    }
    return;
  }
  ctx.drawImage(
    drawable,
    bounds.xPt,
    bounds.yPt,
    bounds.widthPt,
    bounds.heightPt,
  );
}

export function createCanonicalCanvasPaintResourceHandlers(
  threeD?: ChartThreeDRenderer,
  regionMap?: ChartRegionMapRenderer,
  chartImageLookup?: ChartImageLookup,
  chartEx?: ChartExRenderer,
  /** `ChartModel.fontSubstitutions` of the document's loaded font sources. */
  chartFontSubstitutions?: Readonly<Record<string, string>>,
): CanvasPaintResourceHandlers {
  return Object.freeze({
  image(resource, bounds, ctx, coverageBounds) {
    paintImageResource(resource, bounds, ctx, coverageBounds);
  },
  chart(resource, bounds, ctx) {
    // paintLayoutPage has already installed the point-to-device CTM. Passing 1
    // keeps chart font/line point sizes in that same space instead of scaling twice.
    renderChart(
      ctx as CanvasRenderingContext2D,
      withChartFontSubstitutions(
        resource.descriptor.model as import('@silurus/ooxml-core').ChartModel,
        chartFontSubstitutions,
      ),
      { x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt },
      1,
      0,
      threeD,
      regionMap,
      chartImageLookup,
      chartEx,
    );
  },
  math(resource, bounds, ctx) {
    paintDrawableResource(resource, bounds, ctx);
  },
  'picture-bullet'(resource, bounds, ctx) {
    // Picture bullets retain image crop/rotation/reflection. The marker box
    // is already resolved; use the same image painter without resizing it.
    paintImageResource(resource, bounds, ctx);
  },
  });
}

export const canonicalCanvasPaintResourceHandlers: CanvasPaintResourceHandlers =
  createCanonicalCanvasPaintResourceHandlers();

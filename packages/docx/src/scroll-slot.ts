import type { StaticCanvasRenderDispatcher } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { ReadOnlyCommentMarginGeometry } from '@silurus/ooxml-core/internal/read-only-comment-decoration';
import type { DocxTextRunInfo } from './renderer';

/** One mounted page. `canvas` is the drawn page; `textLayer` the optional
 *  per-page selection overlay (both render modes — IX6 ships the worker's run
 *  geometry back beside the bitmap). `renderedPage` guards against
 *  re-rendering a recycled slot for a page whose render is still in flight. */
export interface DocxScrollSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  highlightLayer: HTMLDivElement;
  elementLayer: HTMLDivElement | null;
  commentTintLayer: HTMLDivElement | null;
  commentMargin: HTMLDivElement | null;
  commentDecorationLayer: HTMLDivElement | null;
  commentRuns: readonly Readonly<DocxTextRunInfo>[];
  commentGeometry: ReadOnlyCommentMarginGeometry | null;
  /** page index this slot is currently rendering / has rendered, or -1 when free. */
  renderedPage: number;
  /** The `_scale` at which this slot's on-screen canvas bitmap (and text overlay)
   *  were last rendered, or -1 when unrendered. The flicker-free CSS preview
   *  (design §7) stretches that bitmap to the new layout size on `setScale` and
   *  scales the text overlay by `newScale / renderedScale`; the debounced settle
   *  re-render then repaints at the new scale and updates this to match. */
  renderedScale: number;
  /** Shared single-canvas generation and worker-bitmap ownership primitive. */
  dispatcher: StaticCanvasRenderDispatcher;
}

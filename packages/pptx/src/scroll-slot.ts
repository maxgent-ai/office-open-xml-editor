import type { StaticCanvasRenderDispatcher } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { ReadOnlyCommentMarginGeometry } from '@silurus/ooxml-core/internal/read-only-comment-decoration';
import type { PptxElementBounds } from './element-selection';
import type { PresentationHandle } from './presentation-handle';

/** One mounted slide. `canvas` is the drawn slide; `textLayer` the optional
 *  per-slide selection overlay (both render modes — IX6 ships the worker's run
 *  geometry back beside the bitmap). `renderedSlide` guards against
 *  re-rendering a recycled slot for a slide whose render is still in flight. */
export interface PptxScrollSlot {
  wrapper: HTMLDivElement;
  canvas: HTMLCanvasElement;
  textLayer: HTMLDivElement | null;
  highlightLayer: HTMLDivElement;
  elementLayer: HTMLDivElement | null;
  loadingLayer: HTMLSpanElement;
  commentMarkerLayer: HTMLDivElement | null;
  commentMargin: HTMLDivElement | null;
  commentDecorationLayer: HTMLDivElement | null;
  commentElementBounds: readonly PptxElementBounds[];
  commentGeometry: ReadOnlyCommentMarginGeometry | null;
  commentAnchorSlide: number;
  commentAnchorGeneration: number;
  /** slide index this slot is currently rendering / has rendered, or -1 when free. */
  renderedSlide: number;
  /** The `_scale` at which this slot's on-screen canvas bitmap (and text overlay)
   *  were last rendered, or -1 when unrendered. The flicker-free CSS preview
   *  (design §7) stretches that bitmap to the new layout size on `setScale` and
   *  scales the text overlay by `newScale / renderedScale`; the debounced settle
   *  re-render then repaints at the new scale and updates this to match. */
  renderedScale: number;
  /** Shared single-canvas generation and worker-bitmap ownership primitive. */
  dispatcher: StaticCanvasRenderDispatcher;
  /** Interactive media handle for the canvas currently mounted in this slot. */
  presentationHandle: PresentationHandle | null;
  /** Whether this slot currently owns or awaits an interactive media handle. */
  mediaInteractive: boolean;
  /** Per-slot paint generation across static and interactive canvas ownership. */
  renderGeneration: number;
  /**
   * Per-slot async generation. Unlike the viewer render epoch, this also changes
   * when a pooled slot is recycled and immediately reused for the same slide
   * index, so a late presentSlide() result can never attach to the new owner.
   */
  presentationGeneration: number;
}

/** Indirection module for the render worker. Keep the worker as a standalone
 * module asset: its shared and optional-renderer chunks resolve relative to a
 * network URL, whereas relative imports cannot resolve from an inline Blob URL.
 * This host remains dynamically imported, so main-mode users do not fetch the
 * worker graph. Ordinary OOXML and selected model sources share this asset;
 * only a parse message carrying a source descriptor loads the source sidecar. */
import renderWorkerUrl from './render-worker.ts?worker&url';

export function createRenderWorker(): Worker {
  return new Worker(renderWorkerUrl, { type: 'module' });
}

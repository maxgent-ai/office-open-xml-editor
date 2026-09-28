/** Selected-source render worker, loaded only after source admission. */
import renderWorkerUrl from './render-worker-source.ts?worker&url';

export function createRenderWorker(): Worker {
  return new Worker(renderWorkerUrl, { type: 'module' });
}

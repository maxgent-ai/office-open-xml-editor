/** Selected-source render worker; the ordinary worker asset stays untouched. */
import renderWorkerUrl from './render-worker-source.ts?worker&url';

export function createRenderWorker(): Worker {
  return new Worker(renderWorkerUrl, { type: 'module' });
}

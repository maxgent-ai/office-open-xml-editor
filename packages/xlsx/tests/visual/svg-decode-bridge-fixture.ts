// Browser-side entry for svg-duotone.spec.ts: lets page.evaluate import the
// shared worker SVG bridge by a literal dev-server path.
export {
  WorkerSvgDecodeClient,
  respondToWorkerSvgDecodeRequest,
} from '../../../core/src/worker/svg-decode-bridge.js';

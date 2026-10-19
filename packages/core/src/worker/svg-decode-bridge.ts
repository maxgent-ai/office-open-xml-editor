import { closeImageBitmapIfSupported } from '../image/image-bitmap-lifecycle.js';
import { decodeSvgBlobOnMainThread, type SvgBlobDecoder, type SvgDecodeTarget } from '../image/svg-raster.js';
export { boundedSvgRasterSize, decodeSvgBlobOnMainThread } from '../image/svg-raster.js';
export type { SvgBlobDecoder, SvgDecodeTarget } from '../image/svg-raster.js';

export interface WorkerSvgDecodeRequest extends SvgDecodeTarget {
  readonly kind: 'ooxmlDecodeSvg';
  readonly decodeId: number;
  readonly bytes: ArrayBuffer;
}

export type WorkerSvgDecodeResponse =
  | { readonly kind: 'ooxmlSvgDecoded'; readonly decodeId: number; readonly bitmap: ImageBitmap }
  | { readonly kind: 'ooxmlSvgDecodeFailed'; readonly decodeId: number; readonly message: string };

type WorkerPost = (message: unknown, transfer?: Transferable[]) => void;

function isWorkerSvgDecodeRequest(value: unknown): value is WorkerSvgDecodeRequest {
  return !!value && typeof value === 'object'
    && (value as { kind?: unknown }).kind === 'ooxmlDecodeSvg';
}

export function isWorkerSvgDecodeResponse(value: unknown): value is WorkerSvgDecodeResponse {
  if (!value || typeof value !== 'object') return false;
  const kind = (value as { kind?: unknown }).kind;
  return kind === 'ooxmlSvgDecoded' || kind === 'ooxmlSvgDecodeFailed';
}

/** Main-thread half of the SVG decode bridge. Returns true when `message` was
 * consumed. The async response owns the transferred bitmap; a failed post
 * closes it locally so teardown races cannot leak GPU memory. */
export function respondToWorkerSvgDecodeRequest(
  post: WorkerPost,
  message: unknown,
): boolean {
  if (!isWorkerSvgDecodeRequest(message)) return false;
  void decodeSvgBlobOnMainThread(
    new Blob([message.bytes], { type: 'image/svg+xml' }),
    message,
  ).then((bitmap) => {
    try {
      post({ kind: 'ooxmlSvgDecoded', decodeId: message.decodeId, bitmap }, [bitmap]);
    } catch {
      closeImageBitmapIfSupported(bitmap);
    }
  }).catch((error: unknown) => {
    try {
      post({
        kind: 'ooxmlSvgDecodeFailed',
        decodeId: message.decodeId,
        message: error instanceof Error ? error.message : String(error),
      });
    } catch {
      // The worker has already terminated; no remote waiter remains.
    }
  });
  return true;
}

/** Worker-thread half. It transfers SVG bytes to Window only on a cache miss,
 * correlates the returned ImageBitmap, and exposes a decoder compatible with
 * the shared per-document decoded-bitmap cache. */
export class WorkerSvgDecodeClient {
  private nextId = 1;
  private readonly pending = new Map<
    number,
    { resolve: (bitmap: ImageBitmap) => void; reject: (error: Error) => void }
  >();

  constructor(private readonly post: WorkerPost) {}

  readonly decode: SvgBlobDecoder = async (blob, target = {}) => {
    const decodeId = this.nextId++;
    const bytes = await blob.arrayBuffer();
    return await new Promise<ImageBitmap>((resolve, reject) => {
      this.pending.set(decodeId, { resolve, reject });
      try {
        this.post({ kind: 'ooxmlDecodeSvg', decodeId, bytes, ...target }, [bytes]);
      } catch (error) {
        this.pending.delete(decodeId);
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  };

  accept(message: unknown): boolean {
    if (!isWorkerSvgDecodeResponse(message)) return false;
    const waiter = this.pending.get(message.decodeId);
    if (!waiter) {
      if (message.kind === 'ooxmlSvgDecoded') closeImageBitmapIfSupported(message.bitmap);
      return true;
    }
    this.pending.delete(message.decodeId);
    if (message.kind === 'ooxmlSvgDecoded') waiter.resolve(message.bitmap);
    else waiter.reject(new Error(message.message));
    return true;
  }

  dispose(error = new Error('SVG decode worker disposed')): void {
    for (const waiter of this.pending.values()) waiter.reject(error);
    this.pending.clear();
  }
}

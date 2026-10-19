import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseTypedParserError,
  resourcePolicyForWasm,
} from '@silurus/ooxml-core/worker';
import {
  wasmRuntimeSingleton,
  type WasmArchiveHandle,
  type WasmModuleRuntime,
} from '@silurus/ooxml-core/internal/wasm-runtime-generation';
import { normalizePresentationBootstrap } from '../presentation-preflight.js';
import type { PptxSlideCursorArchive } from '../slide-cursor-operation.js';
import type { PresentationBootstrap } from '../worker-protocol.js';
// @ts-ignore wasm-pack generated module has no declaration entry
import * as pptxWasm from '../wasm/pptx_parser.js';
import { createAbortError, throwIfAborted } from '@silurus/ooxml-core/internal/abort';

const ABORT_MESSAGE = 'PPTX presentation session was aborted';

export interface PptxNodeAcquisitionOptions {
  readonly resourceLimits?: import('@silurus/ooxml-core').OoxmlResourceLimits;
  readonly maxZipEntryBytes?: number;
  readonly debug?: boolean;
  readonly onResourceMetrics?: (metrics: import('@silurus/ooxml-core').OoxmlResourceMetrics) => void;
  readonly signal?: AbortSignal;
}

export interface PptxNodeArchive extends PptxSlideCursorArchive {
  free(): void;
  assert_healthy(): void;
  presentation_bootstrap(): Uint8Array;
  resource_usage(): Uint8Array;
  extract_image(path: string): Uint8Array;
  extract_media(path: string): Uint8Array;
}

interface PptxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
    maxArchiveEntries?: bigint | null,
  ): PptxNodeArchive;
}

const formatRuntime = wasmRuntimeSingleton<PptxNodeArchive>(
  pptxWasm as unknown as WasmModuleRuntime,
  'PPTX',
);

export interface PptxNodeAcquisition {
  readonly archive: PptxNodeArchive;
  readonly bootstrap: PresentationBootstrap;
  readonly metrics: OoxmlResourceMetricsSession;
  closeArchive(): void;
}

/** Format-owned archive acquisition and bootstrap projection for Node. */
export async function acquirePptxNodeSession(
  bytes: Uint8Array,
  wasmModule: WebAssembly.Module,
  options: PptxNodeAcquisitionOptions = {},
): Promise<PptxNodeAcquisition> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'pptx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(bytes.byteLength);
  let handle: WasmArchiveHandle<PptxNodeArchive> | undefined;
  try {
    throwIfAborted(options.signal, ABORT_MESSAGE);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (pptxWasm as unknown as { PptxArchive: PptxArchiveConstructor }).PptxArchive;
    handle = await formatRuntime(wasmModule).open(
      () => new Archive(bytes, maxEntry, maxTotal, maxEntries),
      {
        signal: options.signal,
        abortError: () => createAbortError(ABORT_MESSAGE),
        disposeOnAbort: (archive) => archive.free(),
      },
    );
    throwIfAborted(options.signal, ABORT_MESSAGE);
    const archive = handle.proxy;
    const bootstrap = normalizePresentationBootstrap(JSON.parse(
      new TextDecoder().decode(archive.presentation_bootstrap()),
    ) as PresentationBootstrap);
    metrics.checkpoint('presentation bootstrap ready');
    return {
      archive,
      bootstrap,
      metrics,
      closeArchive: () => handle?.close((current: PptxNodeArchive) => current.free()),
    };
  } catch (error) {
    try { handle?.close((archive: PptxNodeArchive) => archive.free()); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

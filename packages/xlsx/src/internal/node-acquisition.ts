import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  normalizeLoadResourceOptions,
  OoxmlResourceMetricsSession,
  parseTypedParserError,
  resourcePolicyForWasm,
} from '@silurus/ooxml-core/worker';
import {
  WasmRuntimeGenerationHost,
  type WasmArchiveHandle,
  type WasmModuleRuntime,
} from '@silurus/ooxml-core/internal/wasm-runtime-generation';
import type { ParsedWorkbook } from '../types.js';
import { readXlsxArchiveBootstrap } from './archive-bootstrap.js';
// @ts-ignore wasm-pack generated module has no declaration entry
import * as xlsxWasm from '../wasm/xlsx_parser.js';

export interface XlsxNodeAcquisitionOptions {
  readonly resourceLimits?: import('@silurus/ooxml-core').OoxmlResourceLimits;
  readonly maxZipEntryBytes?: number;
  readonly debug?: boolean;
  readonly onResourceMetrics?: (metrics: import('@silurus/ooxml-core').OoxmlResourceMetrics) => void;
  readonly signal?: AbortSignal;
}

export interface XlsxNodeArchive {
  free(): void;
  parse(): Uint8Array;
  resource_usage(): Uint8Array;
  open_sheet_cursor(sheetIndex: number, name: string): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
  sheet_cursor_resource_usage(): Uint8Array;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  close_sheet_cursor(): void;
}

interface XlsxArchiveConstructor {
  new (
    data: Uint8Array,
    maxArchiveEntryBytes?: bigint | null,
    maxTotalInflatedBytes?: bigint | null,
    maxArchiveEntries?: bigint | null,
  ): XlsxNodeArchive;
}

let runtimeModule: WebAssembly.Module | undefined;
let runtimeHost: WasmRuntimeGenerationHost<XlsxNodeArchive> | undefined;

function formatRuntime(wasmModule: WebAssembly.Module): WasmRuntimeGenerationHost<XlsxNodeArchive> {
  if (!runtimeHost) {
    runtimeModule = wasmModule;
    runtimeHost = new WasmRuntimeGenerationHost(
      xlsxWasm as unknown as WasmModuleRuntime,
      wasmModule,
    );
  } else if (runtimeModule !== wasmModule) {
    throw new Error('XLSX runtime was already initialized with another WebAssembly.Module');
  }
  return runtimeHost;
}

export interface XlsxNodeAcquisition {
  readonly archive: XlsxNodeSessionArchive;
  readonly workbookIndex: ParsedWorkbook;
  readonly usage: OoxmlResourceUsageSnapshot | undefined;
  readonly metrics: OoxmlResourceMetricsSession;
  closeArchive(): void;
}

/**
 * The archive a Node XLSX session reads after acquisition: the XLSX parser
 * archive, or a model-source archive whose ZIP accounting is optional.
 */
export interface XlsxNodeSessionArchive
  extends Omit<XlsxNodeArchive, 'free' | 'resource_usage' | 'sheet_cursor_resource_usage'> {
  /** Absent when the source has no ZIP accounting; absence is not zero usage. */
  resource_usage?(): Uint8Array;
  sheet_cursor_resource_usage?(): Uint8Array;
}

/** Format-owned archive acquisition and workbook-index projection for Node. */
export async function acquireXlsxNodeSession(
  bytes: Uint8Array,
  wasmModule: WebAssembly.Module,
  options: XlsxNodeAcquisitionOptions = {},
): Promise<XlsxNodeAcquisition> {
  const resourceOptions = normalizeLoadResourceOptions(options);
  const metrics = new OoxmlResourceMetricsSession({
    enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
    format: 'xlsx',
    mode: 'node',
    scope: 'session',
    policy: resourceOptions.policy,
    onMetrics: resourceOptions.onResourceMetrics,
    emitToConsole: resourceOptions.debug,
  });
  metrics.setSourceBytes(bytes.byteLength);
  let handle: WasmArchiveHandle<XlsxNodeArchive> | undefined;
  try {
    throwIfAborted(options.signal);
    const [maxEntry, maxTotal, maxEntries] = resourcePolicyForWasm(resourceOptions.policy);
    const Archive = (xlsxWasm as unknown as { XlsxArchive: XlsxArchiveConstructor }).XlsxArchive;
    handle = await formatRuntime(wasmModule).open(
      () => new Archive(bytes, maxEntry, maxTotal, maxEntries),
      {
        signal: options.signal,
        abortError: createAbortError,
        disposeOnAbort: (archive) => archive.free(),
      },
    );
    throwIfAborted(options.signal);
    const archive = handle.proxy;
    const { workbook: workbookIndex, usage } = readXlsxArchiveBootstrap(
      () => JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook,
      () => archive.resource_usage(),
    );
    metrics.observeUsage(usage);
    metrics.checkpoint('workbook index ready');
    return {
      archive,
      workbookIndex,
      usage,
      metrics,
      closeArchive: () => handle?.close((current: XlsxNodeArchive) => current.free()),
    };
  } catch (error) {
    try { handle?.close((archive: XlsxNodeArchive) => archive.free()); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  throw createAbortError();
}

function createAbortError(): Error {
  const error = new Error('XLSX workbook session was aborted');
  error.name = 'AbortError';
  return error;
}

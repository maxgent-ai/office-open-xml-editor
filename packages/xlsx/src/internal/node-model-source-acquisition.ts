import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import { normalizeLoadResourceOptions, OoxmlResourceMetricsSession, parseTypedParserError } from '@silurus/ooxml-core/worker';
import type { ParsedWorkbook } from '../types.js';
import { readXlsxArchiveBootstrap } from './archive-bootstrap-source.js';
import type { XlsxNodeAcquisition, XlsxNodeAcquisitionOptions, XlsxNodeSessionArchive } from './node-acquisition.js';

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('XLSX workbook session was aborted');
  error.name = 'AbortError';
  throw error;
}
/** An already-opened archive admitted into the Node XLSX session. */
export interface XlsxOwnedArchiveSource {
  readonly archive: XlsxNodeSessionArchive;
  readonly sourceByteLength: number;
  /** Host layout already configured on the archive (see host-layout.ts). */
  readonly layoutMetrics?: Readonly<{ maximumDigitWidth: number }>;
  closeArchive(): void;
}

/**
 * Admit an already-opened archive (for example one returned by an
 * application-supplied model source) without initializing the XLSX parser
 * WASM. The session owns `closeArchive` from this call on, including failure.
 */
export function acquireXlsxSessionFromArchive(
  owned: XlsxOwnedArchiveSource,
  options: XlsxNodeAcquisitionOptions = {},
): XlsxNodeAcquisition {
  let metrics: OoxmlResourceMetricsSession | undefined;
  try {
    if (!Number.isSafeInteger(owned.sourceByteLength) || owned.sourceByteLength < 0) {
      throw new RangeError('XLSX sourceByteLength must be a non-negative safe integer');
    }
    const resourceOptions = normalizeLoadResourceOptions(options);
    metrics = new OoxmlResourceMetricsSession({
      enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
      format: 'xlsx', mode: 'node', scope: 'session', policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics, emitToConsole: resourceOptions.debug,
    });
    metrics.setSourceBytes(owned.sourceByteLength);
    throwIfAborted(options.signal);
    const archive = owned.archive;
    const { workbook: workbookIndex, usage } = readXlsxArchiveBootstrap(
      () => JSON.parse(new TextDecoder().decode(archive.parse())) as ParsedWorkbook,
      () => archive.resource_usage?.(),
    );
    if (owned.layoutMetrics) {
      workbookIndex.layoutMetrics = { maximumDigitWidth: owned.layoutMetrics.maximumDigitWidth };
    }
    metrics.observeUsage(usage);
    metrics.checkpoint('workbook index ready');
    return { archive, workbookIndex, usage, metrics, closeArchive: () => owned.closeArchive() };
  } catch (error) {
    try { owned.closeArchive(); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics?.fail(normalized);
    throw normalized;
  }
}

import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import { normalizeLoadResourceOptions, OoxmlResourceMetricsSession, parseTypedParserError } from '@silurus/ooxml-core/worker';
import type { PullSessionCommand } from '@silurus/ooxml-core/worker';
import { InProcessPullTransport } from '@silurus/ooxml-core/internal/in-process-pull-transport';
import { DocumentPullWorker, readDocxDocumentCursorUsage } from '../document-pull-worker.js';
import type { AcquiredDocxNodeDocument, DocxNodeAcquisitionOptions, DocxNodePullIdentity, DocxNodePullOptions, DocxNodePullTransport, DocxNodeSessionArchive } from './node-acquisition.js';

export interface DocxOwnedArchiveSource {
  readonly archive: DocxNodeSessionArchive;
  readonly sourceByteLength: number;
  closeArchive(): void;
}
function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('DOCX document session was aborted');
  error.name = 'AbortError';
  throw error;
}

/**
 * Admit an already-opened archive (for example one returned by an
 * application-supplied model source) into the same acknowledged body cursor,
 * accounting and cleanup as an OOXML package. The session owns `closeArchive`
 * from the moment this is called, including on failure.
 */
export async function acquireDocxSessionFromArchive<TResult>(
  source: DocxOwnedArchiveSource,
  options: DocxNodeAcquisitionOptions,
  consume: (
    transport: DocxNodePullTransport,
    identity: DocxNodePullIdentity,
    options: DocxNodePullOptions,
  ) => Promise<TResult>,
): Promise<AcquiredDocxNodeDocument<TResult>> {
  let closed = false;
  const closeArchive = (): void => {
    if (closed) return;
    closed = true;
    source.closeArchive();
  };
  let metrics: OoxmlResourceMetricsSession | undefined;
  try {
    if (!Number.isSafeInteger(source.sourceByteLength) || source.sourceByteLength < 0) {
      throw new RangeError('DOCX sourceByteLength must be a non-negative safe integer');
    }
    const resourceOptions = normalizeLoadResourceOptions(options);
    metrics = new OoxmlResourceMetricsSession({
      enabled: resourceOptions.debug || resourceOptions.onResourceMetrics !== undefined,
      format: 'docx',
      mode: 'node',
      scope: 'session',
      policy: resourceOptions.policy,
      onMetrics: resourceOptions.onResourceMetrics,
      emitToConsole: resourceOptions.debug,
    });
    metrics.setSourceBytes(source.sourceByteLength);
    throwIfAborted(options.signal);
    metrics.checkpoint('container ready');
  } catch (error) {
    try { closeArchive(); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics?.fail(normalized);
    throw normalized;
  }
  return consumeDocxArchive(source.archive, closeArchive, metrics, options, consume);
}

async function consumeDocxArchive<TResult>(
  archive: DocxNodeSessionArchive,
  closeArchive: () => void,
  metrics: OoxmlResourceMetricsSession,
  options: DocxNodeAcquisitionOptions,
  consume: (
    transport: DocxNodePullTransport,
    identity: DocxNodePullIdentity,
    options: DocxNodePullOptions,
  ) => Promise<TResult>,
): Promise<AcquiredDocxNodeDocument<TResult>> {
  let pull: DocumentPullWorker | undefined;
  let transport: DocxNodePullTransport | undefined;
  try {
    pull = new DocumentPullWorker(() => archive);
    const identity = { sessionId: 1, operationId: 1, generation: 1 } as const;
    pull.open(identity);
    transport = new InProcessPullTransport(
      (command, respond) => pull?.dispatch(command as PullSessionCommand<number>, respond),
      () => undefined,
    );
    let usage: OoxmlResourceUsageSnapshot | undefined;
    const result = await consume(transport, identity, {
      signal: options.signal,
      onUsage: (checkpoint) => {
        usage = checkpoint;
        metrics.observeUsage(checkpoint);
      },
    });
    usage ??= readDocxDocumentCursorUsage((operation) => operation(archive));
    metrics.observeUsage(usage);
    metrics.checkpoint('model streamed');
    await pull.reset();
    transport.terminate();
    return { archive, result, usage, metrics, closeArchive };
  } catch (error) {
    await pull?.reset().catch(() => undefined);
    transport?.terminate();
    try { closeArchive(); } catch {}
    const normalized = parseTypedParserError(error) ?? error;
    metrics.fail(normalized);
    throw normalized;
  }
}

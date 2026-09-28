import type { OoxmlResourceUsageSnapshot } from '@silurus/ooxml-core';
import {
  BoundedPullSession,
  PULL_SESSION_PROTOCOL,
  type PullCancelReason,
  type PullSessionIdentity,
  type PullSessionResponse,
  type WorkerBridgeTransport,
} from '@silurus/ooxml-core/worker';
import type { ParsedWorkbook, Row, Worksheet } from './types.js';
import { XLSX_WORKSHEET_PULL_BYTES } from './worksheet-pull-worker.js';
import { decodeWorksheetPullChunk, type WorksheetPreviewBlocker } from './worksheet-pull-codec.js';

export type XlsxWorksheetPullUnit =
  | {
      readonly kind: 'preview';
      readonly worksheet: Worksheet | null;
      readonly reason: WorksheetPreviewBlocker | null;
      readonly maxRow: number;
      readonly maxCol: number;
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    }
  | {
      readonly kind: 'rows';
      readonly rows: Row[];
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    }
  | {
      readonly kind: 'finished';
      readonly worksheet: Worksheet;
      readonly sequence: number;
      readonly wireBytes: number;
      readonly usage?: OoxmlResourceUsageSnapshot;
    };

export interface XlsxWorksheetPullClientOptions {
  readonly generation?: number;
  readonly transport: WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>;
  readonly sharedStrings: ParsedWorkbook['sharedStrings'];
  readonly timeoutMs?: number;
  readonly open: (
    sheetIndex: number,
    sheetName: string,
    identity: PullSessionIdentity<number>,
    timeoutMs?: number,
  ) => Promise<void>;
  readonly onUsage?: (usage: OoxmlResourceUsageSnapshot) => void;
  /** Test/resource hook for transfer-backed payloads. ArrayBuffers need no
   * explicit release, but the coordinator still drops every chunk reference at
   * the same acceptance boundary as transferable resource payloads. */
  readonly disposeTransferred?: (payload: ArrayBuffer) => void;
}

/** Canonical Browser/Node worksheet unit coordinator. It alone owns wire
 * decoding, shared-string normalization, terminal validation, ACK order,
 * cancellation, and transferred-payload disposal. Consumers retain only their
 * output-specific model/accounting policy. */
export class XlsxWorksheetPullClient {
  private readonly active = new Set<BoundedPullSession<ArrayBuffer, number>>();
  private nextSessionId = 1;

  constructor(private readonly options: XlsxWorksheetPullClientOptions) {
    if (
      options.generation !== undefined &&
      (!Number.isSafeInteger(options.generation) || options.generation <= 0)
    ) {
      throw new TypeError('generation must be a positive safe integer');
    }
  }

  async *stream(
    sheetIndex: number,
    sheetName: string,
    signal?: AbortSignal,
  ): AsyncGenerator<XlsxWorksheetPullUnit, void, void> {
    if (!Number.isSafeInteger(sheetIndex) || sheetIndex < 0) {
      throw new RangeError('sheetIndex must be a non-negative safe integer');
    }
    if (!sheetName) throw new TypeError('sheetName must be non-empty');
    throwIfAborted(signal);

    const sessionId = this.nextSessionId++;
    const identity = {
      sessionId,
      operationId: sessionId,
      generation: this.options.generation ?? 1,
    };
    const session = new BoundedPullSession(this.options.transport, {
      ...identity,
      maxByteCredit: XLSX_WORKSHEET_PULL_BYTES,
      timeoutMs: this.options.timeoutMs,
      disposeTransferred: this.options.disposeTransferred,
    });
    this.active.add(session);
    let terminalAcknowledged = false;
    let operationError: unknown;
    try {
      await this.options.open(sheetIndex, sheetName, identity, this.options.timeoutMs);
      for (;;) {
        throwIfAborted(signal);
        const chunk = await session.pull(XLSX_WORKSHEET_PULL_BYTES, { signal });
        try {
          const usage = chunk.usage ?? session.usageCheckpoint;
          if (usage) this.options.onUsage?.(usage);
          const decoded = decodeWorksheetPullChunk(
            chunk.payload,
            chunk.done,
            this.options.sharedStrings,
          );
          const unit: XlsxWorksheetPullUnit = decoded.kind === 'preview'
            ? { ...decoded, sequence: chunk.sequence, wireBytes: chunk.byteLength, usage }
            : decoded.kind === 'rows'
            ? {
                kind: 'rows',
                rows: decoded.rows,
                sequence: chunk.sequence,
                wireBytes: chunk.byteLength,
                usage,
              }
            : {
                kind: 'finished',
                worksheet: decoded.worksheet,
                sequence: chunk.sequence,
                wireBytes: chunk.byteLength,
                usage,
              };
          // The sink accepts the unit while execution is suspended here. Only
          // advancing the iterator ACKs it, preserving rollback on throw/return.
          yield unit;
          await chunk.ack({ signal });
          if (decoded.kind === 'finished') {
            terminalAcknowledged = true;
            return;
          }
        } finally {
          // ACK releases producer staging; this separately releases the
          // consumer's transferred payload at the same canonical boundary.
          chunk.disposeTransferred();
        }
      }
    } catch (error) {
      operationError = error;
      throw error;
    } finally {
      let cleanupError: unknown;
      try {
        if (!terminalAcknowledged) {
          await session.cancel(cancelReason(operationError));
        }
      } catch (error) {
        cleanupError = error;
      } finally {
        this.active.delete(session);
      }
      // Preserve the operation failure when both work and cleanup fail. A lone
      // cleanup failure remains observable to the owner.
      if (operationError === undefined && cleanupError !== undefined) throw cleanupError;
    }
  }

  /** Converge every active operation before reload, close, or worker teardown. */
  async cancelAll(reason: PullCancelReason = 'closed'): Promise<void> {
    const outcomes = await Promise.allSettled(
      [...this.active].map((session) => session.cancel(reason)),
    );
    const failed = outcomes.find(
      (outcome): outcome is PromiseRejectedResult => outcome.status === 'rejected',
    );
    if (failed) throw failed.reason;
  }
}

export function isXlsxWorksheetPullResponse(
  value: unknown,
): value is PullSessionResponse<ArrayBuffer, number> {
  return !!value && typeof value === 'object' &&
    (value as { protocol?: unknown }).protocol === PULL_SESSION_PROTOCOL;
}

function cancelReason(error: unknown): PullCancelReason {
  if (error && typeof error === 'object' && 'name' in error && error.name === 'AbortError') {
    return 'abort';
  }
  return error === undefined ? 'closed' : 'request-error';
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;
  const error = new Error('XLSX workbook session was aborted');
  error.name = 'AbortError';
  throw error;
}

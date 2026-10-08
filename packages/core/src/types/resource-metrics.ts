import type {
  OoxmlFormat,
  OoxmlResourceUsageSnapshot,
} from '../errors/ooxml-error.js';

/** Configured public admission policy used for one measured operation. */
export interface OoxmlResourcePolicySnapshot {
  readonly maxArchiveEntryBytes: number | null;
  readonly maxTotalInflatedBytes: number | null;
  readonly maxArchiveEntries: number | null;
  /** Resolved per-worksheet XLSX logical limits; present only for XLSX. */
  readonly xlsxWorksheetLimits?: Readonly<{
    readonly maxRows: number;
    readonly maxCells: number;
    readonly maxOwnedUtf8Bytes: number;
    readonly maxJsonBytes: number;
  }>;
}

export interface OoxmlResourceMetricsCheckpoint {
  readonly name: string;
  readonly elapsedMs: number;
  readonly usage?: OoxmlResourceUsageSnapshot;
}

/**
 * Content-free, machine-readable resource report for an OOXML load or bounded
 * Node session. Byte counters describe measured package work, not the JavaScript
 * heap, WASM allocator overhead, decoded images, canvas, or GPU. Browser engines
 * and Viewers can return a newer snapshot after lazy package access.
 *
 * No source URL, filename, OOXML part name, document text, password, or raw
 * error message is included. Sizes, counts, and timings are still
 * document-derived metadata; applications apply their own consent, retention,
 * and telemetry policy.
 */
export interface OoxmlResourceMetrics {
  /** Version of this metrics payload, independent of the package version. */
  readonly schemaVersion: 1;
  /**
   * Browser factories report `load`; bounded Node sessions report `session`.
   */
  readonly scope: 'load' | 'session';
  readonly format: OoxmlFormat;
  readonly mode: 'main' | 'worker' | 'node';
  /** Outcome of this measured load/session, not of later render operations. */
  readonly status: 'ok' | 'error';
  /** Compressed or decrypted OOXML container bytes supplied to the parser. */
  readonly sourceBytes?: number;
  readonly elapsedMs: number;
  readonly policy: Readonly<OoxmlResourcePolicySnapshot>;
  /** Last complete observed package-usage checkpoint. */
  readonly usage?: OoxmlResourceUsageSnapshot;
  readonly checkpoints: readonly OoxmlResourceMetricsCheckpoint[];
  readonly outcome?: Readonly<Record<string, number>>;
  readonly error?: Readonly<{
    readonly code?: string;
    readonly stage?: string;
    readonly resource?: string;
    readonly metric?: string;
  }>;
}

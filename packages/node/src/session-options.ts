import type {
  OoxmlResourceLimits,
  OoxmlResourceMetrics,
  XlsxWorksheetLimits,
} from '@silurus/ooxml-core';

/** Resource policy, diagnostics, and cancellation shared by every Node session. */
export interface OoxmlNodeSessionOptions {
  /** Password for an Agile-encrypted OOXML container. */
  password?: string;
  /** Package-level inflated ZIP admission limits. */
  resourceLimits?: OoxmlResourceLimits;
  /** XLSX-only logical worksheet admission limits; ignored by DOCX and PPTX. */
  xlsxWorksheetLimits?: XlsxWorksheetLimits;
  /** @deprecated Use `resourceLimits.maxArchiveEntryBytes`. Scheduled for
   * removal in a future breaking release. */
  maxZipEntryBytes?: number;
  /** Emit one content-free resource report for the terminal session outcome. */
  debug?: boolean;
  /** Receive the same terminal report without enabling console output. */
  onResourceMetrics?: (metrics: OoxmlResourceMetrics) => void;
  /** Cooperatively abort initialization or active work; synchronous WASM cannot be preempted. */
  signal?: AbortSignal;
}

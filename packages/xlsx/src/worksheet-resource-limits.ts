import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS,
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES,
  HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS,
  HARD_MAX_XLSX_WORKSHEET_CELLS,
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES,
  HARD_MAX_XLSX_WORKSHEET_JSON_BYTES,
  HARD_MAX_XLSX_WORKSHEET_ROWS,
  type NormalizedXlsxWorksheetPolicy,
} from '@silurus/ooxml-core/worker';
import {
  OoxmlResourceLimitError,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
import {
  cappedAdd,
  measureStructuralJson,
  utf8Bytes,
} from '@silurus/ooxml-core/internal/resource-measurement';
import type { Row, Worksheet } from './types.js';

export const XLSX_MAX_MATERIALIZED_ROWS = HARD_MAX_XLSX_WORKSHEET_ROWS;
export const XLSX_MAX_MATERIALIZED_CELLS = HARD_MAX_XLSX_WORKSHEET_CELLS;
export const XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES =
  HARD_MAX_XLSX_WORKSHEET_CELL_CONTENT_UTF8_BYTES;
export const XLSX_MAX_MATERIALIZED_JSON_BYTES = HARD_MAX_XLSX_WORKSHEET_JSON_BYTES;
export const XLSX_MAX_CACHED_ROWS = HARD_MAX_XLSX_WORKBOOK_CACHED_ROWS;
export const XLSX_MAX_CACHED_CELLS = HARD_MAX_XLSX_WORKBOOK_CACHED_CELLS;
export const XLSX_MAX_CACHED_OWNED_UTF8_BYTES =
  HARD_MAX_XLSX_WORKBOOK_CACHED_CELL_CONTENT_UTF8_BYTES;
export const XLSX_MAX_CACHED_JSON_BYTES = HARD_MAX_XLSX_WORKBOOK_CACHED_JSON_BYTES;

export interface WorksheetModelUsage {
  rows: number;
  cells: number;
  ownedUtf8Bytes: number;
}

export interface WorksheetCacheUsage {
  rows: number;
  cells: number;
  ownedUtf8Bytes: number;
  jsonBytes: number;
}

const ZERO_RESOURCE_USAGE: OoxmlResourceUsageSnapshot = Object.freeze({
  archiveEntryCount: 0,
  declaredInflatedBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
});

export function measureRows(
  rows: readonly Row[],
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): WorksheetModelUsage {
  const { maxCells, maxOwnedUtf8Bytes } = policy.worksheet;
  const cells = rows.reduce(
    (total, row) => cappedAdd(total, row.cells.length, maxCells),
    0,
  );
  return {
    rows: rows.length,
    cells,
    // This is deliberately cell-scoped. It includes every retained string in
    // Cell.value (rich/phonetic formatting and the discriminator included) and
    // formula text. Shared values must be resolved before calling this helper,
    // so repeated shared-string references are charged once per materialized
    // cell. Ancillary worksheet strings are covered by the exact JSON ceiling.
    ownedUtf8Bytes: rows.reduce((rowTotal, row) => row.cells.reduce((cellTotal, cell) => {
      const valueBytes = measureStructuralJson(
        cell.value,
        maxOwnedUtf8Bytes,
      ).stringValueUtf8Bytes;
      const formulaBytes = cell.formula === undefined
        ? 0
        : utf8Bytes(cell.formula, maxOwnedUtf8Bytes);
      return cappedAdd(
        cellTotal,
        cappedAdd(valueBytes, formulaBytes, maxOwnedUtf8Bytes),
        maxOwnedUtf8Bytes,
      );
    }, rowTotal), 0),
  };
}

export function measureWorksheet(
  worksheet: Worksheet,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): WorksheetModelUsage & { jsonBytes: number } {
  return completeWorksheetUsage(worksheet, measureRows(worksheet.rows, policy), policy);
}

/** Complete an incrementally accumulated row/cell measurement with the exact
 * retained JSON size. Streaming callers already measured each row chunk, so
 * repeating that full traversal at terminal admission adds cost but no safety. */
export function completeWorksheetUsage(
  worksheet: Worksheet,
  model: WorksheetModelUsage,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): WorksheetCacheUsage {
  // The JSON budget is independent of the owned-string budget: measure the
  // canonical JSON.stringify structural graph against maxJsonBytes only.
  const measured = measureStructuralJson(worksheet, policy.worksheet.maxJsonBytes);
  return { ...model, jsonBytes: measured.jsonBytes };
}

export function addWorksheetUsage(
  current: WorksheetModelUsage,
  addition: WorksheetModelUsage,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): WorksheetModelUsage {
  const limits = policy.worksheet;
  return {
    rows: cappedAdd(current.rows, addition.rows, limits.maxRows),
    cells: cappedAdd(current.cells, addition.cells, limits.maxCells),
    ownedUtf8Bytes: cappedAdd(
      current.ownedUtf8Bytes,
      addition.ownedUtf8Bytes,
      limits.maxOwnedUtf8Bytes,
    ),
  };
}

export function addWorksheetCacheUsage(
  current: WorksheetCacheUsage,
  addition: WorksheetModelUsage & { jsonBytes: number },
  subtraction: Partial<WorksheetCacheUsage> = {},
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): WorksheetCacheUsage {
  const limits = policy.cache;
  const baseRows = current.rows - (subtraction.rows ?? 0);
  const baseCells = current.cells - (subtraction.cells ?? 0);
  const baseOwnedUtf8Bytes = current.ownedUtf8Bytes - (subtraction.ownedUtf8Bytes ?? 0);
  const baseJsonBytes = current.jsonBytes - (subtraction.jsonBytes ?? 0);
  if (baseRows < 0 || baseCells < 0 || baseOwnedUtf8Bytes < 0 || baseJsonBytes < 0) {
    throw new Error('worksheet cache accounting underflow');
  }
  return {
    rows: cappedAdd(baseRows, addition.rows, limits.maxRows),
    cells: cappedAdd(baseCells, addition.cells, limits.maxCells),
    ownedUtf8Bytes: cappedAdd(
      baseOwnedUtf8Bytes,
      addition.ownedUtf8Bytes,
      limits.maxOwnedUtf8Bytes,
    ),
    jsonBytes: cappedAdd(baseJsonBytes, addition.jsonBytes, limits.maxJsonBytes),
  };
}

export function worksheetLimitError(
  operation: string,
  part: string | undefined,
  resource:
    | 'delimited-text-source'
    | 'worksheet-model'
    | 'worksheet-cell-content'
    | 'worksheet-json'
    | 'worksheet-cache',
  metric: 'rows' | 'cells' | 'owned-utf8-bytes' | 'bytes',
  limit: number,
  observed: number,
  usage?: OoxmlResourceUsageSnapshot,
  // Worksheet model/cache/JSON budgets are adjustable policy values. The
  // independent delimited-text source ceiling stays non-configurable unless a
  // caller states otherwise explicitly.
  configurable: boolean = resource !== 'delimited-text-source',
): OoxmlResourceLimitError {
  const stage = resource === 'worksheet-json' ? 'serialization' : 'parsing';
  return new OoxmlResourceLimitError(
    `OOXML resource limit exceeded${part ? ` for ${part}` : ''}: ${metric} ${observed} > ${limit}`,
    {
      stage,
      violation: {
        format: 'xlsx',
        operation,
        resource,
        metric,
        ...(part === undefined ? {} : { part }),
        limit,
        observed: Math.min(observed, limit + 1),
        configurable,
        usage: usage ?? ZERO_RESOURCE_USAGE,
      },
    },
  );
}

export function assertWorksheetModelUsage(
  measured: WorksheetModelUsage,
  operation: string,
  part: string | undefined,
  usage?: OoxmlResourceUsageSnapshot,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): void {
  const limits = policy.worksheet;
  const checks = [
    ['rows', measured.rows, limits.maxRows],
    ['cells', measured.cells, limits.maxCells],
    ['owned-utf8-bytes', measured.ownedUtf8Bytes, limits.maxOwnedUtf8Bytes],
  ] as const;
  for (const [metric, observed, limit] of checks) {
    if (observed > limit) {
      throw worksheetLimitError(
        operation,
        part,
        metric === 'owned-utf8-bytes' ? 'worksheet-cell-content' : 'worksheet-model',
        metric,
        limit,
        observed,
        usage,
      );
    }
  }
}

export function assertWorksheetJsonBytes(
  observed: number,
  operation: string,
  part: string | undefined,
  usage?: OoxmlResourceUsageSnapshot,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): void {
  const limit = policy.worksheet.maxJsonBytes;
  if (observed > limit) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-json',
      'bytes',
      limit,
      observed,
      usage,
    );
  }
}

export function assertWorksheetCacheUsage(
  usage: WorksheetCacheUsage,
  operation: string,
  part: string | undefined,
  resourceUsage?: OoxmlResourceUsageSnapshot,
  policy: NormalizedXlsxWorksheetPolicy = DEFAULT_XLSX_WORKSHEET_POLICY,
): void {
  const limits = policy.cache;
  if (usage.rows > limits.maxRows) {
    throw worksheetLimitError(
      operation, part, 'worksheet-cache', 'rows', limits.maxRows, usage.rows, resourceUsage,
    );
  }
  if (usage.cells > limits.maxCells) {
    throw worksheetLimitError(
      operation, part, 'worksheet-cache', 'cells', limits.maxCells, usage.cells, resourceUsage,
    );
  }
  if (usage.ownedUtf8Bytes > limits.maxOwnedUtf8Bytes) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-cache',
      'owned-utf8-bytes',
      limits.maxOwnedUtf8Bytes,
      usage.ownedUtf8Bytes,
      resourceUsage,
    );
  }
  if (usage.jsonBytes > limits.maxJsonBytes) {
    throw worksheetLimitError(
      operation,
      part,
      'worksheet-cache',
      'bytes',
      limits.maxJsonBytes,
      usage.jsonBytes,
      resourceUsage,
    );
  }
}

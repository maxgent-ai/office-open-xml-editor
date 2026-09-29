import type { Worksheet } from '../types.js';
import { addWorksheetCacheUsage, type WorksheetCacheUsage } from '../worksheet-resource-limits.js';
import type { WorksheetViewProjectionCache } from '../worker-protocol.js';

const ZERO_USAGE: WorksheetCacheUsage = {
  rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0,
};

/** Main owns LRU order. Worker eviction is an exact, checked mirror of that
 * decision, and happens before the incoming pull terminal is acknowledged. */
export function evictWorkerWorksheets(
  sheetIndices: readonly number[],
  sheets: Map<number, Worksheet>,
  usages: Map<number, WorksheetCacheUsage>,
  retained: WorksheetCacheUsage,
  projections: WorksheetViewProjectionCache,
): WorksheetCacheUsage {
  for (const sheetIndex of sheetIndices) {
    if (!sheets.has(sheetIndex) || !usages.has(sheetIndex)) {
      throw new Error(`Worksheet cache is inconsistent for sheet ${sheetIndex}`);
    }
  }
  for (const sheetIndex of sheetIndices) {
    const previous = usages.get(sheetIndex)!;
    sheets.delete(sheetIndex);
    usages.delete(sheetIndex);
    retained = addWorksheetCacheUsage(retained, ZERO_USAGE, previous);
    projections.evictSheet(sheetIndex);
  }
  return retained;
}

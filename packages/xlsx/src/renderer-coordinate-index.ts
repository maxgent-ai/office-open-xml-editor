import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES } from '@silurus/ooxml-core/worker';

/** Default finite policy for each independently charged renderer index. */
export const MAX_RENDERER_COORDINATE_INDEX_ENTRIES =
  HARD_MAX_XLSX_RENDERER_COORDINATE_INDEX_ENTRIES;

export interface CoordinateRange {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface CoordinateIndexIdentity {
  resource: string;
  operation: string;
  /**
   * Optional finite entry budget (normally the normalized policy's
   * maxCoordinateIndexEntries). Defaults to MAX_RENDERER_COORDINATE_INDEX_ENTRIES.
   * Must be a positive safe integer no greater than MAX_SAFE_INTEGER - 1.
   */
  limit?: number;
}

const MAX_ACCEPTED_COORDINATE_LIMIT = Number.MAX_SAFE_INTEGER - 1;

function resolveLimit(identity: CoordinateIndexIdentity): number {
  const limit: unknown = identity.limit;
  if (limit === undefined) return MAX_RENDERER_COORDINATE_INDEX_ENTRIES;
  if (typeof limit !== 'number') {
    throw new TypeError('coordinate index limit must be a number');
  }
  if (!Number.isSafeInteger(limit) || limit <= 0 || limit > MAX_ACCEPTED_COORDINATE_LIMIT) {
    throw new RangeError(
      `coordinate index limit must be a positive safe integer no greater than ${MAX_ACCEPTED_COORDINATE_LIMIT}`,
    );
  }
  return limit;
}

/**
 * `configurable` is true only for a finite budget violation of a countable
 * entry total. Invalid or unrepresentable inputs are non-configurable refusals.
 */
function resourceLimitError(
  identity: CoordinateIndexIdentity,
  limit: number,
  observed: number,
  configurable: boolean,
): OoxmlResourceLimitError {
  return new OoxmlResourceLimitError(
    `XLSX renderer ${identity.resource} exceeded its limit of ${limit} entries`,
    {
      stage: 'rendering',
      violation: {
        format: 'xlsx',
        operation: identity.operation,
        resource: identity.resource,
        metric: 'entry-count',
        limit,
        observed: Math.min(observed, limit + 1),
        configurable,
        // Renderer-only failures do not have access to the package session's
        // inflation counters. Keep those independent measurements explicit
        // rather than fabricating package usage from worksheet model data.
        usage: {
          archiveEntryCount: 0,
          declaredInflatedBytes: 0,
          distinctInflatedBytes: 0,
          operationInflatedBytes: 0,
        },
      },
    },
  );
}

/**
 * Verify that one rectangular expansion cannot cross the entry budget before
 * entering its nested coordinate loops. `excludedEntries` covers topology
 * entries such as a merge anchor that are deliberately not inserted. The
 * check is pure arithmetic: no coordinate scan or allocation happens here.
 */
export function assertCoordinateRangeArea(
  range: CoordinateRange,
  identity: CoordinateIndexIdentity,
  excludedEntries = 0,
): number {
  const limit = resolveLimit(identity);
  const { top, bottom, left, right } = range;
  if (
    !Number.isSafeInteger(top)
    || !Number.isSafeInteger(bottom)
    || !Number.isSafeInteger(left)
    || !Number.isSafeInteger(right)
    || !Number.isSafeInteger(excludedEntries)
    || excludedEntries < 0
  ) {
    throw resourceLimitError(identity, limit, limit + 1, false);
  }
  if (bottom < top || right < left) return 0;

  const height = bottom - top + 1;
  const width = right - left + 1;
  if (!Number.isSafeInteger(height) || !Number.isSafeInteger(width)) {
    throw resourceLimitError(identity, limit, limit + 1, false);
  }

  let entries: number;
  if (width <= Math.floor(Number.MAX_SAFE_INTEGER / height)) {
    // Fast path: the area is an exact safe integer.
    const area = height * width;
    if (excludedEntries > area) {
      throw resourceLimitError(identity, limit, limit + 1, false);
    }
    entries = area - excludedEntries;
  } else {
    // Rare path: the area itself is not representable; decide exactly.
    const area = BigInt(height) * BigInt(width);
    const excluded = BigInt(excludedEntries);
    if (excluded > area) {
      throw resourceLimitError(identity, limit, limit + 1, false);
    }
    const exact = area - excluded;
    if (exact > BigInt(Number.MAX_SAFE_INTEGER)) {
      throw resourceLimitError(identity, limit, limit + 1, false);
    }
    entries = Number(exact);
  }

  if (entries > limit) {
    throw resourceLimitError(identity, limit, entries, true);
  }
  return entries;
}

/** Insert a coordinate while charging only an actual new Map entry. */
export function setCoordinateIndexValue<T>(
  map: Map<string, T>,
  key: string,
  value: T,
  identity: CoordinateIndexIdentity,
): void {
  const limit = resolveLimit(identity);
  if (!map.has(key) && map.size >= limit) {
    throw resourceLimitError(identity, limit, map.size + 1, true);
  }
  map.set(key, value);
}

/** Build the renderer's canonical row:column cell lookup. */
export function buildCellCoordinateIndex<T extends { row: number; col: number }>(
  rows: readonly { cells: readonly T[] }[],
  identity: CoordinateIndexIdentity,
): Map<string, T> {
  const map = new Map<string, T>();
  for (const row of rows) {
    for (const cell of row.cells) {
      setCoordinateIndexValue(map, `${cell.row}:${cell.col}`, cell, identity);
    }
  }
  return map;
}

/** Insert a coordinate while charging only an actual new Set entry. */
export function addCoordinateIndexEntry(
  set: Set<string>,
  key: string,
  identity: CoordinateIndexIdentity,
): void {
  const limit = resolveLimit(identity);
  if (!set.has(key) && set.size >= limit) {
    throw resourceLimitError(identity, limit, set.size + 1, true);
  }
  set.add(key);
}

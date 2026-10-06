import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { normalizeXlsxWorksheetPolicy } from '@silurus/ooxml-core/worker';
import { utf8Bytes } from '@silurus/ooxml-core/internal/resource-measurement';
import type { Row, Worksheet } from './types.js';
import {
  XLSX_MAX_CACHED_CELLS,
  XLSX_MAX_CACHED_JSON_BYTES,
  XLSX_MAX_CACHED_OWNED_UTF8_BYTES,
  XLSX_MAX_CACHED_ROWS,
  XLSX_MAX_MATERIALIZED_CELLS,
  XLSX_MAX_MATERIALIZED_JSON_BYTES,
  XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
  XLSX_MAX_MATERIALIZED_ROWS,
  addWorksheetCacheUsage,
  addWorksheetUsage,
  assertWorksheetCacheUsage,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  completeWorksheetUsage,
  measureRows,
  measureWorksheet,
} from './worksheet-resource-limits.js';
import {
  MAX_RENDERER_COORDINATE_INDEX_ENTRIES,
  addCoordinateIndexEntry,
  assertCoordinateRangeArea,
  buildCellCoordinateIndex,
  setCoordinateIndexValue,
} from './renderer-coordinate-index.js';

const W = { rows: 'maxRows', cells: 'maxCells', owned: 'maxOwnedUtf8Bytes', json: 'maxJsonBytes' } as const;
const C = W;
function worksheetPolicy(overrides: import('@silurus/ooxml-core').XlsxWorksheetLimits) {
  return normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: overrides });
}

function violationOf(fn: () => unknown): Record<string, unknown> {
  let caught: unknown;
  try {
    fn();
  } catch (error) {
    caught = error;
  }
  expect(caught).toBeInstanceOf(OoxmlResourceLimitError);
  return (caught as OoxmlResourceLimitError).details.violation as unknown as Record<string, unknown>;
}

function worksheetOf(rows: Row[], name: string): Worksheet {
  return {
    name, rows, colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
  } as Worksheet;
}

describe('finite worksheet policy overrides through production helpers', () => {
  it('measures and indexes 251000 real cells under a raised policy while defaults still refuse', () => {
    const total = 251_000;
    const perRow = 20;
    const rows: Row[] = [];
    for (let r = 1; r <= total / perRow; r += 1) {
      const cells: Row['cells'] = new Array(perRow);
      for (let c = 0; c < perRow; c += 1) {
        cells[c] = { row: r, col: c + 1, value: { type: 'text', text: 'x' } } as Row['cells'][number];
      }
      rows.push({ index: r, height: null, cells });
    }

    const policy = worksheetPolicy({ [W.cells]: 300_000 });
    const measured = measureRows(rows, policy);
    expect(measured.cells).toBe(total);
    expect(measured.rows).toBe(total / perRow);
    expect(() => assertWorksheetModelUsage(measured, 'get-worksheet', 'worksheet/0', undefined, policy))
      .not.toThrow();

    const defaultMeasured = measureRows(rows);
    expect(defaultMeasured.cells).toBe(Math.min(total, XLSX_MAX_MATERIALIZED_CELLS + 1));
    if (total > XLSX_MAX_MATERIALIZED_CELLS) {
      expect(violationOf(() => assertWorksheetModelUsage(defaultMeasured, 'get-worksheet', 'worksheet/0')))
        .toMatchObject({ metric: 'cells', limit: XLSX_MAX_MATERIALIZED_CELLS, configurable: true });
    }

    const coordinatePolicy = policy;
    expect(coordinatePolicy.maxCoordinateIndexEntries).toBe(300_000);
    const index = buildCellCoordinateIndex(rows, {
      resource: 'cell-index', operation: 'render', limit: coordinatePolicy.maxCoordinateIndexEntries,
    });
    expect(index.size).toBe(total);

    expect(MAX_RENDERER_COORDINATE_INDEX_ENTRIES).toBe(250_000);
    expect(violationOf(() => buildCellCoordinateIndex(rows, { resource: 'cell-index', operation: 'render' })))
      .toMatchObject({
        metric: 'entry-count',
        limit: MAX_RENDERER_COORDINATE_INDEX_ENTRIES,
        observed: MAX_RENDERER_COORDINATE_INDEX_ENTRIES + 1,
        configurable: true,
      });
  });

  it('measures owned and JSON bytes independently with limit+1 saturation and exact/one-over budgets', () => {
    const text = 'é"\\😀';
    const formula = 'A1&"x"';
    const rows: Row[] = [{
      index: 1, height: null,
      cells: [{ row: 1, col: 1, value: { type: 'text', text }, formula } as Row['cells'][number]],
    }];
    const owned = utf8Bytes('text') + utf8Bytes(text) + utf8Bytes(formula);

    expect(measureRows(rows, worksheetPolicy({ [W.owned]: owned })).ownedUtf8Bytes).toBe(owned);
    expect(measureRows(rows, worksheetPolicy({ [W.owned]: 3 })).ownedUtf8Bytes).toBe(4);

    const exactOwned = worksheetPolicy({ [W.owned]: owned });
    const underOwned = worksheetPolicy({ [W.owned]: owned - 1 });
    const measured = measureRows(rows, exactOwned);
    expect(() => assertWorksheetModelUsage(measured, 'op', 'worksheet/0', undefined, exactOwned)).not.toThrow();
    expect(violationOf(() => assertWorksheetModelUsage(measured, 'op', 'worksheet/0', undefined, underOwned)))
      .toMatchObject({
        resource: 'worksheet-cell-content', metric: 'owned-utf8-bytes',
        limit: owned - 1, observed: owned, configurable: true,
      });

    // Escaped ancillary metadata changes JSON bytes but never owned cell bytes.
    const plain = worksheetOf(rows, 'S');
    const escaped = worksheetOf(rows, 'S"\n\u000b\\é');
    const plainJson = new TextEncoder().encode(JSON.stringify(plain)).byteLength;
    const escapedJson = new TextEncoder().encode(JSON.stringify(escaped)).byteLength;
    const jsonPolicy = worksheetPolicy({ [W.json]: escapedJson });
    const a = measureWorksheet(plain, jsonPolicy);
    const b = measureWorksheet(escaped, jsonPolicy);
    expect(a.ownedUtf8Bytes).toBe(owned);
    expect(b.ownedUtf8Bytes).toBe(owned);
    expect(a.jsonBytes).toBe(plainJson);
    expect(b.jsonBytes).toBe(escapedJson);
    const tiny = worksheetPolicy({ [W.owned]: 10, [W.json]: 10 });
    expect(completeWorksheetUsage(escaped, { rows: 1, cells: 1, ownedUtf8Bytes: 2 }, tiny))
      .toEqual({ rows: 1, cells: 1, ownedUtf8Bytes: 2, jsonBytes: 11 });
    expect(() => assertWorksheetJsonBytes(escapedJson, 'op', 'worksheet/0', undefined, jsonPolicy)).not.toThrow();
    expect(violationOf(() => assertWorksheetJsonBytes(escapedJson + 1, 'op', 'worksheet/0', undefined, jsonPolicy)))
      .toMatchObject({ resource: 'worksheet-json', metric: 'bytes', limit: escapedJson, observed: escapedJson + 1, configurable: true });

    // Raising rows/cells does not raise byte budgets.
    const raised = worksheetPolicy({ [W.rows]: 200_000, [W.cells]: 300_000 });
    expect(raised.worksheet[W.owned]).toBe(XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES);
    expect(raised.worksheet[W.json]).toBe(XLSX_MAX_MATERIALIZED_JSON_BYTES);
    expect(() => assertWorksheetJsonBytes(XLSX_MAX_MATERIALIZED_JSON_BYTES + 1, 'op', 'worksheet/0', undefined, raised))
      .toThrow(OoxmlResourceLimitError);
    expect(() => assertWorksheetModelUsage(
      { rows: 0, cells: 0, ownedUtf8Bytes: XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES + 1 },
      'op', 'worksheet/0', undefined, raised,
    )).toThrow(OoxmlResourceLimitError);

    // Large row counts accumulate past the old row saturation under policy.
    const zero = { cells: 0, ownedUtf8Bytes: 0 };
    expect(addWorksheetUsage({ rows: 100_000, ...zero }, { rows: 1, ...zero }, raised).rows).toBe(100_001);
    expect(addWorksheetUsage({ rows: 100_000, ...zero }, { rows: 1, ...zero }).rows)
      .toBe(Math.min(100_001, XLSX_MAX_MATERIALIZED_ROWS + 1));
  });

  it('derives cache limits from sheet policy and keeps subtraction and overflow sentinels exact', () => {
    const lowered = worksheetPolicy({ [W.cells]: 10 });
    expect(lowered.cache[C.cells]).toBe(XLSX_MAX_CACHED_CELLS);
    const higher = XLSX_MAX_CACHED_CELLS + 1_000;
    const raised = worksheetPolicy({ [W.cells]: higher });
    expect(raised.cache[C.cells]).toBeGreaterThanOrEqual(higher);
    expect(raised.cache[C.rows]).toBe(XLSX_MAX_CACHED_ROWS);
    expect(raised.cache[C.owned]).toBe(XLSX_MAX_CACHED_OWNED_UTF8_BYTES);
    expect(raised.cache[C.json]).toBe(XLSX_MAX_CACHED_JSON_BYTES);

    const cap = raised.cache[C.cells];
    const current = { rows: 10, cells: cap - 100, ownedUtf8Bytes: 30, jsonBytes: 40 };
    const previous = { rows: 3, cells: 50, ownedUtf8Bytes: 7, jsonBytes: 11 };
    const replacement = { rows: 4, cells: 150, ownedUtf8Bytes: 8, jsonBytes: 12 };
    const replaced = addWorksheetCacheUsage(current, replacement, previous, raised);
    expect(replaced).toEqual({ rows: 11, cells: cap, ownedUtf8Bytes: 31, jsonBytes: 41 });
    expect(() => assertWorksheetCacheUsage(replaced, 'op', 'worksheet/0', undefined, raised)).not.toThrow();

    const overflow = addWorksheetCacheUsage(
      { rows: 0, cells: cap, ownedUtf8Bytes: 0, jsonBytes: 0 },
      { rows: 0, cells: 5_000, ownedUtf8Bytes: 0, jsonBytes: 0 },
      {},
      raised,
    );
    expect(overflow.cells).toBe(cap + 1);
    expect(violationOf(() => assertWorksheetCacheUsage(overflow, 'op', 'worksheet/0', undefined, raised)))
      .toMatchObject({ resource: 'worksheet-cache', metric: 'cells', limit: cap, observed: cap + 1, configurable: true });

    const defaultOver = { rows: 0, cells: XLSX_MAX_CACHED_CELLS + 1, ownedUtf8Bytes: 0, jsonBytes: 0 };
    expect(violationOf(() => assertWorksheetCacheUsage(defaultOver, 'op', 'worksheet/0', undefined, lowered)))
      .toMatchObject({ metric: 'cells', limit: XLSX_MAX_CACHED_CELLS, configurable: true });
  });

  it('computes rectangle exclusions without unsafe limit arithmetic and classifies refusals', () => {
    const max = Number.MAX_SAFE_INTEGER;
    const huge = { resource: 'merge-index', operation: 'render', limit: max - 1 };
    expect(assertCoordinateRangeArea({ top: 1, bottom: 1, left: 1, right: max }, huge, 1)).toBe(max - 1);

    expect(violationOf(() => assertCoordinateRangeArea({ top: 1, bottom: 2, left: 1, right: max }, huge, 1)))
      .toMatchObject({ configurable: false });
    expect(violationOf(() => assertCoordinateRangeArea({ top: Number.NaN, bottom: 2, left: 1, right: 2 }, huge)))
      .toMatchObject({ configurable: false });
    expect(violationOf(() => assertCoordinateRangeArea({ top: 1, bottom: 2, left: 0.5, right: 2 }, huge)))
      .toMatchObject({ configurable: false });
    expect(violationOf(() => assertCoordinateRangeArea({ top: 1, bottom: 1, left: 1, right: 2 }, huge, -1)))
      .toMatchObject({ configurable: false });

    const finite = { resource: 'merge-index', operation: 'render', limit: 99 };
    expect(assertCoordinateRangeArea({ top: 1, bottom: 10, left: 1, right: 10 }, finite, 1)).toBe(99);
    expect(violationOf(() => assertCoordinateRangeArea({ top: 1, bottom: 10, left: 1, right: 10 }, finite)))
      .toMatchObject({ metric: 'entry-count', limit: 99, observed: 100, configurable: true });
    expect(assertCoordinateRangeArea({ top: 5, bottom: 4, left: 1, right: 1 }, finite)).toBe(0);
  });

  it('charges unique coordinate keys only while measurement still charges duplicate records', () => {
    const identity = { resource: 'cell-index', operation: 'render', limit: 2 };
    const set = new Set<string>();
    addCoordinateIndexEntry(set, '1:1', identity);
    addCoordinateIndexEntry(set, '1:2', identity);
    addCoordinateIndexEntry(set, '1:1', identity);
    expect(set.size).toBe(2);
    expect(violationOf(() => addCoordinateIndexEntry(set, '1:3', identity)))
      .toMatchObject({ limit: 2, observed: 3, configurable: true });

    const map = new Map<string, number>();
    setCoordinateIndexValue(map, 'a', 1, identity);
    setCoordinateIndexValue(map, 'b', 2, identity);
    setCoordinateIndexValue(map, 'a', 3, identity);
    expect(map.get('a')).toBe(3);
    expect(() => setCoordinateIndexValue(map, 'c', 4, identity)).toThrow(OoxmlResourceLimitError);

    const cell = (col: number, text: string) => ({ row: 1, col, value: { type: 'text', text } }) as Row['cells'][number];
    const rows: Row[] = [{ index: 1, height: null, cells: [cell(1, 'a'), cell(2, 'b'), cell(1, 'c')] }];
    const index = buildCellCoordinateIndex(rows, identity);
    expect(index.size).toBe(2);
    const policy = worksheetPolicy({ [W.cells]: 2 });
    const measured = measureRows(rows, policy);
    expect(measured.cells).toBe(3);
    expect(violationOf(() => assertWorksheetModelUsage(measured, 'op', 'worksheet/0', undefined, policy)))
      .toMatchObject({ metric: 'cells', limit: 2, observed: 3, configurable: true });
  });
});

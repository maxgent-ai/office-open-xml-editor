import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
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
  assertWorksheetCacheUsage,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  completeWorksheetUsage,
  measureRows,
  measureWorksheet,
  addWorksheetCacheUsage,
} from './worksheet-resource-limits.js';

const row = (value: Row['cells'][number]['value'], formula?: string): Row => ({
  index: 1,
  height: null,
  cells: [{ col: 1, row: 1, value, ...(formula === undefined ? {} : { formula }) }],
});

describe('worksheet retained resource measurements', () => {
  it('uses exact inclusive boundaries and canonical +1 observations', () => {
    expect(() => assertWorksheetModelUsage({
      rows: XLSX_MAX_MATERIALIZED_ROWS,
      cells: XLSX_MAX_MATERIALIZED_CELLS,
      ownedUtf8Bytes: XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES,
    }, 'get-worksheet', 'worksheet/0')).not.toThrow();

    for (const [metric, measured] of [
      ['rows', { rows: XLSX_MAX_MATERIALIZED_ROWS + 1, cells: 0, ownedUtf8Bytes: 0 }],
      ['cells', { rows: 0, cells: XLSX_MAX_MATERIALIZED_CELLS + 1, ownedUtf8Bytes: 0 }],
      ['owned-utf8-bytes', { rows: 0, cells: 0, ownedUtf8Bytes: XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES + 1 }],
    ] as const) {
      try {
        assertWorksheetModelUsage(measured, 'get-worksheet', 'worksheet/0');
        throw new Error('expected resource error');
      } catch (error) {
        expect(error).toBeInstanceOf(OoxmlResourceLimitError);
        expect((error as OoxmlResourceLimitError).details.violation).toMatchObject({
          metric,
          observed: metric === 'rows'
            ? XLSX_MAX_MATERIALIZED_ROWS + 1
            : metric === 'cells'
              ? XLSX_MAX_MATERIALIZED_CELLS + 1
              : XLSX_MAX_MATERIALIZED_OWNED_UTF8_BYTES + 1,
          configurable: true,
          part: 'worksheet/0',
        });
      }
    }
  });

  it('charges resolved shared content per cell, including rich/phonetic/error/formula strings', () => {
    const rich = {
      type: 'text' as const,
      text: '共有',
      runs: [{ text: 'rich', font: { bold: false, italic: false, underline: false, strike: false, name: 'Font' } }],
      phoneticRuns: [{ sb: 0, eb: 2, text: 'キョウユウ' }],
    };
    const rows = [
      row(rich, 'A1&"x"'),
      { ...row(rich), index: 2, cells: [{ ...row(rich).cells[0], row: 2 }] },
      { ...row({ type: 'error', error: '#VALUE!' }), index: 3 },
    ];
    const one = measureRows([rows[0]]).ownedUtf8Bytes;
    const withoutFormula = one - utf8Bytes('A1&"x"');
    expect(measureRows(rows).ownedUtf8Bytes).toBe(
      one + withoutFormula + utf8Bytes('error') + utf8Bytes('#VALUE!'),
    );
  });

  it('counts monolithic JSON exactly without JSON.stringify', () => {
    const worksheet = {
      name: 'é\n\u000b"', rows: [row({ type: 'text', text: '😀\\' })], colWidths: {}, rowHeights: {},
      defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [], freezeRows: 0,
      freezeCols: 0, conditionalFormats: [], images: [], charts: [],
    } as Worksheet;
    const expected = new TextEncoder().encode(JSON.stringify(worksheet)).byteLength;
    expect(measureWorksheet(worksheet).jsonBytes).toBe(expected);
    expect(completeWorksheetUsage(worksheet, {
      rows: 7,
      cells: 9,
      ownedUtf8Bytes: 11,
    })).toEqual({
      rows: 7,
      cells: 9,
      ownedUtf8Bytes: 11,
      jsonBytes: expected,
    });
    expect(() => assertWorksheetJsonBytes(XLSX_MAX_MATERIALIZED_JSON_BYTES, 'x', 'worksheet/0'))
      .not.toThrow();
    expect(() => assertWorksheetJsonBytes(XLSX_MAX_MATERIALIZED_JSON_BYTES + 1, 'x', 'worksheet/0'))
      .toThrow(OoxmlResourceLimitError);
  });

  it('matches JSON.stringify for unpaired surrogate keys and values', () => {
    const worksheet = {
      name: '\ud800', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
      defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
      conditionalFormats: [], images: [], charts: [], ['\udc00']: '\ud800',
    } as unknown as Worksheet;
    expect(measureWorksheet(worksheet).jsonBytes).toBe(
      new TextEncoder().encode(JSON.stringify(worksheet)).byteLength,
    );
    expect(utf8Bytes('\ud800')).toBe(3);
    expect(utf8Bytes('\u0000\u000b\u001f')).toBe(3);
  });

  it('checks workbook cache row and cell totals independently at exact/+1', () => {
    expect(() => assertWorksheetCacheUsage(
      {
        rows: XLSX_MAX_CACHED_ROWS,
        cells: XLSX_MAX_CACHED_CELLS,
        ownedUtf8Bytes: XLSX_MAX_CACHED_OWNED_UTF8_BYTES,
        jsonBytes: XLSX_MAX_CACHED_JSON_BYTES,
      }, 'x', 'worksheet/0',
    )).not.toThrow();
    expect(() => assertWorksheetCacheUsage(
      { rows: XLSX_MAX_CACHED_ROWS + 1, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 }, 'x', 'worksheet/0',
    )).toThrow(OoxmlResourceLimitError);
    expect(() => assertWorksheetCacheUsage(
      { rows: 0, cells: XLSX_MAX_CACHED_CELLS + 1, ownedUtf8Bytes: 0, jsonBytes: 0 }, 'x', 'worksheet/0',
    )).toThrow(OoxmlResourceLimitError);
    expect(() => assertWorksheetCacheUsage(
      { rows: 0, cells: 0, ownedUtf8Bytes: XLSX_MAX_CACHED_OWNED_UTF8_BYTES + 1, jsonBytes: 0 },
      'x',
      'worksheet/0',
    )).toThrow(OoxmlResourceLimitError);
    expect(() => assertWorksheetCacheUsage(
      { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: XLSX_MAX_CACHED_JSON_BYTES + 1 },
      'x',
      'worksheet/0',
    )).toThrow(OoxmlResourceLimitError);
  });

  it('replaces all four cached worksheet metrics without double charging', () => {
    const current = { rows: 10, cells: 20, ownedUtf8Bytes: 30, jsonBytes: 40 };
    const previous = { rows: 3, cells: 5, ownedUtf8Bytes: 7, jsonBytes: 11 };
    const replacement = { rows: 4, cells: 6, ownedUtf8Bytes: 8, jsonBytes: 12 };

    expect(addWorksheetCacheUsage(current, replacement, previous)).toEqual({
      rows: 11,
      cells: 21,
      ownedUtf8Bytes: 31,
      jsonBytes: 41,
    });
  });
});

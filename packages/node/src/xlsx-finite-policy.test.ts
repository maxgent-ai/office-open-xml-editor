import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import type { Row, Worksheet } from '@silurus/ooxml-xlsx';
import { DEFAULT_XLSX_WORKSHEET_POLICY } from '@silurus/ooxml-core/worker';
import { OoxmlError } from '@silurus/ooxml-core';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';
// Test-only direct relative import of the realm-local policy context.
import { getWorksheetPolicy } from '../../xlsx/src/worksheet-policy-context.ts';
import { materializeXlsxWorksheet, openXlsxWorkbook } from './xlsx.ts';

type OpenOptions = NonNullable<Parameters<typeof openXlsxWorkbook>[1]>;
type XlsxWorkbookSession = Awaited<ReturnType<typeof openXlsxWorkbook>>;

const LARGE_TIMEOUT = 30_000;

let directory = '';
let small: Buffer;

beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'ooxml-xlsx-finite-policy-'));
  const fixture = join(directory, 'small.xlsx');
  // The generator emits exactly `rows` <row> elements (no separate header row).
  await generateSyntheticXlsx(fixture, { rows: 3, columns: 2 });
  small = await readFile(fixture);
});

afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

async function collectRows(stream: ReturnType<XlsxWorkbookSession['worksheetRows']>): Promise<Row[]> {
  const rows: Row[] = [];
  for await (const chunk of stream) {
    if (chunk.kind === 'rows') rows.push(...chunk.rows);
  }
  return rows;
}

function cellCount(rows: readonly Row[]): number {
  let total = 0;
  for (const row of rows) total += row.cells.length;
  return total;
}

function limitViolation(limit: number) {
  return {
    name: 'OoxmlResourceLimitError',
    code: 'ooxml-resource-limit',
    details: expect.objectContaining({
      violation: expect.objectContaining({ configurable: true, limit }),
    }),
  };
}

async function rejectionOf(promise: Promise<unknown>): Promise<unknown> {
  return promise.then(
    () => { throw new Error('expected rejection but promise resolved'); },
    (error: unknown) => error,
  );
}

describe('Node XLSX finite worksheet policy', () => {
  it('rejects materialization of 3 rows under maxRows 1', async () => {
    await expect(materializeXlsxWorksheet(small, 0, { xlsxWorksheetLimits: { maxRows: 1 } }))
      .rejects.toMatchObject(limitViolation(1));
  });

  it('streams the 3x2 worksheet unchanged under maxRows 3', async () => {
    const baseline = await openXlsxWorkbook(small);
    let expected: Row[];
    try {
      expected = await collectRows(baseline.worksheetRows(0));
    } finally {
      await baseline.close();
    }
    expect(expected).toHaveLength(3);
    expect(cellCount(expected)).toBe(6);

    const workbook = await openXlsxWorkbook(small, { xlsxWorksheetLimits: { maxRows: 3 } });
    try {
      const rows = await collectRows(workbook.worksheetRows(0));
      expect(rows).toHaveLength(3);
      expect(cellCount(rows)).toBe(6);
      expect(rows).toEqual(expected);
    } finally {
      await workbook.close();
    }
  });

  it('binds the raised session policy to the materialized worksheet context', async () => {
    const worksheet: Worksheet = await materializeXlsxWorksheet(small, 0, {
      xlsxWorksheetLimits: { maxRows: 3 },
    });
    expect(worksheet.rows).toHaveLength(3);
    expect(cellCount(worksheet.rows)).toBe(6);
    const policy = getWorksheetPolicy(worksheet);
    expect(policy).not.toBe(DEFAULT_XLSX_WORKSHEET_POLICY);
    expect(policy.worksheet.maxRows).toBe(3);
    expect(Object.isFrozen(policy)).toBe(true);
    expect(Object.isFrozen(policy.worksheet)).toBe(true);
  });

  it('captures options when open starts; later caller mutation cannot modify the session', async () => {
    const limits: { maxRows: number } = { maxRows: 1 };
    const options = { xlsxWorksheetLimits: limits };
    const pending = openXlsxWorkbook(small, options);
    limits.maxRows = 100;
    options.xlsxWorksheetLimits = { maxRows: 1000 };
    const workbook = await pending;
    try {
      let terminal: Worksheet | undefined;
      let rows = 0;
      for await (const chunk of workbook.worksheetRows(0)) {
        if (chunk.kind === 'rows') rows += chunk.rows.length;
        else terminal = chunk.worksheet;
      }
      expect(rows).toBe(3);
      expect(terminal).toBeDefined();
      expect(getWorksheetPolicy(terminal!).worksheet.maxRows).toBe(1);
    } finally {
      await workbook.close();
    }
  });

  const invalidOptions: ReadonlyArray<readonly [string, OpenOptions, typeof TypeError | typeof RangeError]> = [
    ['negative maxRows', { xlsxWorksheetLimits: { maxRows: -1 } }, RangeError],
    ['zero maxCells', { xlsxWorksheetLimits: { maxCells: 0 } }, RangeError],
    ['NaN maxRows', { xlsxWorksheetLimits: { maxRows: Number.NaN } }, RangeError],
    ['fractional maxJsonBytes', { xlsxWorksheetLimits: { maxJsonBytes: 1.5 } }, RangeError],
    ['infinite maxOwnedUtf8Bytes', { xlsxWorksheetLimits: { maxOwnedUtf8Bytes: Number.POSITIVE_INFINITY } }, RangeError],
    ['MAX_SAFE_INTEGER maxRows', { xlsxWorksheetLimits: { maxRows: Number.MAX_SAFE_INTEGER } }, RangeError],
    ['string maxRows', { xlsxWorksheetLimits: { maxRows: '5' } } as unknown as OpenOptions, TypeError],
    ['null limits', { xlsxWorksheetLimits: null } as unknown as OpenOptions, TypeError],
    ['array limits', { xlsxWorksheetLimits: [] } as unknown as OpenOptions, TypeError],
    ['scalar limits', { xlsxWorksheetLimits: 7 } as unknown as OpenOptions, TypeError],
  ];

  it.each(invalidOptions)('rejects %s before container/WASM work', async (_name, options, ErrorType) => {
    const notAnArchive = new Uint8Array([1, 2, 3, 4]);
    for (const open of [
      () => openXlsxWorkbook(notAnArchive, options),
      () => materializeXlsxWorksheet(notAnArchive, 0, options),
    ]) {
      const error = await rejectionOf(open());
      expect(error).toBeInstanceOf(ErrorType);
      expect(error).not.toBeInstanceOf(OoxmlError);
      expect(error).not.toMatchObject({ code: 'not-ooxml' });
    }
  });

  it('streams sibling sessions with different limits independently; a policy failure does not invalidate the healthy sibling', async () => {
    // Initialize the WASM generation first so this test exercises policy, not initialization.
    const warm = await openXlsxWorkbook(small);
    await warm.close();

    const [low, high] = await Promise.all([
      openXlsxWorkbook(small, { xlsxWorksheetLimits: { maxRows: 1 } }),
      openXlsxWorkbook(small, { xlsxWorksheetLimits: { maxRows: 3 } }),
    ]);
    try {
      const highStream = high.worksheetRows(0);
      const first = await highStream.next();
      expect(first.done).toBe(false);

      await expect(collectRows(low.worksheetRows(0))).resolves.toHaveLength(3);
      await expect(materializeXlsxWorksheet(small, 0, { xlsxWorksheetLimits: { maxRows: 1 } }))
        .rejects.toMatchObject(limitViolation(1));

      const rows: Row[] = [];
      if (first.value?.kind === 'rows') rows.push(...first.value.rows);
      for await (const chunk of highStream) {
        if (chunk.kind === 'rows') rows.push(...chunk.rows);
      }
      expect(rows).toHaveLength(3);
      expect(cellCount(rows)).toBe(6);

      // The healthy session remains usable afterwards.
      await expect(collectRows(high.worksheetRows(0))).resolves.toHaveLength(3);
      // Non-retaining row streams remain usable under the small materialization policy.
      await expect(collectRows(low.worksheetRows(0))).resolves.toHaveLength(3);
    } finally {
      await low.close();
      await high.close();
    }
  });
});

describe('Node XLSX finite worksheet policy on large real fixtures', () => {
  let tall: Buffer;
  let wide: Buffer;

  beforeAll(async () => {
    const tallPath = join(directory, 'tall.xlsx');
    const widePath = join(directory, 'wide.xlsx');
    const tallResult = await generateSyntheticXlsx(tallPath, { rows: 100_001, columns: 1 });
    const wideResult = await generateSyntheticXlsx(widePath, { rows: 12_550, columns: 20 });
    expect(tallResult.rows).toBe(100_001);
    expect(wideResult.cells).toBe(251_000);
    tall = await readFile(tallPath);
    wide = await readFile(widePath);
  }, 120_000);

  it('rejects materializing 100001 rows by default with a configurable row limit', async () => {
    await expect(materializeXlsxWorksheet(tall, 0)).rejects.toMatchObject(
      limitViolation(DEFAULT_XLSX_WORKSHEET_POLICY.worksheet.maxRows),
    );
  }, LARGE_TIMEOUT);

  it('materializes all 100001 rows under raised maxRows with the exact last value', async () => {
    const worksheet = await materializeXlsxWorksheet(tall, 0, {
      xlsxWorksheetLimits: { maxRows: 100_001 },
    });
    expect(worksheet.rows).toHaveLength(100_001);
    expect(cellCount(worksheet.rows)).toBe(100_001);
    const last = worksheet.rows[worksheet.rows.length - 1];
    expect(last.cells).toHaveLength(1);
    // (100001 + 1) % 5 === 2 → inline string `r100001c1`.
    expect(JSON.stringify(last.cells[0])).toContain('r100001c1');
    expect(getWorksheetPolicy(worksheet).worksheet.maxRows).toBe(100_001);
  }, LARGE_TIMEOUT);

  it('rejects 251000 cells by default with a configurable cell limit', async () => {
    expect(DEFAULT_XLSX_WORKSHEET_POLICY.worksheet.maxRows).toBeGreaterThanOrEqual(12_550);
    expect(DEFAULT_XLSX_WORKSHEET_POLICY.worksheet.maxCells).toBeLessThan(251_000);
    await expect(materializeXlsxWorksheet(wide, 0)).rejects.toMatchObject(
      limitViolation(DEFAULT_XLSX_WORKSHEET_POLICY.worksheet.maxCells),
    );
  }, LARGE_TIMEOUT);

  it('materializes 251000 cells under raised maxCells and binds the coordinate cap', async () => {
    const worksheet = await materializeXlsxWorksheet(wide, 0, {
      xlsxWorksheetLimits: { maxCells: 251_000 },
    });
    expect(worksheet.rows).toHaveLength(12_550);
    expect(cellCount(worksheet.rows)).toBe(251_000);
    const lastRow = worksheet.rows[worksheet.rows.length - 1];
    expect(lastRow.cells).toHaveLength(20);
    // (12550 + 20) % 5 === 0 → numeric 12550 * 20 + 20.
    expect(JSON.stringify(lastRow.cells[lastRow.cells.length - 1])).toContain('251020');

    const policy = getWorksheetPolicy(worksheet);
    expect(policy.worksheet.maxCells).toBe(251_000);
    expect(policy.maxCoordinateIndexEntries).toBe(
      Math.max(DEFAULT_XLSX_WORKSHEET_POLICY.maxCoordinateIndexEntries, 251_000),
    );
    expect(policy.maxCoordinateIndexEntries).toBeGreaterThanOrEqual(251_000);
    expect(policy.cache.maxCells).toBeGreaterThanOrEqual(251_000);
  }, LARGE_TIMEOUT);
});

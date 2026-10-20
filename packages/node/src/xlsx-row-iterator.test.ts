import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { spawnSync } from 'node:child_process';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import type { Row, Worksheet } from '@silurus/ooxml-xlsx';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';
import {
  materializeXlsxWorkbook,
  materializeXlsxWorkbookIndex,
  materializeXlsxWorksheet,
  openXlsxWorkbook,
  type OpenXlsxWorkbookOptions,
  type XlsxWorksheetRowChunk,
} from './xlsx.ts';

let directory = '';
let fixture = '';
let bytes: Buffer;
let worksheetBytes = 0;
let corruptBytes: Buffer;

beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'ooxml-xlsx-row-'));
  fixture = join(directory, 'parity.xlsx');
  const generated = await generateSyntheticXlsx(fixture, { rows: 2049, columns: 32 });
  worksheetBytes = generated.worksheetBytes;
  bytes = await readFile(fixture);
  const corrupt = join(directory, 'malformed-tail.xlsx');
  await generateSyntheticXlsx(corrupt, {
    rows: 513,
    columns: 8,
    malformedWorksheetTail: true,
  });
  corruptBytes = await readFile(corrupt);
});

afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

describe('Node bounded XLSX worksheet row iterator', () => {
  it('keeps a recursively frozen workbook index readable after close', async () => {
    const session = await openXlsxWorkbook(bytes);
    const index = session.workbookIndex;
    expect(Object.isFrozen(index)).toBe(true);
    expect(Object.isFrozen(index.workbook.sheets)).toBe(true);
    expect(Object.isFrozen(index.styles.cellXfs)).toBe(true);
    await session.close();
    expect(session.workbookIndex).toBe(index);
    expect(session.workbookIndex.workbook.sheets[0]?.name).toBe('Synthetic');
  });

  it('fans a worksheet trap out to sibling sessions and recovers for the next open', async () => {
    const openCursor = vi.spyOn(archivePrototype(), 'open_sheet_cursor')
      .mockImplementationOnce(() => { throw new RangeError('synthetic trap'); });
    const first = await openXlsxWorkbook(bytes);
    const sibling = await openXlsxWorkbook(bytes);
    try {
      await expect(first.worksheetRows(0).next())
        .rejects.toMatchObject({ name: 'WasmTrapError', code: 'parser-crashed' });
      await expect(sibling.worksheetRows(0).next())
        .rejects.toMatchObject({ name: 'WasmTrapError', code: 'parser-crashed' });
      await first.close();
      await sibling.close();
    } finally {
      openCursor.mockRestore();
    }

    const recovered = await openXlsxWorkbook(bytes);
    await expect(recovered.worksheetRows(0).next())
      .resolves.toMatchObject({ done: false, value: { kind: 'rows' } });
    await recovered.close();
  });

  it('materializes detached index, worksheet, and complete workbook results', async () => {
    const index = await materializeXlsxWorkbookIndex(bytes);
    const worksheet = await materializeXlsxWorksheet(bytes, 0);
    const complete = await materializeXlsxWorkbook(bytes);

    expect(index.workbook.sheets.map((sheet) => sheet.name)).toEqual(['Synthetic']);
    expect(Object.isFrozen(index)).toBe(false);
    expect(worksheet.rows).toHaveLength(2049);
    expect(complete.workbookIndex).toEqual(index);
    expect(complete.worksheets).toEqual([worksheet]);
    expect(Object.isFrozen(complete.workbookIndex)).toBe(false);
    // ~2.0 s on an idle machine; the default 5 s budget timed out when
    // the suite ran on a loaded machine.
  }, 30_000);

  it('matches the synchronous compatibility parse across multiple pulls', async () => {
    const expected = await materializeXlsxWorksheet(bytes, 0);
    const rows: Row[] = [];
    let terminal: Worksheet | undefined;
    let pulls = 0;
    for await (const chunk of streamWorksheetRows(bytes, 0)) {
      pulls++;
      if (chunk.kind === 'rows') rows.push(...chunk.rows);
      else terminal = chunk.worksheet;
    }

    expect(pulls).toBeGreaterThan(2);
    for (const chunk of await collectChunks(bytes)) {
      if (chunk.kind === 'rows') expect(chunk.rows.length).toBeLessThanOrEqual(128);
      expect(chunk.wireBytes).toBeLessThanOrEqual(64 * 1024 * 1024);
    }
    expect(rows).toEqual(expected.rows);
    expect(terminal).toBeDefined();
    expect(terminal?.rows).toEqual([]);
    expect({ ...terminal, rows }).toEqual(expected);
    // ~2.4 s on an idle machine; the default 5 s budget timed out when
    // the suite ran on a loaded machine.
  }, 30_000);

  it('honors an exact per-entry boundary and returns the canonical typed error one byte below it', async () => {
    let rowCount = 0;
    for await (const chunk of streamWorksheetRows(bytes, 0, {
      resourceLimits: { maxArchiveEntryBytes: worksheetBytes, maxTotalInflatedBytes: null },
    })) {
      if (chunk.kind === 'rows') rowCount += chunk.rows.length;
    }
    expect(rowCount).toBe(2049);

    let error: unknown;
    try {
      for await (const _chunk of streamWorksheetRows(bytes, 0, {
        resourceLimits: { maxArchiveEntryBytes: worksheetBytes - 1, maxTotalInflatedBytes: null },
      })) { /* drain */ }
    } catch (caught) {
      error = caught;
    }
    expect(error).toMatchObject({ name: 'OoxmlResourceLimitError', code: 'ooxml-resource-limit' });
    expect(String(error)).toContain('xl/worksheets/sheet1.xml');
  });

  it('cancels cleanly when a consumer returns after its first batch', async () => {
    const prototype = (xlsxWasm as unknown as {
      XlsxArchive: { prototype: { cancel_sheet_cursor(): void } };
    }).XlsxArchive.prototype;
    const cancel = vi.spyOn(prototype, 'cancel_sheet_cursor');
    const iterator = streamWorksheetRows(bytes, 0);
    const first = await iterator.next();
    expect(first.value?.kind).toBe('rows');
    await expect(iterator.return()).resolves.toEqual({ done: true, value: undefined });
    expect(cancel).toHaveBeenCalledOnce();
    cancel.mockRestore();

    // A fresh archive/session can immediately drain after early-return cleanup.
    let rows = 0;
    for await (const chunk of streamWorksheetRows(bytes, 0)) {
      if (chunk.kind === 'rows') rows += chunk.rows.length;
    }
    expect(rows).toBe(2049);
  });

  it('acknowledges the terminal only after the consumer advances past it', async () => {
    const prototype = (xlsxWasm as unknown as {
      XlsxArchive: { prototype: { acknowledge_sheet_cursor_terminal(): void } };
    }).XlsxArchive.prototype;
    const acknowledge = vi.spyOn(prototype, 'acknowledge_sheet_cursor_terminal');
    const iterator = streamWorksheetRows(bytes, 0);
    try {
      for (;;) {
        const result = await iterator.next();
        expect(result.done).toBe(false);
        if (result.value?.kind !== 'finished') continue;
        expect(result.value.worksheet.rows).toEqual([]);
        expect(acknowledge).not.toHaveBeenCalled();
        await expect(iterator.next()).resolves.toEqual({ done: true, value: undefined });
        expect(acknowledge).toHaveBeenCalledOnce();
        break;
      }
    } finally {
      await iterator.return();
      acknowledge.mockRestore();
    }
  });

  it('frees the native archive exactly once after normal completion and an ordinary error', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      for await (const _chunk of streamWorksheetRows(bytes, 0)) { /* drain */ }
      expect(free).toHaveBeenCalledTimes(1);
      await expect(streamWorksheetRows(bytes, 99).next()).rejects.toThrow(/out of range/);
      expect(free).toHaveBeenCalledTimes(2);
    } finally {
      free.mockRestore();
    }
  });

  it('cancels and frees when the consumer body throws', async () => {
    const cancel = vi.spyOn(archivePrototype(), 'cancel_sheet_cursor');
    const free = vi.spyOn(archivePrototype(), 'free');
    const consumerError = new Error('consumer failed');
    try {
      await expect(async () => {
        for await (const _chunk of streamWorksheetRows(bytes, 0)) throw consumerError;
      }).rejects.toBe(consumerError);
      expect(cancel).toHaveBeenCalledOnce();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('cancels without ACK when the consumer returns after receiving the terminal', async () => {
    const acknowledge = vi.spyOn(archivePrototype(), 'acknowledge_sheet_cursor_terminal');
    const cancel = vi.spyOn(archivePrototype(), 'cancel_sheet_cursor');
    const free = vi.spyOn(archivePrototype(), 'free');
    const iterator = streamWorksheetRows(bytes, 0);
    try {
      for (;;) {
        const result = await iterator.next();
        if (result.value?.kind !== 'finished') continue;
        await iterator.return();
        break;
      }
      expect(acknowledge).not.toHaveBeenCalled();
      expect(cancel).toHaveBeenCalledOnce();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      acknowledge.mockRestore();
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('honors an already-aborted signal and an abort after a yielded batch', async () => {
    const cancel = vi.spyOn(archivePrototype(), 'cancel_sheet_cursor');
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const pre = new AbortController();
      pre.abort();
      await expect(streamWorksheetRows(bytes, 0, { signal: pre.signal }).next())
        .rejects.toMatchObject({ name: 'AbortError' });
      expect(cancel).not.toHaveBeenCalled();
      expect(free).not.toHaveBeenCalled();

      const post = new AbortController();
      const iterator = streamWorksheetRows(bytes, 0, { signal: post.signal });
      expect((await iterator.next()).value?.kind).toBe('rows');
      post.abort();
      await expect(iterator.next()).rejects.toMatchObject({ name: 'AbortError' });
      expect(cancel).toHaveBeenCalledOnce();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('cancels and frees after malformed JSON or a terminal-marker mismatch', async () => {
    const prototype = archivePrototype();
    const cancel = vi.spyOn(prototype, 'cancel_sheet_cursor');
    const free = vi.spyOn(prototype, 'free').mockImplementation(() => {
      throw new Error('free failed');
    });
    try {
      const malformed = vi.spyOn(prototype, 'pull_sheet_cursor')
        .mockImplementationOnce(() => new TextEncoder().encode('{'));
      await expect(streamWorksheetRows(bytes, 0).next()).rejects.toBeInstanceOf(SyntaxError);
      malformed.mockRestore();

      const marker = vi.spyOn(prototype, 'sheet_cursor_pull_finished')
        .mockImplementationOnce(() => true);
      await expect(streamWorksheetRows(bytes, 0).next())
        .rejects.toThrow(/terminal marker mismatch/);
      marker.mockRestore();

      expect(cancel).toHaveBeenCalledTimes(2);
      expect(free).toHaveBeenCalledTimes(2);
    } finally {
      vi.restoreAllMocks();
    }
  });

  it('does not mask a resource error when lifecycle cleanup also fails', async () => {
    const prototype = archivePrototype();
    const resourceError = new OoxmlResourceLimitError('synthetic row limit', {
      stage: 'parsing',
      violation: {
        format: 'xlsx',
        operation: 'parse_sheet',
        resource: 'worksheet-row',
        metric: 'projected-bytes',
        part: 'xl/worksheets/sheet1.xml',
        limit: 1,
        observed: 2,
        configurable: false,
        usage: {
          archiveEntryCount: 7,
          declaredInflatedBytes: worksheetBytes,
          distinctInflatedBytes: 1,
          operationInflatedBytes: 1,
        },
      },
    });
    const pull = vi.spyOn(prototype, 'pull_sheet_cursor').mockImplementationOnce(() => {
      throw resourceError;
    });
    const cancel = vi.spyOn(prototype, 'cancel_sheet_cursor').mockImplementation(() => {
      throw new Error('cleanup failed');
    });
    const free = vi.spyOn(prototype, 'free').mockImplementation(() => {
      throw new Error('free failed');
    });
    try {
      await expect(streamWorksheetRows(bytes, 0).next()).rejects.toMatchObject({
        name: 'OoxmlResourceLimitError',
        code: 'ooxml-resource-limit',
      });
      expect(cancel).toHaveBeenCalled();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      pull.mockRestore();
      cancel.mockRestore();
      free.mockRestore();
    }
  });

  it('matches compatibility placeholder semantics for a malformed worksheet tail', async () => {
    const expected = await materializeXlsxWorksheet(corruptBytes, 0);
    const provisionalRows: Row[] = [];
    let terminal: Worksheet | undefined;
    for await (const chunk of streamWorksheetRows(corruptBytes, 0)) {
      if (chunk.kind === 'rows') provisionalRows.push(...chunk.rows);
      else terminal = chunk.worksheet;
    }
    expect(provisionalRows.length).toBeGreaterThan(0);
    expect(terminal?.parseError).toBeTruthy();
    expect(terminal?.rows).toEqual([]);
    // The terminal placeholder invalidates the already-yielded provisional rows.
    provisionalRows.length = 0;
    expect(provisionalRows).toEqual([]);
    expect(terminal).toEqual(expected);
  });

  it('never exposes or routes through the removed materializing parse_sheet export', async () => {
    expect('parse_sheet' in xlsxWasm).toBe(false);
    for await (const _chunk of streamWorksheetRows(bytes, 0)) { /* drain */ }
  });

  it('keeps streaming above the retained cell cap but rejects Node materialization', async () => {
    const output = join(directory, 'over-retained-model-cap.xlsx');
    await generateSyntheticXlsx(output, { rows: 7_813, columns: 32 });
    const large = await readFile(output);
    let rows = 0;
    let cells = 0;
    for await (const chunk of streamWorksheetRows(large, 0)) {
      if (chunk.kind !== 'rows') continue;
      rows += chunk.rows.length;
      for (const row of chunk.rows) cells += row.cells.length;
    }
    expect(rows).toBe(7_813);
    expect(cells).toBe(250_016);
    await expect(materializeXlsxWorksheet(large, 0)).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
      details: {
        violation: expect.objectContaining({
          resource: 'worksheet-model',
          metric: 'cells',
          limit: 250_000,
        }),
      },
    });
  }, 30_000);

  it('materializes the workbook index asynchronously', async () => {
    expect((await materializeXlsxWorkbookIndex(bytes)).workbook.sheets.map((sheet) => sheet.name))
      .toEqual(['Synthetic']);
  });
});

type ArchivePrototype = {
  free(): void;
  open_sheet_cursor(sheetIndex: number, name: string): void;
  acknowledge_sheet_cursor_terminal(): void;
  cancel_sheet_cursor(): void;
  pull_sheet_cursor(rowCredit: number): Uint8Array;
  sheet_cursor_pull_finished(): boolean;
};

function archivePrototype(): ArchivePrototype {
  return (xlsxWasm as unknown as { XlsxArchive: { prototype: ArchivePrototype } })
    .XlsxArchive.prototype;
}

async function collectChunks(input: Buffer) {
  const chunks = [];
  for await (const chunk of streamWorksheetRows(input, 0)) chunks.push(chunk);
  return chunks;
}

async function* streamWorksheetRows(
  input: Buffer,
  sheetIndex: number,
  options: OpenXlsxWorkbookOptions = {},
): AsyncGenerator<XlsxWorksheetRowChunk, void, void> {
  const workbook = await openXlsxWorkbook(input, options);
  let operationError: unknown;
  try {
    yield* workbook.worksheetRows(sheetIndex);
  } catch (error) {
    operationError = error;
    throw error;
  } finally {
    try {
      await workbook.close();
    } catch (cleanupError) {
      if (operationError === undefined) throw cleanupError;
    }
  }
}

describe.runIf(process.env.RUN_XLSX_ISSUE_1102 === '1')('Issue #1102 synthetic boundary and memory benchmark', () => {
  it('rejects an over-default inflated worksheet and records fresh-process metrics', async () => {
    const large = join(directory, 'over-default.xlsx');
    const target = 128 * 1024 * 1024 + 1;
    const generated = await generateSyntheticXlsx(large, {
      rows: 2049,
      columns: 32,
      targetWorksheetBytes: target,
    });
    expect(generated.worksheetBytes).toBe(target);
    const largeBytes = await readFile(large);
    await expect(async () => {
      for await (const _chunk of streamWorksheetRows(largeBytes, 0)) { /* drain */ }
    }).rejects.toMatchObject({ code: 'ooxml-resource-limit' });

    const child = spawnSync(process.execPath, [
      join(import.meta.dirname, '../scripts/bench-xlsx-row-iterator.mjs'),
      large,
      String(target),
    ], { encoding: 'utf8', env: { ...process.env, NODE_NO_WARNINGS: '1' } });
    expect(child.stderr).toBe('');
    expect(child.status).toBe(0);
    const metrics = JSON.parse(child.stdout) as Record<string, number>;
    expect(metrics).toMatchObject({ rows: 2049, cells: 2049 * 32 });
    expect(metrics.pulls).toBeGreaterThan(2);
    expect(metrics.wireBytes).toBeGreaterThan(0);
    expect(metrics.wasmPages).toBeGreaterThan(0);
    expect(metrics.maxRssKiB).toBeGreaterThan(0);
    expect(metrics.elapsedMs).toBeGreaterThan(0);

    const maxRssMiB = process.env.XLSX_BENCH_MAX_RSS_MIB;
    if (maxRssMiB) expect(metrics.maxRssKiB).toBeLessThanOrEqual(Number(maxRssMiB) * 1024);
  }, 120_000);
});

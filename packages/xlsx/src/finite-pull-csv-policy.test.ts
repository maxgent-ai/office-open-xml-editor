import { describe, expect, it, vi } from 'vitest';
import {
  PULL_SESSION_PROTOCOL,
  normalizeXlsxWorksheetPolicy,
  type NormalizedXlsxWorksheetPolicy,
  type PullSessionCommand,
  type PullSessionResponse,
} from '@silurus/ooxml-core/worker';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { parseDelimitedWorksheet, resolveDelimitedTextOptions } from './delimited-text.js';
import { getWorksheetPolicy } from './worksheet-policy-context.js';
import { measureWorksheet } from './worksheet-resource-limits.js';
import { WorksheetPullWorker } from './worksheet-pull-worker.js';
import { WorksheetPullWorker as SourceWorksheetPullWorker } from './worksheet-pull-source-worker.js';
import type { Worksheet } from './types.js';

const csv = resolveDelimitedTextOptions({ format: 'csv' });

function source(text: string): ArrayBuffer {
  const bytes = new TextEncoder().encode(text);
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return buffer;
}

function thrownLimit(operation: () => void): OoxmlResourceLimitError {
  try {
    operation();
  } catch (error) {
    if (error instanceof OoxmlResourceLimitError) return error;
    throw error;
  }
  throw new Error('expected an OoxmlResourceLimitError');
}

describe('parseDelimitedWorksheet finite policy', () => {
  it('admits exactly maxCells logical cells, including blanks, and rejects one more', () => {
    const policy = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxCells: 2 } });
    const result = parseDelimitedWorksheet(source(',\n'), csv, policy);
    expect(result.worksheet.rows).toHaveLength(1);
    expect(getWorksheetPolicy(result.worksheet)).toBe(policy);

    const error = thrownLimit(() => parseDelimitedWorksheet(source(',,\n'), csv, policy));
    expect(error.details).toMatchObject({
      violation: { metric: 'cells', limit: 2, observed: 3, configurable: true },
    });
  });

  it('admits maxRows records and rejects the second record when maxRows is 1', () => {
    const policy = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxRows: 1 } });
    const result = parseDelimitedWorksheet(source('a\n'), csv, policy);
    expect(result.worksheet.rows).toHaveLength(1);
    expect(getWorksheetPolicy(result.worksheet)).toBe(policy);

    const error = thrownLimit(() => parseDelimitedWorksheet(source('a\nb\n'), csv, policy));
    expect(error.details).toMatchObject({
      violation: { metric: 'rows', limit: 1, observed: 2, configurable: true },
    });
  });

  it('counts owned UTF-8 bytes as text discriminator plus content bytes', () => {
    const exact = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxOwnedUtf8Bytes: 6 } });
    const result = parseDelimitedWorksheet(source('ab\n'), csv, exact);
    expect(result.worksheet.rows[0]?.cells).toEqual([
      { row: 1, col: 1, value: { type: 'text', text: 'ab' } },
    ]);
    expect(getWorksheetPolicy(result.worksheet)).toBe(exact);

    const below = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxOwnedUtf8Bytes: 5 } });
    const error = thrownLimit(() => parseDelimitedWorksheet(source('ab\n'), csv, below));
    expect(error.details).toMatchObject({
      violation: { metric: 'owned-utf8-bytes', limit: 5, observed: 6, configurable: true },
    });
  });

  it('applies maxJsonBytes independently at the observed JSON size and one below', () => {
    const text = 'ab,c\nd\n';
    const baseline = parseDelimitedWorksheet(source(text), csv);
    const observed = measureWorksheet(baseline.worksheet).jsonBytes;
    expect(observed).toBeGreaterThan(1);

    const exact = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxJsonBytes: observed } });
    const result = parseDelimitedWorksheet(source(text), csv, exact);
    expect(result.worksheet.rows).toEqual(baseline.worksheet.rows);
    expect(getWorksheetPolicy(result.worksheet)).toBe(exact);

    const below = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxJsonBytes: observed - 1 } });
    const error = thrownLimit(() => parseDelimitedWorksheet(source(text), csv, below));
    expect(error.details).toMatchObject({
      violation: { metric: 'bytes', limit: observed - 1, observed, configurable: true },
    });
  });
});

const identity = { sessionId: 7, operationId: 11, generation: 3 } as const;
const usageBytes = new TextEncoder().encode(JSON.stringify({
  archiveEntryCount: 1,
  declaredInflatedBytes: 2,
  distinctInflatedBytes: 3,
  operationInflatedBytes: 4,
}));
const largeCredit = 64 * 1024 * 1024;

function command(
  requestId: number,
  body:
    | { kind: 'pull'; sequence: number; byteCredit: number }
    | { kind: 'ack'; sequence: number },
): PullSessionCommand<number> {
  return { protocol: PULL_SESSION_PROTOCOL, requestId, ...identity, ...body };
}

function terminalWorksheet(): Worksheet {
  return {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43,
    defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
    conditionalFormats: [], images: [], charts: [],
  };
}

function twoRowArchive() {
  const payloads = [
    new TextEncoder().encode(JSON.stringify({ kind: 'rows', rows: [{ index: 1, height: null, cells: [] }] })),
    new TextEncoder().encode(JSON.stringify({ kind: 'rows', rows: [{ index: 2, height: null, cells: [] }] })),
    new TextEncoder().encode(JSON.stringify({ kind: 'finished', worksheet: terminalWorksheet() })),
  ];
  let pullIndex = 0;
  const archive = {
    open_sheet_cursor: vi.fn(),
    pull_sheet_cursor: vi.fn(() => payloads[pullIndex++]),
    sheet_cursor_pull_finished: vi.fn(() => pullIndex === payloads.length),
    sheet_cursor_resource_usage: vi.fn(() => usageBytes),
    acknowledge_sheet_cursor_terminal: vi.fn(),
    cancel_sheet_cursor: vi.fn(),
    close_sheet_cursor: vi.fn(),
  };
  return { archive, payloads };
}

describe.each([
  ['ordinary', WorksheetPullWorker],
  ['source', SourceWorksheetPullWorker],
] as const)('%s worksheet pull worker finite policy', (_, Worker) => {
  it('rejects the second rows pull under maxRows 1 before any terminal ACK', async () => {
    const policy = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxRows: 1 } });
    const { archive, payloads } = twoRowArchive();
    const accepted = vi.fn();
    const worker = new Worker(() => archive, accepted, undefined, undefined, undefined, () => policy);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    const post = (response: PullSessionResponse<ArrayBuffer, number>) => replies.push(response);

    worker.reserveOpen(identity);
    await worker.open(0, 'Sheet1', identity);
    await worker.dispatch(command(1, { kind: 'pull', sequence: 0, byteCredit: largeCredit }), post);
    expect(replies.at(-1)).toMatchObject({ kind: 'chunk', done: false, byteLength: payloads[0].byteLength });
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), post);
    await worker.dispatch(command(3, { kind: 'pull', sequence: 1, byteCredit: largeCredit }), post)
      .catch(() => undefined);

    expect(replies.at(-1)).not.toMatchObject({ kind: 'chunk', sequence: 1 });
    expect(archive.acknowledge_sheet_cursor_terminal).not.toHaveBeenCalled();
    expect(accepted).not.toHaveBeenCalled();
    const sibling = vi.fn();
    const outcome = worker.run(sibling);
    await expect(outcome).rejects.toBeInstanceOf(OoxmlResourceLimitError);
    await expect(outcome).rejects.toMatchObject({
      details: { violation: { metric: 'rows', limit: 1, observed: 2, configurable: true } },
    });
    expect(sibling).not.toHaveBeenCalled();
  });

  it('binds the policy to the accepted terminal and commits only after main ACK under maxRows 2', async () => {
    const policy: NormalizedXlsxWorksheetPolicy =
      normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxRows: 2 } });
    const { archive, payloads } = twoRowArchive();
    const seen: { policy: NormalizedXlsxWorksheetPolicy; rows: number[]; terminalAcked: boolean }[] = [];
    const accepted = vi.fn((_sheetIndex: number, worksheet: Worksheet) => {
      seen.push({
        policy: getWorksheetPolicy(worksheet),
        rows: worksheet.rows.map((row) => row.index),
        terminalAcked: archive.acknowledge_sheet_cursor_terminal.mock.calls.length > 0,
      });
    });
    const worker = new Worker(() => archive, accepted, undefined, undefined, undefined, () => policy);
    const replies: PullSessionResponse<ArrayBuffer, number>[] = [];
    const post = (response: PullSessionResponse<ArrayBuffer, number>) => replies.push(response);

    worker.reserveOpen(identity);
    await worker.open(0, 'Sheet1', identity);
    await worker.dispatch(command(1, { kind: 'pull', sequence: 0, byteCredit: largeCredit }), post);
    expect(replies.at(-1)).toMatchObject({ kind: 'chunk', done: false, byteLength: payloads[0].byteLength });
    await worker.dispatch(command(2, { kind: 'ack', sequence: 0 }), post);
    await worker.dispatch(command(3, { kind: 'pull', sequence: 1, byteCredit: largeCredit }), post);
    expect(replies.at(-1)).toMatchObject({ kind: 'chunk', done: false, byteLength: payloads[1].byteLength });
    await worker.dispatch(command(4, { kind: 'ack', sequence: 1 }), post);
    await worker.dispatch(command(5, { kind: 'pull', sequence: 2, byteCredit: largeCredit }), post);
    expect(replies.at(-1)).toMatchObject({ kind: 'chunk', done: true, byteLength: payloads[2].byteLength });
    expect(archive.acknowledge_sheet_cursor_terminal).not.toHaveBeenCalled();
    expect(accepted).not.toHaveBeenCalled();

    await worker.dispatch(command(6, { kind: 'ack', sequence: 2 }), post);
    expect(archive.acknowledge_sheet_cursor_terminal).toHaveBeenCalledOnce();
    expect(accepted).toHaveBeenCalledOnce();
    expect(accepted).toHaveBeenCalledWith(
      0,
      expect.objectContaining({
        name: 'Sheet1',
        rows: [expect.objectContaining({ index: 1 }), expect.objectContaining({ index: 2 })],
      }),
      expect.objectContaining({ rows: 2, cells: 0, jsonBytes: expect.any(Number) }),
      expect.objectContaining({ operationInflatedBytes: 4 }),
    );
    expect(seen).toHaveLength(1);
    expect(seen[0]?.policy).toBe(policy);
    expect(seen[0]?.rows).toEqual([1, 2]);
    expect(seen[0]?.terminalAcked).toBe(false);
  });
});

import { normalizeXlsxWorksheetPolicy } from '@silurus/ooxml-core/worker';
import { describe, expect, it } from 'vitest';
import type { Worksheet } from '../types.js';
import { WorksheetViewProjectionCache } from '../worker-protocol.js';
import { evictWorkerWorksheets } from './worksheet-cache.js';

describe('worker worksheet eviction', () => {
  it('removes exactly the requested copies and their retained usage', () => {
    const sheets = new Map<number, Worksheet>([
      [0, { name: 'A' } as Worksheet],
      [1, { name: 'B' } as Worksheet],
    ]);
    const usages = new Map([
      [0, { rows: 70_000, cells: 70_000, ownedUtf8Bytes: 420_000, jsonBytes: 7_328_207 }],
      [1, { rows: 70_000, cells: 70_000, ownedUtf8Bytes: 420_000, jsonBytes: 7_328_207 }],
    ]);
    const retained = evictWorkerWorksheets(
      [0], sheets, usages,
      { rows: 140_000, cells: 140_000, ownedUtf8Bytes: 840_000, jsonBytes: 14_656_414 },
      new WorksheetViewProjectionCache(),
    );
    expect([...sheets.keys()]).toEqual([1]);
    expect([...usages.keys()]).toEqual([1]);
    expect(retained).toEqual(usages.get(1));
    expect(() => evictWorkerWorksheets(
      [1, 0], sheets, usages, retained, new WorksheetViewProjectionCache(),
    )).toThrow('Worksheet cache is inconsistent for sheet 0');
    expect([...sheets.keys()]).toEqual([1]);
  });
});

it('keeps raised-policy residuals exact across multiple LRU victims', () => {
  const policy = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxCells: 1_000_000 } });
  const sheets = new Map<number, Worksheet>([
    [0, { name: 'small first victim' } as Worksheet],
    [1, { name: 'large second victim' } as Worksheet],
    [2, { name: 'survivor' } as Worksheet],
  ]);
  const usage = (cells: number) => ({ rows: 1, cells, ownedUtf8Bytes: 0, jsonBytes: 0 });
  const usages = new Map([[0, usage(100_000)], [1, usage(800_000)], [2, usage(50_000)]]);
  const projections = new WorksheetViewProjectionCache();
  const residual = evictWorkerWorksheets(
    [0, 1], sheets, usages, { rows: 3, cells: 950_000, ownedUtf8Bytes: 0, jsonBytes: 0 },
    projections, policy,
  );
  expect(residual).toEqual(usage(50_000));
  expect([...sheets.keys()]).toEqual([2]);
  expect([...usages.keys()]).toEqual([2]);
});

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

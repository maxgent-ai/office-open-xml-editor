import { describe, expect, it } from 'vitest';
import { UnitFindController } from './find-controller';

const runs = (texts: string[]) => texts.map((text) => ({ text }));
const controller = (units: string[][]) => new UnitFindController(
  () => units.length,
  async (unit: number) => runs(units[unit] ?? []),
  (unit: number) => ({ unit }),
);

describe('UnitFindController', () => {
  it('finds across runs in unit order, retains case, and cycles highlights', async () => {
    const find = controller([['Hel', 'lo'], ['hello']]);
    expect(await find.find('hello')).toEqual([
      { matchIndex: 0, text: 'Hello', location: { unit: 0 } },
      { matchIndex: 1, text: 'hello', location: { unit: 1 } },
    ]);
    expect(find.next()?.location.unit).toBe(0);
    expect(find.unitHighlights(0)).toEqual([{
      slices: [{ runIndex: 0, start: 0, end: 3 }, { runIndex: 1, start: 0, end: 2 }],
      active: true,
    }]);
    expect(find.prev()?.location.unit).toBe(1);
    expect(find.activeUnit()).toBe(1);
  });

  it('rejects stale collection errors but propagates current errors', async () => {
    let reject!: (reason: Error) => void;
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((_resolve, rejectRuns) => { reject = rejectRuns; }),
      (unit) => ({ unit }),
    );
    const pending = find.find('x');
    find.invalidate();
    reject(new Error('previous load closed'));
    await expect(pending).resolves.toEqual([]);

    const current = new UnitFindController(
      () => 1, async (): Promise<{ text: string }[]> => { throw new Error('current load failed'); },
      (unit) => ({ unit }),
    );
    await expect(current.find('x')).rejects.toThrow('current load failed');
  });

  it('does not publish a scan that resolves after invalidation', async () => {
    let resolve!: (runs: { text: string }[]) => void;
    const find = new UnitFindController(
      () => 1,
      () => new Promise<{ text: string }[]>((done) => { resolve = done; }),
      (unit) => ({ unit }),
    );
    const pending = find.find('a');
    find.invalidate();
    resolve([{ text: 'a' }]);
    await expect(pending).resolves.toEqual([]);
    expect(find.unitRuns(0)).toBeUndefined();
    expect(find.matches()).toEqual([]);
  });

  it('commits a scan atomically and keeps newer visible render geometry', async () => {
    let resolveSecond!: (runs: { text: string; x: number }[]) => void;
    const original = [{ text: 'a', x: 1 }];
    const fresh = [{ text: 'a', x: 20 }];
    const find = new UnitFindController(
      () => 2,
      (unit) => unit === 0 ? Promise.resolve(original) : new Promise<{ text: string; x: number }[]>((resolve) => { resolveSecond = resolve; }),
      (unit) => ({ unit }),
    );
    const pending = find.find('a');
    await Promise.resolve();
    await Promise.resolve();
    expect(find.unitRuns(0)).toBeUndefined();
    find.setUnitRuns(0, fresh);
    resolveSecond([{ text: 'a', x: 2 }]);
    await expect(pending).resolves.toHaveLength(2);
    expect(find.unitRuns(0)).toBe(fresh);
  });

  it('clears matches and cached runs on an empty query', async () => {
    const find = controller([['a']]);
    await find.find('a');
    await expect(find.find('')).resolves.toEqual([]);
    expect(find.matches()).toEqual([]);
    expect(find.unitRuns(0)).toBeUndefined();
    expect(find.next()).toBeNull();
  });
});

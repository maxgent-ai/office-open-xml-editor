import { describe, expect, it } from 'vitest';
import { solveExactColumnBalance } from './column-balancing.js';

describe('solveExactColumnBalance', () => {
  it('keeps a terminal column break inside the balanced region when a column is left', () => {
    const result = solveExactColumnBalance({
      columnCount: 2,
      fragments: [
        { extentPt: 10, breakAfter: 'allowed' },
        { extentPt: 10, breakAfter: 'forced' },
      ],
    });
    // Both paragraphs stay in the first column; the break takes the second.
    expect(result.cutIndexes).toEqual([2]);
  });

  it('lets a terminal break in the last column follow it to the next page (control C5)', () => {
    const result = solveExactColumnBalance({
      columnCount: 2,
      fragments: [
        { extentPt: 10, breakAfter: 'forced' },
        { extentPt: 10, breakAfter: 'forced' },
      ],
    });
    expect(result.cutIndexes).toEqual([1, 2]);
  });

  it('minimizes the largest exact legal partition without probing numeric caps', () => {
    const result = solveExactColumnBalance({
      columnCount: 2,
      fragments: [30, 20, 20, 20].map((extentPt) => ({
        extentPt,
        breakAfter: 'allowed' as const,
      })),
    });

    expect(result).toMatchObject({
      targetPt: 50,
      cutIndexes: [2, 4],
    });
  });

  it('honors indivisible keep chains and authored column breaks', () => {
    const result = solveExactColumnBalance({
      columnCount: 3,
      fragments: [
        { extentPt: 20, breakAfter: 'forbidden' },
        { extentPt: 20, breakAfter: 'allowed' },
        { extentPt: 10, breakAfter: 'forced' },
        { extentPt: 30, breakAfter: 'allowed' },
      ],
    });

    expect(result.cutIndexes).toEqual([2, 3, 4]);
    expect(result.targetPt).toBe(40);
  });

  it.each([16, 64, 256])(
    'expands a linear frontier for %i legal boundaries',
    (boundaryCount) => {
      const result = solveExactColumnBalance({
        columnCount: 4,
        fragments: Array.from({ length: boundaryCount }, () => ({
          extentPt: 1,
          breakAfter: 'allowed' as const,
        })),
      });

      expect(result.targetPt).toBe(Math.ceil(boundaryCount / 4));
      expect(result.transitionExpansions).toBeLessThanOrEqual(
        3 * 4 * (boundaryCount + 1),
      );
    },
  );

  it('matches exhaustive legal partitions for small frontiers', () => {
    const bruteForce = (
      extents: readonly number[],
      breakKinds: readonly ('allowed' | 'forbidden' | 'forced')[],
      columnCount: number,
    ) => {
      const optional = breakKinds.flatMap((kind, index) => (
        index + 1 < extents.length && kind === 'allowed' ? [index + 1] : []
      ));
      const forced = breakKinds.flatMap((kind, index) => (
        index + 1 < extents.length && kind === 'forced' ? [index + 1] : []
      ));
      let best = Number.POSITIVE_INFINITY;
      for (let mask = 0; mask < 2 ** optional.length; mask += 1) {
        const cuts = [
          ...forced,
          ...optional.filter((_cut, bit) => (mask & (1 << bit)) !== 0),
          extents.length,
        ].sort((left, right) => left - right);
        if (cuts.length > columnCount) continue;
        let start = 0;
        let maximum = 0;
        for (const end of cuts) {
          maximum = Math.max(
            maximum,
            extents.slice(start, end).reduce((sum, value) => sum + value, 0),
          );
          start = end;
        }
        best = Math.min(best, maximum);
      }
      return best;
    };

    for (const extents of [
      [1, 2, 3, 4],
      [4, 1, 1, 7],
      [3, 5, 2, 6, 1],
    ]) {
      for (const breakKinds of [
        extents.map(() => 'allowed' as const),
        extents.map((_value, index) => index === 0 ? 'forbidden' as const : 'allowed' as const),
        extents.map((_value, index) => index === 1 ? 'forced' as const : 'allowed' as const),
      ]) {
        for (const columnCount of [1, 2, 3, 4]) {
          const expected = bruteForce(extents, breakKinds, columnCount);
          if (!Number.isFinite(expected)) {
            expect(() => solveExactColumnBalance({
              columnCount,
              fragments: extents.map((extentPt, index) => ({
                extentPt,
                breakAfter: breakKinds[index]!,
              })),
            })).toThrow();
            continue;
          }
          expect(solveExactColumnBalance({
            columnCount,
            fragments: extents.map((extentPt, index) => ({
              extentPt,
              breakAfter: breakKinds[index]!,
            })),
          }).targetPt).toBe(expected);
        }
      }
    }
  });
});

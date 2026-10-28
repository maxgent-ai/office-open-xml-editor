import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { buildMergeCoverageIndex } from './merge-coverage-index.js';
import { getSheetRenderCache } from './renderer.js';
import type { MergeCell, Worksheet } from './types.js';

function worksheet(mergeCells: MergeCell[]): Worksheet {
  return {
    name: 'Merges',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells,
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  };
}

/** Count coordinate-set operations while building the sheet's indexes. */
function countSetOperations(run: () => void): number {
  const { add, has } = Set.prototype;
  let operations = 0;
  Set.prototype.add = function countedAdd(this: Set<unknown>, value: unknown) {
    operations++;
    return add.call(this, value);
  };
  Set.prototype.has = function countedHas(this: Set<unknown>, value: unknown) {
    operations++;
    return has.call(this, value);
  };
  try {
    run();
  } finally {
    Set.prototype.add = add;
    Set.prototype.has = has;
  }
  return operations;
}

// A1:ZZ300 from issue #1958: 702 columns × 300 rows.
const LARGE: MergeCell = { top: 1, left: 1, bottom: 300, right: 702 };
const LARGE_AREA = 300 * 702;

describe('merge coverage index', () => {
  it('builds repeated identical ranges in the work of one range (#1958)', () => {
    const once = countSetOperations(() => getSheetRenderCache(worksheet([LARGE])));
    let cache!: ReturnType<typeof getSheetRenderCache>;
    const repeated = countSetOperations(() => {
      cache = getSheetRenderCache(worksheet(Array.from({ length: 200 }, () => ({ ...LARGE }))));
    });
    expect(cache.mergeSkipSet.size).toBe(LARGE_AREA - 1);
    expect(cache.mergeAnchorSet.size).toBe(1);
    // Each repeat may add O(1) bookkeeping, never another walk of its area.
    expect(repeated - once).toBeLessThan(200 * 16);
  });

  it('does not re-walk cells covered by distinct overlapping or contained ranges', () => {
    const nested = Array.from({ length: 200 }, (_, index) => ({
      top: 1 + (index % 100),
      left: 1 + Math.floor(index / 100),
      bottom: 300,
      right: 702,
    }));
    let cache!: ReturnType<typeof getSheetRenderCache>;
    const operations = countSetOperations(() => {
      cache = getSheetRenderCache(worksheet(nested));
    });
    expect(cache.mergeSkipSet.size).toBe(LARGE_AREA - 1);
    // A few operations per distinct covered cell plus per-range bookkeeping;
    // walking each range would cost about 200 × LARGE_AREA.
    expect(operations).toBeLessThan(3 * LARGE_AREA + 200 * 64);
  });

  it('marks exactly the cells some merge covers without being its top-left', () => {
    const expectDefinition = (merges: MergeCell[]) => {
      const expected = new Set<string>();
      for (const m of merges) {
        for (let r = m.top; r <= m.bottom; r++) {
          for (let c = m.left; c <= m.right; c++) {
            if (r !== m.top || c !== m.left) expected.add(`${r}:${c}`);
          }
        }
      }
      const cache = getSheetRenderCache(worksheet(merges));
      expect([...cache.mergeSkipSet].sort()).toEqual([...expected].sort());
      expect([...cache.mergeAnchorSet].sort())
        .toEqual([...new Set(merges.map((m) => `${m.top}:${m.left}`))].sort());
    };
    // Partial overlap, containment, a shared top-left, a top-left inside
    // another range, separated runs persisting across bands, and 1×1 ranges.
    expectDefinition([
      { top: 1, left: 1, bottom: 3, right: 3 },
      { top: 2, left: 2, bottom: 5, right: 4 },
      { top: 1, left: 1, bottom: 2, right: 5 },
      { top: 4, left: 7, bottom: 9, right: 7 },
      { top: 6, left: 10, bottom: 9, right: 10 },
      { top: 7, left: 7, bottom: 7, right: 7 },
      { top: 8, left: 12, bottom: 8, right: 12 },
      { top: 2, left: 2, bottom: 2, right: 2 },
    ]);
    // Deterministic pseudo-random ranges on a small grid (32-bit LCG).
    let seed = 1958;
    const next = (bound: number) => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return 1 + (seed >>> 8) % bound;
    };
    for (let trial = 0; trial < 200; trial++) {
      const merges: MergeCell[] = [];
      for (let i = 0; i < 8; i++) {
        const top = next(10), left = next(10);
        merges.push({ top, left, bottom: top + next(5) - 1, right: left + next(5) - 1 });
      }
      merges.push({ ...merges[0] });
      merges.push({ ...merges[1], bottom: merges[1].bottom + 1 });
      expectDefinition(merges);
    }
  });

  it('reports per-range refusals before the combined coverage budget', () => {
    const anchors = { resource: 'anchors', operation: 'index-anchors', limit: 2 };
    const coverage = { resource: 'coverage', operation: 'expand-coverage', limit: 2 };
    const refusal = (merges: MergeCell[]) => {
      try {
        buildMergeCoverageIndex(merges, anchors, coverage);
      } catch (error) {
        expect(error).toBeInstanceOf(OoxmlResourceLimitError);
        const { violation } = (error as OoxmlResourceLimitError).details as {
          violation: { resource: string; configurable: boolean };
        };
        return [violation.resource, violation.configurable];
      }
      return undefined;
    };
    // Coverage alone (3 continuations) overflows only once combined.
    const overCoverage: MergeCell[] = [
      { top: 1, left: 1, bottom: 1, right: 3 },
      { top: 2, left: 1, bottom: 2, right: 2 },
    ];
    expect(refusal(overCoverage)).toEqual(['coverage', true]);
    // A later range's anchor or malformed coordinates win over it.
    expect(refusal([...overCoverage, { top: 3, left: 1, bottom: 3, right: 1 }]))
      .toEqual(['anchors', true]);
    expect(refusal([
      { top: 1, left: 1, bottom: 1, right: 3 },
      { top: 2, left: 1, bottom: 2, right: 1.5 },
    ])).toEqual(['coverage', false]);
  });

  it('keeps an inverted range as an anchor without coverage', () => {
    const cache = getSheetRenderCache(worksheet([
      { top: 3, left: 3, bottom: 2, right: 2 },
      { top: 1, left: 1, bottom: 1, right: 1 },
    ]));
    expect(cache.mergeSkipSet.size).toBe(0);
    expect([...cache.mergeAnchorSet].sort()).toEqual(['1:1', '3:3']);
  });
});

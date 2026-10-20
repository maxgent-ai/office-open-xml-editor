import { describe, expect, it, vi } from 'vitest';
import { tableColumnLayoutInput } from '../parser-model.js';
import type { DocTable } from '../types.js';
import {
  combineAdjacentTableLayoutInputs,
  type AdjacentTableGroupLayoutInput,
} from './adjacent-table-layout-input.js';
import { resolveTableColumnLayout } from './table-columns.js';
import { stableFingerprint } from './fingerprint.js';
import type {
  TableBorderInput,
  TableEdgeInputs,
  TableLayoutInput,
  TableRowLayoutInput,
} from './types.js';

// This suite verifies the transient union-grid BUILDER output only. The builder
// is a deterministic internal layout policy (§17.4.37 does not define grid
// reconciliation) and is intentionally decoupled from paint/pagination, so the
// tests inspect the produced TableLayoutInput directly.

const source = (tableIndex: number, rowIndex?: number) => ({
  story: 'body' as const,
  storyInstance: 'body',
  path: rowIndex == null ? [tableIndex] : [tableIndex, rowIndex],
});

const border = (widthPt: number, color: string): TableBorderInput => ({
  widthPt, color, authoredStyle: 'single',
});

const edges = (overrides: Partial<TableEdgeInputs> = {}): TableEdgeInputs => ({
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
  ...overrides,
});

function row(tableIndex: number, widthCount: number): TableRowLayoutInput {
  return {
    id: `table:${tableIndex}:row:0`, source: source(tableIndex, 0), logicalRowIndex: 0,
    cantSplit: false, heightPt: 10, heightRule: 'exact', cellSpacingPt: 0,
    exceptionBorders: null, alignment: 'left', indentPt: 0, repeatedHeader: false,
    cells: [{
      id: `table:${tableIndex}:cell:0`, source: { ...source(tableIndex, 0), path: [tableIndex, 0, 0] },
      columnStart: 0, columnSpan: widthCount, verticalMerge: 'none',
      margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
      vAlign: 'top', borders: edges(), blocks: [],
    }],
  };
}

function tableInput(
  tableIndex: number,
  widths: readonly number[],
  tableEdges: TableEdgeInputs = edges(),
): TableLayoutInput {
  return {
    kind: 'table', id: `table:${tableIndex}`, source: source(tableIndex),
    flowDomainId: `table:${tableIndex}`, ordinaryFlow: true,
    alignment: 'left', indentPt: 0, bidiVisual: false,
    columnWidthsPt: widths, borders: tableEdges,
    rows: [row(tableIndex, widths.length)],
  };
}

/** Build a source TableLayoutInput whose column topology carries the exact
 * authored grid identity, exercising the exact-length seam of the union. */
function lexicalGridTableInput(tableIndex: number, widths: readonly string[]): TableLayoutInput {
  const sourceTable = {
    colWidths: [],
    rows: [{
      cells: [{
        content: [], colSpan: widths.length, vMerge: null, borders: edges(),
        background: null, vAlign: 'top', widthPt: null,
      }],
      gridBefore: 0, gridAfter: 0, isHeader: false, cantSplit: false,
    }],
    borders: edges(),
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left',
    __tableLayout: {
      effectiveStyleId: 'SharedStyle', ordinaryFlow: true,
      logicalSequenceId: 'table-sequence:0',
      logicalRowOffset: tableIndex, logicalTotalRows: 2,
      grid: {
        authored: true,
        columns: widths.map((width) => ({ width })),
        requiredColumnCount: widths.length,
      },
      preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
    },
  } as unknown as DocTable;
  const resolved = resolveTableColumnLayout(tableColumnLayoutInput(
    sourceTable,
    1000,
    () => ({ minWidthPt: 0, maxWidthPt: 0 }),
  ));
  return {
    ...tableInput(tableIndex, resolved.widthsPt),
    columnWidthKeys: resolved.widthKeys,
  };
}

function resolvedLexicalGrid(tableIndex: number, widths: readonly string[]): TableLayoutInput {
  return lexicalGridTableInput(tableIndex, widths);
}

/** A source table with one cell per column, for per-cell tiling assertions. */
function tiledTableInput(
  tableIndex: number,
  widths: readonly number[],
  overrides: Partial<TableLayoutInput> = {},
): TableLayoutInput {
  const cells = widths.map((_width, column) => ({
    id: `table:${tableIndex}:cell:${column}`,
    source: { ...source(tableIndex, 0), path: [tableIndex, 0, column] },
    columnStart: column, columnSpan: 1, verticalMerge: 'none' as const,
    margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    vAlign: 'top' as const, borders: edges(), blocks: [],
  }));
  return {
    ...tableInput(tableIndex, widths),
    ...overrides,
    rows: [{ ...row(tableIndex, widths.length), cells }],
  };
}

function cellIntervals(
  combined: AdjacentTableGroupLayoutInput,
  rowIndex: number,
): [number, number][] {
  return (combined.rows[rowIndex]?.cells ?? []).map((cell) => (
    [cell.columnStart, cell.columnStart + cell.columnSpan]
  ));
}

function rowCellWidths(combined: AdjacentTableGroupLayoutInput, rowIndex: number): number[] {
  return (combined.rows[rowIndex]?.cells ?? []).map((cell) => {
    let sum = 0;
    for (let column = cell.columnStart; column < cell.columnStart + cell.columnSpan; column += 1) {
      sum += combined.columnWidthsPt[column] ?? 0;
    }
    return sum;
  });
}

function assertContiguousTiling(combined: AdjacentTableGroupLayoutInput, rowIndex: number): void {
  const sorted = [...(combined.rows[rowIndex]?.cells ?? [])]
    .sort((left, right) => left.columnStart - right.columnStart);
  for (let index = 1; index < sorted.length; index += 1) {
    expect(sorted[index]!.columnStart)
      .toBe(sorted[index - 1]!.columnStart + sorted[index - 1]!.columnSpan);
  }
}

describe('adjacent-table logical layout input (transient union grid)', () => {
  it('rejects an empty group or a positioned member', () => {
    expect(() => combineAdjacentTableLayoutInputs('', [tableInput(0, [80])])).toThrow(/group id/);
    expect(() => combineAdjacentTableLayoutInputs('group:0', [])).toThrow(/at least one/);
    const floating: TableLayoutInput = { ...tableInput(0, [80]), ordinaryFlow: false };
    expect(() => combineAdjacentTableLayoutInputs('group:0', [floating, tableInput(1, [80])]))
      .toThrow(/absolutely positioned/);
  });

  it('uses exact grid identity when equivalent boundaries have different addition paths', () => {
    const split = lexicalGridTableInput(0, ['1', '2']);
    const whole = lexicalGridTableInput(1, ['3']);

    const combined = combineAdjacentTableLayoutInputs('group:exact', [split, whole]);

    expect(combined.columnWidthsPt).toEqual([0.05, 0.1]);
    expect(combined.columnWidthKeys).toEqual(['1/20', '1/10']);
    expect(structuredClone(combined).columnWidthKeys).toEqual(['1/20', '1/10']);
    expect(stableFingerprint('adjacent-grid', combined)).toBe(
      stableFingerprint('adjacent-grid', structuredClone(combined)),
    );
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 2 });
    expect(combined.rows[1]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 2 });
  });

  it('preserves genuinely distinct exact boundaries even when both round to one paint coordinate', () => {
    const lower = {
      ...tableInput(0, [0.15]),
      columnWidthKeys: ['3/20'],
    } as TableLayoutInput;
    const higher = {
      ...tableInput(1, [0.15]),
      columnWidthKeys: ['150000000000000001/1000000000000000000'],
    } as TableLayoutInput;

    const combined = combineAdjacentTableLayoutInputs('group:close', [lower, higher]);

    expect(combined.columnWidthKeys).toEqual([
      '3/20',
      '1/1000000000000000000',
    ]);
    expect(combined.columnWidthsPt).toHaveLength(2);
  });

  it('keeps distinct over-budget tracks unknown and prevents false boundary merges', () => {
    const first = resolvedLexicalGrid(0, [`72.${'0'.repeat(800)}1pt`]);
    const second = resolvedLexicalGrid(1, [`72.${'0'.repeat(800)}2pt`]);

    expect(first.columnWidthKeys).toEqual([null]);
    expect(second.columnWidthKeys).toEqual([null]);
    const combined = combineAdjacentTableLayoutInputs('group:unknown-distinct', [first, second]);

    expect(combined.columnWidthsPt).toEqual([72, 0]);
    expect(combined.columnWidthKeys).toEqual([null, null]);
    expect(combined.rows[0]?.cells[0].columnSpan).toBe(1);
    expect(combined.rows[1]?.cells[0].columnSpan).toBe(2);
  });

  it('fails closed for identical over-budget tracks from different source tables', () => {
    const lexical = `72.${'0'.repeat(800)}1pt`;
    const combined = combineAdjacentTableLayoutInputs('group:unknown-identical', [
      resolvedLexicalGrid(0, [lexical]),
      resolvedLexicalGrid(1, [lexical]),
    ]);

    expect(combined.columnWidthsPt).toEqual([72, 0]);
    expect(combined.columnWidthKeys).toEqual([null, null]);
    expect(combined.rows[0]?.cells[0].columnSpan).not.toBe(
      combined.rows[1]?.cells[0].columnSpan,
    );
  });

  it('does not merge an exact boundary with unknown geometry at the same coordinate', () => {
    const unknown = resolvedLexicalGrid(1, [`72.${'0'.repeat(800)}1pt`]);
    const combined = combineAdjacentTableLayoutInputs('group:exact-unknown', [
      lexicalGridTableInput(0, ['72pt']),
      unknown,
    ]);

    expect(combined.columnWidthsPt).toEqual([72, 0]);
    expect(combined.columnWidthKeys).toEqual(['72/1', null]);
  });

  it('aligns repeated rows from the same unknown source topology', () => {
    const source = resolvedLexicalGrid(0, [`72.${'0'.repeat(800)}1pt`]);
    const repeated: TableLayoutInput = { ...source, rows: [source.rows[0]!, source.rows[0]!] };
    const combined = combineAdjacentTableLayoutInputs('group:unknown-rows', [
      repeated,
      lexicalGridTableInput(1, ['72pt']),
    ]);

    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 1 });
    expect(combined.rows[1]?.cells[0]).toMatchObject({
      columnStart: combined.rows[0]?.cells[0].columnStart,
      columnSpan: combined.rows[0]?.cells[0].columnSpan,
    });
  });

  it('cancels an RTL source unknown symbol when mirrored into its own group frame', () => {
    const unknown = resolvedLexicalGrid(0, [`72.${'0'.repeat(800)}1pt`]);
    const rtl: TableLayoutInput = { ...unknown, bidiVisual: true };
    const combined = combineAdjacentTableLayoutInputs('group:rtl-unknown', [rtl]);

    expect(combined.columnWidthsPt).toEqual([72]);
    expect(combined.columnWidthKeys).toEqual([null]);
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 1 });
  });

  it('uses deterministic symbolic half-boundaries for centered unknown rows that tile', () => {
    const unknown = resolvedLexicalGrid(0, [`72.${'0'.repeat(800)}1pt`]);
    const centered: TableLayoutInput = {
      ...unknown,
      rows: [unknown.rows[0]!, { ...unknown.rows[0]!, alignment: 'center' }],
    };
    const combined = combineAdjacentTableLayoutInputs('group:center-unknown', [
      lexicalGridTableInput(1, ['144pt']),
      centered,
    ]);

    expect(combined.columnWidthsPt).toEqual([36, 36, 36, 36]);
    expect(combined.columnWidthKeys).toEqual([null, null, null, null]);
    assertContiguousTiling(combined, 1);
    assertContiguousTiling(combined, 2);
  });

  it('orders an RTL run of distinct unknown zero boundaries from row constraints', () => {
    const rtl = tiledTableInput(1, [0, 0], {
      bidiVisual: true,
      columnWidthKeys: [null, null],
    });
    const combined = combineAdjacentTableLayoutInputs('group:rtl-zero-unknown', [
      tableInput(0, [80]),
      rtl,
    ]);

    expect(combined.columnWidthsPt).toEqual([0, 0, 80]);
    expect(combined.columnWidthKeys).toEqual([null, null, null]);
    expect(cellIntervals(combined, 1)).toEqual([[1, 2], [0, 1]]);
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 3 });
  });

  it('preserves parser-model over-budget underflow zero tracks through RTL ordering', () => {
    const underflow = `0.${'0'.repeat(1200)}1pt`;
    const resolved = resolvedLexicalGrid(1, [underflow, underflow]);
    expect(resolved.columnWidthsPt).toEqual([0, 0]);
    expect(resolved.columnWidthKeys).toEqual([null, null]);
    const rtl = tiledTableInput(1, resolved.columnWidthsPt, {
      bidiVisual: true,
      columnWidthKeys: resolved.columnWidthKeys,
    });

    const combined = combineAdjacentTableLayoutInputs('group:rtl-parser-zero', [
      tableInput(0, [80]),
      rtl,
    ]);

    expect(combined.columnWidthsPt).toEqual([0, 0, 80]);
    expect(combined.columnWidthKeys).toEqual([null, null, null]);
    expect(cellIntervals(combined, 1)).toEqual([[1, 2], [0, 1]]);
  });

  it('keeps map insertions linear in unknown tracks and row boundaries', () => {
    const insertionsFor = (count: number): number => {
      const unknown = tableInput(1, Array.from({ length: count }, () => 0));
      const input: TableLayoutInput = {
        ...unknown,
        columnWidthKeys: Array.from({ length: count }, () => null),
      };
      const insertions = vi.spyOn(Map.prototype, 'set');
      try {
        combineAdjacentTableLayoutInputs(`group:complexity:${count}`, [tableInput(0, [80]), input]);
        return insertions.mock.calls.length;
      } finally { insertions.mockRestore(); }
    };

    const small = insertionsFor(100);
    const large = insertionsFor(2000);
    // Count all map work, including symbol interning and grid reconciliation.
    // Twenty times the input must stay near twenty times the insertions.
    expect(small).toBeLessThanOrEqual(16 * 100 + 64);
    expect(large).toBeLessThanOrEqual(16 * 2000 + 64);
    expect(large / small).toBeLessThanOrEqual(25);
  });

  it('unifies source grids while retaining authored row and cell identity', () => {
    const combined = combineAdjacentTableLayoutInputs('group:0', [
      tableInput(0, [80]),
      tableInput(1, [80, 80]),
    ]);

    expect(combined.columnWidthsPt).toEqual([80, 80]);
    expect(combined.id).toBe('group:0');
    // Authored SourceRef paths and per-row identity are untouched by the union.
    expect(combined.rows.map((entry) => ({
      path: entry.source.path,
      logicalRowIndex: entry.logicalRowIndex,
      cell: [entry.cells[0]?.columnStart, entry.cells[0]?.columnSpan],
      outer: [entry.sourceOuterColumnStart, entry.sourceOuterColumnEnd],
    }))).toEqual([
      { path: [0, 0], logicalRowIndex: 0, cell: [0, 1], outer: [0, 1] },
      { path: [1, 0], logicalRowIndex: 1, cell: [0, 2], outer: [0, 2] },
    ]);
  });

  it('folds each source seam into row-scoped insideH exception borders', () => {
    const inside = border(1, '#111111');
    const combined = combineAdjacentTableLayoutInputs('group:0', [
      tableInput(0, [80], edges({ bottom: border(4, '#ff0000'), insideH: inside })),
      tableInput(1, [80], edges({ top: border(5, '#0000ff'), insideH: inside })),
    ]);

    // The group grid has no whole-table border; each source table's borders are
    // folded onto its own row in the dedicated sourceTableEdges layer, and the
    // row's own exceptionBorders slot is fixed null.
    expect(combined.rows[0]?.exceptionBorders).toBeNull();
    expect(combined.rows[0]?.sourceTableEdges).toMatchObject({
      bottom: { color: '#ff0000', widthPt: 4 }, insideH: { color: '#111111', widthPt: 1 },
    });
    expect(combined.rows[1]?.sourceTableEdges).toMatchObject({
      top: { color: '#0000ff', widthPt: 5 }, insideH: { color: '#111111', widthPt: 1 },
    });
  });

  it('cascades a row-exception none border through while nil blocks source-table fallback', () => {
    const base = tableInput(0, [80], edges({ right: border(2, '#ff0000') }));
    const withNone: TableLayoutInput = {
      ...base,
      rows: base.rows.map((entry) => ({
        ...entry,
        exceptionBorders: edges({ right: { widthPt: 0, color: '#000000', authoredStyle: 'none' } }),
      })),
    };
    const withNil: TableLayoutInput = {
      ...base,
      rows: base.rows.map((entry) => ({
        ...entry,
        exceptionBorders: edges({ right: { widthPt: 0, color: '#000000', authoredStyle: 'nil' } }),
      })),
    };

    const noneCombined = combineAdjacentTableLayoutInputs('group:none', [withNone, tableInput(1, [80])]);
    // [MS-OI29500] 2.1.169: none falls through to the source table border.
    expect(noneCombined.rows[0]?.sourceTableEdges.right).toMatchObject({ color: '#ff0000', widthPt: 2 });

    const nilCombined = combineAdjacentTableLayoutInputs('group:nil', [withNil, tableInput(1, [80])]);
    // nil is retained instead of being replaced by the source table border;
    // downstream conflict resolution can still display an opposing visible edge.
    expect(nilCombined.rows[0]?.sourceTableEdges.right).toMatchObject({
      authoredStyle: 'nil', widthPt: 0,
    });
  });

  it('maps a narrower centered row into the shared grid without changing its width', () => {
    const narrow = tableInput(0, [80]);
    const centered: TableLayoutInput = {
      ...narrow,
      rows: narrow.rows.map((entry) => ({ ...entry, alignment: 'center' })),
    };
    const combined = combineAdjacentTableLayoutInputs('group:0', [
      centered,
      tableInput(1, [80, 80]),
    ]);

    expect(combined.columnWidthsPt).toEqual([40, 40, 40, 40]);
    // The centered narrow row occupies the middle physical interval [1,3).
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 1, columnSpan: 2 });
    expect([combined.rows[0]?.sourceOuterColumnStart, combined.rows[0]?.sourceOuterColumnEnd])
      .toEqual([1, 3]);
    expect(combined.rows[1]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 4 });
  });

  it('normalizes a source row indent into the group frame, mirroring it only on a bidi mismatch', () => {
    const wide = tableInput(0, [80, 80]); // group frame is LTR (first table)
    const ltrNarrow: TableLayoutInput = {
      ...tableInput(1, [80]),
      rows: tableInput(1, [80]).rows.map((entry) => ({ ...entry, indentPt: 6 })),
    };
    // Same bidi as the group frame: the indent is already group-oriented.
    const ltr = combineAdjacentTableLayoutInputs('group:ltr', [wide, ltrNarrow]);
    expect(ltr.rows[1]?.indentPt).toBe(6);
    expect('sourcePhysicalIndentPt' in ltr.rows[1]!).toBe(false);

    // Source bidi differs from the group frame: re-orient into the group frame.
    const rtlNarrow: TableLayoutInput = { ...ltrNarrow, bidiVisual: true };
    const rtl = combineAdjacentTableLayoutInputs('group:rtl', [wide, rtlNarrow]);
    expect(rtl.rows[1]?.indentPt).toBe(-6);
  });

  it('preserves a source table zero-width track through the multiset union', () => {
    // A set-based union would collapse the two coincident 80pt boundaries and
    // drop the middle zero-width column; the multiset keeps its multiplicity.
    const withZeroTrack = tableInput(0, [80, 0, 80]);
    const plain = tableInput(1, [160]);

    const combined = combineAdjacentTableLayoutInputs('group:zero', [withZeroTrack, plain]);

    expect(combined.columnWidthsPt).toEqual([80, 0, 80]);
    expect(combined.columnWidthKeys).toEqual(['80/1', '0/1', '80/1']);
    // Both source rows span the full three-column union including the zero track.
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 3 });
    expect([combined.rows[0]?.sourceOuterColumnStart, combined.rows[0]?.sourceOuterColumnEnd])
      .toEqual([0, 3]);
    expect(combined.rows[1]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 3 });
  });

  it('maps a resolved gridBefore/gridSpan cell to the correct union interval', () => {
    // Simulate a row whose acquisition already resolved gridBefore=1 and a
    // gridSpan=2 cell into columnStart=1, columnSpan=2 over a 3-column grid.
    const spanned: TableLayoutInput = {
      ...tableInput(0, [40, 40, 40]),
      rows: [{
        ...row(0, 3),
        cells: [{
          id: 'table:0:cell:0', source: source(0, 0),
          columnStart: 1, columnSpan: 2, verticalMerge: 'none',
          margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
          vAlign: 'top', borders: edges(), blocks: [],
        }],
      }],
    };
    const combined = combineAdjacentTableLayoutInputs('group:span', [spanned, tableInput(1, [120])]);

    expect(combined.columnWidthsPt).toEqual([40, 40, 40]);
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 1, columnSpan: 2 });
    expect([combined.rows[0]?.sourceOuterColumnStart, combined.rows[0]?.sourceOuterColumnEnd])
      .toEqual([0, 3]);
  });

  it('retains physical outer-edge column ownership when the group is bidiVisual', () => {
    const rtlWide: TableLayoutInput = { ...tableInput(0, [80, 80]), bidiVisual: true };
    const ltrNarrow = tableInput(1, [80], edges({
      left: border(2, '#ff0000'),
      right: border(3, '#0000ff'),
    }));
    const rightAligned: TableLayoutInput = {
      ...ltrNarrow,
      rows: ltrNarrow.rows.map((entry) => ({ ...entry, alignment: 'right' })),
    };

    const combined = combineAdjacentTableLayoutInputs('group:0', [rtlWide, rightAligned]);
    // In the bidiVisual group grid, the right-aligned narrow source table's
    // outer interval maps to logical columns [0,1) (its physical right edge is
    // the group's logical start). Its left/right borders were mirrored to stay
    // physical.
    expect([combined.rows[1]?.sourceOuterColumnStart, combined.rows[1]?.sourceOuterColumnEnd])
      .toEqual([0, 1]);
    expect(combined.rows[1]?.sourceTableEdges).toMatchObject({
      left: { color: '#0000ff' },
      right: { color: '#ff0000' },
    });
  });

  it('maps a mirrored zero-track source table by orientation, not by coordinate', () => {
    // [40,0,0,40] has three coincident interior boundaries. The mapping must
    // depend on bidi orientation (source vs group), never on the coordinates.
    const ascending = combineAdjacentTableLayoutInputs('group:asc', [
      tableInput(0, [80]),
      tiledTableInput(1, [40, 0, 0, 40]),
    ]);
    expect(ascending.columnWidthsPt).toEqual([40, 0, 0, 40]);
    expect(cellIntervals(ascending, 1)).toEqual([[0, 1], [1, 2], [2, 3], [3, 4]]);
    assertContiguousTiling(ascending, 1);

    // Same source, but bidi-mismatched with the group frame: its source-order
    // boundaries descend, so it takes the group's HIGH occurrences and its cells
    // tile the union in reversed physical order.
    const descending = combineAdjacentTableLayoutInputs('group:desc', [
      tableInput(0, [80]),
      tiledTableInput(1, [40, 0, 0, 40], { bidiVisual: true }),
    ]);
    expect(descending.columnWidthsPt).toEqual([40, 0, 0, 40]);
    expect(cellIntervals(descending, 1)).toEqual([[3, 4], [2, 3], [1, 2], [0, 1]]);
    assertContiguousTiling(descending, 1);
  });

  it('tiles an all-zero-width source table consistently in both orientations', () => {
    const ascending = combineAdjacentTableLayoutInputs('group:z-asc', [
      tableInput(0, [80]),
      tiledTableInput(1, [0, 0, 0]),
    ]);
    const descending = combineAdjacentTableLayoutInputs('group:z-desc', [
      tableInput(0, [80]),
      tiledTableInput(1, [0, 0, 0], { bidiVisual: true }),
    ]);
    expect(ascending.columnWidthsPt).toEqual([0, 0, 0, 80]);
    expect(descending.columnWidthsPt).toEqual([0, 0, 0, 80]);
    expect(cellIntervals(ascending, 1)).toEqual([[0, 1], [1, 2], [2, 3]]);
    expect(cellIntervals(descending, 1)).toEqual([[2, 3], [1, 2], [0, 1]]);
    assertContiguousTiling(ascending, 1);
    assertContiguousTiling(descending, 1);
  });

  it('preserves each source tiling when adjacent tables have mismatched zero-track multiplicity', () => {
    // Table A places two boundaries at 40; table B places three. The union keeps
    // the max (three); each source still tiles its own row without gaps.
    const combined = combineAdjacentTableLayoutInputs('group:mult', [
      tiledTableInput(0, [40, 0, 40]),
      tiledTableInput(1, [40, 0, 0, 40]),
    ]);
    expect(combined.columnWidthsPt).toEqual([40, 0, 0, 40]);
    assertContiguousTiling(combined, 0);
    assertContiguousTiling(combined, 1);
    // Each row's cell widths sum to the group width and each authored width
    // survives (multiplicities differ, geometry does not).
    expect(rowCellWidths(combined, 0)).toEqual([40, 0, 40]);
    expect(rowCellWidths(combined, 1)).toEqual([40, 0, 0, 40]);
  });

  it('carries acquisition-resolved gridBefore/gridAfter zero tracks through the union', () => {
    // Full path: authored:false grid with an omitted-width gridBefore and a
    // symmetric gridAfter -> tableColumnLayoutInput -> resolveTableColumnLayout
    // -> union builder. The empty leading/trailing tracks are zero-width and
    // must survive; the content cell keeps its offset; both tables share a seam.
    const contentCell = {
      content: [], colSpan: 1, vMerge: null, borders: edges(),
      background: null, vAlign: 'top', widthPt: 40, widthPct: null,
    };
    const acquired = {
      colWidths: [],
      rows: [{
        gridBefore: 1, gridAfter: 1,
        cells: [contentCell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
      }],
      borders: edges(),
      cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0, jc: 'left',
      __tableLayout: {
        effectiveStyleId: 'SharedStyle', ordinaryFlow: true,
        logicalSequenceId: 'table-sequence:0', logicalRowOffset: 0, logicalTotalRows: 2,
        grid: { authored: false, columns: [], requiredColumnCount: 0 },
        preferredWidth: null, layout: null, cellSpacing: null,
      },
    } as unknown as DocTable;

    const resolved = resolveTableColumnLayout(tableColumnLayoutInput(
      acquired, 200, () => ({ minWidthPt: 0, maxWidthPt: 0 }),
    ));
    // Solver keys: empty before/content/after tracks around one 40pt cell.
    expect(resolved.widthsPt).toEqual([0, 40, 0]);
    expect(resolved.widthKeys).toEqual(['0/1', '40/1', '0/1']);

    // Rebuild the source TableLayoutInput as acquisition would: the content cell
    // sits after the gridBefore track (columnStart 1), the gridAfter track is
    // empty.
    const sourceInput: TableLayoutInput = {
      ...tableInput(0, resolved.widthsPt),
      columnWidthKeys: resolved.widthKeys,
      rows: [{
        ...row(0, 1),
        cells: [{
          id: 'table:0:cell:0', source: { ...source(0, 0), path: [0, 0, 0] },
          columnStart: 1, columnSpan: 1, verticalMerge: 'none',
          margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
          vAlign: 'top', borders: edges(), blocks: [],
        }],
      }],
    };

    const combined = combineAdjacentTableLayoutInputs('group:acquired', [sourceInput, tableInput(1, [40])]);

    // Zero tracks survive; the shared seam is the exact 40pt column identity.
    expect(combined.columnWidthsPt).toEqual([0, 40, 0]);
    expect(combined.columnWidthKeys).toEqual(['0/1', '40/1', '0/1']);
    // The content cell keeps its offset over the surviving leading zero track.
    expect(combined.rows[0]?.cells[0]).toMatchObject({ columnStart: 1, columnSpan: 1 });
    expect([combined.rows[0]?.sourceOuterColumnStart, combined.rows[0]?.sourceOuterColumnEnd])
      .toEqual([0, 3]);
    // The plain [40] table absorbs the leading zero track into its single cell.
    expect(combined.rows[1]?.cells[0]).toMatchObject({ columnStart: 0, columnSpan: 2 });
  });

  it('produces a discriminated group grid that is not assignable to TableLayoutInput', () => {
    const combined = combineAdjacentTableLayoutInputs('group:kind', [
      tableInput(0, [80]),
      tableInput(1, [80]),
    ]);
    expect(combined.kind).toBe('adjacent-table-group-grid');
    expect('ordinaryFlow' in combined).toBe(false);
    expect('borders' in combined).toBe(false);
    // @ts-expect-error the group grid is a distinct kind, not a TableLayoutInput.
    const asTable: TableLayoutInput = combined;
    expect(asTable.kind).toBe('adjacent-table-group-grid');
  });
});

import { describe, expect, it } from 'vitest';
import { resolveTableColumnLayout, resolveTableColumnWidths } from './table-columns.js';
import type { TableColumnLayoutInput } from './types.js';

function input(overrides: Partial<TableColumnLayoutInput> = {}): TableColumnLayoutInput {
  return {
    layout: 'fixed',
    availableWidthPt: 200,
    gridWidthsPt: [40, 60],
    tablePreferredWidthPt: null,
    rows: [],
    ...overrides,
  };
}

describe('ECMA-376 §17.18.87 table column solver', () => {
  it('fits two growing unpreferred tracks between their content minima and maxima', () => {
    // Word's simultaneous-growth controls distinguish content-interval sharing
    // from sharing deficits over the saved grid or sharing absolute maxima.
    const solve = (availableWidthPt: number, gridWidthsPt = [25, 25]) => resolveTableColumnWidths(input({
      layout: 'autofit', availableWidthPt, gridWidthsPt, growUnpreferredColumns: true,
      rows: [{ before: null, after: null, cells: [
        { columnStart: 0, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 20, maxContentWidthPt: 300 },
        { columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 20, maxContentWidthPt: 200 },
      ] }],
    }));
    const binding = solve(300);
    expect(binding[0]).toBeCloseTo(20 + 260 * 280 / 460, 8);
    expect(binding[1]).toBeCloseTo(20 + 260 * 180 / 460, 8);
    expect(solve(600)).toEqual([300, 200]);
    // An oversized saved short track must release room for the long track.
    expect(solve(300, [280, 30])[0]).toBeCloseTo(20 + 260 * 280 / 460, 8);
  });

  it('releases unpreferred saved width beyond the content maximum for a growing neighbor', () => {
    const result = resolveTableColumnWidths(input({
      layout: 'autofit', availableWidthPt: 430, gridWidthsPt: [340, 85],
      growUnpreferredColumns: true,
      rows: [{ before: null, after: null, cells: [
        { columnStart: 0, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 96, maxContentWidthPt: 96 },
        { columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 45, maxContentWidthPt: 1400 },
      ] }],
    }));
    expect(result).toEqual([96, 334]);
  });

  it('retains the established solver result outside the two-track simultaneous-growth evidence', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit', availableWidthPt: 300, gridWidthsPt: [25, 25, 25],
      growUnpreferredColumns: true,
      rows: [{ before: null, after: null, cells: [300, 200, 100].map((maximum, column) => ({
        columnStart: column, columnSpan: 1, preferredWidth: null,
        minContentWidthPt: 20, maxContentWidthPt: maximum,
      })) }],
    }))).toEqual([25, 25, 25]);
  });

  it('constructs a zero-width grid when tblGrid is omitted and extends it for gridSpan', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [],
      rows: [{ before: null, after: null, cells: [{
        columnStart: 0, columnSpan: 3,
        preferredWidth: { kind: 'dxa', value: 90 },
        minContentWidthPt: 0, maxContentWidthPt: 0,
      }] }],
    }))).toEqual([0, 0, 90]);
  });

  it('treats a missing gridCol width as zero instead of inventing a default', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [36, 0],
      rows: [],
    }))).toEqual([36, 0]);
  });

  it('applies fixed tcW and skipped-column preferences as constraints over the initial grid', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [10, 10, 10, 10],
      rows: [{
        before: { columnSpan: 1, preferredWidth: { kind: 'dxa', value: 20 } },
        after: { columnSpan: 1, preferredWidth: { kind: 'dxa', value: 15 } },
        cells: [{
          columnStart: 1, columnSpan: 2,
          preferredWidth: { kind: 'dxa', value: 50 },
          minContentWidthPt: 0, maxContentWidthPt: 0,
        }],
      }],
    }))).toEqual([20, 25, 25, 15]);
  });

  it('proportionally reduces fixed constraints only when an authored table width requires it', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [70, 30],
      tablePreferredWidthPt: 50,
    }))).toEqual([35, 15]);
  });

  it('proportionally fits fixed tracks to the caller-projected physical boundary', () => {
    expect(resolveTableColumnWidths(input({
      availableWidthPt: 75,
      gridWidthsPt: [70, 30],
    }))).toEqual([52.5, 22.5]);
  });

  it('keeps fixed tracks unconstrained when the caller owns no width ceiling', () => {
    expect(resolveTableColumnWidths(input({
      availableWidthPt: null,
      gridWidthsPt: [70, 30],
    }))).toEqual([70, 30]);
  });

  it('gives solver-changed tracks exact keys for their final numeric definitions', () => {
    expect(resolveTableColumnLayout(input({
      availableWidthPt: 75,
      gridWidthsPt: [70, 30],
      gridWidthKeys: [null, '30/1'],
    }))).toEqual({
      widthsPt: [52.5, 22.5],
      widthKeys: ['105/2', '45/2'],
    });
  });

  it('distributes a preferred table width when every declared grid track starts at zero', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [0, 0],
      tablePreferredWidthPt: 80,
    }))).toEqual([40, 40]);
  });

  it('uses tcW even when a preferred table width exists', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [70, 30],
      tablePreferredWidthPt: 100,
      rows: [{ before: null, after: null, cells: [
        {
          columnStart: 0, columnSpan: 1,
          preferredWidth: { kind: 'dxa', value: 40 },
          minContentWidthPt: 0, maxContentWidthPt: 0,
        },
        {
          columnStart: 1, columnSpan: 1,
          preferredWidth: { kind: 'dxa', value: 60 },
          minContentWidthPt: 0, maxContentWidthPt: 0,
        },
      ] }],
    }))).toEqual([40, 60]);
  });

  it('autofit grows a spanning constraint to its minimum content width', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit',
      gridWidthsPt: [20, 20, 20],
      rows: [{ before: null, after: null, cells: [{
        columnStart: 0, columnSpan: 2, preferredWidth: null,
        minContentWidthPt: 70, maxContentWidthPt: 120,
      }] }],
    }))).toEqual([35, 35, 0]);
  });

  it('autofit may override a preferred table width up to the available band', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit',
      availableWidthPt: 120,
      gridWidthsPt: [30, 30],
      tablePreferredWidthPt: 60,
      rows: [{ before: null, after: null, cells: [
        {
          columnStart: 0, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 80, maxContentWidthPt: 100,
        },
        {
          columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 20, maxContentWidthPt: 20,
        },
      ] }],
    }))).toEqual([80, 20]);
  });

  it('resolves circular tcW percentages against the resulting table width', () => {
    expect(resolveTableColumnWidths(input({
      gridWidthsPt: [50, 50],
      rows: [
        { before: null, after: null, cells: [
          {
            columnStart: 0, columnSpan: 1,
            preferredWidth: { kind: 'pct', value: 0.5 },
            minContentWidthPt: 0, maxContentWidthPt: 0,
          },
          {
            columnStart: 1, columnSpan: 1,
            preferredWidth: { kind: 'pct', value: 0.5 },
            minContentWidthPt: 0, maxContentWidthPt: 0,
          },
        ] },
        { before: null, after: null, cells: [
          {
            columnStart: 0, columnSpan: 1,
            preferredWidth: { kind: 'dxa', value: 100 },
            minContentWidthPt: 0, maxContentWidthPt: 0,
          },
          {
            columnStart: 1, columnSpan: 1,
            preferredWidth: { kind: 'dxa', value: 50 },
            minContentWidthPt: 0, maxContentWidthPt: 0,
          },
        ] },
      ],
    }))).toEqual([100, 100]);
  });

  it('uses maximum content width when reallocating autofit slack to a deficient cell', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit',
      availableWidthPt: 100,
      gridWidthsPt: [20, 80],
      rows: [{ before: null, after: null, cells: [
        {
          columnStart: 0, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 40, maxContentWidthPt: 60,
        },
        {
          columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 20, maxContentWidthPt: 80,
        },
      ] }],
    }))).toEqual([60, 40]);
  });

  it('shrinks autofit slack before forcing widths below content minimums', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit',
      availableWidthPt: 100,
      gridWidthsPt: [90, 60],
      rows: [{ before: null, after: null, cells: [
        {
          columnStart: 0, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 70, maxContentWidthPt: 90,
        },
        {
          columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 30, maxContentWidthPt: 60,
        },
      ] }],
    }))).toEqual([70, 30]);
  });

  it('preserves a satisfiable spanning minimum while fitting autofit to the page band', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit',
      availableWidthPt: 160,
      gridWidthsPt: [100, 100, 100],
      rows: [{ before: null, after: null, cells: [
        {
          columnStart: 0, columnSpan: 2, preferredWidth: null,
          minContentWidthPt: 160, maxContentWidthPt: 200,
        },
      ] }],
    }))).toEqual([80, 80, 0]);
  });

  it('uses the first preferred cell width as the single-column maximum', () => {
    expect(resolveTableColumnWidths(input({
      layout: 'autofit',
      availableWidthPt: 100,
      gridWidthsPt: [20, 80],
      rows: [{ before: null, after: null, cells: [
        {
          columnStart: 0, columnSpan: 1,
          preferredWidth: { kind: 'dxa', value: 50 },
          minContentWidthPt: 60, maxContentWidthPt: 70,
        },
        {
          columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 20, maxContentWidthPt: 80,
        },
      ] }],
    }))).toEqual([60, 40]);
  });

  it('protects a noWrap dxa preference until the competing column reaches its minimum', () => {
    const makeInput = (noWrap: boolean) => input({
      layout: 'autofit',
      availableWidthPt: 150,
      gridWidthsPt: [100, 100],
      rows: [{ before: null, after: null, cells: [
        { columnStart: 0, columnSpan: 1, preferredWidth: { kind: 'dxa', value: 100 },
          noWrap, minContentWidthPt: 20, maxContentWidthPt: 100 },
        { columnStart: 1, columnSpan: 1, preferredWidth: { kind: 'dxa', value: 100 },
          minContentWidthPt: 20, maxContentWidthPt: 100 },
      ] }],
    });
    expect(resolveTableColumnWidths(makeInput(true))).toEqual([100, 50]);
    expect(resolveTableColumnWidths(makeInput(false))).toEqual([75, 75]);

    // Once the other column reaches its absolute minimum, the protected
    // preference can also shrink to fit the remaining available width.
    expect(resolveTableColumnWidths({ ...makeInput(true), availableWidthPt: 90 }))
      .toEqual([70, 20]);
  });

  it('protects the aggregate dxa preference of a spanning noWrap cell', () => {
    const result = resolveTableColumnWidths(input({
      layout: 'autofit', availableWidthPt: 150, gridWidthsPt: [60, 40, 100],
      rows: [{ before: null, after: null, cells: [
        { columnStart: 0, columnSpan: 2, preferredWidth: { kind: 'dxa', value: 100 },
          noWrap: true, minContentWidthPt: 20, maxContentWidthPt: 100 },
        { columnStart: 2, columnSpan: 1, preferredWidth: { kind: 'dxa', value: 100 },
          minContentWidthPt: 20, maxContentWidthPt: 100 },
      ] }],
    }));
    expect(result).toEqual([60, 40, 50]);
  });

  it('reclaims protected dxa width only to satisfy another cell minimum', () => {
    const result = resolveTableColumnWidths(input({
      layout: 'autofit', availableWidthPt: 220, gridWidthsPt: [100, 100],
      rows: [{ before: null, after: null, cells: [
        { columnStart: 0, columnSpan: 1, preferredWidth: { kind: 'dxa', value: 100 },
          noWrap: true, minContentWidthPt: 20, maxContentWidthPt: 100 },
        { columnStart: 1, columnSpan: 1, preferredWidth: null,
          minContentWidthPt: 120, maxContentWidthPt: 150 },
      ] }],
    }));
    expect(result).toEqual([80, 120]);
  });
});

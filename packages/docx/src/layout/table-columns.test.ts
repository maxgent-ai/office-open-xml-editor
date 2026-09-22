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
});

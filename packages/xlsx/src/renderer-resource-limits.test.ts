import { describe, expect, it } from 'vitest';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { deserializeWorkerError, serializeWorkerError } from '@silurus/ooxml-core/worker';
import { buildTableStyleMap, getSheetRenderCache, renderViewport } from './renderer.js';
import {
  assertCoordinateRangeArea,
  MAX_RENDERER_COORDINATE_INDEX_ENTRIES,
} from './renderer-coordinate-index.js';
import type { Cell, SparklineGroup, Styles, TableInfo, Worksheet } from './types.js';

const LIMIT = MAX_RENDERER_COORDINATE_INDEX_ENTRIES;
const EMPTY_VALUE = { type: 'empty' } as const;
const STYLES: Styles = {
  fonts: [{
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 11,
    color: null,
    name: null,
  }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function worksheet(overrides: Partial<Worksheet> = {}): Worksheet {
  return {
    name: 'Limits',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    ...overrides,
  };
}

function expectRendererPolicyLimit(run: () => unknown, resource: string, operation: string, configurable = true): void {
  try {
    run();
    throw new Error('expected renderer resource limit');
  } catch (error) {
    expect(error).toBeInstanceOf(OoxmlResourceLimitError);
    const typed = error as OoxmlResourceLimitError;
    expect(typed.code).toBe('ooxml-resource-limit');
    expect(typed.details).toMatchObject({
      stage: 'rendering',
      violation: {
        format: 'xlsx',
        resource,
        operation,
        metric: 'entry-count',
        limit: LIMIT,
        observed: LIMIT + 1,
        configurable,
      },
    });
  }
}

function table(range: TableInfo['range']): TableInfo {
  return {
    range,
    styleName: 'TableStyleLight1',
    headerRowCount: 1,
    totalsRowCount: 0,
    showRowStripes: true,
    showColumnStripes: false,
    showFirstColumn: false,
    showLastColumn: false,
    accentColor: '#4472C4',
    columns: [],
  };
}

function sparklineGroup(count: number, duplicateLast = false): SparklineGroup {
  const sparklines = Array.from({ length: count }, (_, index) => ({
    row: duplicateLast && index === count - 1 ? count - 1 : index + 1,
    col: 1,
    values: [] as (number | null)[],
  }));
  return {
    kind: 'line',
    markers: false,
    high: false,
    low: false,
    first: false,
    last: false,
    negative: false,
    displayXAxis: false,
    displayEmptyCellsAs: 'gap',
    minAxisType: 'individual',
    maxAxisType: 'individual',
    lineWeight: 0.75,
    sparklines,
  };
}

function recordingCtx(width = 200, height = 100): CanvasRenderingContext2D {
  const ctx: Record<string, unknown> = {
    canvas: { width, height },
    font: '11px sans-serif',
    fillStyle: '#000',
    strokeStyle: '#000',
    lineWidth: 1,
    textBaseline: 'alphabetic',
    textAlign: 'left',
    letterSpacing: '0px',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: text.length * 8 }),
    fillText: () => {},
    strokeText: () => {},
    fillRect: () => {},
    strokeRect: () => {},
    clearRect: () => {},
    beginPath: () => {},
    closePath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    rect: () => {},
    arc: () => {},
    fill: () => {},
    stroke: () => {},
    clip: () => {},
    save: () => {},
    restore: () => {},
    translate: () => {},
    rotate: () => {},
    scale: () => {},
    setLineDash: () => {},
    setTransform: () => {},
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

describe('XLSX renderer coordinate-index hard limits', () => {
  it('allows exactly 250,000 unique cells, counts duplicate coordinates once, and shares the lookup with CF', () => {
    const cells = Array.from({ length: LIMIT }, (_, index): Cell => ({
      row: 1,
      col: index + 1,
      value: EMPTY_VALUE,
    }));
    cells.push({ row: 1, col: LIMIT, value: EMPTY_VALUE });

    const cache = getSheetRenderCache(worksheet({ rows: [{ index: 1, height: null, cells }] }));

    expect(cache.cellMap.size).toBe(LIMIT);
    expect(cache.cfContext.cellIndex).toBe(cache.cellMap);
  });

  it('rejects the 250,001st unique cell with a fatal typed rendering error', () => {
    const cells = Array.from({ length: LIMIT + 1 }, (_, index): Cell => ({
      row: 1,
      col: index + 1,
      value: EMPTY_VALUE,
    }));

    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({ rows: [{ index: 1, height: null, cells }] })),
      'worksheet-cell-index',
      'index-worksheet-cells',
    );
  });

  it('allows the merge-skip boundary and overlapping coordinates without double charging', () => {
    const exact = { top: 1, bottom: 1, left: 1, right: LIMIT + 1 };
    const cache = getSheetRenderCache(worksheet({ mergeCells: [exact, exact] }));
    expect(cache.mergeSkipSet.size).toBe(LIMIT);
  });

  it('rejects a merge-skip range one entry above the boundary before expansion', () => {
    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({
        mergeCells: [{ top: 1, bottom: 1, left: 1, right: LIMIT + 2 }],
      })),
      'worksheet-merge-skip-index',
      'expand-merged-cell-coordinates',
    );
  });

  it('rejects cumulative unique merge-skip entries across individually safe disjoint ranges', () => {
    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({
        mergeCells: [
          { top: 1, bottom: 1, left: 1, right: 125_001 },
          { top: 2, bottom: 2, left: 1, right: 125_002 },
        ],
      })),
      'worksheet-merge-skip-index',
      'expand-merged-cell-coordinates',
    );
  });

  it('allows the styled-table boundary and an overlapping coordinate without double charging', () => {
    const ws = worksheet({
      tables: [
        table({ top: 1, bottom: 500, left: 1, right: 500 }),
        table({ top: 1, bottom: 1, left: 1, right: 1 }),
      ],
    });
    expect(buildTableStyleMap(ws).size).toBe(LIMIT);
  });

  it('rejects a styled-table range one entry above the boundary before expansion', () => {
    expectRendererPolicyLimit(
      () => buildTableStyleMap(worksheet({
        tables: [table({ top: 1, bottom: 1, left: 1, right: LIMIT + 1 })],
      })),
      'worksheet-table-style-index',
      'expand-styled-table-coordinates',
    );
  });

  it('rejects cumulative unique table-style entries across individually safe disjoint ranges', () => {
    expectRendererPolicyLimit(
      () => buildTableStyleMap(worksheet({
        tables: [
          table({ top: 1, bottom: 250, left: 1, right: 500 }),
          table({ top: 251, bottom: 500, left: 1, right: 500 }),
          table({ top: 501, bottom: 501, left: 1, right: 1 }),
        ],
      })),
      'worksheet-table-style-index',
      'expand-styled-table-coordinates',
    );
  });

  it('allows the auto-filter boundary and rejects one additional coordinate', () => {
    expect(getSheetRenderCache(worksheet({
      autoFilter: { top: 1, bottom: 1, left: 1, right: LIMIT },
    })).autoFilterCells.size).toBe(LIMIT);

    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({
        autoFilter: { top: 1, bottom: 1, left: 1, right: LIMIT + 1 },
      })),
      'worksheet-auto-filter-index',
      'expand-auto-filter-coordinates',
    );
  });

  it('counts duplicate hyperlinks once at the boundary and rejects the next unique coordinate', () => {
    const exactWithDuplicate = Array.from({ length: LIMIT + 1 }, (_, index) => ({
      row: index === LIMIT ? LIMIT : index + 1,
      col: 1,
      url: 'https://example.test',
    }));
    expect(getSheetRenderCache(worksheet({ hyperlinks: exactWithDuplicate })).hyperlinkMap.size)
      .toBe(LIMIT);

    const over = Array.from({ length: LIMIT + 1 }, (_, index) => ({
      row: index + 1,
      col: 1,
      url: 'https://example.test',
    }));
    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({ hyperlinks: over })),
      'worksheet-hyperlink-index',
      'index-hyperlink-coordinates',
    );
  });

  it('counts duplicate comments once at the boundary and rejects the next unique coordinate', () => {
    const exactWithDuplicate = Array.from(
      { length: LIMIT + 1 },
      (_, index) => `A${index === LIMIT ? LIMIT : index + 1}`,
    );
    expect(getSheetRenderCache(worksheet({ commentRefs: exactWithDuplicate })).commentCells.size)
      .toBe(LIMIT);

    const over = Array.from({ length: LIMIT + 1 }, (_, index) => `A${index + 1}`);
    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({ commentRefs: over })),
      'worksheet-comment-index',
      'index-comment-coordinates',
    );
  });

  it('counts duplicate sparklines once at the boundary and rejects the next unique coordinate', () => {
    expect(getSheetRenderCache(worksheet({
      sparklineGroups: [sparklineGroup(LIMIT + 1, true)],
    })).sparklineMap.size).toBe(LIMIT);

    expectRendererPolicyLimit(
      () => getSheetRenderCache(worksheet({
        sparklineGroups: [sparklineGroup(LIMIT + 1)],
      })),
      'worksheet-sparkline-index',
      'index-sparkline-coordinates',
    );
  });

  it('allows exactly 250,000 per-frame merge anchors through a real render call', () => {
    const mergeCells = Array.from({ length: LIMIT }, (_, index) => ({
      top: index + 1,
      bottom: index + 1,
      left: 1,
      right: 1,
    }));
    expect(() => renderViewport(
      recordingCtx(),
      worksheet({ mergeCells }),
      STYLES,
      { row: 1, col: 1, rows: 0, cols: 0 },
    )).not.toThrow();
  });

  it('propagates the 250,001st per-frame merge anchor from a real render call', () => {
    const mergeCells = Array.from({ length: LIMIT + 1 }, (_, index) => ({
      top: index + 1,
      bottom: index + 1,
      left: 1,
      right: 1,
    }));
    expectRendererPolicyLimit(
      () => renderViewport(
        recordingCtx(),
        worksheet({ mergeCells }),
        STYLES,
        { row: 1, col: 1, rows: 0, cols: 0 },
      ),
      'worksheet-merge-anchor-index',
      'index-merge-anchor-coordinates',
    );
  });

  it('handles malformed and maximum-safe ranges without arithmetic overflow or allocation', () => {
    const identity = {
      resource: 'worksheet-table-style-index',
      operation: 'expand-styled-table-coordinates',
    };
    expect(assertCoordinateRangeArea(
      { top: 2, bottom: 1, left: 1, right: Number.MAX_SAFE_INTEGER },
      identity,
    )).toBe(0);
    expect(assertCoordinateRangeArea(
      {
        top: Number.MAX_SAFE_INTEGER,
        bottom: Number.MAX_SAFE_INTEGER,
        left: Number.MAX_SAFE_INTEGER,
        right: Number.MAX_SAFE_INTEGER,
      },
      identity,
    )).toBe(1);
    expectRendererPolicyLimit(
      () => assertCoordinateRangeArea(
        { top: 1, bottom: Number.MAX_SAFE_INTEGER, left: 1, right: 1 },
        identity,
      ),
      identity.resource,
      identity.operation,
    );
    expectRendererPolicyLimit(
      () => assertCoordinateRangeArea(
        { top: 1, bottom: 1, left: 1, right: Number.POSITIVE_INFINITY },
        identity,
      ),
      identity.resource,
      identity.operation,
      false,
    );
    expectRendererPolicyLimit(
      () => assertCoordinateRangeArea(
        { top: 1, bottom: 1.5, left: 1, right: 1 },
        identity,
      ),
      identity.resource,
      identity.operation,
      false,
    );
  });

  it('retries cache construction after a limit throw instead of reusing a partial entry', () => {
    const ws = worksheet({
      autoFilter: { top: 1, bottom: 1, left: 1, right: LIMIT + 1 },
    });
    expectRendererPolicyLimit(
      () => getSheetRenderCache(ws),
      'worksheet-auto-filter-index',
      'expand-auto-filter-coordinates',
    );

    ws.autoFilter = { top: 1, bottom: 1, left: 1, right: 1 };
    const cache = getSheetRenderCache(ws);
    expect(cache.autoFilterCells).toEqual(new Set(['1:1']));
    expect(cache.cfContext.cellIndex).toBe(cache.cellMap);
  });

  it('preserves renderer limit errors as typed fatal errors across the worker wire', () => {
    let thrown: unknown;
    try {
      buildTableStyleMap(worksheet({
        tables: [table({ top: 1, bottom: 1, left: 1, right: LIMIT + 1 })],
      }));
    } catch (error) {
      thrown = error;
    }

    expect(thrown).toBeInstanceOf(OoxmlResourceLimitError);
    const restored = deserializeWorkerError(serializeWorkerError(thrown));
    expect(restored).toBeInstanceOf(OoxmlResourceLimitError);
    expect((restored as OoxmlResourceLimitError).details).toEqual(
      (thrown as OoxmlResourceLimitError).details,
    );
    expect((restored as OoxmlResourceLimitError).details.violation).not.toHaveProperty('part');
  });
});

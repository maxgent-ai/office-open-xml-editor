import { afterEach, describe, expect, it, vi } from 'vitest';
import type { Worksheet } from '../types.js';
import {
  colWidthToPx,
  getGridGeometryForWorksheet,
  getMdwForWorksheet,
  rowHeightToPx,
} from '../renderer.js';
import { GridAxisGeometry, GridGeometry } from './grid-geometry.js';
import { parseDelimitedWorksheet, resolveDelimitedTextOptions } from '../delimited-text.js';

function worksheet(): Worksheet {
  return {
    name: 'Geometry',
    rows: [],
    colWidths: { 1: 12, 2: 20, 4: 30, 16_384: 5 },
    rowHeights: { 1: 25, 2: 40, 5: 60, 1_048_576: 12 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 2,
    freezeCols: 2,
    conditionalFormats: [],
    images: [],
    charts: [],
  } as Worksheet;
}

afterEach(() => vi.unstubAllGlobals());

/** Future parser fact (XLSX1709). The intersection lets the RED phase compile. */
type UiWorksheet = Worksheet & { defaultColWidthIsUi?: boolean };

/** Model shape for a part with no `<sheetFormatPr>`. Column 4 is authored. */
function absentFormatWorksheet(): UiWorksheet {
  const ws: UiWorksheet = worksheet();
  ws.colWidths = { 4: 20 };
  ws.defaultColWidthIsUi = true;
  return ws;
}

const IMPLICIT_COL = 3;
// ECMA-376 §18.3.1.13: 8.43 UI characters at MDW 6/7/8 are 56/64/72 px. The
// same 8.43 authored as a stored width decodes to 51/59/67 px.
const UNIT_CASES = [
  { mdw: 6, ui: 56, stored: 51 },
  { mdw: 7, ui: 64, stored: 59 },
  { mdw: 8, ui: 72, stored: 67 },
] as const;

describe('XLSX1709 absent sheetFormatPr fallback width', () => {
  for (const platform of ['MacIntel', 'Win32'] as const) {
    it(`encodes the flagged 8.43 fallback as UI characters (${platform})`, () => {
      vi.stubGlobal('navigator', { platform, maxTouchPoints: 0 });
      for (const { mdw, ui } of UNIT_CASES) {
        const geometry = GridGeometry.forWorksheet(absentFormatWorksheet(), mdw);
        expect(geometry.col.sizeOf(IMPLICIT_COL)).toBe(ui);
      }
    });

    it(`keeps an unflagged 8.43 default as a stored width (${platform})`, () => {
      vi.stubGlobal('navigator', { platform, maxTouchPoints: 0 });
      for (const { mdw, stored } of UNIT_CASES) {
        const ws = absentFormatWorksheet();
        delete ws.defaultColWidthIsUi;
        expect(GridGeometry.forWorksheet(ws, mdw).col.sizeOf(IMPLICIT_COL)).toBe(stored);
      }
    });

    it(`keeps explicit column 4 = 20 as a stored width (${platform})`, () => {
      vi.stubGlobal('navigator', { platform, maxTouchPoints: 0 });
      expect(GridGeometry.forWorksheet(absentFormatWorksheet(), 7).col.sizeOf(4)).toBe(140);
    });

    it(`lets the existing baseColWidth branch win over the flag (${platform})`, () => {
      vi.stubGlobal('navigator', { platform, maxTouchPoints: 0 });
      const flagged = absentFormatWorksheet();
      flagged.baseColWidth = 8;
      const plain = worksheet();
      plain.colWidths = { 4: 20 };
      plain.baseColWidth = 8;
      expect(GridGeometry.forWorksheet(flagged, 7).col.sizeOf(IMPLICIT_COL))
        .toBe(GridGeometry.forWorksheet(plain, 7).col.sizeOf(IMPLICIT_COL));
    });
  }

  it('preserves the flag and its geometry through JSON and structuredClone', () => {
    const ws = absentFormatWorksheet();
    const copies: UiWorksheet[] = [JSON.parse(JSON.stringify(ws)), structuredClone(ws)];
    for (const copy of copies) {
      expect(copy).toMatchObject({ defaultColWidth: 8.43, defaultColWidthIsUi: true });
      const geometry = GridGeometry.forWorksheet(copy, 7);
      expect(geometry.col.sizeOf(IMPLICIT_COL)).toBe(64);
      expect(geometry.col.sizeOf(4)).toBe(140);
    }
  });

  it('builds a distinct cached geometry when MDW changes', () => {
    const ws = absentFormatWorksheet();
    const at7 = GridGeometry.forWorksheet(ws, 7);
    const at8 = GridGeometry.forWorksheet(ws, 8);
    expect(at8).not.toBe(at7);
    expect([at7.col.sizeOf(IMPLICIT_COL), at8.col.sizeOf(IMPLICIT_COL)]).toEqual([64, 72]);
  });

  it('leaves CSV worksheets unflagged with stored-width semantics', () => {
    const bytes = new TextEncoder().encode('a,b\r\n').buffer as ArrayBuffer;
    const { worksheet: csv } = parseDelimitedWorksheet(bytes, resolveDelimitedTextOptions({ format: 'csv' }));
    expect((csv as UiWorksheet).defaultColWidthIsUi).toBeUndefined();
    expect(GridGeometry.forWorksheet(csv, 7).col.sizeOf(1)).toBe(59);
  });
});

describe('GridGeometry', () => {
  it('derives implicit base-width columns without overriding an authored default width', () => {
    vi.stubGlobal('navigator', { platform: 'MacIntel', maxTouchPoints: 0 });
    const ws = worksheet();
    ws.colWidths = { 1: 12 };
    ws.baseColWidth = 10;
    expect(GridGeometry.forWorksheet(ws, 8).col.sizeOf(2)).toBe(87);

    ws.baseColWidth = undefined;
    GridGeometry.invalidate(ws);
    expect(GridGeometry.forWorksheet(ws, 8).col.sizeOf(2)).toBe(colWidthToPx(8.43, 8));
  });

  it('uses the specification pixel width for implicit base columns outside Mac Excel', () => {
    vi.stubGlobal('navigator', { platform: 'Win32', maxTouchPoints: 0 });
    const ws = worksheet();
    ws.baseColWidth = 10;
    expect(GridGeometry.forWorksheet(ws, 8).col.sizeOf(3)).toBe(85);
  });
  it('keeps offsets finite when a malformed default size reaches the geometry boundary', () => {
    const axis = new GridAxisGeometry({ 1: 10 }, Number.NaN, (value) => value, 10);

    expect(axis.offsetOf(1)).toBe(0);
    expect(axis.offsetOf(2)).toBe(10);
    expect(axis.offsetOf(10)).toBe(10);
  });

  it('jumps over a worksheet-limit run of default-hidden bands in bounded output', () => {
    const axis = new GridAxisGeometry(
      { 1_048_576: 20 },
      0,
      (value) => value,
      1_048_576,
    );

    expect(axis.bandsToCover(1, 1_048_576, 100)).toEqual([
      { index: 1_048_576, size: 20 },
    ]);
  });

  it('keeps an explicit visible row inside a zero-height default sheet', () => {
    const ws = worksheet();
    ws.defaultRowHeight = 0;
    ws.rowHeights = { 3: 22 };
    const geometry = GridGeometry.forWorksheet(ws, 8);

    expect(geometry.row.bandsToCover(1, 1_048_576, 100)).toEqual([
      { index: 3, size: rowHeightToPx(22) },
    ]);
  });

  it('reuses one rounded axis pair for every consumer in the same scale', () => {
    const geometry = GridGeometry.forWorksheet(worksheet(), 8);
    const first = geometry.axesAtScale(1.25);

    expect(geometry.axesAtScale(1.25)).toBe(first);
    expect(geometry.axesAtScale(1)).not.toBe(first);
  });

  it('measures MDW once per worksheet geometry lifetime', () => {
    let canvases = 0;
    vi.stubGlobal('OffscreenCanvas', class {
      constructor() { canvases++; }
      getContext() {
        return { font: '', measureText: () => ({ width: 8 }) };
      }
    });
    const ws = worksheet();
    ws.defaultFontFamily = 'Metric Test';
    ws.defaultFontSize = 11;

    expect(getGridGeometryForWorksheet(ws)).toBe(getGridGeometryForWorksheet(ws));
    expect(canvases).toBe(1);
    GridGeometry.invalidate(ws);
    getGridGeometryForWorksheet(ws);
    expect(canvases).toBe(2);
  });

  it('rebuilds cached column axes when the active font realm changes MDW', () => {
    const ws = worksheet();
    const narrow = GridGeometry.forWorksheet(ws, 6);
    const wide = GridGeometry.forWorksheet(ws, 8);

    expect(wide).not.toBe(narrow);
    expect(wide.col.sizeOf(1)).not.toBe(narrow.col.sizeOf(1));
    expect(GridGeometry.forWorksheet(ws, 8)).toBe(wide);
  });

  it('applies compact authored column-width ranges that omit customWidth', () => {
    const ws = worksheet();
    ws.colWidths = { 1: 30.1640625 };
    ws.colWidthRanges = [{ min: 2, max: 16_384, width: 10.83203125 }];
    ws.defaultColWidth = 8.43;
    const geometry = GridGeometry.forWorksheet(ws, 8);

    expect(geometry.col.sizeOf(1)).toBe(241);
    expect(geometry.col.sizeOf(2)).toBe(87);
    expect(geometry.col.sizeOf(16_384)).toBe(87);
    const anchorWidth = geometry.col.offsetOf(10) - geometry.col.offsetOf(1)
      + (539_749 - 298_449) / 9_525;
    expect(anchorWidth).toBeCloseTo(962.33, 2);
  });

  it('resolves many overlapping full-sheet ranges without expanding each declaration', () => {
    const ws = worksheet();
    ws.colWidths = {};
    ws.colWidthRanges = Array.from({ length: 50_000 }, (_, index) => ({
      min: 1,
      max: 16_384,
      width: 9 + (index % 2),
    }));

    const geometry = GridGeometry.forWorksheet(ws, 8);
    expect(geometry.col.sizeOf(1)).toBe(colWidthToPx(10, 8));
    expect(geometry.col.sizeOf(16_384)).toBe(colWidthToPx(10, 8));
  });

  it('skips a leading run of zero-sized rows and columns at offset zero', () => {
    const ws = worksheet();
    ws.rowHeights = { 1: 0, 2: 0 };
    ws.colWidths = { 1: 0, 2: 0 };
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));

    expect(geometry.row.indexAt(0)).toEqual({ index: 3, partial: 0 });
    expect(geometry.col.indexAt(0)).toEqual({ index: 3, partial: 0 });
    expect(geometry.cellAt(0, 0, { scrollX: 0, scrollY: 0 })).toEqual({ row: 3, col: 3 });
  });

  it('computes far-cell rectangles from cumulative axes while preserving per-cell scale rounding', () => {
    const ws = worksheet();
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));
    const scale = 1.25;
    const scrollX = 231.5;
    const scrollY = 417.25;
    const rect = geometry.cellRect(1_048_576, 16_384, {
      scale,
      scrollX,
      scrollY,
      headerWidth: 48,
      headerHeight: 24,
    });

    const sp = (value: number) => Math.round(value * scale);
    const defaultCol = sp(colWidthToPx(ws.defaultColWidth, getMdwForWorksheet(ws)));
    const defaultRow = sp(rowHeightToPx(ws.defaultRowHeight));
    const scaledCol = (index: number) => sp(colWidthToPx(ws.colWidths[index] ?? ws.defaultColWidth, getMdwForWorksheet(ws)));
    const scaledRow = (index: number) => sp(rowHeightToPx(ws.rowHeights[index] ?? ws.defaultRowHeight));
    const sparseColDelta = [1, 2, 4]
      .reduce((sum, index) => sum + scaledCol(index) - defaultCol, 0);
    const sparseRowDelta = [1, 2, 5]
      .reduce((sum, index) => sum + scaledRow(index) - defaultRow, 0);
    const frozenW = scaledCol(1) + scaledCol(2);
    const frozenH = scaledRow(1) + scaledRow(2);
    const scaledColAxis = geometry.col.scaled(scale);
    const scaledRowAxis = geometry.row.scaled(scale);
    const scaledColStart = scaledColAxis.indexAt(scrollX + scaledColAxis.offsetOf(3));
    const scaledRowStart = scaledRowAxis.indexAt(scrollY + scaledRowAxis.offsetOf(3));
    const expectedX = sp(48) + frozenW - scaledColStart.partial
      + (16_384 - 1) * defaultCol + sparseColDelta
      - ((scaledColStart.index - 1) * defaultCol
        + [1, 2, 4].filter((index) => index < scaledColStart.index)
          .reduce((sum, index) => sum + scaledCol(index) - defaultCol, 0));
    const expectedY = sp(24) + frozenH - scaledRowStart.partial
      + (1_048_576 - 1) * defaultRow + sparseRowDelta
      - ((scaledRowStart.index - 1) * defaultRow
        + [1, 2, 5].filter((index) => index < scaledRowStart.index)
          .reduce((sum, index) => sum + scaledRow(index) - defaultRow, 0));

    expect(rect).toEqual({
      x: expectedX,
      y: expectedY,
      w: scaledCol(16_384),
      h: scaledRow(1_048_576),
    });
  });

  it('rejects coordinates outside the worksheet limits', () => {
    const geometry = GridGeometry.forWorksheet(worksheet(), 7);
    const options = {
      scale: 1,
      scrollX: 0,
      scrollY: 0,
      headerWidth: 48,
      headerHeight: 24,
    };
    expect(geometry.cellRect(0, 1, options)).toBeNull();
    expect(geometry.cellRect(1_048_577, 1, options)).toBeNull();
    expect(geometry.cellRect(1, 0, options)).toBeNull();
    expect(geometry.cellRect(1, 16_385, options)).toBeNull();
  });

  it('derives hit tests and visible ranges from the same frozen-pane axes', () => {
    const ws = worksheet();
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));
    const frozen = geometry.logicalFrozenExtent();
    const scrollX = 51;
    const scrollY = 37;

    expect(geometry.cellAt(frozen.width + 1, frozen.height + 1, { scrollX, scrollY })).toEqual({
      row: geometry.row.scrollableIndexAt(1 + scrollY, 3),
      col: geometry.col.scrollableIndexAt(1 + scrollX, 3),
    });
    const visible = geometry.visibleRange({
      width: 640,
      height: 480,
      scale: 1,
      scrollX,
      scrollY,
      headerWidth: 48,
      headerHeight: 24,
      buffer: 2,
    });
    expect(visible.range.row).toBeGreaterThanOrEqual(3);
    expect(visible.range.col).toBeGreaterThanOrEqual(3);
    expect(visible.range.rows).toBeGreaterThan(2);
    expect(visible.range.cols).toBeGreaterThan(2);
  });

  it('matches the former visible-band walk across scales and partial scroll offsets', () => {
    const ws = worksheet();
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));
    const oldCount = (
      axis: typeof geometry.row,
      start: number,
      partial: number,
      available: number,
      max: number,
    ) => {
      let accumulated = -partial;
      let index = start;
      let count = 0;
      while (accumulated < available + partial && index <= max) {
        accumulated += axis.sizeOf(index);
        count++;
        index++;
      }
      return count + 2;
    };

    for (const scale of [0.75, 1, 1.25]) {
      for (const [scrollX, scrollY] of [[0, 0], [51, 37], [1234.5, 987.25]]) {
        const width = 641;
        const height = 479;
        const visible = geometry.visibleRange({
          width,
          height,
          scale,
          scrollX,
          scrollY,
          headerWidth: 48,
          headerHeight: 24,
          buffer: 2,
        });
        const frozen = geometry.logicalFrozenExtent();
        const availableWidth = width / scale - 48 - frozen.width;
        const availableHeight = height / scale - 24 - frozen.height;
        expect(visible.range.cols).toBe(oldCount(
          geometry.col,
          visible.range.col,
          visible.offsetX,
          availableWidth,
          16_384,
        ));
        expect(visible.range.rows).toBe(oldCount(
          geometry.row,
          visible.range.row,
          visible.offsetY,
          availableHeight,
          1_048_576,
        ));
      }
    }
  });

  it('does not inspect every preceding band when locating a far cell', () => {
    let rowReads = 0;
    const ws = worksheet();
    ws.rowHeights = new Proxy(ws.rowHeights, {
      get(target, key, receiver) {
        if (typeof key === 'string' && /^\d+$/.test(key)) rowReads++;
        return Reflect.get(target, key, receiver);
      },
    });
    const geometry = GridGeometry.forWorksheet(ws, getMdwForWorksheet(ws));

    expect(geometry.cellRect(1_048_576, 16_384, {
      scale: 1.25,
      scrollX: 0,
      scrollY: 0,
      headerWidth: 48,
      headerHeight: 24,
    })).not.toBeNull();
    expect(rowReads).toBeLessThan(100);
  });
});

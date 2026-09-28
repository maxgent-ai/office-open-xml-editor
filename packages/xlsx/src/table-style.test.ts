import { describe, it, expect } from 'vitest';
import { renderViewport, tableOverlayBorder, type TableCellStyle } from './renderer.js';
import type { CellFont, Dxf, PivotTableMetadata, Styles, Worksheet } from './types.js';

/**
 * Regression tests for spurious table borders on *custom* `<tableStyle>`s
 * (ECMA-376 §18.8.83 / §18.5.1.2).
 *
 * A custom style contributes only the dxfs of its declared
 * `<tableStyleElement>`s. When none of those dxfs define a border, Excel draws
 * no table-level border at all — the only structure lines come from theme
 * borders baked into each cell `xf`. The renderer previously synthesized
 * accent-colored rules in that case (an approximation meant only for built-in
 * `TableStyle{Light,Medium,Dark}N` styles whose definitions are absent from the
 * file), which produced gray lines Excel never draws (sample-6 B23:E23 etc.).
 */

function cell(overrides: Partial<TableCellStyle>): TableCellStyle {
  return {
    accent: '#808080',
    isCustom: false,
    isHeader: false,
    isTotals: false,
    isBanded: false,
    isFirstCol: false,
    isLastCol: false,
    isTopEdge: false,
    isBottomEdge: false,
    ...overrides,
  };
}

const borderDxf: Dxf = {
  font: null,
  fill: null,
  border: { left: null, right: null, top: null, bottom: { style: 'thin', color: '#000000' } },
};

describe('tableOverlayBorder', () => {
  it('custom style with no border dxf draws nothing (no accent synthesis)', () => {
    // sample-6: custom "交通費" header row, dxf has fill only, no border.
    const ts = cell({ isCustom: true, isHeader: true, isTopEdge: true });
    const overlay = tableOverlayBorder(ts, undefined, undefined, 0);
    expect(overlay.kind).toBe('none');
  });

  it('custom style with no border dxf on a data row draws nothing', () => {
    const ts = cell({ isCustom: true, isBanded: true });
    const overlay = tableOverlayBorder(ts, undefined, undefined, 2);
    expect(overlay.kind).toBe('none');
  });

  it('custom style WITH a header-row border dxf draws that border', () => {
    const ts = cell({ isCustom: true, isHeader: true });
    const overlay = tableOverlayBorder(ts, undefined, borderDxf, 0);
    expect(overlay.kind).toBe('dxf');
    if (overlay.kind === 'dxf') {
      expect(overlay.border.bottom?.style).toBe('thin');
    }
  });

  it('built-in style with no border dxf still synthesizes accent rules', () => {
    // Built-in TableStyleLight* etc. are not in the file; the accent
    // approximation must remain until we ship a real preset catalog.
    const ts = cell({ isCustom: false, isHeader: true, isTopEdge: true, accent: '#4472C4' });
    const overlay = tableOverlayBorder(ts, undefined, undefined, 0);
    expect(overlay.kind).toBe('accent');
    if (overlay.kind === 'accent') {
      expect(overlay.color).toBe('#4472C4');
      expect(overlay.topEdge).toBe(true);
    }
  });
});

/**
 * Font color precedence between a custom table style (ECMA-376 §18.8.40
 * tableStyle / §18.8.41 tableStyleElement) and the cell font, as measured in
 * Excel: a cell whose font color is its own formatting (`ownFontColor`, set by
 * the parser) keeps that color, including an automatic one; any other cell
 * takes the table element's font color.
 */
describe('custom table style font color vs the cell font', () => {
  const font = (color: string | null): CellFont => ({
    bold: false, italic: false, underline: false, strike: false, size: 11, color, name: 'Arial',
  });
  const xf = (fontId: number, ownFontColor: boolean) => ({
    fontId, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false,
    ...(ownFontColor ? { ownFontColor } : {}),
  });
  const styles: Styles = {
    fonts: [font('#000000'), font('#FFFFFF'), font(null)],
    fills: [],
    borders: [],
    cellXfs: [xf(0, false), xf(1, true), xf(2, true)],
    numFmts: [],
    dxfs: [{ font: font('#00B050'), fill: null, border: null }],
  };
  const text = (col: number, styleIndex: number, value: string) =>
    ({ row: 1, col, styleIndex, value: { type: 'text' as const, text: value } });
  const ws = {
    name: 'T',
    rows: [{ index: 1, height: null, cells: [text(1, 0, 'inherited'), text(2, 1, 'own'), text(3, 2, 'own auto')] }],
    colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
    tables: [{
      range: { top: 1, left: 1, bottom: 2, right: 3 }, styleName: 'Custom', headerRowCount: 1, totalsRowCount: 0,
      showRowStripes: false, showColumnStripes: false, showFirstColumn: false, showLastColumn: false,
      accentColor: '#808080', isCustom: true, headerRowDxf: 0, columns: [],
    }],
  } as unknown as Worksheet;

  it('keeps a cell-owned color and applies the table color to inherited ones', () => {
    let fillStyle = '';
    const drawn = new Map<string, string>();
    const noop = () => {};
    const ctx = new Proxy({
      canvas: { width: 400, height: 100 },
      font: '11px Arial',
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string) { fillStyle = value; },
      measureText: (t: string) => ({ width: t.length * 7 }) as TextMetrics,
      fillText: (t: string) => { drawn.set(t, fillStyle); },
    } as Record<string | symbol, unknown>, {
      get: (target, key) => (key in target ? target[key] : noop),
      set: (target, key, value) => { target[key] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 1, cols: 3 });
    const color = (t: string) => drawn.get(t)?.replace(/\s/g, '').toLowerCase();
    expect(color('inherited')).toMatch(/^(#00b050|rgba\(0,176,80,1\))$/);
    expect(color('own')).toMatch(/^(#ffffff|rgba\(255,255,255,1\))$/);
    expect(color('own auto')).toMatch(/^(#000000|rgba\(0,0,0,1\))$/);
  });
});

/**
 * Rich-text runs under a custom table style font color, as measured in Excel
 * (body rows under a blue wholeTable dxf):
 * - a run without <rPr> takes the cell's drawn color (the table's, unless the
 *   cell's own font color wins);
 * - a run whose <rPr> color is authored like Normal's (`normalColor`) takes
 *   the table color, but only when the cell's font color is not its own;
 * - a run with its own <rPr> color keeps it, and an <rPr> without <color> is
 *   automatic black, even in a red cell.
 */
describe('rich-text runs under a custom table style font color', () => {
  const font = (color: string | null): CellFont => ({
    bold: false, italic: false, underline: false, strike: false, size: 11, color, name: 'Arial',
  });
  const runFont = (color: string | null, normalColor: boolean) => ({
    bold: false, italic: false, underline: false, strike: false, size: 11,
    ...(color ? { color } : {}), ...(normalColor ? { normalColor } : {}),
  });
  const styles: Styles = {
    fonts: [font('#000000'), font('#FF0000')],
    fills: [],
    borders: [],
    cellXfs: [
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
      { fontId: 1, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false, ownFontColor: true },
    ],
    numFmts: [],
    dxfs: [{ font: font('#0070C0'), fill: null, border: null }],
  };
  // [column, cell xf, runs]
  const cases: Array<[number, number, Array<{ text: string; font?: ReturnType<typeof runFont> }>]> = [
    [1, 0, [{ text: 'd1' }, { text: 'd2', font: runFont('#000000', true) }]],
    [2, 0, [{ text: 'e1', font: runFont('#000000', false) }, { text: 'e2', font: runFont('#000000', true) }]],
    [3, 0, [{ text: 'c1', font: runFont(null, false) }]],
    [4, 1, [{ text: 'g1', font: runFont('#000000', true) }, { text: 'h1', font: runFont(null, false) }, { text: 'k1' }]],
    // A run not confirmed as Normal-colored (e.g. Normal unresolvable) keeps
    // its own color rather than being painted over.
    [5, 0, [{ text: 'u1', font: runFont('#FF0000', false) }]],
  ];
  const ws = {
    name: 'T',
    rows: [{
      index: 2, height: null,
      cells: cases.map(([col, styleIndex, runs]) => ({
        row: 2, col, styleIndex,
        value: { type: 'text' as const, text: runs.map((r) => r.text).join(''), runs },
      })),
    }],
    colWidths: { 1: 20, 2: 20, 3: 20, 4: 20, 5: 20 }, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
    tables: [{
      range: { top: 1, left: 1, bottom: 2, right: 5 }, styleName: 'Custom', headerRowCount: 1, totalsRowCount: 0,
      showRowStripes: false, showColumnStripes: false, showFirstColumn: false, showLastColumn: false,
      accentColor: '#808080', isCustom: true, wholeTableDxf: 0, columns: [],
    }],
  } as unknown as Worksheet;

  it('paints each run with the measured color', () => {
    let fillStyle = '';
    const drawn = new Map<string, string>();
    const noop = () => {};
    const ctx = new Proxy({
      canvas: { width: 800, height: 100 },
      font: '11px Arial',
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string) { fillStyle = value; },
      measureText: (t: string) => ({ width: t.length * 7 }) as TextMetrics,
      fillText: (t: string) => { drawn.set(t, fillStyle); },
    } as Record<string | symbol, unknown>, {
      get: (target, key) => (key in target ? target[key] : noop),
      set: (target, key, value) => { target[key] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 2, cols: 5 });
    const hex = (t: string) => {
      const c = drawn.get(t)?.replace(/\s/g, '').toLowerCase() ?? '';
      const m = /^rgba\((\d+),(\d+),(\d+),1\)$/.exec(c);
      return m ? '#' + m.slice(1, 4).map((n) => Number(n).toString(16).padStart(2, '0')).join('') : c;
    };
    const blue = '#0070c0', black = '#000000', red = '#ff0000';
    expect(Object.fromEntries(['d1', 'd2', 'e1', 'e2', 'c1', 'g1', 'h1', 'k1', 'u1'].map((t) => [t, hex(t)]))).toEqual({
      d1: blue, d2: blue, e1: black, e2: blue, c1: black, g1: black, h1: black, k1: red, u1: red,
    });
  });
});

/**
 * PivotTable style font color vs the cell font, measured in Excel: a value
 * cell given Automatic, RGB 0,0,0 or red keeps that color, while one given
 * the theme "Black, Text 1" (authored like Normal) shows the PivotTable
 * style's color. The parser's `ownFontColor` flag carries that distinction.
 */
describe('PivotTable style font color vs the cell font', () => {
  const font = (color: string | null): CellFont => ({
    bold: false, italic: false, underline: false, strike: false, size: 11, color, name: 'Arial',
  });
  const styles: Styles = {
    fonts: [font('#000000'), font('#FF0000')],
    fills: [],
    borders: [],
    cellXfs: [
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false },
      { fontId: 1, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false, ownFontColor: true },
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false, ownFontColor: true },
    ],
    numFmts: [],
    dxfs: [],
  };
  const green: Dxf = { font: font('#00B050'), fill: null, border: null, fontToggles: {} };
  const pivot = {
    name: 'P', cacheId: 1,
    location: { top: 1, left: 1, bottom: 1, right: 3, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 },
    rowFields: [0], columnFields: [], pageFields: [], dataFields: [],
    status: { state: 'complete' },
    rowItems: [{ kind: 'data', depth: 0 }], columnItems: [],
    style: {
      name: 'S', showRowHeaders: true, showColumnHeaders: true,
      showRowStripes: false, showColumnStripes: false, showLastColumn: false,
      elements: [{ kind: 'wholeTable', size: 1, dxf: green }],
    },
  } as unknown as PivotTableMetadata;
  const text = (col: number, styleIndex: number, value: string) =>
    ({ row: 1, col, styleIndex, value: { type: 'text' as const, text: value } });
  const ws = {
    name: 'P',
    rows: [{ index: 1, height: null, cells: [text(1, 0, 'theme'), text(2, 1, 'red'), text(3, 2, 'auto')] }],
    colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
    pivotTables: [pivot],
  } as unknown as Worksheet;

  it('keeps a cell-owned color and applies the PivotTable color otherwise', () => {
    let fillStyle = '';
    const drawn = new Map<string, string>();
    const noop = () => {};
    const ctx = new Proxy({
      canvas: { width: 400, height: 100 },
      font: '11px Arial',
      get fillStyle() { return fillStyle; },
      set fillStyle(value: string) { fillStyle = value; },
      measureText: (t: string) => ({ width: t.length * 7 }) as TextMetrics,
      fillText: (t: string) => { drawn.set(t, fillStyle); },
    } as Record<string | symbol, unknown>, {
      get: (target, key) => (key in target ? target[key] : noop),
      set: (target, key, value) => { target[key] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 1, cols: 3 });
    const color = (t: string) => drawn.get(t)?.replace(/\s/g, '').toLowerCase();
    expect(color('theme')).toMatch(/^(#00b050|rgba\(0,176,80,1\))$/);
    expect(color('red')).toMatch(/^(#ff0000|rgba\(255,0,0,1\))$/);
    expect(color('auto')).toMatch(/^(#000000|rgba\(0,0,0,1\))$/);
  });
});

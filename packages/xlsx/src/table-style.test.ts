import { describe, it, expect } from 'vitest';
import { renderViewport } from './renderer.js';
import { buildStructuredStyleMap } from './pivot-style.js';
import type { CellFont, Dxf, PivotTableMetadata, PivotTableStyleElement, Styles, TableInfo, Worksheet } from './types.js';

/**
 * Table style layering (ECMA-376 §18.8.41 Table Style Element Order over the
 * §18.18.77 Table regions; [MS-XLS] 2.5.176 gating). Built-in and custom
 * styles arrive alike as inline elements (`styleElements`).
 */
describe('table style element layering', () => {
  const fill = (color: string): Dxf => ({
    font: null, border: null,
    fill: { patternType: 'solid', fgColor: color, bgColor: color },
  } as unknown as Dxf);
  const edge = { style: 'thin', color: '#000000' };
  const element = (kind: string, dxf: Dxf, size = 1): PivotTableStyleElement => ({ kind, size, dxf });
  const table = (over: Partial<TableInfo>, styleElements: PivotTableStyleElement[]): Worksheet => ({
    tables: [{
      range: { top: 1, left: 1, bottom: 6, right: 4 }, styleName: 'S', headerRowCount: 1, totalsRowCount: 1,
      showRowStripes: true, showColumnStripes: true, showFirstColumn: true, showLastColumn: true,
      accentColor: '#808080', columns: [], styleElements, ...over,
    }],
  } as unknown as Worksheet);
  const fg = (map: ReturnType<typeof buildStructuredStyleMap>, r: number, c: number) =>
    (map.get(`${r}:${c}`)?.fill as { fgColor?: string } | undefined)?.fgColor;

  it('applies the elements in the Table style element order', () => {
    const map = buildStructuredStyleMap(table({}, [
      // Listed in reverse: the order of application is the spec's, not the file's.
      element('lastTotalCell', fill('#00000C')),
      element('firstHeaderCell', fill('#000009')),
      element('totalRow', fill('#000008')),
      element('headerRow', fill('#000007')),
      element('firstColumn', fill('#000006')),
      element('lastColumn', fill('#000005')),
      element('firstRowStripe', fill('#000003')),
      element('firstColumnStripe', fill('#000001')),
      element('wholeTable', fill('#000000')),
    ]));
    expect(fg(map, 1, 1)).toBe('#000009'); // first header cell
    expect(fg(map, 1, 2)).toBe('#000007'); // header row
    expect(fg(map, 6, 4)).toBe('#00000C'); // last total cell
    expect(fg(map, 6, 2)).toBe('#000008'); // total row
    expect(fg(map, 2, 1)).toBe('#000006'); // first column over the stripes
    expect(fg(map, 2, 4)).toBe('#000005'); // last column
    expect(fg(map, 2, 3)).toBe('#000003'); // row stripe over column stripe
    expect(fg(map, 3, 3)).toBe('#000001'); // second row band: column stripe
    expect(fg(map, 3, 2)).toBe('#000000'); // second band of both: whole table
  });

  it('gates first/last column and the stripes by tableStyleInfo, not the end cells', () => {
    const map = buildStructuredStyleMap(table(
      { showRowStripes: false, showColumnStripes: false, showFirstColumn: false, showLastColumn: false },
      [
        element('wholeTable', fill('#000000')),
        element('firstRowStripe', fill('#000003')),
        element('firstColumnStripe', fill('#000001')),
        element('firstColumn', fill('#000006')),
        element('firstHeaderCell', fill('#000009')),
      ],
    ));
    expect(fg(map, 2, 1)).toBe('#000000');
    expect(fg(map, 1, 1)).toBe('#000009');
  });

  it('honours stripe band sizes over the data body only', () => {
    const map = buildStructuredStyleMap(table({ showColumnStripes: false, totalsRowCount: 0 }, [
      element('firstRowStripe', fill('#000003'), 2),
      element('secondRowStripe', fill('#000004')),
    ]));
    expect([1, 2, 3, 4, 5, 6].map((r) => fg(map, r, 2))).toEqual(
      [undefined, '#000003', '#000003', '#000004', '#000003', '#000003'],
    );
  });

  it('bands the visible rows only, around a hidden data row', () => {
    const ws = table({ showColumnStripes: false, totalsRowCount: 0 }, [
      element('firstRowStripe', fill('#000003')),
    ]) as Worksheet & { rows: unknown[] };
    (ws as unknown as { rows: unknown[] }).rows = [{ index: 2, hidden: true, cells: [] }];
    const map = buildStructuredStyleMap(ws);
    expect([2, 3, 4, 5, 6].map((r) => fg(map, r, 2))).toEqual(
      [undefined, '#000003', undefined, '#000003', undefined],
    );
  });

  it('draws region border edges on the outline and horizontal/vertical inside', () => {
    const border = {
      font: null, fill: null,
      border: { left: edge, right: edge, top: edge, bottom: edge, horizontal: { style: 'dashed', color: '#000000' } },
    } as unknown as Dxf;
    const map = buildStructuredStyleMap(table({ showRowStripes: false, showColumnStripes: false }, [
      element('wholeTable', border),
    ]));
    expect(map.get('1:2')?.top?.style).toBe('thin');
    expect(map.get('1:2')?.bottom?.style).toBe('dashed');
    expect(map.get('6:2')?.bottom?.style).toBe('thin');
    expect(map.get('3:1')?.left?.style).toBe('thin');
    expect(map.get('3:2')?.left).toBeUndefined();
  });

  it('lets an explicit automatic color clear an earlier element color', () => {
    const blue = { font: { color: '#0000FF' }, fill: null, border: null } as unknown as Dxf;
    const auto = { font: { color: null }, fontToggles: { automaticColor: true }, fill: null, border: null } as unknown as Dxf;
    const map = buildStructuredStyleMap(table({}, [element('wholeTable', blue), element('headerRow', auto)]));
    expect(map.get('1:2')?.fontColor).toBeUndefined();
    expect(map.get('1:2')?.fontColorAuto).toBe(true);
    expect(map.get('2:2')?.fontColor).toBe('#0000FF');
  });

  it('paints a style element solid fill with its bgColor', () => {
    const solid = {
      font: null, border: null, fill: { patternType: 'solid', fgColor: '#D9D9D9', bgColor: '#F2F2F2' },
    } as unknown as Dxf;
    const map = buildStructuredStyleMap(table({}, [element('wholeTable', solid)]));
    expect(fg(map, 2, 2)).toBe('#F2F2F2');
  });

  it('resolves legacy dxf indices for a model without styleElements', () => {
    const ws = table({ headerRowDxf: 0 } as Partial<TableInfo>, []);
    delete (ws.tables![0] as { styleElements?: unknown }).styleElements;
    const map = buildStructuredStyleMap(ws, [fill('#000007')]);
    expect(fg(map, 1, 2)).toBe('#000007');
    expect(buildStructuredStyleMap(ws).size).toBe(0);
  });

  it('applies nothing for a table without style elements', () => {
    expect(buildStructuredStyleMap(table({}, [])).size).toBe(0);
  });
});

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
      accentColor: '#808080', isCustom: true, columns: [],
      styleElements: [{ kind: 'headerRow', size: 1, dxf: { font: font('#00B050'), fill: null, border: null } }],
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

describe('an automatic table-element color over an inherited font color', () => {
  it('draws automatic black, not the inherited red', () => {
    const font = (color: string | null): CellFont => ({
      bold: false, italic: false, underline: false, strike: false, size: 11, color, name: 'Arial',
    });
    const styles: Styles = {
      fonts: [font('#FF0000')], fills: [], borders: [],
      cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false }],
      numFmts: [], dxfs: [],
    };
    const ws = {
      name: 'T',
      rows: [{ index: 1, height: null, cells: [{ row: 1, col: 1, styleIndex: 0, value: { type: 'text' as const, text: 'head' } }] }],
      colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
      mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
      tables: [{
        range: { top: 1, left: 1, bottom: 2, right: 1 }, styleName: 'Custom', headerRowCount: 1, totalsRowCount: 0,
        showRowStripes: false, showColumnStripes: false, showFirstColumn: false, showLastColumn: false,
        accentColor: '#808080', columns: [],
        styleElements: [
          { kind: 'wholeTable', size: 1, dxf: { font: font('#0000FF'), fill: null, border: null } },
          { kind: 'headerRow', size: 1, dxf: { font: font(null), fontToggles: { automaticColor: true }, fill: null, border: null } },
        ],
      }],
    } as unknown as Worksheet;
    let fillStyle = '';
    const drawn = new Map<string, string>();
    const noop = () => {};
    const ctx = new Proxy({
      canvas: { width: 400, height: 100 }, font: '11px Arial',
      get fillStyle() { return fillStyle; }, set fillStyle(value: string) { fillStyle = value; },
      measureText: (t: string) => ({ width: t.length * 7 }) as TextMetrics,
      fillText: (t: string) => { drawn.set(t, fillStyle); },
    } as Record<string | symbol, unknown>, {
      get: (target, key) => (key in target ? target[key] : noop),
      set: (target, key, value) => { target[key] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 1, cols: 1 });
    expect(drawn.get('head')?.replace(/\s/g, '').toLowerCase()).toMatch(/^(#000000|rgba\(0,0,0,1\))$/);
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
      accentColor: '#808080', isCustom: true, columns: [],
      styleElements: [{ kind: 'wholeTable', size: 1, dxf: { font: font('#0070C0'), fill: null, border: null } }],
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

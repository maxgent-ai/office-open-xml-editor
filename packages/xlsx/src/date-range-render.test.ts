import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

function paint(width: number, merged = false) {
  const texts: string[] = [];
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 600, height: 300 },
    measureText: (text: string) => ({ width: text.length * 7 }),
    fillText: (text: string) => texts.push(text),
  } as Record<string | symbol, unknown>, {
    get: (target, key) => key in target ? target[key] : noop,
    set: (target, key, value) => { target[key] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const styles: Styles = {
    fonts: [{ name: 'Arial', size: 11, bold: false, italic: false, underline: false, strike: false, color: null }],
    fills: [], borders: [], dxfs: [],
    cellXfs: [
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 14, wrapText: true, textRotation: 45 },
      { fontId: 0, fillId: 0, borderId: 0, numFmtId: 164 },
    ] as Styles['cellXfs'],
    numFmts: [{ numFmtId: 164, formatCode: '"#"' }],
  };
  const ws = {
    name: 'Dates',
    rows: [{ index: 1, height: null, cells: [
      { row: 1, col: 1, styleIndex: 0, value: { type: 'number', number: 12345678901 } },
      { row: 1, col: 3, styleIndex: 1, value: { type: 'number', number: 12345678901 } },
    ] }],
    colWidths: { 1: width, 2: width, 3: width }, rowHeights: {},
    defaultColWidth: width, defaultRowHeight: 15,
    mergeCells: merged ? [{ top: 1, left: 1, bottom: 1, right: 2 }] : [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
  } as Worksheet;
  renderViewport(ctx, ws, styles, { row: 1, col: merged ? 2 : 1, rows: 1, cols: merged ? 2 : 3 });
  return texts.filter(text => /^#+$/.test(text));
}

describe('invalid date cell painting', () => {
  it('fills available width independently of wrapping/rotation and preserves a literal hash', () => {
    const narrow = paint(10);
    const wide = paint(20);
    expect(narrow).toHaveLength(2);
    expect(wide).toHaveLength(2);
    expect(narrow[0].length).toBeGreaterThan(3);
    expect(wide[0].length).toBeGreaterThan(narrow[0].length);
    expect(wide[1]).toBe('#');
  });
  it('fills a merged cell whose invalid-date anchor is outside the viewport', () => {
    const hashes = paint(20, true);
    expect(hashes).toHaveLength(2);
    expect(hashes[0].length).toBeGreaterThan(3);
    expect(hashes[1]).toBe('#');
  });
});

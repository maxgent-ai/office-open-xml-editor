import { afterEach, describe, expect, it, vi } from 'vitest';
import { bindXlsxOfficeFontRoutes, drawResolvedRichLine, renderViewport } from './renderer.js';
import type { CellFont, CellValue, CellXf, Styles, Worksheet } from './types.js';

afterEach(() => vi.unstubAllGlobals());

const base: CellFont = { name: 'Calibri', scheme: 'minor', size: 11, bold: false,
  italic: false, underline: false, strike: false, color: '#123456' };

type Paint = { text: string; font: string; x: number; y: number;
  color: unknown; lineWidth: number; dash: number[] };

function renderCells(
  cells: { font: CellFont; value?: CellValue; xf?: Partial<CellXf>; showPhonetic?: boolean }[],
  theme: Partial<Worksheet> = {},
) {
  const fonts = cells.map((cell) => cell.font);
  const styles: Styles = { fonts, fills: [], borders: [], numFmts: [], dxfs: [],
    cellXfs: cells.map((cell, fontId) => ({ fontId, fillId: 0, borderId: 0, numFmtId: 0,
      alignH: null, alignV: null, wrapText: false, ...cell.xf })) };
  const worksheet = { name: 'Fonts', defaultColWidth: 20, defaultRowHeight: 30,
    colWidths: {}, rowHeights: {}, mergeCells: [], freezeRows: 0, freezeCols: 0,
    themeJapaneseMajorFont: 'Yu Gothic Light', themeJapaneseMinorFont: 'Yu Mincho',
    conditionalFormats: [], images: [], charts: [], ...theme,
    rows: cells.map((cell, index) => ({ index: index + 1, height: null,
      cells: [{ row: index + 1, col: 1, styleIndex: index, showPhonetic: cell.showPhonetic,
        value: cell.value ?? { type: 'text', text: `cell${index}` } }] })) } as Worksheet;
  const fills: Paint[] = [], strokes: Paint[] = [];
  const state: Record<string, unknown> = { canvas: { width: 400, height: 800 },
    font: '', lineWidth: 3, fillStyle: '#000', strokeStyle: '#abc', dash: [2, 3] };
  const stack: Record<string, unknown>[] = [];
  const record = (target: Paint[], text: string, x: number, y: number, color: unknown) => {
    target.push({ text, x, y, font: state.font as string, color,
      lineWidth: state.lineWidth as number, dash: [...state.dash as number[]] });
  };
  const ctx = new Proxy({ ...state,
    save: () => stack.push({ ...state }),
    restore: () => Object.assign(state, stack.pop()),
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    setLineDash: (dash: number[]) => { state.dash = dash; },
    fillText: (text: string, x: number, y: number) => record(fills, text, x, y, state.fillStyle),
    strokeText: (text: string, x: number, y: number) => record(strokes, text, x, y, state.strokeStyle),
  }, {
    get: (target, key) => key in state ? state[key as string]
      : key in target ? target[key as keyof typeof target] : () => {},
    set: (_, key, value) => { state[key as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, worksheet, styles, { row: 1, col: 1, rows: cells.length, cols: 1 });
  return { fills, strokes, ctx, worksheet };
}

describe('theme cell synthetic bold paint', () => {
  it('strokes bold Light and Regular faces, leaving real Bold, Demibold and regular cells alone', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    const { fills, strokes } = renderCells([
      { font: { ...base, scheme: 'major', bold: true } },
      { font: { ...base, bold: true, size: 22 } },
      { font: base },
    ]);
    expect(strokes.map((paint) => paint.text)).toEqual(['cell0', 'cell1']);
    expect(strokes[0].font).toMatch(/^300 15px "Yu Gothic Light"/);
    expect(strokes[1].font).toMatch(/^29px "Yu Mincho"/);
    // Excel's PDF uses 0.31543 pt at 11.04 pt: 1/35 em. The stroke follows
    // the rounded Canvas em, and paint keeps the selected face and color.
    expect(strokes[0].lineWidth).toBeCloseTo(15 / 35);
    expect(strokes[1].lineWidth).toBeCloseTo(29 / 35);
    for (const stroke of strokes) {
      const fill = fills.find((paint) => paint.text === stroke.text);
      expect(stroke).toMatchObject({ font: fill?.font, x: fill?.x, y: fill?.y, color: fill?.color, dash: [] });
    }
    const real = renderCells([
      { font: { ...base, bold: true } },
      { font: { ...base, scheme: 'major', bold: true } },
      { font: { ...base, scheme: undefined, bold: true } },
    ], { themeJapaneseMinorFont: 'Yu Gothic', themeJapaneseMajorFont: 'Yu Mincho Demibold' });
    expect(real.strokes).toEqual([]);
    expect(real.fills.find((paint) => paint.text === 'cell0')?.font).toMatch(/^700 /);
    expect(real.fills.find((paint) => paint.text === 'cell1')?.font).toMatch(/^600 /);
  });

  it('uses the cell painter for rich runs, furigana, wrapped, rotated and stacked text', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    const bold = { ...base, bold: true };
    const { fills, strokes } = renderCells([
      { font: bold, value: { type: 'text', text: 'rich', runs: [{ text: 'rich' }] } },
      { font: bold, value: { type: 'text', text: 'base',
        phoneticRuns: [{ sb: 0, eb: 4, text: 'hint' }], phoneticPr: { fontId: 0 } }, showPhonetic: true },
      { font: bold, xf: { wrapText: true }, value: { type: 'text', text: 'wrap\nline' } },
      { font: bold, xf: { textRotation: 45 }, value: { type: 'text', text: 'rotate' } },
      { font: bold, xf: { textRotation: 255 }, value: { type: 'text', text: 'Aー' } },
    ]);
    expect(strokes.map((paint) => paint.text)).toEqual(['rich', 'base', 'hint', 'wrap', 'line', 'rotate', 'A', 'ー']);
    for (const stroke of strokes) {
      expect(stroke.lineWidth).toBeCloseTo(15 / 35);
      const fill = fills.find((paint) => paint.text === stroke.text);
      expect(stroke).toMatchObject({ font: fill?.font, x: fill?.x, y: fill?.y, color: fill?.color });
    }
  });

  it('replaces compact context font state when rebinding to another sheet', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    const { ctx, worksheet, fills, strokes } = renderCells([{ font: { ...base, bold: true } }]);
    // A binding snapshots only theme names and language. Later worksheet
    // mutation cannot keep its cell graph involved in context font resolution.
    worksheet.themeJapaneseMinorFont = 'Yu Gothic';
    const paint = () => drawResolvedRichLine(ctx,
      [{ text: 'bound', font: { ...base, bold: true }, width: 35 }],
      0, 0, 'top', 1, 1, {});
    ctx.lineWidth = 3;
    ctx.strokeStyle = '#abc';
    paint();
    expect(strokes.at(-1)?.text).toBe('bound');
    expect(fills.at(-1)?.font).toMatch(/^15px "Yu Mincho"/);
    expect(ctx.lineWidth).toBe(3);
    expect(ctx.strokeStyle).toBe('#abc');
    bindXlsxOfficeFontRoutes(ctx, worksheet);
    const count = strokes.length;
    paint();
    expect(strokes).toHaveLength(count);
    expect(fills.at(-1)?.font).toMatch(/^700 15px "Yu Gothic"/);
  });
});

import { afterEach, describe, expect, it, vi } from 'vitest';
import { renderViewport, resolvedThemeCellFont } from './renderer.js';
import type { CellFont, Styles, Worksheet } from './types.js';

afterEach(() => vi.unstubAllGlobals());

function paintedFonts(platform: string, language = 'ja-JP', withJpan = true, minorFace = 'メイリオ'): Map<string, string> {
  vi.stubGlobal('navigator', { platform, language, maxTouchPoints: 0 });
  const base: CellFont = { name: null, size: 11, bold: false, italic: false,
    underline: false, strike: false, color: null };
  const fonts: CellFont[] = [
    { ...base, name: 'Calibri', scheme: 'minor', bold: true, italic: false },
    { ...base, name: 'Arial', scheme: 'major', bold: false, italic: true },
    { ...base, name: '游ゴシック', bold: false, italic: false },
    { ...base, name: null, scheme: 'minor', bold: false, italic: false },
  ];
  const styles: Styles = {
    fonts, fills: [], borders: [], numFmts: [], dxfs: [],
    cellXfs: fonts.map((_, fontId) => ({ fontId, fillId: 0, borderId: 0, numFmtId: 0,
      alignH: null, alignV: null, wrapText: false })),
  };
  const ws = {
    name: 'Authored fonts', defaultColWidth: 20, defaultRowHeight: 15,
    colWidths: {}, rowHeights: {}, mergeCells: [], freezeRows: 0, freezeCols: 0,
    themeJapaneseMajorFont: withJpan ? '游ゴシック Light' : undefined,
    themeJapaneseMinorFont: withJpan ? minorFace : undefined,
    conditionalFormats: [], images: [], charts: [],
    rows: fonts.map((_, styleIndex) => ({ index: styleIndex + 1, height: null,
      cells: [{ row: styleIndex + 1, col: 1, styleIndex,
        value: { type: 'text', text: `cell${styleIndex}` } }] })),
  } as Worksheet;
  let font = '';
  const drawn = new Map<string, string>();
  const noop = () => {};
  const ctx = new Proxy({
    canvas: { width: 400, height: 200 },
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: () => ({ width: 35 }),
    fillText: (text: string) => drawn.set(text, font),
  } as Record<string, unknown>, {
    get: (target, key) => key in target ? target[key as string] : noop,
    set: (target, key, value) => { target[key as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, ws, styles, { row: 1, col: 1, rows: 4, cols: 1 });
  return drawn;
}

describe('Japanese Excel theme cell fonts independent of viewer OS', () => {
  it('paints the same Jpan families and bold/italic styles on Japanese Mac and Windows', () => {
    const windows = paintedFonts('Win32');
    const mac = paintedFonts('MacIntel');
    expect(mac).toEqual(windows);
    expect(mac.get('cell0')).toMatch(/^700 15px "メイリオ", "Meiryo", /);
    expect(mac.get('cell1')).toMatch(/^italic 300 15px "游ゴシック Light", "YuGothic", /);
    expect(mac.get('cell2')).toMatch(/^15px "游ゴシック", /);
    expect(mac.get('cell3')).toMatch(/^15px "メイリオ", "Meiryo", /);
  });

  it('keeps authored families on English Mac and Windows', () => {
    const windows = paintedFonts('Win32', 'en-US');
    expect(paintedFonts('MacIntel', 'en-US')).toEqual(windows);
    expect(windows.get('cell0')).toMatch(/^700 15px "Calibri", /);
    expect(windows.get('cell1')).toMatch(/^italic 15px "Arial", /);
    expect(windows.get('cell3')).toMatch(/^15px "Calibri", /);
  });

  it('keeps authored families when the theme has no Jpan entry', () => {
    const fonts = paintedFonts('Win32', 'ja-JP', false);
    expect(fonts.get('cell0')).toMatch(/^700 15px "Calibri", /);
    expect(fonts.get('cell1')).toMatch(/^italic 15px "Arial", /);
  });

  it('paints bold minor Mincho with the Regular face weight rather than selecting Demibold', () => {
    const fonts = paintedFonts('Win32', 'ja-JP', true, 'Yu Mincho');
    expect(fonts.get('cell0')).toMatch(/^15px "Yu Mincho", /);
  });
});

describe('theme cell font face metadata', () => {
  const theme = {
    themeJapaneseMajorFont: '游ゴシック Light',
    themeJapaneseMinorFont: '游ゴシック',
  };
  const font = { name: 'Calibri', scheme: 'minor' as const, bold: true, italic: false };

  it('resolves different authored names through the same minor face and requested style', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    expect(resolvedThemeCellFont(font, theme))
      .toEqual({ family: '游ゴシック', weight: 700, fallbackAlias: 'YuGothic' });
    expect(resolvedThemeCellFont({ ...font, name: 'Arial', bold: false }, theme))
      .toEqual({ family: '游ゴシック', weight: 400, fallbackAlias: 'YuGothic' });
    expect(resolvedThemeCellFont({ ...font, italic: true }, theme))
      .toEqual({ family: '游ゴシック', weight: 700, fallbackAlias: 'YuGothic' });
    expect(resolvedThemeCellFont({ ...font, bold: false, italic: true }, theme))
      .toEqual({ family: '游ゴシック', weight: 400, fallbackAlias: 'YuGothic' });
  });

  it.each([
    { scheme: 'minor', face: 'Yu Gothic', regular: 400, bold: 700, alias: 'YuGothic' },
    { scheme: 'major', face: 'Yu Gothic Light', regular: 300, bold: 300, alias: 'YuGothic' },
    { scheme: 'minor', face: 'Yu Mincho', regular: 400, bold: 400, alias: 'YuMincho' },
    { scheme: 'major', face: 'Yu Mincho Demibold', regular: 600, bold: 600, alias: 'Yu Mincho' },
  ] as const)('selects catalog weights for $scheme $face in regular and bold cells', ({ scheme, face, regular, bold, alias }) => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    const worksheet = { themeJapaneseMajorFont: face, themeJapaneseMinorFont: face };
    for (const isBold of [false, true]) {
      expect(resolvedThemeCellFont({ ...font, scheme, bold: isBold }, worksheet))
        .toEqual({ family: face, weight: isBold ? bold : regular,
          ...(alias ? { fallbackAlias: alias } : {}),
          ...(isBold && (bold === 300 || bold === 400) ? { syntheticBold: true } : {}) });
    }
  });

  it('uses uncatalogued theme faces without restricting authored names', () => {
    vi.stubGlobal('navigator', { language: 'ja' });
    expect(resolvedThemeCellFont({ ...font, name: 'Another authored face' }, {
      themeJapaneseMinorFont: 'Custom theme face',
    })).toEqual({ family: 'Custom theme face', weight: 700 });
    expect(resolvedThemeCellFont({ ...font, scheme: undefined }, theme))
      .toEqual({ family: 'Calibri', weight: 700 });
  });

  it('keeps the authored font when no viewer language is available', () => {
    vi.stubGlobal('navigator', undefined);
    expect(resolvedThemeCellFont(font, { themeJapaneseMinorFont: '游ゴシック' }))
      .toEqual({ family: 'Calibri', weight: 700 });
  });
});

import { rasterizeMathSvg } from '@silurus/ooxml-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { drawShapeText, renderViewport, prepareWorksheetMath } from './renderer.js';
import type { ShapeText, Styles, Worksheet } from './types.js';

vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  rasterizeMathSvg: vi.fn(async () => ({ source: {} })),
  tintMathRaster: () => ({}),
}));

afterEach(() => vi.restoreAllMocks());

function record() {
  const calls: { text: string; x: number; y: number; angle: number }[] = [];
  const clips: number[][] = [];
  const images: { x: number; y: number; w: number; h: number; angle: number }[] = [];
  let font = '24px Arial';
  let angle = 0;
  let sx = 1, sy = 1;
  const frames: [number, number, number][] = [];
  const state = {
    get font() { return font; }, set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: text.length * 10,
      fontBoundingBoxAscent: 24, fontBoundingBoxDescent: 6 }),
    fillText: (text: string, x: number, y: number) => calls.push({ text, x, y, angle: angle + (sx < 0 ? Math.PI : 0) }),
    drawImage(_image: unknown, x: number, y: number, w: number, h: number) { images.push({ x, y, w, h, angle }); },
    save() { frames.push([angle, sx, sy]); },
    restore() { [angle, sx, sy] = frames.pop()!; },
    rotate(value: number) { angle += sx * sy * value; },
    scale(x: number, y: number) { sx *= x; sy *= y; },
    translate() {}, beginPath() {}, clip() {},
    canvas: { width: 800, height: 400 },
    rect: (...args: number[]) => clips.push(args),
    textAlign: 'left', textBaseline: 'alphabetic', fillStyle: '#000',
  };
  const ctx = new Proxy(state, { get: (target, key) => target[key as keyof typeof target] ?? (() => {}) }) as unknown as CanvasRenderingContext2D;
  return { ctx, calls, clips, images };
}

const body = (vert: string, extra: Partial<ShapeText> = {}): ShapeText => ({
  vert, anchor: 't', wrap: 'none', lIns: 0, rIns: 0, tIns: 0, bIns: 0,
  paragraphs: [{ align: 'l', runs: [{ type: 'text', text: 'AB', size: 18,
    bold: false, italic: false, fontFace: 'Arial' }, { type: 'break' },
    { type: 'text', text: 'CD', size: 18, bold: false, italic: false, fontFace: 'Arial' }] }],
  ...extra,
});

describe('Excel DrawingML stacked WordArt', () => {
  it.each(['wordArtVert', 'wordArtVertRtl'])(
    '%s paints upright letters down each column, with the specified column order', (vert) => {
      const { ctx, calls } = record();
      drawShapeText(ctx, body(vert), 200, 200, 1);
      expect(calls.map((c) => c.text)).toEqual(['A', 'B', 'C', 'D']);
      expect(calls.every((c) => c.angle === 0)).toBe(true);
      expect(calls[1].x).toBe(calls[0].x);
      expect(calls[1].y).toBeGreaterThan(calls[0].y);
      expect(calls[2].y).toBe(calls[0].y);
      expect(Math.sign(calls[2].x - calls[0].x)).toBe(vert === 'wordArtVert' ? 1 : -1);
    },
  );
  it('applies the authored normAutofit fontScale before stacked measurement', () => {
    const normal = record(), scaled = record();
    drawShapeText(normal.ctx, body('wordArtVert'), 200, 200, 1);
    drawShapeText(scaled.ctx, body('wordArtVert', { autoFit: 'norm', fontScale: 0.5 }), 200, 200, 1);
    expect(scaled.calls.map((c) => c.text)).toEqual(normal.calls.map((c) => c.text));
    expect(scaled.calls[1].y - scaled.calls[0].y).toBeCloseTo(
      (normal.calls[1].y - normal.calls[0].y) / 2,
    );
  });
  it('keeps physical asymmetric insets and lets the worksheet clip overflow', () => {
    const { ctx, calls, clips } = record();
    drawShapeText(ctx, body('wordArtVert', { lIns: 95250, tIns: 190500 }), 100, 30, 1);
    expect(calls[0].x).toBeGreaterThan(10);
    expect(calls[0].y).toBeGreaterThan(20);
    expect(calls[1].y).toBeGreaterThan(30);
    expect(clips).toEqual([]);
  });
  it.each([[30, false, false, 30], [90, false, false, 90],
    [0, true, false, 0], [0, false, true, 180]])(
    'worksheet WordArt rotation=%s flipH=%s flipV=%s stays readable', (rot, flipH, flipV, expected) => {
      const { ctx, calls } = record();
      const worksheet = { name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
        defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [], freezeRows: 0, freezeCols: 0,
        conditionalFormats: [], images: [], charts: [], defaultFontFamily: 'Arial', defaultFontSize: 11,
        shapeGroups: [{ fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
          toCol: 4, toColOff: 0, toRow: 8, toRowOff: 0, nativeExtCx: 0, nativeExtCy: 0,
          shapes: [{ x: 0, y: 0, w: 1, h: 1, rot, flipH, flipV, strokeWidth: 0,
            geom: { type: 'preset', name: 'rect' }, text: body('wordArtVert') }] }] } as unknown as Worksheet;
      const styles = { fonts: [{ size: 11, name: 'Arial', bold: false, italic: false }],
        fills: [], borders: [], cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }],
        numFmts: [], dxfs: [] } as unknown as Styles;
      renderViewport(ctx, worksheet, styles, { row: 1, col: 1, rows: 20, cols: 10 });
      const glyph = calls.find((c) => c.text === 'A')!;
      expect(glyph).toBeDefined();
      expect(Math.cos(glyph.angle)).toBeCloseTo(Math.cos(expected * Math.PI / 180));
      expect(Math.sin(glyph.angle)).toBeCloseTo(Math.sin(expected * Math.PI / 180));
    },
  );

});

// Compare cached and unavailable OMML geometry with the horizontal host rule.
it.each(['cached', 'no engine', 'conversion', 'rasterization'] as const)('%s preserves horizontal equation geometry in both stacked modes', async (failure) => {
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
  const error = vi.spyOn(console, 'error').mockImplementation(() => {});
  if (failure === 'rasterization') vi.mocked(rasterizeMathSvg).mockRejectedValue(new Error('rasterization failed'));
  try {
    for (const display of [false, true]) {
      const outcomes = [];
      let horizontalImage: { w: number; h: number } | undefined;
      for (const vert of ['horz', 'wordArtVert', 'wordArtVertRtl']) {
        warn.mockClear(); error.mockClear();
        const equation = { type: 'math' as const, nodes: [{ kind: 'run' as const, text: 'x', style: 'italic' as const }], display, fontSize: 18 };
        const engine = {
          loadMathJax: async () => {},
          mathMLToSvg: async () => {
            if (failure === 'conversion') throw new Error('conversion failed');
            return { svg: '<svg/>', widthEm: 3, ascentEm: 1.5, descentEm: .5 };
          },
        };
        const b = body(vert);
        const text = b.paragraphs[0].runs[0] as Extract<import('./types.js').ShapeTextRun, { type: 'text' }>;
        b.paragraphs[0].runs = [{ ...text, text: 'A' }, equation, { ...text, text: 'B' }];
        if (failure !== 'no engine') {
          await prepareWorksheetMath({ shapeGroups: [{ shapes: [{ text: b }] }] } as unknown as Worksheet, engine);
        }
        const { ctx, calls, images } = record();
        drawShapeText(ctx, b, 300, 300, 1);
        if (failure === 'cached') {
          expect(images).toHaveLength(1);
          const image = images[0];
          expect(image).toMatchObject({ w: 72, h: 48, angle: 0 });
          if (vert === 'horz') {
            horizontalImage = image;
            if (!display) {
              expect(image.x - calls[0].x).toBeCloseTo(10);
              expect(calls[1].x - image.x).toBeCloseTo(image.w);
              expect(calls[1].y).toBeCloseTo(calls[0].y);
            } else {
              expect(image.y).toBeGreaterThan(calls[0].y);
              expect(calls[1].y).toBeGreaterThan(image.y + image.h);
            }
          } else {
            expect(image.w).toBe(horizontalImage!.w);
            expect(image.h).toBe(horizontalImage!.h);
            const cell = (7 / 6) * (2288 / 2048) * 24;
            const sign = vert === 'wordArtVert' ? 1 : -1;
            expect(calls[1].x - calls[0].x).toBeCloseTo(display ? sign * (cell + image.w) : 0);
            expect(calls[1].y - calls[0].y).toBeCloseTo(display ? 0 : cell + image.h);
            expect(image.x + image.w / 2 - calls[0].x).toBeCloseTo(display ? sign * (cell + image.w) / 2 : 0);
            expect(image.y).toBeCloseTo(display ? 0 : cell);
          }
        } else {
          // Horizontal omits unavailable math, including its display boundary.
          const control = record();
          drawShapeText(control.ctx, { ...b, paragraphs: [{ ...b.paragraphs[0], runs: b.paragraphs[0].runs.filter((r) => r.type !== 'math') }] }, 300, 300, 1);
          expect(calls).toEqual(control.calls);
          if (vert === 'horz') {
            expect(calls[0].x).toBeCloseTo(0);
            expect(calls[0].y).toBeCloseTo(control.calls[0].y);
          } else {
            expect(calls[1].x).toBeCloseTo(calls[0].x);
            expect(calls[1].y - calls[0].y).toBeCloseTo((7 / 6) * (2288 / 2048) * 24);
          }
        }
        const outcome = { text: calls.map((c) => c.text).join(''), images: images.length,
          warnings: [...warn.mock.calls], errors: [...error.mock.calls] };
        expect(outcome).toEqual({ text: 'AB', images: failure === 'cached' ? 1 : 0, warnings: [], errors: [] });
        outcomes.push(outcome);
      }
      expect(outcomes.slice(1)).toEqual([outcomes[0], outcomes[0]]);
    }
  } finally {
    vi.mocked(rasterizeMathSvg).mockResolvedValue({ source: {} as CanvasImageSource, widthPx: 3, heightPx: 2 });
  }
});

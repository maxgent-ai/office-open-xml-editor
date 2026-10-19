import { describe, expect, it } from 'vitest';
import { loadSkiaForTests } from './test-imports';
import { resolvePathShade, type ShadeWork } from '../../core/src/shape/path-gradient';
import { playWmf } from '../../core/src/image/wmf';
import type { GradientFill } from '../../core/src/types/common';

const skia = await loadSkiaForTests();
const { Canvas } = (skia ?? {}) as typeof import('skia-canvas');

const words = (...values: number[]) => values.flatMap(v => [v & 255, (v >> 8) & 255]);
const record = (fn: number, data: number[] = []) => [...words(3 + data.length / 2, 0, fn), ...data];
const wmf = (...records: number[][]) => new Uint8Array([
  ...words(1, 9, 0x300, 0, 0, 8, 0, 0, 0),
  ...records.flat(), ...record(0),
]);
const windowExt = record(0x020c, words(100, 100));
const pen = record(0x02fa, words(0, 2, 0, 0, 0));
const brush = record(0x02fc, words(0, 0xff, 0, 0));
const select = (index: number) => record(0x012d, words(index));
const rect = (l: number, t: number, r: number, b: number, fn = 0x041b) => record(fn, words(b, r, t, l));
const pixel = (ctx: CanvasRenderingContext2D, x: number, y: number) => Array.from(ctx.getImageData(x, y, 1, 1).data);
function render(...records: number[][]) {
  const canvas = new Canvas(100, 100);
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  playWmf(wmf(windowExt, ...records), ctx, 100, 100);
  return ctx;
}

describe.skipIf(!skia)('shared paint regressions', () => {
  it.each(['shape', 'rect'] as const)('shades more than 32768 real edges without replacing path=%s', path => {
    const ctx = new Canvas(96, 96).getContext('2d') as unknown as CanvasRenderingContext2D;
    const fill: GradientFill = { fillType: 'gradient', gradType: 'radial', path, angle: 0,
      fillToRect: path === 'rect' ? { l: .25, t: .25, r: .25, b: .25 } : undefined,
      stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }] };
    const work: ShadeWork = { edgeRows: 0, solves: 0, rejected: 0, pixels: 0 };
    const edges = 32770;
    const pattern = resolvePathShade(fill, ctx, { x: 0, y: 0, w: 96, h: 96 }, { x: 0, y: 0, w: 96, h: 96 }, c => {
      for (let i = 0; i < edges; i++) {
        const a = i * Math.PI * 2 / edges; const r = i % 2 ? 28 : 48;
        (i ? c.lineTo.bind(c) : c.moveTo.bind(c))(48 + r * Math.cos(a), 48 + r * Math.sin(a));
      }
      c.closePath();
    }, undefined, work);
    expect(pattern).not.toBeNull();
    ctx.fillStyle = pattern as CanvasPattern;
    ctx.fillRect(0, 0, 96, 96);
    expect(pixel(ctx, 48, 48)[0]).toBeLessThan(10);
    expect(pixel(ctx, 90, 48)[0]).toBeGreaterThan(120);
    // Each pixel is solved once, even when many narrow bands overlap it.
    expect(work.solves).toBeLessThanOrEqual(work.pixels);
    expect(work.edgeRows).toBeLessThanOrEqual(edges * 100);
  });

  it('honors R2_NOP without suppressing later legitimate borders or losing the drawing position', () => {
    const ctx = render(pen, select(0), brush, select(1),
      record(0x0104, words(11, 0)), rect(0, 0, 100, 100),
      record(0x0538, words(1, 3, 0, 0, 100, 0, 0, 100)),
      record(0x0325, words(2, 0, 0, 100, 0)),
      record(0x0214, words(50, 10)), record(0x0213, words(50, 50)),
      record(0x0104, words(13)), record(0x0213, words(50, 90)));
    expect(pixel(ctx, 0, 0)[3]).toBe(0);
    expect(pixel(ctx, 20, 20)[3]).toBe(0);
    expect(pixel(ctx, 20, 50)[3]).toBe(0);
    expect(pixel(ctx, 60, 50)).toEqual([0, 0, 0, 255]);
  });

  it('keeps legitimate boundary strokes and clips a frame only when requested by the metafile', () => {
    const visible = render(pen, select(0), rect(0, 0, 100, 100));
    expect(pixel(visible, 0, 50)[3]).toBeGreaterThan(0);
    const clipped = render(pen, select(0), rect(2, 2, 98, 98, 0x0416), rect(0, 0, 100, 100));
    expect(pixel(clipped, 0, 50)[3]).toBe(0);
  });

  it.each([1, -2])('subtracts clip rectangles and restores clip, mapping and selected objects at DC level %i', level => {
    const ctx = render(brush, select(0),
      rect(10, 10, 90, 90, 0x0416), record(0x001e),
      rect(30, 30, 70, 70, 0x0415), rect(0, 0, 100, 100),
      record(0x001e), record(0x020d, words(0, 50)),
      record(0x02fc, words(0, 0, 0xff, 0)), select(1),
      record(0x0127, words(level)),
      rect(40, 40, 60, 60));
    expect(pixel(ctx, 5, 50)[3]).toBe(0);
    expect(pixel(ctx, 20, 50)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 35, 50)[3]).toBe(0);
    expect(pixel(ctx, 50, 50)).toEqual([255, 0, 0, 255]);
  });

  it('preserves the caller clip when replacing WMF clips and releases its own clip on return', () => {
    const ctx = new Canvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D;
    ctx.beginPath(); ctx.rect(0, 0, 50, 100); ctx.clip();
    const region = record(0x06ff, words(0, 6, 0, 0, 34, 1, 2, 0, 0, 100, 100,
      2, 20, 80, 10, 90, 2));
    playWmf(wmf(windowExt, brush, select(0), region,
      rect(40, 0, 60, 100, 0x0416), record(0x012c, words(1)), rect(0, 0, 100, 100)), ctx, 100, 100);
    expect(pixel(ctx, 20, 50)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 80, 50)[3]).toBe(0);
    ctx.fillStyle = '#0000FF'; ctx.fillRect(0, 0, 100, 100);
    expect(pixel(ctx, 5, 5)).toEqual([0, 0, 255, 255]);
    expect(pixel(ctx, 80, 50)[3]).toBe(0);
  });

  it('holds clips in device space across mapping changes, and offsets them only on OFFSETCLIPRGN', () => {
    const ctx = render(brush, select(0),
      record(0x020b, words(10, 10)), record(0x020c, words(-100, 100)),
      record(0x020d, words(100, 0)),
      rect(20, 20, 40, 40, 0x0416),
      record(0x020b, words(0, 0)), record(0x020c, words(100, 100)),
      record(0x020d, words(0, 0)),
      record(0x0220, words(0, 40)), rect(0, 0, 100, 100));
    expect(pixel(ctx, 20, 80)[3]).toBe(0);
    expect(pixel(ctx, 60, 80)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 60, 20)[3]).toBe(0);
  });

  it('retains off-surface selected region contents through subtraction, translation and saved DC restoration', () => {
    const region = record(0x06ff, words(0, 6, 0, 0, 34, 1, 2, 0, 0, 200, 100,
      2, 0, 100, 0, 200, 2));
    const ctx = render(brush, select(0), region, record(0x012c, words(1)),
      rect(50, 0, 60, 100, 0x0415), record(0x001e),
      record(0x0220, words(0, -100)), rect(0, 0, 100, 100),
      record(0x0127, words(-1)), record(0x02fc, words(0, 0, 0xff, 0)), select(2),
      rect(0, 0, 100, 100));
    expect(pixel(ctx, 50, 50)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 70, 50)).toEqual([0, 0, 255, 255]);
  });

  it('keeps the implicit clip independent of the playback surface after exclusion and offset', () => {
    const ctx = render(brush, select(0), rect(50, 0, 60, 100, 0x0415),
      record(0x0220, words(0, -100)), rect(0, 0, 100, 100));
    expect(pixel(ctx, 50, 50)).toEqual([255, 0, 0, 255]);
  });

  it('keeps fragmented region pieces exact through offsets, intersection and restoration', () => {
    const holes = Array.from({ length: 128 }, (_, i) => ({ x: (i * 13) % 180, y: (i * 7) % 90 }));
    const operations = holes.flatMap(({ x, y }) => [rect(x, y, x + 2, y + 2, 0x0415),
      record(0x0220, words(0, 1))]);
    const ctx = render(brush, select(0), rect(0, 0, 200, 100, 0x0416), ...operations,
      record(0x001e), rect(0, 0, 0, 0, 0x0416), record(0x0127, words(-1)),
      record(0x0220, words(0, -128)), rect(0, 0, 100, 100));
    // Independent point-membership oracle: each hole moves with every later
    // offset. The initial intersection returns to [0,200] after the final move.
    const expected = new Uint8ClampedArray(100 * 100 * 4);
    for (let y = 0; y < 100; y++) for (let x = 0; x < 100; x++) {
      const excluded = holes.some((hole, i) => x >= hole.x - i && x < hole.x - i + 2 &&
        y >= hole.y && y < hole.y + 2);
      if (!excluded) { expected[(y * 100 + x) * 4] = 255; expected[(y * 100 + x) * 4 + 3] = 255; }
    }
    expect(ctx.getImageData(0, 0, 100, 100).data).toEqual(expected);
  });

  it('honors MM_TEXT device pixels despite extent and scale records, including the clip mapping', () => {
    const ctx = render(brush, select(0), record(0x0103, words(1)),
      record(0x020c, words(200, 200)), record(0x020e, words(50, 50)),
      record(0x0410, words(1, 3, 1, 3)), record(0x0412, words(1, 2, 1, 2)),
      rect(20, 20, 35, 40, 0x0416), rect(20, 20, 40, 40));
    expect(pixel(ctx, 30, 30)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 15, 15)[3]).toBe(0);
    expect(pixel(ctx, 37, 30)[3]).toBe(0);
  });

  it.each([7, 8])('uses mode %i for axis scaling and restores it with the DC', mode => {
    const ctx = render(brush, select(0), record(0x0103, words(mode)),
      record(0x020c, words(100, 100)),
      record(0x020d, words(100, 0)), record(0x001e),
      record(0x0103, words(1)), record(0x0127, words(-1)),
      record(0x020e, words(-100, 50)), rect(20, 20, 40, 40));
    expect(pixel(ctx, 15, mode === 7 ? 85 : 70)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 15, mode === 7 ? 70 : 85)[3]).toBe(0);
  });

  it.each([[2, 254], [3, 2540], [4, 100], [5, 1000], [6, 1440]])(
    'maps fixed physical mode %i to the square 96-dpi Canvas device, ignoring extent changes', (mode, inch) => {
      const ctx = render(brush, select(0), record(0x0103, words(mode)),
        record(0x020d, words(96, 0)), record(0x020c, words(10, 10)),
        record(0x020e, words(10, 10)), record(0x0410, words(1, 2, 1, 2)),
        record(0x0412, words(1, 3, 1, 3)),
        rect(0, 0, inch / 2, inch / 2, 0x0416), rect(0, 0, inch, inch));
      expect(pixel(ctx, 24, 72)).toEqual([255, 0, 0, 255]);
      expect(pixel(ctx, 24, 24)[3]).toBe(0);
      expect(pixel(ctx, 72, 72)[3]).toBe(0);
    });

  it('keeps isotropic axes equal after window and viewport scaling', () => {
    const ctx = render(brush, select(0), record(0x0103, words(7)),
      record(0x020c, words(100, 100)), record(0x020e, words(100, 50)),
      record(0x0410, words(1, 1, 1, 2)), record(0x0412, words(1, 2, 1, 3)),
      rect(20, 20, 40, 40));
    expect(pixel(ctx, 15, 15)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 25, 15)[3]).toBe(0);
    expect(pixel(ctx, 15, 25)[3]).toBe(0);
  });

  it('selects scan regions, replacing an existing clip, and maps subsequent geometry through the viewport', () => {
    // MS-WMF 2.2.1.5 / 2.2.2.21: two intervals in a single scan band.
    const region = record(0x06ff, words(0, 6, 0, 0, 38, 1, 4, 0, 0, 100, 100,
      4, 20, 80, 10, 30, 70, 90, 4));
    const ctx = render(brush, select(0), region,
      rect(40, 0, 60, 100, 0x0416), record(0x012c, words(1)),
      record(0x020d, words(0, 10)), record(0x020e, words(100, 50)),
      rect(0, 0, 100, 100));
    expect(pixel(ctx, 20, 50)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 50, 50)[3]).toBe(0);
    expect(pixel(ctx, 80, 50)[3]).toBe(0);
  });

  it('subtracts and offsets selected scan bands under a captured mirrored mapping', () => {
    const region = record(0x06ff, words(0, 6, 0, 0, 38, 1, 4, 0, 0, 100, 100,
      4, 20, 80, 10, 30, 70, 90, 4));
    const ctx = render(brush, select(0), region,
      record(0x020c, words(100, -100)), record(0x020d, words(0, 100)),
      record(0x012c, words(1)), rect(70, 20, 80, 80, 0x0415),
      record(0x0220, words(0, -10)),
      record(0x020c, words(100, 100)), record(0x020d, words(0, 0)), rect(0, 0, 100, 100));
    expect(pixel(ctx, 25, 50)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 85, 50)).toEqual([255, 0, 0, 255]);
    expect(pixel(ctx, 15, 50)[3]).toBe(0);
    expect(pixel(ctx, 35, 50)[3]).toBe(0);
  });
});

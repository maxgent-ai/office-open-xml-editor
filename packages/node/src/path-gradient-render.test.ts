import { describe, expect, it, vi } from 'vitest';
import {
  paintDrawingMLShape, resolveFill,
  type GradientFill, type DrawingMLShapeGeometry,
} from '@silurus/ooxml-core';
import type { Presentation, ShapeElement } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { renderViewport } from '../../xlsx/src/renderer';
import { renderSlideNode } from './render';
import { resolvePathShade, type ShadeWork } from '../../core/src/shape/path-gradient';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
// Three stops keep interpolation linear, so pixel values read the geometric
// gradient position directly: shade ≈ 255·s.
const fill: GradientFill = {
  fillType: 'gradient', gradType: 'radial', path: 'rect', angle: 0,
  fillToRect: { l: .5, r: .5, t: .5, b: .5 },
  stops: [{ position: 0, color: '000000' }, { position: .5, color: '808080' },
    { position: 1, color: 'FFFFFF' }],
};
const twoStop = [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }];

/** Independent reference: triangle fan from the focus over the outline edges
 * in path order, the last containing triangle wins; s = 1 - λ(focus). */
function fanShade(polygon: number[][], focus: number[], p: number[]): number {
  let shade = 1;
  for (let i = 0; i < polygon.length; i++) {
    const a = polygon[i]; const b = polygon[(i + 1) % polygon.length];
    const det = (a[0] - focus[0]) * (b[1] - focus[1]) - (b[0] - focus[0]) * (a[1] - focus[1]);
    if (Math.abs(det) < 1e-12) continue;
    const la = ((p[0] - focus[0]) * (b[1] - focus[1]) - (b[0] - focus[0]) * (p[1] - focus[1])) / det;
    const lb = ((a[0] - focus[0]) * (p[1] - focus[1]) - (p[0] - focus[0]) * (a[1] - focus[1])) / det;
    if (la >= 0 && lb >= 0 && la + lb <= 1) shade = la + lb;
  }
  return shade;
}

/** Native solid, gradient and pattern strokes use different Skia coverage paths.
 * Linux can round their antialiased edge coverage one 8-bit unit apart. The
 * contract here is complete stroke/dash coverage, not identical AA rounding:
 * interior ink and empty regions match exactly; only edge pixels may differ
 * by one quantization unit. Large hull clipping, missing caps and filled dash
 * gaps still fail, including a single missing fully covered pixel. */
function expectStrokeCoverage(
  actual: Uint8ClampedArray, expected: Uint8ClampedArray, width: number, label = '',
): void {
  expect(actual.length, label).toBe(expected.length);
  const height = expected.length / width;
  let ink = 0; let gaps = 0;
  const differences: string[] = [];
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const i = y * width + x;
    const coverage = expected[i];
    // A fully painted or empty 3x3 neighbourhood is independent of AA at
    // the boundary. No threshold can hide a missing interior pixel or gap.
    const uniform = (coverage === 0 || coverage === 255)
      && [-1, 0, 1].every(dy => [-1, 0, 1].every(dx => {
        const u = x + dx; const v = y + dy;
        return u >= 0 && u < width && v >= 0 && v < height
          && expected[v * width + u] === coverage;
      }));
    if (uniform) { if (coverage === 255) ink++; else gaps++; }
    if (Math.abs(actual[i] - coverage) > (uniform ? 0 : 1) && differences.length < 10) {
      differences.push(`${x},${y}: ${actual[i]} != ${coverage}`);
    }
  }
  expect(ink, label).toBeGreaterThan(0);
  expect(gaps, label).toBeGreaterThan(0);
  expect(differences, label).toEqual([]);
}

function expectRedStrokePixels(
  actual: Uint8ClampedArray, expected: Uint8ClampedArray, width: number,
  opaqueBackground: boolean, label = '',
): void {
  const coverage = (rgba: Uint8ClampedArray) => rgba
    .filter((_, i) => i % 4 === (opaqueBackground ? 1 : 3))
    .map(value => opaqueBackground ? 255 - value : value);
  expectStrokeCoverage(coverage(actual), coverage(expected), width, label);
  const incorrectColors: number[] = [];
  // Retain the original colour protection as well as coverage. Ignore RGB
  // under zero alpha, where Canvas may canonicalize transparent colour bytes.
  for (let i = 0; i < actual.length; i += 4) {
    if (actual[i + 3] === 0) continue;
    if ((actual[i] !== 255 || (opaqueBackground
      ? actual[i + 1] !== actual[i + 2] || actual[i + 3] !== 255
      : actual[i + 1] !== 0 || actual[i + 2] !== 0)) && incorrectColors.length < 10) {
      incorrectColors.push(i / 4);
    }
  }
  expect(incorrectColors, label).toEqual([]);
}

describe.skipIf(!skia)('DrawingML path-gradient pixels', () => {
  function canvas(w = 240, h = 120) {
    const c = new (skia as NonNullable<typeof skia>).Canvas(w, h);
    return { c, ctx: c.getContext('2d') as unknown as CanvasRenderingContext2D };
  }
  function pixel(ctx: CanvasRenderingContext2D, x: number, y: number) {
    return [...ctx.getImageData(x, y, 1, 1).data];
  }
  function paint(recipe: GradientFill, ctx: CanvasRenderingContext2D, w: number, h: number) {
    ctx.fillStyle = resolveFill(recipe, ctx, 0, 0, w, h) as CanvasPattern;
    ctx.fillRect(0, 0, w, h);
  }

  it.each(['docx', 'pptx', 'xlsx'] as const)('assigns centered ellipse seam pixels through %s painting', async format => {
    const { c, ctx } = canvas(500, 500);
    const recipe: GradientFill = { ...fill, path: 'shape', tileRect: { l: 0, t: 0, r: 0, b: 0 } };
    if (format === 'docx') paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 500, h: 500 }, geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
      fill: recipe, stroke: null, transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    if (format === 'pptx') await renderSlideNode(c, {
      slideWidth: 500 * 9525, slideHeight: 500 * 9525,
      slides: [{ index: 0, slideNumber: 1, background: null, elements: [{
        type: 'shape', x: 0, y: 0, width: 500 * 9525, height: 500 * 9525,
        rotation: 0, flipH: false, flipV: false, geometry: 'ellipse', fill: recipe,
        stroke: null, textBody: null, custGeom: null,
      } as ShapeElement] }], defaultTextColor: null, majorFont: null, minorFont: null,
    } as Presentation, 0, { width: 500, dpr: 1 });
    if (format === 'xlsx') renderViewport(ctx, {
      name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {}, freezeRows: 0, freezeCols: 0,
      defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [], conditionalFormats: [], images: [], charts: [],
      defaultFontFamily: 'Calibri', defaultFontSize: 11,
      shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0, toCol: 1, toRow: 1,
        toColOff: 0, toRowOff: 0, editAs: 'oneCell', nativeExtCx: 500 * 9525, nativeExtCy: 500 * 9525,
        shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 0, strokeColor: undefined, strokeWidth: 0,
          fill: recipe, geom: { type: 'preset', name: 'ellipse', adj: [] } }],
      }],
    } as Worksheet, { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles,
    { row: 1, col: 1, rows: 1, cols: 1 });
    expect(Math.abs(pixel(ctx, 245, 245)[0] - 255 * Math.hypot(4.5, 4.5) / 250)).toBeLessThanOrEqual(1);
    expect(pixel(ctx, 245, 245)[3]).toBe(255);
    const data = ctx.getImageData(0, 0, 500, 500).data;
    let speckles = 0;
    for (let y = 0; y < 500; y++) for (let x = 0; x < 500; x++) {
      if (Math.hypot(x + .5 - 250, y + .5 - 250) < 212.5 && data[(y * 500 + x) * 4] === 255) speckles++;
    }
    expect(speckles).toBe(0);
  });

  it.each([170, 340, 1000])('keeps ellipse seams shaded at size %i under reflected/sheared raster scaling', size => {
    const { ctx } = canvas(800, 800);
    ctx.translate(700, 50); ctx.transform(-.6, .05, .08, .6, 0, 0);
    const inverse = ctx.getTransform().inverse();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: size, h: size }, geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
      fill: { ...fill, path: 'shape' }, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    const data = ctx.getImageData(0, 0, 800, 800).data;
    let checked = 0; let speckles = 0;
    for (let y = 0; y < 800; y++) for (let x = 0; x < 800; x++) {
      const u = inverse.a * (x + .5) + inverse.c * (y + .5) + inverse.e;
      const v = inverse.b * (x + .5) + inverse.d * (y + .5) + inverse.f;
      if (Math.hypot(u - size / 2, v - size / 2) >= size * .4) continue;
      checked++;
      if (data[(y * 800 + x) * 4] === 255 || data[(y * 800 + x) * 4 + 3] !== 255) speckles++;
    }
    expect(checked).toBeGreaterThan(1000);
    expect(speckles).toBe(0);
  });

  it('assigns all interior pixels for seeded random outlines, sizes, foci and affine raster frames', () => {
    let seed = 1599;
    const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed / 2 ** 32; };
    let checked = 0;
    for (let trial = 0; trial < 96; trial++) {
      const w = 30 + random() * 130; const h = 30 + random() * 130;
      const n = 4 + Math.floor(random() * 24);
      // Alternating radii exercise concave center-kernel polygons as well as convex outlines.
      const polygon = Array.from({ length: n }, (_, i) => {
        const angle = i / n * 2 * Math.PI;
        const r = trial % 2 && i % 2 ? .2 : .5;
        return [w * (.5 + r * Math.cos(angle)), h * (.5 + r * Math.sin(angle))];
      });
      const ox = random() * 1.5 - .25; const oy = random() * 1.5 - .25;
      // Raster support covers areas and shape point foci; segments are native.
      const kx = trial % 6 === 0 ? 0 : .01 + random() * .79;
      const ky = trial % 6 === 0 ? 0 : .01 + random() * .79;
      const { ctx } = canvas(240, 240);
      ctx.translate(trial % 2 ? 200 : 35, 35);
      ctx.transform((trial % 2 ? -1 : 1) * (.6 + random() * .4), random() * .2, random() * .2, .6 + random() * .4, 0, 0);
      const inverse = ctx.getTransform().inverse();
      const outline = (target: CanvasRenderingContext2D) => {
        polygon.forEach(([x, y], i) => target[i ? 'lineTo' : 'moveTo'](x, y)); target.closePath();
      };
      // A discontinuous final stop makes only s=1 (including an unassigned
      // pixel) white. Every assigned interior shade must remain opaque black.
      const recipe: GradientFill = { ...fill, path: trial % 6 === 0 || trial % 3 ? 'shape' : 'rect',
        fillToRect: { l: ox, r: 1 - ox - kx, t: oy, b: 1 - oy - ky },
        stops: [{ position: 0, color: '000000' }, { position: 1, color: '000000' },
          { position: 1, color: 'FFFFFF' }],
      };
      const work: ShadeWork = { edgeRows: 0, solves: 0, rejected: 0, pixels: 0 };
      const box = { x: 0, y: 0, w, h };
      const shade = resolvePathShade(recipe, ctx, box, box, outline, undefined, work);
      expect(shade, `trial ${trial}`).not.toBeNull();
      ctx.fillStyle = shade as CanvasPattern; ctx.fillRect(-1000, -1000, 2000, 2000);
      const data = ctx.getImageData(0, 0, 240, 240).data;
      const contains = (x: number, y: number) => {
        let inside = false;
        for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
          const a = polygon[i]; const b = polygon[j];
          if ((a[1] > y) !== (b[1] > y) && x < a[0] + (y - a[1]) * (b[0] - a[0]) / (b[1] - a[1])) inside = !inside;
        }
        return inside;
      };
      const holes: number[][] = [];
      for (let y = 0; y < 240; y++) for (let x = 0; x < 240; x++) {
        // Test every pixel center inside the outline and host box, with an
        // independent ray-crossing test rather than the band's row intervals.
        const u = inverse.a * (x + .5) + inverse.c * (y + .5) + inverse.e;
        const v = inverse.b * (x + .5) + inverse.d * (y + .5) + inverse.f;
        if (!(u > 0 && u < w && v > 0 && v < h && contains(u, v))) continue;
        checked++;
        const offset = (y * 240 + x) * 4;
        if (data[offset] !== 0 || data[offset + 3] !== 255) holes.push([x, y, ...data.slice(offset, offset + 4)]);
      }
      expect(holes.slice(0, 10), JSON.stringify({ trial, w, h, ox, oy, kx, ky, work, count: holes.length })).toEqual([]);
      expect(work.solves).toBeLessThanOrEqual(work.pixels);
    }
    expect(checked).toBeGreaterThan(50000);
  });

  it('shades a rectangular host outline with shape-path box isolines', () => {
    const { ctx } = canvas();
    paint({ ...fill, path: 'shape' }, ctx, 240, 120);
    // Horizontal, vertical and diagonal half-box points share one isoline.
    for (const [x, y] of [[60, 60], [120, 30], [60, 30], [180, 90]]) {
      expect(Math.abs(pixel(ctx, x, y)[0] - 128)).toBeLessThanOrEqual(3);
    }
  });

  it.each(['rect', 'shape'] as const)('%s shades explicit whole-shape tile rectangles like omitted tiles', path => {
    const expected = canvas().ctx;
    paint({ ...fill, path }, expected, 240, 120);
    for (const tileRect of [{}, { l: 0, t: 0, r: 0, b: 0 }]) {
      const actual = canvas().ctx;
      paint({ ...fill, path, tileRect }, actual, 240, 120);
      expect(Buffer.from(actual.getImageData(0, 0, 240, 120).data)
        .equals(Buffer.from(expected.getImageData(0, 0, 240, 120).data))).toBe(true);
    }
  });

  it('fills the inscribed focus area with the first stop and shades the band to the outline', () => {
    const area = canvas().ctx;
    paint({ ...fill, fillToRect: { l: .2, r: .4, t: .25, b: .25 } }, area, 240, 120);
    // Focus rectangle x 48..144, y 30..90 is flat first stop (§20.1.8.31).
    for (const [x, y] of [[50, 32], [80, 60], [141, 87]]) expect(pixel(area, x, y)[0]).toBeLessThan(4);
    // Halfway from the focus edge to the outline: s = .5 → 128.
    expect(Math.abs(pixel(area, 24, 60)[0] - 125)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(area, 192, 60)[0] - 129)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(area, 96, 105)[0] - 130)).toBeLessThanOrEqual(3);
    const zero = canvas().ctx;
    paint({ ...fill, fillToRect: { l: 0, t: 0, r: 0, b: 0 } }, zero, 240, 120);
    for (const [x, y] of [[2, 2], [120, 60], [237, 117]]) expect(pixel(zero, x, y)).toEqual([0, 0, 0, 255]);
    const omitted = canvas().ctx;
    paint({ ...fill, fillToRect: undefined }, omitted, 240, 120);
    expect(pixel(omitted, 120, 60)[0]).toBeLessThan(6);
  });

  it('keeps Canvas interpolation for two-stop gradients', () => {
    const { ctx } = canvas(400, 10);
    ctx.fillStyle = resolveFill({ fillType: 'gradient', gradType: 'linear', angle: 0, stops: twoStop },
      ctx, 0, 0, 400, 10) as CanvasGradient;
    ctx.fillRect(0, 0, 400, 10);
    expect(Math.abs(pixel(ctx, 200, 5)[0] - 128)).toBeLessThanOrEqual(1);
  });

  it('follows star-shaped outlines with a focus fan in path order', () => {
    const outline = [[0, 0], [100, 40], [200, 0], [160, 60], [200, 120], [100, 80], [0, 120], [40, 60]];
    const custom: DrawingMLShapeGeometry = { kind: 'custom', subpaths: [[
      ...outline.map(([x, y], index) => ({ cmd: index ? 'lineTo' as const : 'moveTo' as const, x: x / 200, y: y / 120 })),
      { cmd: 'close' },
    ]] };
    for (const [fillToRect, focus] of [
      [undefined, [100, 60]], [{ l: 0, t: 0, r: 1, b: 1 }, [0, 0]], [{ l: .75, t: .5, r: .25, b: .5 }, [150, 60]],
    ] as const) {
      const { ctx } = canvas(200, 120);
      paintDrawingMLShape(ctx, {
        rect: { x: 0, y: 0, w: 200, h: 120 }, geometry: custom, stroke: null,
        fill: { ...fill, path: 'shape', fillToRect }, transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
      for (const p of [[100, 60], [70, 55], [130, 70], [60, 30], [150, 95], [100, 50], [45, 80]]) {
        const expected = 255 * fanShade(outline, focus as unknown as number[], [p[0] + .5, p[1] + .5]);
        expect(Math.abs(pixel(ctx, p[0], p[1])[0] - expected), JSON.stringify({ focus, p })).toBeLessThanOrEqual(5);
      }
    }
  });

  it("retains main radial fallback beyond the supported outline boundary", () => {
    const arrow: DrawingMLShapeGeometry = { kind: 'preset', name: 'rightArrow', adjustments: [] };
    for (const [fillToRect, focus] of [
      [{ l: 0, t: 0, r: 0, b: 0 }, [60, 60]], [{ l: 0, t: 0, r: 1, b: 1 }, [0, 0]],
      [{ l: .2, t: 0, r: 0, b: 0 }, [72, 60]], [{ l: .2, r: .2, t: .2, b: .2 }, [60, 60]], [undefined, [60, 60]],
    ] as const) {
      const { ctx } = canvas(120, 120);
      paintDrawingMLShape(ctx, {
        rect: { x: 0, y: 0, w: 120, h: 120 }, geometry: arrow, stroke: null,
        fill: { ...fill, path: 'shape', fillToRect }, transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
      const radius = Math.max(...[[0, 0], [120, 0], [0, 120], [120, 120]]
        .map(([x, y]) => Math.hypot(x - focus[0], y - focus[1])));
      for (const [x, y] of [[10, 40], [50, 60], [90, 50], [72, 30], [100, 62]]) {
        const expected = 255 * Math.min(1, Math.hypot(x + .5 - focus[0], y + .5 - focus[1]) / radius);
        expect(Math.abs(pixel(ctx, x, y)[0] - expected), JSON.stringify({ fillToRect, x, y })).toBeLessThanOrEqual(3);
      }
    }
    // A convex ellipse instead keeps its inscribed focus area.
    const ellipse = canvas(120, 120).ctx;
    paintDrawingMLShape(ellipse, {
      rect: { x: 0, y: 0, w: 120, h: 120 }, stroke: null,
      geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
      fill: { ...fill, path: 'shape', fillToRect: { l: .2, r: .2, t: .2, b: .2 } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ellipse, 60, 60)[0]).toBeLessThan(4);
    expect(pixel(ellipse, 60 + 33, 60)[0]).toBeLessThan(4);
    // Contour radius .3 + .2s: radius .4 of the box is s = .5.
    expect(Math.abs(pixel(ellipse, 60 + 48, 60)[0] - 133)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ellipse, 60, 60 - 48)[0] - 123)).toBeLessThanOrEqual(3);
  });

  it('keeps wide gradient strokes opaque outside the host box', () => {
    for (const path of ['rect', 'shape'] as const) {
      const { ctx } = canvas(200, 200);
      ctx.lineWidth = 20;
      ctx.strokeStyle = resolveFill({ ...fill, path }, ctx, 50, 50, 100, 100) as CanvasPattern;
      ctx.strokeRect(50, 50, 100, 100);
      // main painted these pixels opaque; the shade beyond the frame is the outer stop.
      expect(pixel(ctx, 41, 100)).toEqual([255, 255, 255, 255]);
      expect(pixel(ctx, 100, 158)).toEqual([255, 255, 255, 255]);
      expect(pixel(ctx, 158, 158)[3]).toBe(255);
    }
  });

  it('retains main authored tile flips with a tile-local native focus', () => {
    const { ctx } = canvas();
    paint({ ...fill, tileRect: { r: .5 }, flip: 'x', fillToRect: { l: 0, r: 1, t: 0, b: 1 } }, ctx, 240, 120);
    expect(pixel(ctx, 3, 3)[0]).toBeLessThan(12);
    expect(pixel(ctx, 123, 3)[0]).toBeGreaterThan(243);
    expect(pixel(ctx, 117, 117)[0]).toBeGreaterThan(243);
    const inset = canvas().ctx;
    paint({ ...fill, tileRect: { l: .25, t: .25, r: .25, b: .25 },
      fillToRect: { l: .25, t: .25, r: .75, b: .75 } }, inset, 240, 120);
    // Main puts this tile's quarter-point focus at (90,45) in the host.
    expect(pixel(inset, 90, 45)[0]).toBeLessThan(4);
    const expected = canvas().ctx;
    const tile = canvas(120, 60).ctx;
    const native = tile.createRadialGradient(30, 15, 0, 30, 15, 90);
    native.addColorStop(0, '#000000'); native.addColorStop(.5, '#808080'); native.addColorStop(1, '#FFFFFF');
    tile.fillStyle = native; tile.fillRect(0, 0, 120, 60);
    const pattern = expected.createPattern(tile.canvas, 'repeat') as CanvasPattern;
    pattern.setTransform({ a: 1, b: 0, c: 0, d: 1, e: 60, f: 30 });
    expected.fillStyle = pattern; expected.fillRect(0, 0, 240, 120);
    expect(inset.getImageData(0, 0, 240, 120).data).toEqual(expected.getImageData(0, 0, 240, 120).data);
  });

  it('keeps out-of-box rect shading opaque without moving the authored gradient frame', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 120, h: 120 },
      geometry: { kind: 'custom', subpaths: [[
        { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 2, y: 1 },
        { cmd: 'lineTo', x: 0, y: 1 }, { cmd: 'close' },
      ]] },
      fill: { ...fill, fillToRect: { l: .25, r: .25, t: .25, b: .25 } }, stroke: null,
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(ctx, 180, 105)).toEqual([255, 255, 255, 255]);
    expect(pixel(ctx, 60, 60)[0]).toBeLessThan(6);
    expect(Math.abs(pixel(ctx, 105, 105)[0] - 132)).toBeLessThanOrEqual(3);
  });

  it('observes coverage without changing the caller path or canvas state', () => {
    const { ctx } = canvas();
    ctx.translate(10, 5);
    ctx.globalAlpha = .5;
    ctx.lineWidth = 7;
    ctx.beginPath();
    ctx.rect(20, 20, 10, 10);
    ctx.fillStyle = resolveFill(fill, ctx, 0, 0, 120, 120, 0, undefined, undefined,
      (target, x, y, w, h) => {
        target.save(); target.translate(1, 1); target.restore();
        target.moveTo(x, y);
        target.lineTo(x + w * 2, y + h);
        target.lineTo(x, y + h);
        target.closePath();
      }) as CanvasPattern;
    ctx.fill();
    expect(pixel(ctx, 34, 29)[3]).toBe(128);
    expect(pixel(ctx, 80, 60)[3]).toBe(0);
    expect(ctx.globalAlpha).toBe(.5);
    expect(ctx.lineWidth).toBe(7);
    expect(ctx.getTransform().e).toBe(10);
  });

  it('interpolates translucent stops once', () => {
    const { ctx } = canvas();
    paint({ ...fill, path: 'shape', stops: [{ position: 0, color: 'FF0000FF' }, { position: .5, color: 'FF000080' },
      { position: 1, color: 'FF000000' }] }, ctx, 240, 120);
    expect(pixel(ctx, 120, 60)[3]).toBeGreaterThan(250);
    expect(Math.abs(pixel(ctx, 60, 60)[3] - 128)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ctx, 120, 30)[3] - 128)).toBeLessThanOrEqual(3);
  });

  it('keeps the previous local radial frame for rotWithShape=false', () => {
    const recipe = { ...fill, fillToRect: { l: .1, r: .5, t: .4, b: .1 } };
    const contexts = [canvas(160, 160).ctx, canvas(160, 160).ctx];
    for (const [i, ctx] of contexts.entries()) {
      ctx.translate(80, 80); ctx.rotate(Math.PI / 4); ctx.scale(-1, 1); ctx.translate(-50, -50);
      ctx.fillStyle = resolveFill({ ...recipe, rotWithShape: i === 0 }, ctx, 0, 0, 100, 100, 45) as CanvasPattern;
      ctx.fillRect(0, 0, 100, 100);
    }
    expect(contexts[1].getImageData(0, 0, 160, 160).data).toEqual(contexts[0].getImageData(0, 0, 160, 160).data);
  });

  it('bounds auxiliary allocation even for very large authored extents', () => {
    const { ctx } = canvas(2, 2);
    const allocations: number[][] = [];
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      constructor(w: number, h: number) {
        allocations.push([w, h]);
        // The shade raster, or the 1024×1 stop ramp.
        expect(w * h).toBeLessThanOrEqual(512 * 512); expect(Math.max(w, h)).toBeLessThanOrEqual(1024);
        super(w, h);
      }
    });
    try {
      paint({ ...fill, path: 'shape' }, ctx, 1e9, 1e9);
      expect(pixel(ctx, 0, 0)[3]).toBe(255);
      expect(allocations.some(([w, h]) => w > 500 && h > 500)).toBe(true);
    } finally { vi.unstubAllGlobals(); }
  });

  it('keeps repaint work to one outline pass and one bounded pixel write', () => {
    const { ctx } = canvas(512, 512);
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const work = { surfaces: 0, readbacks: 0, writes: 0, outlines: 0 };
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      constructor(w: number, h: number) { super(w, h); work.surfaces++; }
      getContext(type: '2d') {
        const target = super.getContext(type);
        for (const method of ['getImageData', 'putImageData'] as const) {
          const original = target[method].bind(target);
          vi.spyOn(target, method).mockImplementation((...args: unknown[]) => {
            if (method === 'getImageData') work.readbacks++; else work.writes++;
            return Reflect.apply(original, target, args);
          });
        }
        return target;
      }
    });
    const edges = 720;
    const outline = (target: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) => {
      work.outlines++;
      for (let i = 0; i < edges; i++) {
        const angle = i / edges * Math.PI * 2; const r = i % 2 ? .5 : .3;
        target[i ? 'lineTo' : 'moveTo'](x + w * (.5 + r * Math.cos(angle)), y + h * (.5 + r * Math.sin(angle)));
      }
      target.closePath();
    };
    try {
      for (let i = 0; i < 3; i++) {
        ctx.fillStyle = resolveFill({ ...fill, path: 'shape' }, ctx, 0, 0, 512, 512, 0,
          undefined, undefined, outline) as CanvasPattern;
        ctx.fillRect(0, 0, 512, 512);
        expect(pixel(ctx, 256, 256)[0]).toBeLessThan(4);
      }
      // Per paint: the shade raster plus the 1024×1 stop ramp and its readback.
      expect(work).toEqual({ surfaces: 6, readbacks: 3, writes: 3, outlines: 3 });
    } finally { vi.unstubAllGlobals(); vi.restoreAllMocks(); }
  });

  it('bounds band solves and covers interiors for corner and nonnested area foci', () => {
    const { ctx } = canvas(512, 512);
    for (const path of ['rect', 'shape'] as const) for (const edges of [720, 20000, 32768, 32769]) {
      const outline = (target: CanvasRenderingContext2D, x: number, y: number, w: number, h: number) => {
        for (let i = 0; i < edges; i++) {
          const angle = i / edges * Math.PI * 2; const r = i % 2 ? .5 : .03;
          target[i ? 'lineTo' : 'moveTo'](x + w * (.5 + r * Math.cos(angle)), y + h * (.5 + r * Math.sin(angle)));
        }
        target.closePath();
      };
      for (const focus of [{ l: 0, t: 0, r: 1, b: 1 }, { l: 1, t: 1, r: -.5, b: -.5 }]) {
        const work: ShadeWork = { edgeRows: 0, solves: 0, rejected: 0, pixels: 0 };
        // Keep device pixels aligned with raster pixels; no downsampling.
        const size = 507;
        const box = { x: 0, y: 0, w: size, h: size };
        const paint = resolvePathShade({ ...fill, path, fillToRect: focus,
          stops: [{ position: 0, color: '000000' }, { position: 1, color: '000000' },
            { position: 1, color: 'FFFFFF' }],
        }, ctx, box, box, outline, undefined, work);
        if (edges > 32768 || (path === 'rect' && focus.r === 1)) {
          expect(paint).toBeNull();
          expect(work).toEqual({ edgeRows: 0, solves: 0, rejected: 0, pixels: 0 });
          continue;
        }
        expect(paint).not.toBeNull();
        expect(work.edgeRows).toBeLessThanOrEqual((edges + 1) * 512);
        expect(work.solves, JSON.stringify({ path, edges, focus, work })).toBeLessThanOrEqual(work.pixels);
        expect(work.rejected).toBe(0);
        expect(work.pixels).toBeLessThanOrEqual(512 * 512);
        if (path === 'shape' && edges === 32768) {
          ctx.fillStyle = paint as CanvasPattern; ctx.fillRect(0, 0, size, size);
          const data = ctx.getImageData(0, 0, size, size).data;
          let checked = 0; let holes = 0;
          // Independent radial intersection with the authored polygon edge:
          // angle selects an edge; its cross product gives the outline radius.
          // This avoids repeating the shader's band/row/solve implementation.
          for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
            const px = (x + .5) / size - .5; const py = (y + .5) / size - .5;
            const angle = (Math.atan2(py, px) + 2 * Math.PI) % (2 * Math.PI);
            const i = Math.floor(angle / (2 * Math.PI) * edges);
            const a = i / edges * 2 * Math.PI; const b = (i + 1) / edges * 2 * Math.PI;
            const r1 = i % 2 ? .5 : .03; const r2 = (i + 1) % 2 ? .5 : .03;
            const ax = r1 * Math.cos(a); const ay = r1 * Math.sin(a);
            const bx = r2 * Math.cos(b); const by = r2 * Math.sin(b);
            const radius = (ax * by - ay * bx) / (Math.cos(angle) * (by - ay) - Math.sin(angle) * (bx - ax));
            if (Math.hypot(px, py) >= radius) continue;
            checked++;
            const offset = (y * size + x) * 4;
            if (data[offset] !== 0 || data[offset + 3] !== 255) holes++;
          }
          expect(checked).toBeGreaterThan(10000);
          expect(holes, JSON.stringify({ focus, work })).toBe(0);
        }
      }
    }
  }, 60000);

  it.each(['rect', 'shape'] as const)('preserves main %s fallback bytes for edge limits and unavailable allocation', path => {
    const outline = (target: CanvasRenderingContext2D) => {
      for (let i = 0; i < 32769; i++) {
        const angle = i / 32769 * Math.PI * 2;
        target[i ? 'lineTo' : 'moveTo'](100 + 100 * Math.cos(angle), 50 + 50 * Math.sin(angle));
      }
      target.closePath();
    };
    for (const reason of ['edge budget', 'allocation'] as const) for (const tiled of [false, true]) {
      const actual = canvas(300, 200).ctx; const expected = canvas(300, 200).ctx;
      for (const ctx of [actual, expected]) { ctx.translate(70, 30); ctx.rotate(.3); }
      const recipe = { ...fill, path, fillToRect: { l: .1, r: .5, t: .4, b: .1 },
        tileRect: tiled ? { r: .5 } : undefined };
      const unavailable = new Proxy(actual, {
        get(target, key) {
          if (key === 'canvas') return { constructor: null };
          const value = Reflect.get(target, key, target);
          return typeof value === 'function' ? value.bind(target) : value;
        },
      });
      actual.fillStyle = resolveFill(recipe, reason === 'allocation' ? unavailable : actual,
        0, 0, 200, 100, 0, undefined, undefined, reason === 'edge budget' ? outline : undefined) as CanvasGradient;
      // Independent native Canvas oracle from main: midpoint (60,65),
      // farthest axis distances (140,65), rect uses max; shape uses diagonal.
      // A half-width tile halves x distances. If allocation is unavailable,
      // main paints the same native field directly instead of repeating it.
      const reference = tiled && reason === 'edge budget' ? canvas(100, 100).ctx : expected;
      const cx = tiled ? 30 : 60; const rx = tiled ? 70 : 140;
      const gradient = reference.createRadialGradient(cx, 65, 0, cx, 65,
        path === 'rect' ? Math.max(rx, 65) : Math.sqrt(rx ** 2 + 65 ** 2));
      gradient.addColorStop(0, '#000000'); gradient.addColorStop(.5, '#808080'); gradient.addColorStop(1, '#FFFFFF');
      reference.fillStyle = gradient;
      if (reference !== expected) {
        reference.fillRect(0, 0, 100, 100);
        expected.fillStyle = expected.createPattern(reference.canvas, 'repeat') as CanvasPattern;
      }
      for (const ctx of [actual, expected]) ctx.fillRect(0, 0, 200, 100);
      expect(actual.getImageData(0, 0, 300, 200).data, JSON.stringify({ reason, tiled }))
        .toEqual(expected.getImageData(0, 0, 300, 200).data);
    }
  });

  it.each(['rect', 'shape'] as const)('retains main %s midpoint behavior for inverted focus axes, including tiles', path => {
    for (const focus of [
      { l: .8, r: .4, t: .2, b: .2 },
      { l: .8, r: .8, t: .2, b: .2 },
      { l: .8, r: .4, t: .9, b: .3 },
    ]) for (const tiled of [false, true]) {
      const actual = canvas(200, 100).ctx; const expected = canvas(200, 100).ctx;
      const recipe = { ...fill, path, fillToRect: focus, tileRect: tiled ? { r: .5 } : undefined };
      paint(recipe, actual, 200, 100);
      const reference = tiled ? canvas(100, 100).ctx : expected;
      const width = tiled ? 100 : 200;
      const cx = width * (1 + focus.l - focus.r) / 2;
      const cy = 100 * (1 + focus.t - focus.b) / 2;
      const rx = Math.max(cx, width - cx); const ry = Math.max(cy, 100 - cy);
      const gradient = reference.createRadialGradient(cx, cy, 0, cx, cy,
        path === 'rect' ? Math.max(rx, ry) : Math.sqrt(rx * rx + ry * ry));
      gradient.addColorStop(0, '#000000'); gradient.addColorStop(.5, '#808080'); gradient.addColorStop(1, '#FFFFFF');
      reference.fillStyle = gradient; reference.fillRect(0, 0, width, 100);
      if (tiled) { expected.fillStyle = expected.createPattern(reference.canvas, 'repeat') as CanvasPattern; expected.fillRect(0, 0, 200, 100); }
      expect(actual.getImageData(0, 0, 200, 100).data, JSON.stringify({ focus, tiled })).toEqual(expected.getImageData(0, 0, 200, 100).data);
    }
  });

  it('preserves every interior pixel of a displaced concave focus copy', () => {
    const points = [[0, 0], [100, 40], [200, 0], [160, 60], [200, 120], [100, 80], [0, 120], [40, 60]];
    const { ctx } = canvas(200, 120);
    const outline = (target: CanvasRenderingContext2D) => {
      points.forEach(([x, y], i) => target[i ? 'lineTo' : 'moveTo'](x, y)); target.closePath();
    };
    ctx.fillStyle = resolveFill({ ...fill, path: 'shape', fillToRect: { l: .1, r: .5, t: .4, b: .1 } },
      ctx, 0, 0, 200, 120, 0, undefined, undefined, outline) as CanvasPattern;
    ctx.beginPath(); outline(ctx); ctx.fill();
    const contains = (polygon: number[][], x: number, y: number) => {
      let inside = false;
      for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
        const a = polygon[i]; const b = polygon[j];
        if ((a[1] > y) !== (b[1] > y) && x < (b[0] - a[0]) * (y - a[1]) / (b[1] - a[1]) + a[0]) inside = !inside;
      }
      return inside;
    };
    const focus = points.map(([x, y]) => [20 + .4 * x, 48 + .5 * y]);
    const pixels = ctx.getImageData(0, 0, 200, 120).data;
    let interior = 0;
    // Exclude AA boundaries independently by requiring the four neighbouring
    // corners to belong to both polygons; compare every remaining pixel.
    for (let y = 0; y < 120; y++) for (let x = 0; x < 200; x++) {
      if (![points, focus].every(p => [[x, y], [x + 1, y], [x, y + 1], [x + 1, y + 1]].every(([u, v]) => contains(p, u, v)))) continue;
      interior++;
      expect([...pixels.slice((y * 200 + x) * 4, (y * 200 + x) * 4 + 4)]).toEqual([0, 0, 0, 255]);
    }
    expect(interior).toBeGreaterThan(1500);
    expect(pixel(ctx, 83, 86).slice(0, 3)).toEqual([0, 0, 0]);
  });

  it.each(['rect', 'shape'] as const)('retains main %s native paint for degenerate tile rectangles', path => {
    for (const [tileRect, cx, cy, rx, ry] of [
      [{ l: .5, r: .5 }, 100, 65, 0, 65],
    ] as const) {
      const actual = canvas(200, 100).ctx; const expected = canvas(200, 100).ctx;
      paint({ ...fill, path, tileRect, fillToRect: { l: .1, r: .5, t: .4, b: .1 } }, actual, 200, 100);
      const native = expected.createRadialGradient(cx, cy, 0, cx, cy,
        path === 'rect' ? Math.max(rx, ry) : Math.sqrt(rx * rx + ry * ry));
      native.addColorStop(0, '#000000'); native.addColorStop(.5, '#808080'); native.addColorStop(1, '#FFFFFF');
      expected.fillStyle = native; expected.fillRect(0, 0, 200, 100);
      expect(actual.getImageData(0, 0, 200, 100).data, JSON.stringify(tileRect))
        .toEqual(expected.getImageData(0, 0, 200, 100).data);
    }
  });

  it('does no outline or raster pixel work for tiled rect/shape paints, including slow outlines', () => {
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const work = { surfaces: 0, readbacks: 0, writes: 0, outlines: 0 };
    vi.stubGlobal('OffscreenCanvas', class extends Canvas {
      constructor(w: number, h: number) { super(w, h); work.surfaces++; }
      getContext(type: '2d') {
        const ctx = super.getContext(type);
        for (const method of ['getImageData', 'putImageData'] as const) {
          const original = ctx[method].bind(ctx);
          vi.spyOn(ctx, method).mockImplementation((...args: unknown[]) => {
            if (method === 'getImageData') work.readbacks++; else work.writes++;
            return Reflect.apply(original, ctx, args);
          });
        }
        return ctx;
      }
    });
    try {
      for (const path of ['rect', 'shape'] as const) for (const edges of [720, 20000, 32768]) {
        const ctx = canvas(512, 512).ctx;
        const outline = (target: CanvasRenderingContext2D) => {
          work.outlines++;
          for (let i = 0; i < edges; i++) {
            const angle = i / edges * Math.PI * 2; const r = i % 2 ? 256 : 15.36;
            target[i ? 'lineTo' : 'moveTo'](256 + r * Math.cos(angle), 256 + r * Math.sin(angle));
          }
          target.closePath();
        };
        const before = { ...work };
        const paint = resolveFill({ ...fill, path, tileRect: { r: .5 }, flip: 'xy',
          fillToRect: { l: 1, t: 1, r: -.5, b: -.5 } }, ctx, 0, 0, 512, 512, 0,
          undefined, undefined, outline);
        expect(paint).not.toBeNull();
        // Main's exact work: one native base tile and one flip repeat surface.
        expect(work).toEqual({ ...before, surfaces: before.surfaces + 2 });
      }
    } finally { vi.unstubAllGlobals(); vi.restoreAllMocks(); }
  });

  it('covers affine strokes and large line decorations from their painted geometry', () => {
    const affine = canvas(512, 200).ctx;
    affine.scale(100, 1); affine.lineWidth = 1; affine.miterLimit = 1;
    affine.strokeStyle = resolveFill({ ...fill, path: 'shape' }, affine, 2, 50, 2, 100) as CanvasPattern;
    affine.strokeRect(2, 50, 2, 100);
    expect(pixel(affine, 160, 100)).toEqual([255, 255, 255, 255]);
    const decorated = canvas(500, 400).ctx;
    paintDrawingMLShape(decorated, {
      rect: { x: 200, y: 200, w: 50, h: 1 },
      geometry: { kind: 'preset', name: 'line', adjustments: [] }, fill: null,
      stroke: { color: 'FFFFFF', width: 20, fill: { ...fill, path: 'shape' }, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(pixel(decorated, 150, 190)).toEqual([255, 255, 255, 255]);
  });

  it('covers out-of-box paths, acute miters and curved stroke hulls', () => {
    const geometry: DrawingMLShapeGeometry = { kind: 'custom', subpaths: [[
      { cmd: 'moveTo', x: 0, y: 1 }, { cmd: 'lineTo', x: 2, y: 0 },
      { cmd: 'lineTo', x: 0, y: .9 },
    ], [
      { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'cubicBezTo', x1: -1, y1: -1, x2: 3, y2: 2, x: 1, y: 1 },
    ]] };
    const contexts = [canvas(500, 350).ctx, canvas(500, 350).ctx];
    for (const [i, ctx] of contexts.entries()) {
      ctx.translate(160, 100); ctx.transform(1.5, .1, .4, 1, 0, 0);
      paintDrawingMLShape(ctx, { rect: { x: 0, y: 0, w: 80, h: 100 }, geometry, fill: null,
        stroke: { color: 'FFFFFF', width: 10, lineJoin: 'miter', miterLimit: 20,
          lineCap: 'square', ...(i ? { fill } : {}) },
        transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
    }
    const alpha = (ctx: CanvasRenderingContext2D) => ctx.getImageData(0, 0, 500, 350).data.filter((_, i) => i % 4 === 3);
    expectStrokeCoverage(alpha(contexts[1]), alpha(contexts[0]), 500);
  });

  it.each(['docx', 'pptx', 'xlsx'] as const)('covers curved dashed strokes with every cap through %s painting', async format => {
    const red: GradientFill = { ...fill, stops: [0, .5, 1].map(position => ({ position, color: 'FF0000' })) };
    for (const geometry of ['ellipse', 'roundRect']) for (const lineCap of ['square', 'round', 'butt'] as const) {
      const contexts = [canvas(300, 300), canvas(300, 300)];
      for (const [i, { c, ctx }] of contexts.entries()) {
        const stroke = { color: 'FF0000', width: 80, lineCap, dashStyle: 'sysDot', ...(i ? { fill: red } : {}) };
        if (format === 'docx') paintDrawingMLShape(ctx, {
          rect: { x: 130, y: 130, w: 40, h: 40 }, geometry: { kind: 'preset', name: geometry, adjustments: [] },
          fill: null, stroke, transform: { rotationDeg: 0, flipH: false, flipV: false },
        }, 1);
        if (format === 'pptx') await renderSlideNode(c, {
          slideWidth: 300 * 9525, slideHeight: 300 * 9525,
          slides: [{ index: 0, slideNumber: 1, background: null, elements: [{
            type: 'shape', x: 130 * 9525, y: 130 * 9525, width: 40 * 9525, height: 40 * 9525,
            rotation: 0, flipH: false, flipV: false, geometry, fill: null,
            stroke: { ...stroke, width: 80 * 9525 }, textBody: null, custGeom: null,
          } as ShapeElement] }], defaultTextColor: null, majorFont: null, minorFont: null,
        } as Presentation, 0, { width: 300, dpr: 1 });
        if (format === 'xlsx') renderViewport(ctx, {
          name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {}, freezeRows: 0, freezeCols: 0,
          defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [], conditionalFormats: [], images: [], charts: [],
          defaultFontFamily: 'Calibri', defaultFontSize: 11,
          shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0, toCol: 1, toRow: 1,
            toColOff: 0, toRowOff: 0, editAs: 'oneCell', nativeExtCx: 300 * 9525, nativeExtCy: 300 * 9525,
            shapes: [{ x: 130 / 300, y: 130 / 300, w: 40 / 300, h: 40 / 300, rot: 0,
              strokeColor: 'FF0000', strokeWidth: 80 * 9525, strokeLineCap: lineCap, strokeDashStyle: 'sysDot',
              strokeFill: i ? red : undefined, fill: undefined, geom: { type: 'preset', name: geometry, adj: [] } }],
          }],
        } as Worksheet, { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles,
        { row: 1, col: 1, rows: 1, cols: 1 });
      }
      expectRedStrokePixels(contexts[1].ctx.getImageData(0, 0, 300, 300).data,
        contexts[0].ctx.getImageData(0, 0, 300, 300).data, 300, format !== 'docx',
        JSON.stringify({ geometry, lineCap }));
    }
    // Multiple native/raster paints run under full-suite CPU contention.
  }, 30000);

  it('covers square dash tangents under nonuniform scale and shear', () => {
    const contexts = [canvas(400, 300).ctx, canvas(400, 300).ctx];
    for (const [i, ctx] of contexts.entries()) {
      ctx.translate(180, 120); ctx.transform(1.7, .3, .6, .7, 0, 0);
      paintDrawingMLShape(ctx, {
        rect: { x: 0, y: 0, w: 40, h: 40 }, geometry: { kind: 'preset', name: 'ellipse', adjustments: [] },
        fill: null, stroke: { color: 'FF0000', width: 80, lineCap: 'square', dashStyle: 'sysDot',
          ...(i ? { fill: { ...fill, stops: [0, .5, 1].map(position => ({ position, color: 'FF0000' })) } } : {}) },
        transform: { rotationDeg: 0, flipH: false, flipV: false },
      }, 1);
    }
    expectRedStrokePixels(contexts[1].getImageData(0, 0, 400, 300).data,
      contexts[0].getImageData(0, 0, 400, 300).data, 400, false);
  });

  it('covers PPTX gradient line decorations through the complete slide painter', async () => {
    const { c, ctx } = canvas(400, 400);
    const px = (n: number) => n * 9525;
    const shape = { type: 'shape', x: px(200), y: px(200), width: px(50), height: px(1),
      rotation: 0, flipH: false, flipV: false, geometry: 'line', fill: null, textBody: null, custGeom: null,
      stroke: { color: 'FFFFFF', width: px(20), fill: { ...fill, path: 'shape', tileRect: {} }, tailEnd: { type: 'triangle', w: 'lg', len: 'lg' } },
    } as ShapeElement;
    const presentation = { slideWidth: px(400), slideHeight: px(400),
      slides: [{ index: 0, slideNumber: 1, background: { fillType: 'solid', color: '00FF00' }, elements: [shape] }],
      defaultTextColor: null, majorFont: null, minorFont: null } as Presentation;
    await renderSlideNode(c, presentation, 0, { width: 400, dpr: 1 });
    expect(pixel(ctx, 150, 190)).toEqual([255, 255, 255, 255]);
    shape.stroke = { color: 'FFFFFF', width: px(20), fill: { ...fill, path: 'shape', tileRect: {} }, cmpd: 'dbl' };
    await renderSlideNode(c, presentation, 0, { width: 400, dpr: 1 });
    for (const y of [193, 207]) expect(pixel(ctx, 225, y)).toEqual([255, 255, 255, 255]);
  });

  it('wires rect fills and host-box shape-path strokes through the retained DOCX painter', () => {
    const { ctx } = canvas();
    paintDrawingMLShape(ctx, {
      rect: { x: 0, y: 0, w: 240, h: 120 },
      geometry: { kind: 'preset', name: 'triangle', adjustments: [] },
      fill: { ...fill, tileRect: {}, fillToRect: { l: .25, r: .25, t: .25, b: .25 } }, stroke: { color: '000000', width: 8, fill: { ...fill, path: 'shape', tileRect: {} } },
      transform: { rotationDeg: 0, flipH: false, flipV: false },
    }, 1);
    expect(Math.abs(pixel(ctx, 120, 15)[0] - 123)).toBeLessThanOrEqual(3);
    expect(Math.abs(pixel(ctx, 120, 117)[0] - 242)).toBeLessThanOrEqual(3);
  });

  it('wires PPTX rectangular shading through the complete slide painter', async () => {
    const { c, ctx } = canvas();
    const px = (n: number) => n * 9525;
    const shape = {
      type: 'shape', x: 0, y: 0, width: px(240), height: px(120),
      rotation: 0, flipH: false, flipV: false, geometry: 'triangle',
      fill: { ...fill, tileRect: {}, fillToRect: { l: .25, r: .25, t: .25, b: .25 } }, stroke: { color: '000000', width: px(8), fill: { ...fill, path: 'shape', tileRect: {} } }, textBody: null,
      custGeom: null, shadow: null,
    } as ShapeElement;
    const presentation = {
      slideWidth: px(240), slideHeight: px(120),
      slides: [{ index: 0, slideNumber: 1, background: null, elements: [shape] }],
      defaultTextColor: null, majorFont: null, minorFont: null,
    } as Presentation;
    await renderSlideNode(c, presentation, 0, { width: 240, dpr: 1 });
    expect(Math.abs(pixel(ctx, 120, 15)[0] - 123)).toBeLessThanOrEqual(3);
    // Shape-path strokes have no fill silhouette: every format shades the
    // host box, whose rectangle is star-shaped (box isolines, s = .95).
    expect(Math.abs(pixel(ctx, 120, 117)[0] - 242)).toBeLessThanOrEqual(3);
  });

  it('wires XLSX custom coordinates and excludes unfilled decorative paths', () => {
    const { ctx } = canvas();
    const triangle = [
      { op: 'moveTo', x: 120, y: 0 }, { op: 'lineTo', x: 240, y: 120 },
      { op: 'lineTo', x: 0, y: 120 }, { op: 'close' },
    ];
    const worksheet = {
      name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {},
      freezeRows: 0, freezeCols: 0,
      defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [], conditionalFormats: [],
      images: [], charts: [], defaultFontFamily: 'Calibri', defaultFontSize: 11,
      shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: 0, fromRowOff: 0,
        toCol: 1, toRow: 1, toColOff: 0, toRowOff: 0, editAs: 'oneCell',
        nativeExtCx: 240 * 9525, nativeExtCy: 120 * 9525,
        shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 0, strokeColor: '000000', strokeWidth: 8 * 9525,
          strokeFill: { ...fill, path: 'shape', tileRect: {} },
          fill: { ...fill, tileRect: {}, fillToRect: { l: .25, r: .25, t: .25, b: .25 } }, geom: { type: 'custom', paths: [
            { w: 240, h: 120, commands: triangle },
            { w: 240, h: 120, fill: 'none', stroke: false, commands: [
              { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 240, y: 0 },
              { op: 'lineTo', x: 240, y: 120 }, { op: 'lineTo', x: 0, y: 120 }, { op: 'close' },
            ] },
          ] } }],
      }],
    } as Worksheet;
    const styles = { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles;
    renderViewport(ctx, worksheet, styles, { row: 1, col: 1, rows: 1, cols: 1 });
    expect(Math.abs(pixel(ctx, 120, 15)[0] - 123)).toBeLessThanOrEqual(3);
    // Shape-path strokes have no fill silhouette: every format shades the
    // host box, whose rectangle is star-shaped (box isolines, s = .95).
    expect(Math.abs(pixel(ctx, 120, 117)[0] - 242)).toBeLessThanOrEqual(3);
    expect(pixel(ctx, 20, 10)).toEqual([255, 255, 255, 255]);
  });
});

import { describe, it, expect } from 'vitest';
import {
  drawArrowHead,
  lineEndPaintExtent,
  lineEndRetract,
  retractLineEndpoint,
} from './arrow';
import { lineEndMetrics, resolveArrowGeometry, trimGeometryPathEnds } from './arrow-geometry';
import type { GeometryPath } from './path-data';
import type { ArrowEnd, Stroke } from '../types/common';

// Expectations are PowerPoint PDF measurements (points) of a synthetic control
// deck, not values derived from the implementation. Strokes are in EMU with a
// scale of 1 px per point, so device px read directly as points.
const EMU = 12700;
const SCALE = 1 / EMU;
const pt = (width: number): Stroke => ({ color: '#000000', width: width * EMU });
const end = (type: string, w = 'med', len = 'med'): ArrowEnd => ({ type, w, len });
const geometry = (e: ArrowEnd, stroke: Stroke) => resolveArrowGeometry(e, stroke, SCALE, EMU)!;
const points = (e: ArrowEnd, stroke: Stroke) =>
  geometry(e, stroke).path.filter(c => c.op === 'moveTo' || c.op === 'lineTo').map(c => c.args.map(v => +v.toFixed(3)));

describe('line-end size table (ST_LineEndWidth / ST_LineEndLength)', () => {
  it('spans sm 2, med 3, lg 5 line widths, independently across and along the line', () => {
    const cases: Array<[string, string, number, number]> = [
      ['sm', 'sm', 4.5, 4.5], ['sm', 'lg', 4.5, 11.25], ['med', 'med', 6.75, 6.75],
      ['lg', 'sm', 11.25, 4.5], ['lg', 'med', 11.25, 6.75], ['lg', 'lg', 11.25, 11.25],
    ];
    for (const [w, len, width, length] of cases) {
      const m = lineEndMetrics(end('triangle', w, len), pt(2.25), SCALE, EMU);
      expect([m.width, m.length]).toEqual([expect.closeTo(width, 6), expect.closeTo(length, 6)]);
    }
    const thick = lineEndMetrics(end('oval', 'lg', 'sm'), pt(9), SCALE, EMU);
    expect([thick.width, thick.length]).toEqual([expect.closeTo(45, 6), expect.closeTo(18, 6)]);
  });

  it('sizes lines thinner than 2 pt, including hairlines, as a 2 pt line', () => {
    for (const width of [0, 0.25, 0.75, 1, 1.5]) {
      const m = lineEndMetrics(end('diamond', 'lg', 'med'), pt(width), SCALE, EMU);
      expect([m.width, m.length]).toEqual([expect.closeTo(10, 6), expect.closeTo(6, 6)]);
    }
    // The floor is absolute: the same 2 pt holds for point-authored strokes.
    const docx = lineEndMetrics(end('triangle', 'sm', 'sm'), { color: '#000', width: 0.75 }, 1, 1);
    expect(docx.length).toBeCloseTo(4, 6);
  });
});

describe('line-end outlines', () => {
  it('triangle: tip on the line end, base one length behind', () => {
    expect(points(end('triangle', 'lg', 'lg'), pt(2.25))).toEqual([[0, 0], [-11.25, -5.625], [-11.25, 5.625]]);
  });

  it('stealth: triangle wings notched on the axis at 1/2/3 units behind the tip', () => {
    expect(points(end('stealth', 'sm', 'sm'), pt(2.25))).toEqual([[0, 0], [-4.5, -2.25], [-2.25, 0], [-4.5, 2.25]]);
    expect(points(end('stealth', 'med', 'med'), pt(9))).toEqual([[0, 0], [-27, -13.5], [-18, 0], [-27, 13.5]]);
    expect(points(end('stealth', 'lg', 'lg'), pt(0.25))).toEqual([[0, 0], [-10, -5], [-6, 0], [-10, 5]]);
  });

  it('diamond and oval are centred on the line end', () => {
    expect(points(end('diamond', 'lg', 'sm'), pt(4.5))).toEqual([[4.5, 0], [0, -11.25], [-4.5, 0], [0, 11.25]]);
    const oval = geometry(end('oval', 'sm', 'lg'), pt(2.25)).path[0];
    expect(oval.op).toBe('ellipse');
    expect(oval.args.slice(0, 4)).toEqual([0, 0, expect.closeTo(5.625, 6), expect.closeTo(2.25, 6)]);
  });

  it('open arrow: line-width arms whose outer miter lands on the line end', () => {
    // lg/lg at 2.25 pt: arm centrelines 4.5 units along, 2.5 across.
    const g = geometry(end('arrow', 'lg', 'lg'), pt(2.25));
    expect(g.paint).toBe('stroke');
    expect(g.stroke).toMatchObject({ lineWidth: 2.25, lineCap: 'round', lineJoin: 'miter' });
    const [[ax, ay], [vx, vy], [bx, by]] = points(end('arrow', 'lg', 'lg'), pt(2.25));
    expect([vy, ax - vx, ay, by]).toEqual([0, expect.closeTo(-10.125, 2), expect.closeTo(-5.625, 3), expect.closeTo(5.625, 3)]);
    // Measured: vertex 2.31 pt behind the tip (1.125 / sin 29.05°).
    expect(vx).toBeCloseTo(-2.316, 2);
    // sm/sm on a thin line: arms use the 2 pt unit but the line's own width.
    const [[tx, ty], [tvx]] = points(end('arrow', 'sm', 'sm'), pt(0.75));
    expect([tx - tvx, ty]).toEqual([expect.closeTo(-4, 6), expect.closeTo(-2.5, 6)]);
    expect(geometry(end('arrow', 'sm', 'sm'), pt(0.75)).stroke.lineWidth).toBeCloseTo(0.75, 6);
  });
});

describe('lineEndRetract — where the line stops behind its end', () => {
  it('matches the measured line ends for every type', () => {
    const s = pt(2.25);
    expect(lineEndRetract(end('triangle', 'sm', 'sm'), s, SCALE, EMU)).toBeCloseTo(3.375, 6);
    expect(lineEndRetract(end('triangle', 'lg', 'lg'), s, SCALE, EMU)).toBeCloseTo(10.125, 6);
    expect(lineEndRetract(end('stealth', 'lg', 'lg'), s, SCALE, EMU)).toBeCloseTo(6.75, 6);
    expect(lineEndRetract(end('arrow', 'sm', 'lg'), s, SCALE, EMU)).toBeCloseTo(4.21, 1);
    expect(lineEndRetract(end('diamond'), s, SCALE, EMU)).toBe(0);
    expect(lineEndRetract(end('oval'), s, SCALE, EMU)).toBe(0);
    expect(lineEndRetract(end('none'), s, SCALE, EMU)).toBe(0);
    // Thin lines retract against the 2 pt unit: triangle sm/lg 10 − 1.
    expect(lineEndRetract(end('triangle', 'sm', 'lg'), pt(0.25), SCALE, EMU)).toBeCloseTo(9, 6);
  });
});

describe('lineEndPaintExtent', () => {
  it('bounds every painted pixel around the tip', () => {
    expect(lineEndPaintExtent(end('triangle', 'lg', 'lg'), pt(2.25), SCALE, EMU)).toBeCloseTo(Math.hypot(11.25, 5.625), 6);
    expect(lineEndPaintExtent(end('oval', 'lg', 'lg'), pt(2.25), SCALE, EMU)).toBeCloseTo(Math.hypot(5.625, 5.625), 6);
    const arrow = lineEndPaintExtent(end('arrow', 'lg', 'lg'), pt(2.25), SCALE, EMU);
    expect(arrow).toBeGreaterThan(Math.hypot(2.316 + 10.125 + 1.125, 5.625 + 1.125) - 0.01);
    expect(lineEndPaintExtent(end('none'), pt(2.25), SCALE, EMU)).toBe(0);
  });
});

describe('trimGeometryPathEnds — stop a curved or custom line short of its ends', () => {
  const path = (...commands: Array<[GeometryPath[number]['op'], number[]]>): GeometryPath =>
    commands.map(([op, args]) => ({ op, args }));

  const ends = (trimmed: GeometryPath) => {
    const first = trimmed[0].args, last = trimmed.at(-1)!.args;
    return [first[0], first[1], last.at(-2)!, last.at(-1)!];
  };

  it('matches straight retraction on a line segment', () => {
    expect(ends(trimGeometryPathEnds(path(['moveTo', [0, 0]], ['lineTo', [10, 0]]), 2, 3)))
      .toEqual([expect.closeTo(2, 9), 0, expect.closeTo(7, 9), 0]);
  });

  it('measures from the end point across short final segments', () => {
    const trimmed = trimGeometryPathEnds(path(['moveTo', [0, 0]], ['lineTo', [1, 0]], ['lineTo', [99, 0]], ['lineTo', [100, 0]]), 4, 4);
    expect(trimmed.map(c => c.op)).toEqual(['moveTo', 'lineTo']);
    expect(ends(trimmed)).toEqual([expect.closeTo(4, 9), 0, expect.closeTo(96, 9), 0]);
  });

  it('cuts a cubic where it leaves the circle around each end, keeping its shape', () => {
    const p = [{ x: 0, y: 0 }, { x: 0, y: 10 }, { x: 10, y: 10 }, { x: 10, y: 0 }];
    const at = (t: number) => {
      const u = 1 - t;
      return { x: u ** 3 * p[0].x + 3 * u * u * t * p[1].x + 3 * u * t * t * p[2].x + t ** 3 * p[3].x,
        y: u ** 3 * p[0].y + 3 * u * u * t * p[1].y + 3 * u * t * t * p[2].y + t ** 3 * p[3].y };
    };
    const onCurve = (x: number, y: number) =>
      Math.min(...Array.from({ length: 2001 }, (_, i) => at(i / 2000)).map(q => Math.hypot(q.x - x, q.y - y)));
    const trimmed = trimGeometryPathEnds(path(['moveTo', [0, 0]], ['bezierCurveTo', [0, 10, 10, 10, 10, 0]]), 2, 4);
    const [sx, sy] = trimmed[0].args, [, , , , ex, ey] = trimmed[1].args;
    expect(Math.hypot(sx, sy)).toBeCloseTo(2, 6);
    expect(Math.hypot(ex - 10, ey)).toBeCloseTo(4, 6);
    expect(onCurve(sx, sy)).toBeLessThan(0.02);
    expect(onCurve(ex, ey)).toBeLessThan(0.02);
    // The kept piece is the same curve: its midpoint stays on the original.
    const [, c1x, c1y, c2x, c2y] = [0, ...trimmed[1].args];
    const mid = { x: (sx + 3 * c1x + 3 * c2x + ex) / 8, y: (sy + 3 * c1y + 3 * c2y + ey) / 8 };
    expect(onCurve(mid.x, mid.y)).toBeLessThan(0.02);
  });

  it('cuts arcs by angle and their straight lead', () => {
    // Clockwise half circle around (5,0) from (0,0) to (10,0), entered by a lead from (0,-20).
    const arc = trimGeometryPathEnds(path(['moveTo', [0, -20]], ['arc', [5, 0, 5, Math.PI, 0, 0]]), 4, 3);
    expect(arc.map(c => c.op)).toEqual(['moveTo', 'arc']);
    const [sx, sy] = arc[0].args, [cx, cy, r, , to] = arc[1].args;
    const [ex, ey] = [cx + r * Math.cos(to), cy + r * Math.sin(to)];
    expect([sx, sy]).toEqual([expect.closeTo(0, 9), expect.closeTo(-16, 9)]);
    expect(Math.hypot(ex - 10, ey)).toBeCloseTo(3, 6);
    expect(Math.hypot(ex - 5, ey)).toBeCloseTo(5, 9);
    expect(ey).toBeLessThan(0);
  });

  it('drops a path both ends consume', () => {
    expect(trimGeometryPathEnds(path(['moveTo', [0, 0]], ['lineTo', [15, 0]]), 10, 10)).toEqual([]);
  });

  it('leaves closed ends alone', () => {
    const closed = path(['moveTo', [0, 0]], ['lineTo', [10, 0]], ['lineTo', [10, 10]], ['closePath', []]);
    expect(trimGeometryPathEnds(closed, 2, 3)).toEqual(closed);
  });
});

describe('retractLineEndpoint — pull a point toward its neighbour by `amount`', () => {
  it('returns the point unchanged when amount is 0', () => {
    expect(retractLineEndpoint({ x: 0, y: 0 }, { x: 10, y: 0 }, 0)).toEqual({ x: 0, y: 0 });
  });

  it('moves the point toward the neighbour by `amount`', () => {
    const p = retractLineEndpoint({ x: 0, y: 0 }, { x: 10, y: 0 }, 3);
    expect(p.x).toBeCloseTo(3, 5);
    expect(p.y).toBeCloseTo(0, 5);
  });

  it('does not overshoot past the neighbour', () => {
    const p = retractLineEndpoint({ x: 0, y: 0 }, { x: 10, y: 0 }, 25);
    expect(p.x).toBeCloseTo(10, 5);
    expect(p.y).toBeCloseTo(0, 5);
  });

  it('works on a diagonal (length 5 → 3-4-5 triangle)', () => {
    const p = retractLineEndpoint({ x: 0, y: 0 }, { x: 3, y: 4 }, 2.5);
    expect(p.x).toBeCloseTo(1.5, 5);
    expect(p.y).toBeCloseTo(2.0, 5);
  });
});

describe('drawArrowHead — open arrow paint', () => {
  it('strokes one mitred path with rounded exposed ends', () => {
    const operations: Array<{ name: string; args: unknown[] }> = [];
    let lineCap: CanvasLineCap = 'butt';
    let lineJoin: CanvasLineJoin = 'miter';
    const ctx = {
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      get lineCap() { return lineCap; },
      set lineCap(value: CanvasLineCap) {
        lineCap = value;
        operations.push({ name: 'lineCap', args: [value] });
      },
      get lineJoin() { return lineJoin; },
      set lineJoin(value: CanvasLineJoin) {
        lineJoin = value;
        operations.push({ name: 'lineJoin', args: [value] });
      },
      save() { operations.push({ name: 'save', args: [] }); },
      restore() { operations.push({ name: 'restore', args: [] }); },
      translate(...args: unknown[]) { operations.push({ name: 'translate', args }); },
      rotate(...args: unknown[]) { operations.push({ name: 'rotate', args }); },
      setLineDash(...args: unknown[]) { operations.push({ name: 'setLineDash', args }); },
      beginPath() { operations.push({ name: 'beginPath', args: [] }); },
      moveTo(...args: unknown[]) { operations.push({ name: 'moveTo', args }); },
      lineTo(...args: unknown[]) { operations.push({ name: 'lineTo', args }); },
      stroke() { operations.push({ name: 'stroke', args: [] }); },
    } as unknown as CanvasRenderingContext2D;

    drawArrowHead(ctx, 20, 30, 0, end('arrow'), pt(2.25), SCALE, EMU);

    expect(operations).toEqual(expect.arrayContaining([
      { name: 'lineCap', args: ['round'] },
      { name: 'lineJoin', args: ['miter'] },
      { name: 'stroke', args: [] },
    ]));
    expect(operations.filter(({ name }) => name === 'moveTo')).toHaveLength(1);
    expect(operations.filter(({ name }) => name === 'lineTo')).toHaveLength(2);
  });

  it('uses the resolved stroke paint for the arrow head', () => {
    const gradient = { addColorStop() {} } as unknown as CanvasGradient;
    const ctx = {
      fillStyle: '',
      strokeStyle: '',
      lineWidth: 1,
      lineCap: 'butt',
      lineJoin: 'miter',
      save() {},
      restore() {},
      translate() {},
      rotate() {},
      setLineDash() {},
      beginPath() {},
      moveTo() {},
      lineTo() {},
      closePath() {},
      fill() {},
    } as unknown as CanvasRenderingContext2D;

    drawArrowHead(ctx, 20, 30, 0, end('triangle'), pt(2.25), SCALE, EMU, gradient);

    expect(ctx.fillStyle).toBe(gradient);
    expect(ctx.strokeStyle).toBe(gradient);
  });
});

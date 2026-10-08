import type { GradientFill, GradientStop } from '../types/common';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';

/** Append the fill-bearing outline; coordinates belong to the supplied box.
 * The caller owns beginPath. Decorative, unfilled paths must be excluded. */
export type FillOutline = (
  ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number,
) => void;

export interface ShadeBox { x: number; y: number; w: number; h: number }
type Point = [number, number];
type Matrix = [number, number, number, number, number, number];

// Resource policy, not an OOXML rule: at most 512² raster pixels. Exact
// row intersections of bilinear bands avoid probing their bounding boxes.
// For E flattened edges, R <= 512 rows and P <= 512² pixels, geometry work is
// O(N + E R (log(E + 1) + alpha(P)) + P alpha(P)):
// focus crossings are sorted per row,
// N is input outline work; E <= 32768 after the explicit support budget.
// Each band has constant-degree row extrema, and successor links remove
// assigned pixels. Geometry memory is O(N + P); no shade geometry is
// clustered or cached. Stop sorting/ramp construction is separate.
// This is an inherent additive cost over main's native point-focus brush,
// not a constant-factor latency promise. A sequential Node/Skia resolve-only
// probe at 512px with 32768 alternating-radius edges measured medians of
// 0.14s (center), 0.35s (corner), 0.45s (outside-area focus) after replacing
// rational extrema with shared-edge scans for convex bands. The previous
// renderer measured 1.00s/2.20s/2.25s in the same three-run comparison.
// Hardware-dependent times include outline traversal and raster generation,
// not the subsequent paint; folded bands still require rational extrema.
// Work guards cover the maximum supported edge count and reject larger inputs
// before raster allocation. Tiled paints retain the native brush entirely.
const MAX_EDGE = 512;
const EDGE_LIMIT = 32768;
const MARGIN = 2;
const BEZIER_SEGMENTS = 16;
const ARC_STEP = Math.PI / 32;

/** Internal deterministic work accounting, used by resource regression tests. */
export interface ShadeWork { edgeRows: number; solves: number; rejected: number; pixels: number }

const apply = (m: Matrix, x: number, y: number): Point =>
  [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const multiply = (m: Matrix, n: Matrix): Matrix => [
  m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
  m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
  m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
];
function invert(m: Matrix): Matrix | null {
  const det = m[0] * m[3] - m[1] * m[2];
  const scale = Math.max(Math.abs(m[0]), Math.abs(m[1]), Math.abs(m[2]), Math.abs(m[3]));
  if (!Number.isFinite(det) || !(Math.abs(det) > scale * scale * 1e-12)) return null;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det,
    (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}

/** Flatten the fill outline into closed polygons (fill subpaths are closed
 * implicitly). The recorder tracks the outline's own transforms and never
 * forwards drawing or state calls to the caller's context. */
export function flattenFillOutline(
  outline: FillOutline, ctx: CanvasRenderingContext2D, box: ShadeBox,
): Point[][] {
  const polygons: Point[][] = [];
  let current: Point[] = [];
  let matrix: Matrix = [1, 0, 0, 1, 0, 0];
  const stack: Matrix[] = [];
  const finish = () => {
    // Duplicate joins and straight subdivisions carry no fill geometry. Remove
    // only exact, forward collinearity (never a reversing edge): this preserves
    // winding and avoids zero-length/collinear bands, in linear input work.
    const straight = (a: Point, b: Point, c: Point) =>
      (b[0] - a[0]) * (c[1] - b[1]) === (b[1] - a[1]) * (c[0] - b[0])
      && (b[0] - a[0]) * (c[0] - b[0]) + (b[1] - a[1]) * (c[1] - b[1]) >= 0;
    const clean: Point[] = [];
    for (const p of current) {
      const last = clean[clean.length - 1];
      if (last && p[0] === last[0] && p[1] === last[1]) continue;
      while (clean.length >= 2 && straight(clean[clean.length - 2], clean[clean.length - 1], p)) clean.pop();
      clean.push(p);
    }
    if (clean.length > 1 && clean[0][0] === clean[clean.length - 1][0]
      && clean[0][1] === clean[clean.length - 1][1]) clean.pop();
    let start = 0;
    while (clean.length - start >= 3) {
      if (straight(clean[clean.length - 2], clean[clean.length - 1], clean[start])) clean.pop();
      else if (straight(clean[clean.length - 1], clean[start], clean[start + 1])) start++;
      else break;
    }
    if (clean.length - start >= 3) polygons.push(clean.slice(start));
    current = [];
  };
  const lastRaw = (): Point | undefined => {
    const last = current[current.length - 1];
    const inverse = last ? invert(matrix) : null;
    return last && inverse ? apply(inverse, last[0], last[1]) : undefined;
  };
  const point = (x: number, y: number) => {
    if (Number.isFinite(x) && Number.isFinite(y)) current.push(apply(matrix, x, y));
  };
  const ellipse = (cx: number, cy: number, rx: number, ry: number, rotation: number,
    start: number, end: number, ccw: boolean) => {
    const full = Math.PI * 2;
    let sweep = end - start;
    if (!ccw && sweep >= full) sweep = full;
    else if (ccw && -sweep >= full) sweep = -full;
    else if (!ccw) { sweep %= full; if (sweep < 0) sweep += full; }
    else { sweep %= full; if (sweep > 0) sweep -= full; }
    const steps = Math.max(1, Math.ceil(Math.abs(sweep) / ARC_STEP));
    const cos = Math.cos(rotation); const sin = Math.sin(rotation);
    for (let i = 0; i <= steps; i++) {
      const angle = start + sweep * i / steps;
      const ex = rx * Math.cos(angle); const ey = ry * Math.sin(angle);
      point(cx + ex * cos - ey * sin, cy + ex * sin + ey * cos);
    }
  };
  const curve = (order: 2 | 3, c: number[]) => {
    const p0 = lastRaw() ?? [c[0], c[1]];
    if (current.length === 0) point(c[0], c[1]);
    for (let i = 1; i <= BEZIER_SEGMENTS; i++) {
      const t = i / BEZIER_SEGMENTS; const u = 1 - t;
      if (order === 2) {
        point(u * u * p0[0] + 2 * u * t * c[0] + t * t * c[2],
          u * u * p0[1] + 2 * u * t * c[1] + t * t * c[3]);
      } else {
        point(u * u * u * p0[0] + 3 * u * u * t * c[0] + 3 * u * t * t * c[2] + t * t * t * c[4],
          u * u * u * p0[1] + 3 * u * u * t * c[1] + 3 * u * t * t * c[3] + t * t * t * c[5]);
      }
    }
  };
  const roundRect = (x: number, y: number, w: number, h: number, radii?: unknown) => {
    const list = Array.isArray(radii) ? radii : [radii ?? 0];
    const radius = (value: unknown): number => {
      const r = typeof value === 'number' ? value
        : typeof value === 'object' && value !== null ? Number((value as { x?: number }).x ?? 0) : 0;
      return Math.max(0, Math.min(Math.abs(w) / 2, Math.abs(h) / 2, Number.isFinite(r) ? r : 0));
    };
    const [tl, tr, br, bl] = list.length === 1 ? [list[0], list[0], list[0], list[0]]
      : list.length === 2 ? [list[0], list[1], list[0], list[1]]
        : list.length === 3 ? [list[0], list[1], list[2], list[1]] : list;
    finish();
    const corners: Array<[number, number, number, number]> = [
      [x + w - radius(tr), y + radius(tr), radius(tr), -Math.PI / 2],
      [x + w - radius(br), y + h - radius(br), radius(br), 0],
      [x + radius(bl), y + h - radius(bl), radius(bl), Math.PI / 2],
      [x + radius(tl), y + radius(tl), radius(tl), Math.PI],
    ];
    for (const [cx, cy, r, start] of corners) ellipse(cx, cy, r, r, 0, start, start + Math.PI / 2, false);
    finish();
    point(x, y);
  };
  const methods: Record<string, (...a: number[]) => void> = {
    beginPath: () => { polygons.length = 0; current = []; },
    moveTo: (x, y) => { finish(); point(x, y); },
    lineTo: (x, y) => point(x, y),
    closePath: () => { const first = current[0]; finish(); if (first) current.push(first); },
    bezierCurveTo: (...c) => curve(3, c),
    quadraticCurveTo: (...c) => curve(2, c),
    // Not used by DrawingML geometry builders; keeping the corner point keeps
    // the flattened outline conservative rather than dropping the segment.
    arcTo: (x1, y1) => point(x1, y1),
    arc: (cx, cy, r, start, end, ccw) => ellipse(cx, cy, r, r, 0, start, end, Boolean(ccw)),
    ellipse: (cx, cy, rx, ry, rotation, start, end, ccw) =>
      ellipse(cx, cy, rx, ry, rotation, start, end, Boolean(ccw)),
    rect: (x, y, w, h) => {
      finish(); point(x, y); point(x + w, y); point(x + w, y + h); point(x, y + h); finish(); point(x, y);
    },
    roundRect: roundRect as unknown as (...a: number[]) => void,
    save: () => { stack.push(matrix); },
    restore: () => { matrix = stack.pop() ?? matrix; },
    translate: (x, y) => { matrix = multiply(matrix, [1, 0, 0, 1, x, y]); },
    scale: (x, y) => { matrix = multiply(matrix, [x, 0, 0, y, 0, 0]); },
    rotate: (angle) => {
      const c = Math.cos(angle); const s = Math.sin(angle);
      matrix = multiply(matrix, [c, s, -s, c, 0, 0]);
    },
    transform: (a, b, c, d, e, f) => { matrix = multiply(matrix, [a, b, c, d, e, f]); },
    setTransform: (a, b, c, d, e, f) => { matrix = [a, b, c, d, e, f]; },
    resetTransform: () => { matrix = [1, 0, 0, 1, 0, 0]; },
  };
  const recorder = new Proxy(ctx, {
    get(target, property) {
      if (typeof property === 'string' && Object.prototype.hasOwnProperty.call(methods, property)) return methods[property];
      const value = Reflect.get(target, property);
      return typeof value === 'function' ? () => undefined : value;
    },
    set() { return true; },
  });
  outline(recorder, box.x, box.y, box.w, box.h);
  finish();
  return polygons;
}

/** Strictly star-shaped about `center`: every edge turns the same way around
 * it and the outline winds exactly once (one region, center strictly inside
 * its kernel). Holes, several loops and centers on the outline fail. */
export function isStrictlyStarShaped(polygons: Point[][], center: Point, box: ShadeBox): boolean {
  const size = Math.hypot(box.w, box.h);
  if (!(size > 0)) return false;
  const epsilon = size * 1e-9;
  let sign = 0; let turning = 0;
  for (const polygon of polygons) {
    for (let i = 0; i < polygon.length; i++) {
      const a = polygon[i]; const b = polygon[(i + 1) % polygon.length];
      const length = Math.hypot(b[0] - a[0], b[1] - a[1]);
      // Arc joins leave floating-point slivers; they carry no direction.
      if (length <= epsilon) continue;
      const ax = a[0] - center[0]; const ay = a[1] - center[1];
      const bx = b[0] - center[0]; const by = b[1] - center[1];
      const cross = ax * by - ay * bx;
      // The center must not lie on any edge's line (distance cross/length).
      if (Math.abs(cross) / length <= epsilon) return false;
      if (sign !== 0 && Math.sign(cross) !== sign) return false;
      sign = Math.sign(cross);
      turning += Math.atan2(cross, ax * bx + ay * by);
    }
  }
  return sign !== 0 && Math.abs(Math.abs(turning) - Math.PI * 2) < 1e-6;
}

/** Crossed focus edges are outside the measured area/segment/point model.
 * ECMA-376 §20.1.8.31 and [MS-OE376] §2.1.1377 define offsets and inscribed
 * regions, but do not determine a replacement region for negative extents.
 * Preserve the previous native midpoint approximation for this input class;
 * choosing a collapse point requires new Office boundary evidence (#1599). */
export function hasInvertedPathShadeFocus(fill: GradientFill): boolean {
  const rect = fill.fillToRect;
  return !!rect && ((rect.l ?? 0) + (rect.r ?? 0) > 1 || (rect.t ?? 0) + (rect.b ?? 0) > 1);
}

/** Authored focus rectangle, normalized to its frame. Keep signed extents:
 * the raster support gate rejects inversion rather than inventing a collapse.
 * Omitted fillToRect uses the measured frame center point. */
export function pathShadeFocusRect(fill: GradientFill): { origin: Point; size: Point } {
  const rect = fill.fillToRect;
  if (!rect) return { origin: [0.5, 0.5], size: [0, 0] };
  return { origin: [rect.l ?? 0, rect.t ?? 0],
    size: [1 - ((rect.l ?? 0) + (rect.r ?? 0)), 1 - ((rect.t ?? 0) + (rect.b ?? 0))] };
}

const rectangle = (box: ShadeBox): Point[] =>
  [[box.x, box.y], [box.x + box.w, box.y], [box.x + box.w, box.y + box.h], [box.x, box.y + box.h]];

function boundsOf(points: Point[]): ShadeBox {
  let left = Infinity; let top = Infinity; let right = -Infinity; let bottom = -Infinity;
  for (const [x, y] of points) {
    left = Math.min(left, x); top = Math.min(top, y);
    right = Math.max(right, x); bottom = Math.max(bottom, y);
  }
  return { x: left, y: top, w: right - left, h: bottom - top };
}

/** Linear stop lookup sampled from a native Canvas ramp, so interpolation
 * (including translucent stops) matches the linear and circle gradients. */
function colorTable(ctx: CanvasRenderingContext2D, stops: readonly GradientStop[]): Uint8ClampedArray | null {
  const ramp = createAuxCanvasForContext(ctx, 1024, 1);
  const rampCtx = ramp?.getContext('2d') as CanvasRenderingContext2D | null | undefined;
  if (!ramp || !rampCtx) return null;
  const gradient = rampCtx.createLinearGradient(0, 0, 1024, 0);
  for (const stop of stops) {
    const hex = stop.color.replace(/^#/, '');
    gradient.addColorStop(Math.max(0, Math.min(1, stop.position)), `#${hex}`);
  }
  rampCtx.fillStyle = gradient;
  rampCtx.fillRect(0, 0, 1024, 1);
  return rampCtx.getImageData(0, 0, 1024, 1).data;
}

function rgba(color: string): [number, number, number, number] {
  const hex = color.replace(/^#/, '');
  const value = (offset: number) => Number.parseInt(hex.slice(offset, offset + 2), 16);
  return [value(0), value(2), value(4), hex.length >= 8 ? value(6) : 255];
}

/** Gradient position of point (u, v) in the band of edge a→b, or -1. The band
 * joins the edge to its image in the focus rectangle; its contour at position
 * s is the edge under M_s(q) = (s + (1-s)k)∘q + (1-s)o. Solved exactly. */
function bandPosition(u: number, v: number, ax: number, ay: number, ex: number, ey: number,
  ox: number, oy: number, kx: number, ky: number): number {
  // Per axis: (a + t e)·g(s) = p - h(s), g = k + s(1-k), h = (1-s)o.
  const x0 = u - ox - ax * kx; const x1 = ox - ax * (1 - kx);
  const y0 = v - oy - ay * ky; const y1 = oy - ay * (1 - ky);
  const gx0 = kx; const gx1 = 1 - kx; const gy0 = ky; const gy1 = 1 - ky;
  const c2 = ey * x1 * gy1 - ex * y1 * gx1;
  const c1 = ey * (x0 * gy1 + x1 * gy0) - ex * (y0 * gx1 + y1 * gx0);
  const c0 = ey * x0 * gy0 - ex * y0 * gx0;
  const roots: number[] = [];
  if (Math.abs(c2) < 1e-12) {
    if (Math.abs(c1) > 1e-12) roots.push(-c0 / c1);
  } else {
    const disc = c1 * c1 - 4 * c2 * c0;
    if (disc >= 0) {
      const root = Math.sqrt(disc);
      roots.push((-c1 - root) / (2 * c2), (-c1 + root) / (2 * c2));
    }
  }
  let best = -1;
  const length2 = ex * ex + ey * ey;
  // A degenerate (segment or point) focus: the s = 0 contour is the edge's
  // image A(a + t e) itself, where the general solve divides by zero.
  if (kx <= 1e-12 || ky <= 1e-12) {
    const along = kx > 1e-12 && Math.abs(ex) > 1e-12 ? ((u - ox) / kx - ax) / ex
      : ky > 1e-12 && Math.abs(ey) > 1e-12 ? ((v - oy) / ky - ay) / ey : 0;
    if (along >= -1e-9 && along <= 1 + 1e-9
      && Math.abs(ox + kx * (ax + along * ex) - u) <= 1e-9
      && Math.abs(oy + ky * (ay + along * ey) - v) <= 1e-9) return 0;
  }
  for (const s of roots) {
    if (!(s >= -1e-9 && s <= 1 + 1e-9)) continue;
    const gx = gx0 + gx1 * s; const gy = gy0 + gy1 * s;
    if (gx <= 1e-12 || gy <= 1e-12) continue;
    const qx = (u - (1 - s) * ox) / gx; const qy = (v - (1 - s) * oy) / gy;
    const t = ((qx - ax) * ex + (qy - ay) * ey) / length2;
    if (t < -1e-9 || t > 1 + 1e-9) continue;
    const clamped = Math.min(1, Math.max(0, s));
    if (best < 0 || clamped < best) best = clamped;
  }
  return best;
}

/** Shared edges must produce the identical row coordinate in either path
 * direction. Closed x intervals then share pixel-center seams; successor sets
 * give each pixel to the last band exactly once, without a raster tolerance. */
function rowCrossing(p: Point, q: Point, y: number): number {
  if (p[1] > q[1]) [p, q] = [q, p];
  if (y === p[1]) return p[0];
  if (y === q[1]) return q[0];
  return p[0] + (y - p[1]) / (q[1] - p[1]) * (q[0] - p[0]);
}

/** A bilinear band with a convex boundary fills that polygon. Point foci
 * reduce to triangles. Only folded/concave bands need rational extrema. */
function convexBand(corners: Point[]): Point[] | null {
  const polygon = corners.filter((p, i) => {
    const q = corners[(i + 1) % corners.length];
    return p[0] !== q[0] || p[1] !== q[1];
  });
  let sign = 0;
  for (let i = 0; i < polygon.length; i++) {
    const a = polygon[i]; const b = polygon[(i + 1) % polygon.length]; const c = polygon[(i + 2) % polygon.length];
    const cross = (b[0] - a[0]) * (c[1] - b[1]) - (b[1] - a[1]) * (c[0] - b[0]);
    if (cross === 0) continue;
    if (sign && Math.sign(cross) !== sign) return null;
    sign = Math.sign(cross);
  }
  return sign ? polygon : [];
}

function convexRowSpan(polygon: Point[], y: number): Array<[number, number]> {
  let low = Infinity; let high = -Infinity;
  for (let i = 0; i < polygon.length; i++) {
    const p = polygon[i]; const q = polygon[(i + 1) % polygon.length];
    if (y < Math.min(p[1], q[1]) || y > Math.max(p[1], q[1])) continue;
    if (p[1] === q[1]) {
      low = Math.min(low, p[0], q[0]); high = Math.max(high, p[0], q[0]);
    } else {
      const x = rowCrossing(p, q, y);
      low = Math.min(low, x); high = Math.max(high, x);
    }
  }
  return low <= high ? [[low, high]] : [];
}

/** Exact horizontal range(s) of a bilinear edge band. Its four corners
 * are A(a), A(b), b, a. For contour s, endpoints a(s), b(s) are linear.
 * Intersecting y gives x(s) = quadratic / linear. Split at endpoint/row
 * crossings and the denominator zero; extrema are endpoints or roots of
 * the quadratic derivative. This includes folded bands' interior envelope,
 * which a quadrilateral bounding-box/edge crossing scan misses. */
export function bandRowSpans(corners: Point[], y: number): Array<[number, number]> {
  const [a, b, outerB, outerA] = corners;
  const ax = outerA[0] - a[0]; const ay = outerA[1] - a[1];
  const bx = b[0] - a[0]; const by = b[1] - a[1];
  const dx = outerB[0] - outerA[0] - bx; const dy = outerB[1] - outerA[1] - by;
  const cy = y - a[1];
  if (by === 0 && dy === 0 && ay !== 0) {
    const s = cy / ay;
    if (s < 0 || s > 1) return [];
    const x = rowCrossing(a, outerA, y); const end = rowCrossing(b, outerB, y);
    return [[Math.min(x, end), Math.max(x, end)]];
  }
  // Carry the endpoint coordinate with each cut. Re-evaluating x(s) as a
  // quadratic/linear quotient introduces cancellation on shared connectors,
  // allowing both neighbours to round a pixel center out of their spans.
  const cuts: Array<{ s: number; x?: number }> = [{ s: 0 }, { s: 1 }];
  for (const [p, q] of [[a, outerA], [b, outerB]]) {
    if (p[1] === q[1]) continue;
    const s = (y - p[1]) / (q[1] - p[1]);
    if (s > 0 && s < 1) cuts.push({ s, x: rowCrossing(p, q, y) });
  }
  if (dy !== 0) {
    const s = -by / dy;
    if (s > 0 && s < 1) cuts.push({ s });
  }
  cuts.sort((m, n) => m.s - n.s);
  const n0 = a[0] * by + bx * cy;
  const n1 = a[0] * dy + ax * by + dx * cy - bx * ay;
  const n2 = ax * dy - dx * ay;
  const extrema: number[] = [];
  const c2 = n2 * dy; const c1 = 2 * n2 * by; const c0 = n1 * by - n0 * dy;
  if (c2 === 0) { if (c1 !== 0) extrema.push(-c0 / c1); }
  else {
    const disc = c1 * c1 - 4 * c2 * c0;
    if (disc >= 0) { const r = Math.sqrt(disc); extrema.push((-c1 - r) / (2 * c2), (-c1 + r) / (2 * c2)); }
  }
  const at = (s: number): number => {
    const denominator = by + dy * s;
    if (Math.abs(denominator) > 1e-12) return (n0 + s * (n1 + s * n2)) / denominator;
    // Removable singularity at a collapsed contour; use the derivative ratio.
    return dy !== 0 ? (n1 + 2 * n2 * s) / dy : a[0] + ax * s;
  };
  const spans: Array<[number, number]> = [];
  for (let i = 0; i < cuts.length - 1; i++) {
    const lo = cuts[i].s; const hi = cuts[i + 1].s; const mid = (lo + hi) / 2;
    const denominator = by + dy * mid;
    if (Math.abs(denominator) <= 1e-12) {
      if (Math.abs(cy - ay * mid) <= 1e-12) {
        const xs = [lo, hi].flatMap(s => [a[0] + ax * s, a[0] + ax * s + bx + dx * s]);
        spans.push([Math.min(...xs), Math.max(...xs)]);
      }
      continue;
    }
    const t = (cy - ay * mid) / denominator;
    if (t < 0 || t > 1) continue;
    const endpoint = (cut: { s: number; x?: number }) => {
      if (cut.x !== undefined) return cut.x;
      if (cut.s === 0 && a[1] !== b[1]) return rowCrossing(a, b, y);
      if (cut.s === 1 && outerA[1] !== outerB[1]) return rowCrossing(outerA, outerB, y);
      return at(cut.s);
    };
    const xs = [endpoint(cuts[i]), endpoint(cuts[i + 1]),
      ...extrema.filter(s => s > lo && s < hi).map(at)];
    spans.push([Math.min(...xs), Math.max(...xs)]);
  }
  return spans;
}

/**
 * Untiled path shade for `path="rect"` and `path="shape"` (ECMA-376 §20.1.8.46).
 *
 * Normative/documented model: the center shade (first stop) fills the focus
 * rectangle — in Office, the gradient path inscribed in fillToRect relative to
 * the shape box ([MS-OE376] §2.1.1377 a/b) — and the shade runs to the last
 * stop at the outline. Contour s is the outline under the affine blend
 * M_s = s·I + (1-s)·A, A mapping the frame onto the focus rectangle, so the
 * band between each outline edge and its inscribed image is shaded exactly.
 * `path="rect"` uses the frame rectangle (the inscribed copy is fillToRect).
 *
 * Office observation (#1599, PowerPoint 16.113 screen/print preview and its
 * point-focus rasters): an explicit all-zero fillToRect is solid first stop;
 * an inset area is flat; supported shape point foci shade toward the point.
 * For degenerate point foci, measured export bands paint in path order,
 * later bands winning. For area foci the complete inscribed copy is painted
 * last: point-focus export order cannot override the center-shade contract.
 * Rect point foci and one-axis segment foci retain native paint (see the
 * support gate below). The center-kernel predicate applies to both path types
 * and is a library support boundary, not an Office brush classifier.
 * Unsupported outlines keep the previous path-specific native resolver;
 * adjustment/aspect/topology boundaries lack live Office evidence (#1599).
 * Rotation framing also keeps the previous local frame pending live evidence.
 * Returns null when the host cannot allocate the shade raster.
 */
export function resolvePathShade(
  fill: GradientFill, ctx: CanvasRenderingContext2D,
  frame: ShadeBox, shapeBox: ShadeBox, outline?: FillOutline,
  paintBounds?: ShadeBox,
  work?: ShadeWork,
): CanvasPattern | null {
  if (hasInvertedPathShadeFocus(fill)) return null;
  const focus = pathShadeFocusRect(fill);
  // Compatibility support policy (#1599), not an inferred Office formula:
  // rect point/default foci mapped to box isolines regress centered, quarter
  // and corner foci on plus/star outlines and translucent ellipse controls.
  // ECMA-376 §20.1.8.31 specifies a point focus but does not settle the band
  // mapping for those hosts. Retain the previous native brush for ALL rect
  // point foci rather than fitting individual outlines or stop colours.
  // A one-axis segment has a singular affine image: collinear bands can
  // compete at its collapsed contour. The area evidence does not establish
  // segment shading (rect/roundRect shape paths and plus/star rect paths).
  // Keep the previous midpoint brush for both segment orientations until
  // Office boundary evidence settles the mapping; no partial segment shader.
  const [fw, fh] = focus.size;
  if ((fw === 0) !== (fh === 0) || (fill.path === 'rect' && fw === 0 && fh === 0)) return null;
  if (![frame, shapeBox].every(box => [box.x, box.y, box.w, box.h].every(Number.isFinite)
    && box.w > 0 && box.h > 0)) return null;
  if (typeof ctx.getTransform !== 'function' || fill.stops.length === 0) return null;
  const outlinePolygons = outline ? flattenFillOutline(outline, ctx, shapeBox) : [];
  // Explicit resource policy: above this edge budget retain main's native
  // circle approximation. Never cluster geometry or silently alter its shape.
  if (outlinePolygons.reduce((n, polygon) => n + polygon.length, 0) > EDGE_LIMIT) return null;
  const coverage = outlinePolygons.length > 0 ? outlinePolygons : [rectangle(shapeBox)];
  // Host topology constrains rect shades too: rect only changes the field,
  // not the evidence boundary. Non-kernel outlines (arrows, chevrons, holes,
  // L-shaped custom paths) retain main's bytes. Live right-arrow captures do
  // not establish the rule for other focus/aspect/transform combinations;
  // all fallback hosts remain out of scope, with no shape-name exceptions.
  if (!isStrictlyStarShaped(
    coverage, [shapeBox.x + shapeBox.w / 2, shapeBox.y + shapeBox.h / 2], shapeBox,
  )) return null;

  const t = ctx.getTransform();
  const user: Matrix = [t.a, t.b, t.c, t.d, t.e, t.f];
  const userInverse = user.every(Number.isFinite) ? invert(user) : null;
  if (!userInverse) return null;
  const coverageField = coverage;
  const fieldFrame = frame;
  if (!(fieldFrame.w > 0 && fieldFrame.h > 0)) return null;

  // Coverage is supplied in device space from the actual paint geometry.
  // The authored shade frame stays separate from these raster bounds.
  const painted = boundsOf([...coverageField.flat(), ...rectangle(fieldFrame)]
    .map(p => apply(user, p[0], p[1])));
  const device = paintBounds ? boundsOf([...rectangle(painted), ...rectangle(paintBounds)]) : painted;
  if (![device.x, device.y, device.w, device.h].every(Number.isFinite)) return null;
  const k = Math.min(1, (MAX_EDGE - 2 * MARGIN - 1) / Math.max(device.w, 1e-9),
    (MAX_EDGE - 2 * MARGIN - 1) / Math.max(device.h, 1e-9));
  const originX = k === 1 ? Math.floor(device.x) : device.x;
  const originY = k === 1 ? Math.floor(device.y) : device.y;
  const inner = MAX_EDGE - 2 * MARGIN;
  const bw = Math.max(1, Math.min(inner, Math.ceil((device.x + device.w - originX) * k))) + 2 * MARGIN;
  const bh = Math.max(1, Math.min(inner, Math.ceil((device.y + device.h - originY) * k))) + 2 * MARGIN;
  const auxToDevice: Matrix = [1 / k, 0, 0, 1 / k, originX - MARGIN / k, originY - MARGIN / k];
  const deviceToAux = invert(auxToDevice);
  if (!deviceToAux) return null;
  const fieldToAux = multiply(deviceToAux, user);
  const normToField: Matrix = [fieldFrame.w, 0, 0, fieldFrame.h, fieldFrame.x, fieldFrame.y];
  const normToAux = multiply(fieldToAux, normToField);
  const auxToNorm = invert(normToAux);
  if (!auxToNorm) return null;

  const surface = createAuxCanvasForContext(ctx, bw, bh);
  const target = surface?.getContext('2d') as CanvasRenderingContext2D | null | undefined;
  const stops = [...fill.stops].sort((a, b) => a.position - b.position);
  const ramp = colorTable(ctx, stops);
  if (!surface || !target || !ramp || typeof target.createImageData !== 'function') return null;
  const first = rgba(stops[0].color);
  const last = rgba(stops[stops.length - 1].color);
  // Keep the solver's double precision: Float32 can round s<1 to exactly 1
  // near a thin band's outline, turning an interior pixel into the final stop
  // (visible with a discontinuous final stop). This adds at most 512²·4 bytes.
  const shade = new Float64Array(bw * bh).fill(Number.NaN);

  if (work) { work.edgeRows = 0; work.solves = 0; work.rejected = 0; work.pixels = bw * bh; }
  {
    const { origin: [ox, oy], size: [kx, ky] } = pathShadeFocusRect(fill);
    const toNorm = (p: Point): Point => [(p[0] - fieldFrame.x) / fieldFrame.w, (p[1] - fieldFrame.y) / fieldFrame.h];
    const polygons = fill.path === 'shape' ? coverageField.map(polygon => polygon.map(toNorm)) : [rectangle({ x: 0, y: 0, w: 1, h: 1 })];
    let solveCount = 0;
    // Row-wise successor sets, with rank + path compression for the stated
    // amortized bound. The rightmost member is the next unassigned column.
    const next = new Int32Array(bh * (bw + 1));
    const rightmost = new Int32Array(next.length);
    const rank = new Uint8Array(next.length);
    for (let row = 0; row < bh; row++) for (let col = 0; col <= bw; col++) next[row * (bw + 1) + col] = rightmost[row * (bw + 1) + col] = col;
    const rootOf = (row: number, col: number): number => {
      const base = row * (bw + 1);
      let root = col;
      while (next[base + root] !== root) root = next[base + root];
      while (next[base + col] !== root) { const up = next[base + col]; next[base + col] = root; col = up; }
      return root;
    };
    const find = (row: number, col: number): number => rightmost[row * (bw + 1) + rootOf(row, col)];
    const erase = (row: number, col: number) => {
      const base = row * (bw + 1);
      let a = rootOf(row, col); let b = rootOf(row, col + 1);
      if (a === b) return;
      if (rank[base + a] < rank[base + b]) [a, b] = [b, a];
      next[base + b] = a;
      rightmost[base + a] = Math.max(rightmost[base + a], rightmost[base + b]);
      if (rank[base + a] === rank[base + b]) rank[base + a]++;
    };
    const edges: Array<[Point, Point]> = [];
    for (const polygon of polygons) {
      for (let i = 0; i < polygon.length; i++) edges.push([polygon[i], polygon[(i + 1) % polygon.length]]);
    }
    const inset = (p: Point): Point => [ox + kx * p[0], oy + ky * p[1]];
    // Reserve the entire first-stop copy. Bands are clipped to its exterior
    // through the successor sets, equivalent to painting it last even for
    // concave outlines. This also removes zero-area identity-band probes.
    if (kx > 0 && ky > 0) {
      const innerEdges = edges.map(([a, b]) => [inset(a), inset(b)].map(p => apply(normToAux, p[0], p[1])));
      for (let row = 0; row < bh; row++) {
        const y = row + .5;
        const crossings: Array<[number, number]> = [];
        for (const [p, q] of innerEdges) {
          if ((p[1] <= y) !== (q[1] <= y)) {
            crossings.push([rowCrossing(p, q, y), q[1] > p[1] ? 1 : -1]);
          }
        }
        crossings.sort((m, n) => m[0] - n[0]);
        let winding = 0;
        for (let i = 0; i < crossings.length - 1; i++) {
          winding += crossings[i][1];
          if (winding === 0) continue;
          const start = Math.max(0, Math.ceil(crossings[i][0] - .5));
          const end = Math.min(bw - 1, Math.floor(crossings[i + 1][0] - .5));
          if (start > end) continue;
          for (let col = find(row, start); col <= end; col = find(row, col)) {
            shade[row * bw + col] = 0;
            erase(row, col);
          }
        }
      }
    }
    // Later bands win: visit bands in reverse path order, first visit fixes a pixel.
    for (let index = edges.length - 1; index >= 0; index--) {
      const [a, b] = edges[index];
      const ex = b[0] - a[0]; const ey = b[1] - a[1];
      if (ex === 0 && ey === 0) continue;
      const corners = [inset(a), inset(b), b, a].map(p => apply(normToAux, p[0], p[1]));
      const convex = convexBand(corners);
      if (convex?.length === 0) continue;
      const ys = corners.map(p => p[1]);
      const rowStart = Math.max(0, Math.ceil(Math.min(...ys) - .5));
      const rowEnd = Math.min(bh - 1, Math.floor(Math.max(...ys) - .5));
      for (let row = rowStart; row <= rowEnd; row++) {
        if (work) work.edgeRows++;
        for (const [low, high] of convex ? convexRowSpan(convex, row + .5) : bandRowSpans(corners, row + .5)) {
          // Both ends are closed and shared crossings are computed identically.
          // The successor set supplies ownership, so rounding cannot leave a seam.
          const colEnd = Math.min(bw - 1, Math.floor(high - .5));
          const colStart = Math.max(0, Math.ceil(low - .5));
          if (colStart > colEnd) continue;
          let col = find(row, colStart);
          while (col <= colEnd) {
            const [u, v] = apply(auxToNorm, col + .5, row + .5);
            // Numerical degeneracies must not reintroduce E × P work. The
            // exact spans ordinarily require no rejections; if roundoff defeats
            // that contract, fail closed to the previous native circle resolver.
            if (solveCount >= bw * bh) return null;
            solveCount++;
            if (work) work.solves++;
            const s = bandPosition(u, v, a[0], a[1], ex, ey, ox, oy, kx, ky);
            if (s >= 0) {
              shade[row * bw + col] = s;
              erase(row, col);
            }
            if (s < 0 && work) work.rejected++;
            col = find(row, col + 1);
          }
        }
      }
    }

  }

  const pixels = target.createImageData(bw, bh);
  const data = pixels.data;
  for (let i = 0, offset = 0; i < shade.length; i++, offset += 4) {
    // Bands and the reserved focus copy cover the outline; only its exterior
    // uses the last stop (including the host stroke/paint coverage margin).
    const s = Number.isNaN(shade[i]) ? 1 : shade[i];
    const color = s <= 0 ? first : s >= 1 ? last : undefined;
    if (color) {
      data[offset] = color[0]; data[offset + 1] = color[1]; data[offset + 2] = color[2]; data[offset + 3] = color[3];
    } else {
      // Canvas samples the ramp at texel centers.
      const entry = Math.min(1023, Math.floor(s * 1024)) * 4;
      data[offset] = ramp[entry]; data[offset + 1] = ramp[entry + 1];
      data[offset + 2] = ramp[entry + 2]; data[offset + 3] = ramp[entry + 3];
    }
  }
  target.putImageData(pixels, 0, 0);
  const pattern = ctx.createPattern(surface, 'no-repeat');
  if (!pattern || typeof pattern.setTransform !== 'function') return null;
  const auxToUser = multiply(userInverse, auxToDevice);
  pattern.setTransform({
    a: auxToUser[0], b: auxToUser[1], c: auxToUser[2], d: auxToUser[3], e: auxToUser[4], f: auxToUser[5],
  });
  return pattern;
}

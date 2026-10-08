import type { ShadeBox } from './path-gradient';

type Point = [number, number];
type Segment = { points: Point[]; start: Point; end: Point; curved: boolean };
type Subpath = { origin: Point; segments: Segment[]; closed: boolean };
type PathState = { paths: Subpath[]; current?: Subpath; pen?: Point };
const states = new WeakMap<CanvasRenderingContext2D, PathState>();
const sources = new WeakMap<CanvasRenderingContext2D, CanvasRenderingContext2D>();
export const paintPathSource = (ctx: CanvasRenderingContext2D): CanvasRenderingContext2D => sources.get(ctx) ?? ctx;
const wrappers = new WeakMap<CanvasRenderingContext2D, CanvasRenderingContext2D>();

/** Observe the actual Canvas path without changing its commands or state.
 * Canvas exposes no current-path bounds API. Keep only the current path,
 * recording device-space control hulls and exact join tangents. Curves use
 * their conservative convex hull; straight segments, caps and miters use
 * the actual stroked geometry. No line-width/decoration-size guess is used. */
export function trackPaintPath(ctx: CanvasRenderingContext2D): CanvasRenderingContext2D {
  if (states.has(ctx)) return ctx;
  const existing = wrappers.get(ctx);
  if (existing) return existing;
  const state: PathState = { paths: [] };
  const transform = () => typeof ctx.getTransform === 'function' ? ctx.getTransform() : { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const point = (x: number, y: number): Point => {
    const m = transform(); return [m.a * x + m.c * y + m.e, m.b * x + m.d * y + m.f];
  };
  const vector = (x: number, y: number): Point => {
    const m = transform(); return [m.a * x + m.c * y, m.b * x + m.d * y];
  };
  const move = (p: Point) => {
    const path: Subpath = { origin: p, segments: [], closed: false };
    state.paths.push(path); state.current = path; state.pen = p;
  };
  const line = (p: Point) => {
    if (!state.pen || !state.current) { move(p); return; }
    const tangent: Point = [p[0] - state.pen[0], p[1] - state.pen[1]];
    state.current.segments.push({ points: [state.pen, p], start: tangent, end: tangent, curved: false });
    state.pen = p;
  };
  const close = () => {
    if (!state.current) return;
    line(state.current.origin); state.current.closed = true;
  };
  const ellipse = (cx: number, cy: number, rx: number, ry: number, rotation: number, start: number, end: number, ccw: boolean) => {
    const c = Math.cos(rotation); const s = Math.sin(rotation);
    const at = (angle: number) => point(cx + rx * Math.cos(angle) * c - ry * Math.sin(angle) * s,
      cy + rx * Math.cos(angle) * s + ry * Math.sin(angle) * c);
    const tangent = (angle: number) => {
      const direction = ccw ? -1 : 1;
      return vector(direction * (-rx * Math.sin(angle) * c - ry * Math.cos(angle) * s),
        direction * (-rx * Math.sin(angle) * s + ry * Math.cos(angle) * c));
    };
    const first = at(start); if (!state.current) move(first); else line(first);
    const center = point(cx, cy); const vx = vector(rx * c, rx * s); const vy = vector(-ry * s, ry * c);
    const hx = Math.hypot(vx[0], vy[0]); const hy = Math.hypot(vx[1], vy[1]);
    state.current?.segments.push({ points: [first, at(end), [center[0] - hx, center[1] - hy], [center[0] + hx, center[1] + hy]],
      start: tangent(start), end: tangent(end), curved: true });
    state.pen = at(end);
  };
  const observe = (name: string, a: number[]) => {
    switch (name) {
      case 'beginPath': state.paths = []; state.current = undefined; state.pen = undefined; break;
      case 'moveTo': move(point(a[0], a[1])); break;
      case 'lineTo': line(point(a[0], a[1])); break;
      case 'closePath': close(); break;
      case 'rect':
      case 'roundRect':
        // A round rectangle is contained by the rectangle; its smooth corners
        // cannot add miters. Keeping the rectangle is conservative for both.
        move(point(a[0], a[1])); line(point(a[0] + a[2], a[1]));
        line(point(a[0] + a[2], a[1] + a[3])); line(point(a[0], a[1] + a[3])); close();
        if (name === 'roundRect' && state.current) {
          for (const segment of state.current.segments) segment.curved = true;
        }
        break;
      case 'quadraticCurveTo':
      case 'bezierCurveTo': {
        const controls: Point[] = [];
        for (let i = 0; i < a.length; i += 2) controls.push(point(a[i], a[i + 1]));
        if (!state.pen) move(controls[0]);
        const points = [state.pen as Point, ...controls]; const last = points.length - 1;
        const tangent = (from: number, step: number): Point => {
          for (let i = from + step; i >= 0 && i <= last; i += step) {
            const v: Point = [(points[i][0] - points[from][0]) * step, (points[i][1] - points[from][1]) * step];
            if (v[0] !== 0 || v[1] !== 0) return v;
          }
          return [0, 0];
        };
        state.current?.segments.push({ points, start: tangent(0, 1), end: tangent(last, -1), curved: true });
        state.pen = points[last]; break;
      }
      case 'arc': ellipse(a[0], a[1], a[2], a[2], 0, a[3], a[4], Boolean(a[5])); break;
      case 'ellipse': ellipse(a[0], a[1], a[2], a[3], a[4], a[5], a[6], Boolean(a[7])); break;
    }
  };
  const proxy = new Proxy(ctx, {
    get(target, property) {
      const value = Reflect.get(target, property, target);
      if (typeof value !== 'function') return value;
      return (...args: number[]) => {
        if (typeof property === 'string') observe(property, args);
        return Reflect.apply(value, target, args);
      };
    },
    set(target, property, value) { return Reflect.set(target, property, value, target); },
  });
  states.set(proxy, state); sources.set(proxy, ctx); wrappers.set(ctx, proxy);
  return proxy;
}

/** Conservative device bounds of the recorded path's stroked outline.
 * Every dash has caps, including interior dashes on closed curves. A square
 * cap's corners are p + r(±t ±n), with unit tangent t and normal n. For a
 * device-axis row v of the affine transform, their support is
 * r(|v·t| + |v·n|) <= r sqrt(2) |v|. This envelopes every possible end tangent
 * on a curved control hull without flattening or approximating dash lengths.
 * Straight segments use their exact tangent. Round caps use circular support;
 * flat caps add no tangent extension. Neither envelope changes the shade frame. */
export function currentStrokeBounds(ctx: CanvasRenderingContext2D): ShadeBox | undefined {
  const state = states.get(ctx);
  if (!state || state.paths.length === 0) return undefined;
  if (typeof ctx.getTransform !== 'function') return undefined;
  const m = ctx.getTransform(); const det = m.a * m.d - m.b * m.c;
  if (!Number.isFinite(det) || det === 0) return undefined;
  const r = ctx.lineWidth / 2;
  const hx = r * Math.hypot(m.a, m.c); const hy = r * Math.hypot(m.b, m.d);
  const dashed = typeof ctx.getLineDash === 'function' && ctx.getLineDash().some(length => length > 0);
  const squareDash = dashed && ctx.lineCap === 'square';
  let left = Infinity; let right = -Infinity; let top = Infinity; let bottom = -Infinity;
  const add = (p: Point, x = 0, y = 0) => {
    left = Math.min(left, p[0] - x); right = Math.max(right, p[0] + x);
    top = Math.min(top, p[1] - y); bottom = Math.max(bottom, p[1] + y);
  };
  const localUnit = (p: Point): Point => {
    const x = (m.d * p[0] - m.c * p[1]) / det; const y = (-m.b * p[0] + m.a * p[1]) / det;
    const length = Math.hypot(x, y); return length > 0 ? [x / length, y / length] : [0, 0];
  };
  const offset = (p: Point, x: number, y: number) => {
    add([p[0] + m.a * x + m.c * y, p[1] + m.b * x + m.d * y]);
  };
  for (const path of state.paths) {
    const segments = path.segments.filter(segment => segment.start[0] !== 0 || segment.start[1] !== 0);
    if (segments.length === 0 && path.segments.length > 0 && !path.closed && ctx.lineCap === 'round') {
      add(path.origin, hx, hy);
    }
    for (const segment of segments) {
      if (segment.curved) {
        for (const p of segment.points) add(p, squareDash ? Math.SQRT2 * hx : hx, squareDash ? Math.SQRT2 * hy : hy);
      }
      else {
        const [x, y] = localUnit(segment.start);
        for (const p of segment.points) {
          offset(p, -y * r, x * r); offset(p, y * r, -x * r);
          if (squareDash) {
            for (const direction of [-1, 1]) {
              offset(p, r * (direction * x - y), r * (direction * y + x));
              offset(p, r * (direction * x + y), r * (direction * y - x));
            }
          } else if (dashed && ctx.lineCap === 'round') add(p, hx, hy);
        }
      }
    }
    for (let i = 0; i < segments.length; i++) {
      const incoming = segments[i]; const outgoing = segments[(i + 1) % segments.length];
      if (!path.closed && i === segments.length - 1) break;
      const p = incoming.points[incoming.points.length - 1];
      if (ctx.lineJoin === 'round') { add(p, hx, hy); continue; }
      if (ctx.lineJoin !== 'miter') continue;
      const a = localUnit(incoming.end); const b = localUnit(outgoing.start);
      const divisor = 1 + a[0] * b[0] + a[1] * b[1];
      if (divisor <= 0 || Math.sqrt(2 / divisor) > ctx.miterLimit) continue;
      const x = -r * (a[1] + b[1]) / divisor; const y = r * (a[0] + b[0]) / divisor;
      offset(p, x, y); offset(p, -x, -y);
    }
    if (!path.closed && segments.length > 0) {
      const first = segments[0]; const last = segments[segments.length - 1];
      for (const [p, tangent, direction] of [[first.points[0], first.start, -1], [last.points[last.points.length - 1], last.end, 1]] as const) {
        if (ctx.lineCap === 'round') add(p, hx, hy);
        else if (ctx.lineCap === 'square') {
          const [x, y] = localUnit(tangent);
          offset(p, r * (direction * x - y), r * (direction * y + x));
          offset(p, r * (direction * x + y), r * (direction * y - x));
        }
      }
    }
  }
  return Number.isFinite(left) ? { x: left, y: top, w: right - left, h: bottom - top } : undefined;
}

/** Public resolver compatibility when no current path was supplied: coverage
 * of a stroked host rectangle, transformed as geometry (including its joins).
 * Custom strokes must pass their recorded current-path bounds. */
export function hostStrokeBounds(ctx: CanvasRenderingContext2D, box: ShadeBox): ShadeBox | undefined {
  if (typeof ctx.getTransform !== 'function') return undefined;
  const m = ctx.getTransform(); const r = (ctx.lineWidth ?? 1) / 2;
  const miter = ctx.lineJoin === 'miter' && ctx.miterLimit >= Math.SQRT2;
  const hx = r * (miter ? Math.abs(m.a) + Math.abs(m.c) : Math.hypot(m.a, m.c));
  const hy = r * (miter ? Math.abs(m.b) + Math.abs(m.d) : Math.hypot(m.b, m.d));
  const points = [[box.x, box.y], [box.x + box.w, box.y], [box.x, box.y + box.h], [box.x + box.w, box.y + box.h]];
  const xs = points.map(([x, y]) => m.a * x + m.c * y + m.e);
  const ys = points.map(([x, y]) => m.b * x + m.d * y + m.f);
  return { x: Math.min(...xs) - hx, y: Math.min(...ys) - hy,
    w: Math.max(...xs) - Math.min(...xs) + 2 * hx, h: Math.max(...ys) - Math.min(...ys) + 2 * hy };
}

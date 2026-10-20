import type { ArrowEnd, Stroke } from '../types/common';
import { createGeometryPath, type GeometryPath, type GeometryPathCommand } from './path-data';
import type { ResolvedStrokeGeometry } from './stroke-geometry';

/** A 2D point in canvas pixels. */
export interface Point {
  x: number;
  y: number;
}

/*
 * DrawingML line-end decorations: `<a:headEnd>` / `<a:tailEnd>`
 * (CT_LineEndProperties, ECMA-376 Part 1 §20.1.8.38 / §20.1.8.57) with
 * ST_LineEndType (§20.1.10.33) and ST_LineEndWidth / ST_LineEndLength
 * (§20.1.10.34 / §20.1.10.32).
 *
 * Normative part: the spec names `sm`/`med`/`lg` only as sizes "relative to the
 * width of the line itself"; it gives neither ratios nor outlines.
 *
 * Observed Office behaviour (PowerPoint, "Best for electronic distribution"
 * PDF export of a synthetic control deck). The deck covers every type ×
 * w sm/med/lg × len sm/med/lg at line widths 0.25, 0.75, 1.5, 2.25, 4.5 and
 * 9 pt, plus hairline (w=0), 0.5 pt and 1 pt. Each case appears as a
 * straightConnector1 head and tail, a flipped `line` tail and a
 * borderCallout1 leader tail. Mixed head≠tail, dashed and diagonal cases were
 * also exported. Every filled outline fitted the rules below within the PDF's
 * 0.01 pt coordinate precision. Open-arrow arm centrelines, derived from the
 * round-cap centres, fitted within 0.05 e. One exporter limitation: it paints any
 * stroke thinner than 1 pt (the line and the open arrow's arms alike) at 1 pt.
 * Outline sizes are unaffected (the 2 pt floor already covers them). As a
 * library policy, the arms keep the line's own painted width, so they always
 * match the shaft, as they do at every measured width of 1 pt and above.
 *
 *  - Base unit `e = max(line width, 2 pt)`. Lines thinner than 2 pt
 *    (including hairlines) get the decoration of a 2 pt line; 1.5 pt is still
 *    floored and 2.25 pt is not.
 *  - Full width across the line: sm 2e, med 3e, lg 5e (`w`). Length along
 *    the line: sm 2e, med 3e, lg 5e (`len`). The two attributes are independent.
 *  - triangle: tip on the line end, base `len` behind it. The line stops
 *    `len - e/2` behind the tip, inside the body.
 *  - stealth: the triangle's tip and wings, notched on the axis at
 *    {sm 1e, med 2e, lg 3e} (by `len`) behind the tip. The line stops at the
 *    notch.
 *  - diamond / oval: centred on the line end with diagonals or axes `len` × `w`.
 *    The line runs to the centre.
 *  - arrow (open): two arms stroked at the line width with round caps and a
 *    miter join. Arm centrelines run {sm 2e, med 3e, lg 4.5e} along (by `len`)
 *    and {sm 1.25e, med 1.75e, lg 2.5e} across (by `w`) from their vertex. The
 *    vertex sits `(lw/2)/sin(θ)` behind the line end, so the outer miter point
 *    lands exactly on it. The line stops at the vertex.
 *  - Line kind (connector, line, callout leader), head vs tail, orientation
 *    and dash style do not change the decoration; decorations are never dashed.
 *
 * The retraction was measured on straight leaders. Paths that are not
 * re-stroked from straight vertices (curved connectors, custom geometry,
 * compound lines) apply the same distance through
 * {@link trimGeometryPathEnds}. This is a library generalisation; it was not
 * separately measured.
 *
 * Positions use a local frame: +x points out of the line through its end at
 * (0, 0), and y spans the line's width.
 */
const SPAN: Readonly<Record<string, number>> = { sm: 2, med: 3, lg: 5 };
const STEALTH_NOTCH: Readonly<Record<string, number>> = { sm: 1, med: 2, lg: 3 };
const OPEN_ARM_ALONG: Readonly<Record<string, number>> = { sm: 2, med: 3, lg: 4.5 };
const OPEN_ARM_ACROSS: Readonly<Record<string, number>> = { sm: 1.25, med: 1.75, lg: 2.5 };
/** Line widths below this (in points) size their decorations as this width. */
const LINE_END_MIN_BASE_PT = 2;

const step = (value: string | undefined): string => (value === 'sm' || value === 'lg' ? value : 'med');

/** Device-space decoration metrics. `lw` is the painted line width; `e` the size unit. */
export interface LineEndMetrics {
  readonly lw: number;
  readonly e: number;
  /** Full span across the line. */
  readonly width: number;
  /** Extent along the line. */
  readonly length: number;
}

/**
 * Resolve a line end's size. `scale` converts stroke-width units to device px;
 * `unitsPerPoint` is the number of stroke-width units in one point (12700 for
 * EMU strokes, 1 for point-authored plans). It is required because the 2 pt
 * floor is absolute.
 */
export function lineEndMetrics(
  arrowEnd: ArrowEnd, stroke: Stroke, scale: number, unitsPerPoint: number,
): LineEndMetrics {
  const lw = Math.max(0.5, stroke.width * scale);
  const e = Math.max(stroke.width, LINE_END_MIN_BASE_PT * unitsPerPoint) * scale;
  return { lw, e, width: SPAN[step(arrowEnd.w)] * e, length: SPAN[step(arrowEnd.len)] * e };
}

/** Open-arrow arm centreline: vertex x (≤ 0), arm end offset `along` behind it, `across` off-axis. */
function openArm(end: ArrowEnd, m: LineEndMetrics): { vertex: number; along: number; across: number } {
  const along = OPEN_ARM_ALONG[step(end.len)] * m.e;
  const across = OPEN_ARM_ACROSS[step(end.w)] * m.e;
  const sin = across / Math.hypot(along, across);
  return { vertex: -(m.lw / 2) / sin, along, across };
}

/**
 * How far (device px) the line is pulled back from its end so it stops where
 * Office stops it (see the table above). Zero for diamond, oval and none.
 */
export function lineEndRetract(arrowEnd: ArrowEnd, stroke: Stroke, scale: number, unitsPerPoint: number): number {
  const m = lineEndMetrics(arrowEnd, stroke, scale, unitsPerPoint);
  switch (arrowEnd.type) {
    case 'triangle': return m.length - m.e / 2;
    case 'stealth': return STEALTH_NOTCH[step(arrowEnd.len)] * m.e;
    case 'arrow': return -openArm(arrowEnd, m).vertex;
    default: return 0;
  }
}

/** Local axis-aligned box of the painted decoration, including the open arrow's pen. */
export function lineEndLocalBounds(
  arrowEnd: ArrowEnd, stroke: Stroke, scale: number, unitsPerPoint: number,
): { readonly x0: number; readonly x1: number; readonly halfY: number } | null {
  if (arrowEnd.type === 'none') return null;
  const m = lineEndMetrics(arrowEnd, stroke, scale, unitsPerPoint);
  switch (arrowEnd.type) {
    case 'diamond': case 'oval': return { x0: -m.length / 2, x1: m.length / 2, halfY: m.width / 2 };
    case 'arrow': {
      const arm = openArm(arrowEnd, m);
      return { x0: arm.vertex - arm.along - m.lw / 2, x1: 0, halfY: arm.across + m.lw / 2 };
    }
    default: return { x0: -m.length, x1: 0, halfY: m.width / 2 };
  }
}

/** Conservative radius around a line-end tip occupied by its painted pixels. */
export function lineEndPaintExtent(arrowEnd: ArrowEnd, stroke: Stroke, scale: number, unitsPerPoint: number): number {
  const b = lineEndLocalBounds(arrowEnd, stroke, scale, unitsPerPoint);
  return b ? Math.hypot(Math.max(-b.x0, b.x1), b.halfY) : 0;
}

/**
 * Pull `p` toward its neighbour `toward` by `amount` px, clamped so it never
 * passes the neighbour. Used to retract a polyline's terminal vertex before
 * stroking, so a decorated end stops at the decoration's base.
 */
export function retractLineEndpoint(p: Point, toward: Point, amount: number): Point {
  if (amount <= 0) return { x: p.x, y: p.y };
  const dx = toward.x - p.x;
  const dy = toward.y - p.y;
  const d = Math.hypot(dx, dy);
  if (d < 1e-9) return { x: p.x, y: p.y };
  const t = Math.min(amount, d) / d;
  return { x: p.x + dx * t, y: p.y + dy * t };
}


export interface ResolvedArrowGeometry {
  readonly path: GeometryPath;
  readonly paint: 'fill' | 'stroke';
  readonly stroke: ResolvedStrokeGeometry;
}
export function resolveArrowGeometry(end: ArrowEnd, stroke: Stroke, scale: number, unitsPerPoint: number,
  charge: (units: number) => void = () => {}): ResolvedArrowGeometry | null {
  if (end.type === 'none') return null;
  const m = lineEndMetrics(end, stroke, scale, unitsPerPoint);
  const { lw, length: len } = m, halfW = m.width / 2;
  const path = createGeometryPath(charge), ctx = path.sink;
  switch (end.type) {
    case 'triangle':
      ctx.moveTo(0, 0); ctx.lineTo(-len, -halfW); ctx.lineTo(-len, halfW); ctx.closePath(); break;
    case 'stealth': {
      const notch = -STEALTH_NOTCH[step(end.len)] * m.e;
      ctx.moveTo(0, 0); ctx.lineTo(-len, -halfW); ctx.lineTo(notch, 0); ctx.lineTo(-len, halfW); ctx.closePath(); break;
    }
    case 'arrow': {
      const arm = openArm(end, m);
      ctx.moveTo(arm.vertex - arm.along, -arm.across); ctx.lineTo(arm.vertex, 0); ctx.lineTo(arm.vertex - arm.along, arm.across);
      break;
    }
    case 'diamond':
      ctx.moveTo(len / 2, 0); ctx.lineTo(0, -halfW); ctx.lineTo(-len / 2, 0); ctx.lineTo(0, halfW); ctx.closePath(); break;
    case 'oval': ctx.ellipse(0, 0, len / 2, halfW, 0, 0, Math.PI * 2); break;
    default: return null;
  }
  const open = end.type === 'arrow';
  return Object.freeze({ path: Object.freeze(path.commands), paint: open ? 'stroke' : 'fill',
    stroke: Object.freeze({ lineWidth: lw, lineCap: open ? 'round' : 'butt',
      lineJoin: 'miter', miterLimit: 10, dash: Object.freeze([]) }) });
}

type Cmd = { op: GeometryPathCommand['op']; args: number[] };
/** One drawn piece of a path: a line/curve command, an arc/ellipse, or the
 *  straight lead canvas draws from the pen to an arc's start point. */
interface TrimSegment {
  /** Index of the command that draws it. */
  readonly index: number;
  at(t: number): Point;
  /** Commands drawing [0, t] with the pen at the segment start. */
  left(t: number): Cmd[];
  /** Commands drawing [t, 1] with the pen at `at(t)`. */
  right(t: number): Cmd[];
}

const lerp = (a: Point, b: Point, t: number): Point => ({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t });
/** De Casteljau split of control polygon `p` at `t`: [0,t] and [t,1] polygons. */
function splitPolygon(p: readonly Point[], t: number): [Point[], Point[]] {
  const left: Point[] = [], right: Point[] = [];
  let q = p.slice();
  while (q.length > 0) {
    left.push(q[0]); right.unshift(q[q.length - 1]);
    q = q.slice(1).map((b, i) => lerp(q[i], b, t));
  }
  return [left, right];
}
const polygonCommand = (op: Cmd['op'], p: readonly Point[]): Cmd => ({ op, args: p.slice(1).flatMap(q => [q.x, q.y]) });
const TAU = Math.PI * 2;
/** Canvas arc/ellipse sweep (signed radians) from its start/end angles and direction. */
function arcSweep(start: number, end: number, ccw: boolean): number {
  const raw = ccw ? start - end : end - start;
  const sweep = raw >= TAU ? TAU : ((raw % TAU) + TAU) % TAU;
  return ccw ? -sweep : sweep;
}

/** Split a path into subpaths of trimmable segments (pen tracked per canvas semantics). */
function trimSegments(commands: readonly Cmd[]): { segments: TrimSegment[]; closed: boolean }[] {
  const subpaths: { segments: TrimSegment[]; closed: boolean }[] = [];
  let pen: Point | null = null, current: { segments: TrimSegment[]; closed: boolean } | null = null;
  const open = () => { current = { segments: [], closed: false }; subpaths.push(current); return current; };
  commands.forEach(({ op, args: a }, index) => {
    switch (op) {
      case 'moveTo': pen = { x: a[0], y: a[1] }; open(); return;
      case 'closePath': if (current) current.closed = true; return;
      case 'rect': case 'roundRect': pen = { x: a[0], y: a[1] }; open().closed = true; return;
      case 'lineTo': case 'quadraticCurveTo': case 'bezierCurveTo': {
        const target = current ?? open();
        const p: Point[] = [pen ?? { x: a[0], y: a[1] }];
        for (let k = 0; k < a.length; k += 2) p.push({ x: a[k], y: a[k + 1] });
        pen = p[p.length - 1];
        target.segments.push({ index,
          at: t => splitPolygon(p, t)[0].at(-1)!,
          left: t => [polygonCommand(op, splitPolygon(p, t)[0])],
          right: t => [polygonCommand(op, splitPolygon(p, t)[1])] });
        return;
      }
      case 'arc': case 'ellipse': {
        const target = current ?? open();
        const ellipse = op === 'ellipse';
        const [cx, cy] = a, rx = a[2], ry = ellipse ? a[3] : a[2], rot = ellipse ? a[4] : 0;
        const a0 = ellipse ? a[5] : a[3], a1 = ellipse ? a[6] : a[4], ccw = Boolean(ellipse ? a[7] : a[5]);
        const sweep = arcSweep(a0, a1, ccw), c = Math.cos(rot), s = Math.sin(rot);
        const point = (angle: number): Point => {
          const ex = rx * Math.cos(angle), ey = ry * Math.sin(angle);
          return { x: cx + ex * c - ey * s, y: cy + ex * s + ey * c };
        };
        const arcCommand = (from: number, to: number): Cmd => ({ op,
          args: ellipse ? [cx, cy, rx, ry, rot, from, to, Number(ccw)] : [cx, cy, rx, from, to, Number(ccw)] });
        const startPoint = point(a0);
        if (pen && Math.hypot(pen.x - startPoint.x, pen.y - startPoint.y) > 1e-9) {
          const from = pen;
          target.segments.push({ index, at: t => lerp(from, startPoint, t),
            left: t => [{ op: 'lineTo', args: [lerp(from, startPoint, t).x, lerp(from, startPoint, t).y] }],
            right: () => [arcCommand(a0, a1)] });
        }
        target.segments.push({ index, at: t => point(a0 + sweep * t),
          left: t => [arcCommand(a0, a0 + sweep * t)],
          right: t => [arcCommand(a0 + sweep * t, a0 + sweep)] });
        pen = point(a0 + sweep);
        return;
      }
    }
  });
  return subpaths;
}

/** Parameter where `segment`, walked from its `from` end, first reaches
 *  distance `d` from `anchor`; null when the whole segment stays inside. */
function reachParameter(segment: TrimSegment, from: 0 | 1, anchor: Point, d: number): number | null {
  const dist = (t: number) => { const q = segment.at(t); return Math.hypot(q.x - anchor.x, q.y - anchor.y); };
  const at = (k: number) => (from === 1 ? 1 - k : k);
  const STEPS = 64;
  let inside = 0;
  for (let i = 0; i <= STEPS; i++) {
    const k = i / STEPS;
    if (dist(at(k)) >= d) {
      if (i === 0) return at(0);
      let lo = inside, hi = k;
      for (let j = 0; j < 50; j++) { const mid = (lo + hi) / 2; if (dist(at(mid)) >= d) hi = mid; else lo = mid; }
      return at(hi);
    }
    inside = k;
  }
  return null;
}

function trimTail(commands: Cmd[], d: number): Cmd[] {
  const last = trimSegments(commands).at(-1);
  if (!last || last.closed || last.segments.length === 0) return commands;
  const anchor = last.segments.at(-1)!.at(1);
  for (let k = last.segments.length - 1; k >= 0; k--) {
    const segment = last.segments[k];
    const t = reachParameter(segment, 1, anchor, d);
    if (t !== null) return [...commands.slice(0, segment.index), ...segment.left(t)];
  }
  // The whole subpath lies within reach: nothing of it is stroked.
  return commands.slice(0, last.segments[0].index);
}

function trimHead(commands: Cmd[], d: number): Cmd[] {
  const first = trimSegments(commands)[0];
  if (!first || first.closed || first.segments.length === 0 || commands[0]?.op !== 'moveTo') return commands;
  const anchor = { x: commands[0].args[0], y: commands[0].args[1] };
  for (const segment of first.segments) {
    const t = reachParameter(segment, 0, anchor, d);
    if (t !== null) {
      const start = segment.at(t);
      return [{ op: 'moveTo', args: [start.x, start.y] }, ...segment.right(t), ...commands.slice(segment.index + 1)];
    }
  }
  return commands.slice(first.segments.at(-1)!.index + 1);
}

/**
 * Shorten a stroked path at its open ends so the line stops `head` / `tail`
 * device px from its first / last point. The distance is measured straight
 * from that point: the path is cut where it last (tail) or first (head)
 * crosses the circle of that radius. Earlier segments that lie wholly inside
 * the circle are dropped. Each kept piece keeps its shape: curves are split
 * with de Casteljau and arcs by angle. On a straight leader this equals
 * {@link retractLineEndpoint}. Use it for paths that cannot be re-stroked
 * from straight vertices (curved connectors, custom geometry, compound
 * lines), so the measured line-end retraction applies to every path kind.
 * Closed subpaths are left alone.
 */
export function trimGeometryPathEnds(path: GeometryPath, head: number, tail: number): GeometryPath {
  if (!(head > 0) && !(tail > 0)) return path;
  let commands: Cmd[] = path.map(c => ({ op: c.op, args: [...c.args] }));
  if (tail > 0) commands = trimTail(commands, tail);
  if (head > 0) commands = trimHead(commands, head);
  return Object.freeze(commands.map(c => Object.freeze({ op: c.op, args: Object.freeze(c.args) })));
}

import { GeometryWorkBudgetError } from '@silurus/ooxml-core';
import type { GeometryPath } from '@silurus/ooxml-core';
import type { PathCmd } from '../types.js';
type Point = readonly [number, number];

/** Acquire bounded output polygons for VML vector effects. Canvas has no
 * projective path transform; use 1/1024pt in output space (numerical precision,
 * independent of samples).
 * Lines project exactly. Curves subdivide adaptively; work is charged against
 * the caller's geometry allowance, with bounded recursion for singular input. */
export function projectPath(path: GeometryPath, project: (p: Point) => Point, charge: () => void): PathCmd[] {
  const out: PathCmd[] = [];
  let pen: Point = [0, 0], start: Point = pen;
  const move = (p: Point) => { const [x, y] = project(p); out.push({ cmd: 'moveTo', x, y }); pen = p; start = p; };
  const line = (p: Point) => { const [x, y] = project(p); out.push({ cmd: 'lineTo', x, y }); pen = p; };
  const curve = (at: (t: number) => Point) => {
    const subdivide = (a: number, b: number, pa: Point, pb: Point, depth: number): void => {
      charge();
      const samples = [0.25, 0.5, 0.75].map(f => project(at(a + (b - a) * f)));
      const error = Math.max(...samples.map((p, i) => {
        const f = (i + 1) / 4;
        return Math.hypot(p[0] - pa[0] - (pb[0] - pa[0]) * f, p[1] - pa[1] - (pb[1] - pa[1]) * f);
      }));
      if (error <= 1 / 1024) { out.push({ cmd: 'lineTo', x: pb[0], y: pb[1] }); return; }
      if (depth >= 24) throw new GeometryWorkBudgetError('VML projected vector cannot meet numerical precision');
      const mid = (a + b) / 2;
      subdivide(a, mid, pa, samples[1], depth + 1);
      subdivide(mid, b, samples[1], pb, depth + 1);
    };
    subdivide(0, 1, project(at(0)), project(at(1)), 0);
    pen = at(1);
  };
  const ellipse = (cx: number, cy: number, rx: number, ry: number, rotation: number, a: number, b: number, ccw: boolean) => {
    let sweep = b - a;
    if (ccw && sweep > 0) sweep = sweep >= 2 * Math.PI ? -2 * Math.PI : sweep - 2 * Math.PI;
    if (!ccw && sweep < 0) sweep = sweep <= -2 * Math.PI ? 2 * Math.PI : sweep + 2 * Math.PI;
    const at = (t: number): Point => {
      const angle = a + sweep * t, vx = rx * Math.cos(angle), vy = ry * Math.sin(angle);
      return [cx + vx * Math.cos(rotation) - vy * Math.sin(rotation), cy + vx * Math.sin(rotation) + vy * Math.cos(rotation)];
    };
    line(at(0)); curve(at);
  };
  for (const { op, args: a } of path) {
    charge();
    switch (op) {
      case 'moveTo': move([a[0], a[1]]); break;
      case 'lineTo': line([a[0], a[1]]); break;
      case 'closePath': out.push({ cmd: 'close' }); pen = start; break;
      case 'quadraticCurveTo': {
        const p = pen;
        curve(t => { const s = 1 - t; return [s * s * p[0] + 2 * s * t * a[0] + t * t * a[2], s * s * p[1] + 2 * s * t * a[1] + t * t * a[3]]; }); break;
      }
      case 'bezierCurveTo': {
        const p = pen;
        curve(t => { const s = 1 - t; return [s ** 3 * p[0] + 3 * s * s * t * a[0] + 3 * s * t * t * a[2] + t ** 3 * a[4], s ** 3 * p[1] + 3 * s * s * t * a[1] + 3 * s * t * t * a[3] + t ** 3 * a[5]]; }); break;
      }
      case 'rect': move([a[0], a[1]]); line([a[0] + a[2], a[1]]); line([a[0] + a[2], a[1] + a[3]]); line([a[0], a[1] + a[3]]); out.push({ cmd: 'close' }); pen = start; break;
      case 'arc': ellipse(a[0], a[1], a[2], a[2], 0, a[3], a[4], !!a[5]); break;
      case 'ellipse': ellipse(a[0], a[1], a[2], a[3], a[4], a[5], a[6], !!a[7]); break;
      case 'roundRect': {
        const radii: Point[] = a[4] === 0
          ? Array.from({ length: 4 }, () => [a[5], a[5]] as const)
          : Array.from({ length: (a.length - 5) / 2 }, (_, i) => [a[5 + 2 * i], a[6 + 2 * i]] as const);
        const at = (i: number): Point => radii[i % radii.length];
        const tl = at(0), tr = radii.length === 1 ? tl : at(1);
        const br = radii.length === 2 ? tl : at(2), bl = radii.length === 1 ? tl : radii.length === 2 ? tr : radii.length === 3 ? tr : at(3);
        const left = a[0], top = a[1], right = left + a[2], bottom = top + a[3];
        const f = Math.min(1, a[2] / (tl[0] + tr[0] || 1), a[2] / (bl[0] + br[0] || 1), a[3] / (tl[1] + bl[1] || 1), a[3] / (tr[1] + br[1] || 1));
        const [lx, ly] = [tl[0] * f, tl[1] * f], [rx, ry] = [tr[0] * f, tr[1] * f];
        const [bx, by] = [br[0] * f, br[1] * f], [ex, ey] = [bl[0] * f, bl[1] * f];
        move([left + lx, top]); line([right - rx, top]); ellipse(right - rx, top + ry, rx, ry, 0, -Math.PI / 2, 0, false);
        line([right, bottom - by]); ellipse(right - bx, bottom - by, bx, by, 0, 0, Math.PI / 2, false);
        line([left + ex, bottom]); ellipse(left + ex, bottom - ey, ex, ey, 0, Math.PI / 2, Math.PI, false);
        line([left, top + ly]); ellipse(left + lx, top + ly, lx, ly, 0, Math.PI, 1.5 * Math.PI, false);
        out.push({ cmd: 'close' }); pen = start; break;
      }
    }
  }
  return out;
}

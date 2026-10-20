import { buildCustomPath, samplePolyline, GeometryWorkBudgetError } from '@silurus/ooxml-core';
import type { ShapeRun } from '../types.js';
import type {
  DeepReadonly,
  DrawingPaintCommand,
  LayoutRect,
  VmlTextPathAcquisitionInput,
} from './types.js';

type TextCommand = Extract<DrawingPaintCommand, { kind: 'watermark-text' }>;
type Point = { x: number; y: number };

/** Part 4 §19.1.2.23: acquire the baseline and glyph transforms here, never
 * during paint. Curves use 48 chords, the shared WordArt numerical policy;
 * Canvas glyphs are locally affine, as with DrawingML text-warp envelopes. */
export function planVmlPathText(
  shape: DeepReadonly<Pick<ShapeRun, 'subpaths' | 'presetGeometry'>>,
  bounds: LayoutRect,
  facts: VmlTextPathAcquisitionInput,
  command: Readonly<{ sourceBounds: LayoutRect; spans: readonly Readonly<{text: string}>[] }>,
  maximumWork = Infinity,
): TextCommand['pathSlices'] {
  maximumWork = Math.min(maximumWork, 1_048_576);
  let work = 0;
  const charge = (units = 1) => {
    work += units;
    if (work > maximumWork) throw new GeometryWorkBudgetError('VML text path work budget exceeded');
  };
  const paths: Point[][] = [];
  let current: Point[] = [];
  let pen: Point = { x: 0, y: 0 };
  const point = (x: number, y: number) => {
    charge();
    pen = { x, y };
    current.push(pen);
  };
  const finish = () => {
    if (current.length > 1) paths.push(current);
    current = [];
  };
  const curve = (evaluate: (t: number) => Point) => {
    for (let i = 1; i <= 48; i++) {
      const p = evaluate(i / 48);
      point(p.x, p.y);
    }
  };
  const ellipse = (
    x: number,
    y: number,
    rx: number,
    ry: number,
    rotation: number,
    start: number,
    end: number,
    ccw = false,
  ) => {
    let delta = end - start;
    if (ccw && delta > 0) delta -= Math.PI * 2;
    if (!ccw && delta < 0) delta += Math.PI * 2;
    curve((t) => {
      const a = start + delta * t;
      return {
        x: x + rx * Math.cos(a) * Math.cos(rotation) - ry * Math.sin(a) * Math.sin(rotation),
        y: y + rx * Math.cos(a) * Math.sin(rotation) + ry * Math.sin(a) * Math.cos(rotation),
      };
    });
  };
  if (!facts.xScale && shape.presetGeometry === 'line') {
    // Part 4 §§19.1.2.12/.23: a primitive line is its actual baseline,
    // including vertical and diagonal lines. A zero-width allocation is not
    // a zero-length path, and must not become the box's horizontal edge.
    paths.push([{x:0,y:0},{x:bounds.widthPt,y:bounds.heightPt}]);
  } else if (facts.xScale || shape.subpaths.length === 0) {
    paths.push([
      { x: 0, y: bounds.heightPt },
      { x: bounds.widthPt, y: bounds.heightPt },
    ]);
  } else {
    buildCustomPath(
      {
        moveTo: (x, y) => {
          finish();
          point(x, y);
        },
        lineTo: point,
        bezierCurveTo: (x1, y1, x2, y2, x, y) => {
          const p = pen;
          curve((t) => ({
            x:
              (1 - t) ** 3 * p.x +
              3 * (1 - t) ** 2 * t * x1 +
              3 * (1 - t) * t * t * x2 +
              t ** 3 * x,
            y:
              (1 - t) ** 3 * p.y +
              3 * (1 - t) ** 2 * t * y1 +
              3 * (1 - t) * t * t * y2 +
              t ** 3 * y,
          }));
        },
        quadraticCurveTo: (x1, y1, x, y) => {
          const p = pen;
          curve((t) => ({
            x: (1 - t) ** 2 * p.x + 2 * (1 - t) * t * x1 + t * t * x,
            y: (1 - t) ** 2 * p.y + 2 * (1 - t) * t * y1 + t * t * y,
          }));
        },
        ellipse,
        arc: (x, y, r, a, b, ccw) => ellipse(x, y, r, r, 0, a, b, ccw),
        closePath: () => {
          if (current[0]) point(current[0].x, current[0].y);
        },
        rect: (x, y, w, h) => {
          finish();
          point(x, y);
          point(x + w, y);
          point(x + w, y + h);
          point(x, y + h);
          point(x, y);
        },
        roundRect: () => {
          throw new Error('Custom VML paths cannot contain roundRect');
        },
      },
      shape.subpaths,
      0,
      0,
      bounds.widthPt,
      bounds.heightPt,
    );
    finish();
  }
  const baseline = paths[paths.length - 1];
  if (!baseline) return [];
  const lengths = (points: Point[]) => {
    const values = [0];
    for (let i = 1; i < points.length; i++)
      values.push(
        values[i - 1]! +
          Math.hypot(points[i]!.x - points[i - 1]!.x, points[i]!.y - points[i - 1]!.y),
      );
    return values;
  };
  const cumulative = lengths(baseline);
  const total = cumulative[cumulative.length - 1]!;
  const top = paths.length > 1 ? paths[0] : undefined;
  const topLengths = top ? lengths(top) : undefined;
  const scale = facts.fitPath || facts.fitShape ? total / command.sourceBounds.widthPt : 1;
  // Keep complete shaped spans intact. Affine strips deform the retained
  // run instead of re-shaping code points, preserving joining, ligatures,
  // combining clusters and bidi within each service-owned span. At most 256
  // strips bound replay cost; straight paths need only one exact transform.
  const affine =
    baseline.length === 2 &&
    (!top ||
      (top.length === 2 &&
        top[1]!.x - top[0]!.x === baseline[1]!.x - baseline[0]!.x &&
        top[1]!.y - top[0]!.y === baseline[1]!.y - baseline[0]!.y));
  const count = affine ? 1 : 256;
  // Each strip replays the complete shaped spans. Charge that amplification
  // as well as sampled geometry, so a long run on a curved path stays bounded.
  charge(count * command.spans.reduce((sum, span) => sum + span.text.length, 0));
  return Array.from({ length: count }, (_, i) => {
    charge();
    const sourceWidth = command.sourceBounds.widthPt / count;
    const sourceX = command.sourceBounds.xPt + (i + 0.5) * sourceWidth;
    const u = total > 0 ? ((i + 0.5) * sourceWidth * scale) / total : 0;
    const b = samplePolyline(baseline, cumulative, u);
    const t = top && topLengths ? samplePolyline(top, topLengths, u) : undefined;
    const scaleY = facts.fitShape
      ? (t ? Math.hypot(b.x - t.x, b.y - t.y) : bounds.heightPt) / command.sourceBounds.heightPt
      : facts.fitPath
        ? scale
        : 1;
    const descent = command.sourceBounds.yPt + command.sourceBounds.heightPt;
    const c = facts.fitShape && t ? (b.x - t.x) / command.sourceBounds.heightPt : -b.ty * scaleY;
    const d = facts.fitShape && t ? (b.y - t.y) / command.sourceBounds.heightPt : b.tx * scaleY;
    return {
      sourceX,
      sourceWidth,
      xPt: b.x - (facts.fitShape ? descent * c : 0),
      yPt: b.y - (facts.fitShape ? descent * d : 0),
      a: b.tx * scale,
      b: b.ty * scale,
      c,
      d,
    };
  });
}

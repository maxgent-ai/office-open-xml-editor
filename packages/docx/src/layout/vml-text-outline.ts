import type { PathCmd } from '../types.js';
import type { DrawingPaintCommand, LayoutRect } from './types.js';
type TextCommand = Extract<DrawingPaintCommand, { kind: 'watermark-text' }>;
type Point = Readonly<{ x: number; y: number }>;

/** Part 4 §19.1.2.23: deform native glyph contours through the already
 * acquired text fitting map, before ordinary point-space stroke acquisition. */
export function acquireVmlTextContours(
  contours: readonly (readonly Point[])[],
  command: Pick<TextCommand, 'sourceBounds' | 'fitShape' | 'pathSlices' | 'glyphTransform'>,
  bounds: LayoutRect,
  maximumWork = Infinity,
): Readonly<{ subpaths: PathCmd[][]; bounds: LayoutRect; workUnits: number }> {
  const { sourceBounds, fitShape } = command;
  const slices = command.pathSlices;
  const point = (p: Point) => {
    if (slices?.length) {
      const i = Math.max(0, Math.min(slices.length - 1, Math.floor((p.x - sourceBounds.xPt) / slices[0]!.sourceWidth)));
      const t = slices[i]!;
      return { x: (t.xPt + t.a * (p.x - t.sourceX) + t.c * p.y),
        y: (t.yPt + t.b * (p.x - t.sourceX) + t.d * p.y) };
    }
    return { x: (p.x - sourceBounds.xPt) * (fitShape ? bounds.widthPt / sourceBounds.widthPt : 1),
      y: (p.y - sourceBounds.yPt) * (fitShape ? bounds.heightPt / sourceBounds.heightPt : 1) };
  };
  let contourWork = 0;
  const contourBudget = Math.min(maximumWork ?? Infinity, 1_048_576);
  const subpaths = contours.map(contour => {
    const mapped: ReturnType<typeof point>[] = [];
    const add = (p: Point) => {
      if (++contourWork > contourBudget)
        throw new RangeError('VML fitted glyph contour work budget exceeded');
      const q = point(p), m = command.glyphTransform;
      const x = bounds.xPt + q.x, y = bounds.yPt + q.y;
      mapped.push(m ? { x: m.a * x + m.c * y + m.e, y: m.b * x + m.d * y + m.f } : { x, y });
    };
    for (let i = 0; i < contour.length; i++) {
      const p = contour[i]!, q = contour[(i + 1) % contour.length]!;
      add(p);
      // Straight source edges must bend with the baseline too. Split
      // every crossed strip; endpoint-only mapping would cut a chord
      // across a curved E/F/I edge while its fill follows the curve.
      if (slices?.length && p.x !== q.x) {
        const width = slices[0]!.sourceWidth;
        const first = Math.max(1, Math.floor((Math.min(p.x, q.x) - sourceBounds.xPt) / width) + 1);
        const last = Math.min(slices.length - 1, Math.ceil((Math.max(p.x, q.x) - sourceBounds.xPt) / width) - 1);
        // Visit only crossed boundaries. Scanning every strip for every tiny
        // contour edge would amplify acquisition work without adding geometry.
        for (let k = 0; k <= last - first; k++) {
          const t = slices[q.x > p.x ? first + k : last - k]!;
          const x = t.sourceX - t.sourceWidth / 2;
          add({ x, y: p.y + (q.y - p.y) * (x - p.x) / (q.x - p.x) });
        }
      }
    }
    return [...mapped.map((p, i) => ({ cmd: i === 0 ? 'moveTo' as const : 'lineTo' as const, ...p })),
      { cmd: 'close' as const }];
  });
  // Part 4 §§19.1.2.12/.23: a line's allocation may have a zero
  // dimension while fitted glyph ink does not. Give the pen and its brush an
  // invertible ink frame, keeping the host centre (hence its rotation/flip
  // pivot) unchanged. Never normalize glyphs by a zero line extent.
  let frame = bounds;
  if (bounds.widthPt === 0 || bounds.heightPt === 0) {
    const cx = bounds.xPt + bounds.widthPt / 2, cy = bounds.yPt + bounds.heightPt / 2;
    let rx = 0, ry = 0;
    for (const path of subpaths)
      for (const p of path)
        if ('x' in p && 'y' in p) {
          rx = Math.max(rx, Math.abs(p.x - cx));
          ry = Math.max(ry, Math.abs(p.y - cy));
        }
    if (rx === 0 || ry === 0)
      throw new RangeError('VML fitted glyph contours have no ink frame');
    frame = { xPt: cx - rx, yPt: cy - ry, widthPt: 2 * rx, heightPt: 2 * ry };
  }
  return { subpaths: subpaths.map(path => path.map(p => 'x' in p && 'y' in p
      ? { ...p, x: (p.x - frame.xPt) / frame.widthPt, y: (p.y - frame.yPt) / frame.heightPt } : p)),
    bounds: frame, workUnits: contourWork };
}

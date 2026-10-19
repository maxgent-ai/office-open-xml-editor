import { drawingMLGeometryBounds } from '@silurus/ooxml-core';
import type { DrawingMLShapePaintPlan } from '@silurus/ooxml-core';
import type { DrawingPaintCommand } from './types.js';

type Concentric = NonNullable<
  Extract<DrawingPaintCommand, { kind: 'drawingml-shape' }>['concentricFill']
>;

/** Part 4 §19.1.2.14: gradientshapeok requests scaled copies of the actual
 * path. 256 rings are numerical colour precision (one per
 * 8-bit channel level); ring clipping avoids repeated alpha accumulation. */
export function acquireVmlConcentricFill(plan: DrawingMLShapePaintPlan): Concentric | undefined {
  const fill = plan.fill;
  if (fill?.fillType !== 'gradient' || fill.path !== 'shape') return undefined;
  const stops = [...fill.stops].sort((a, b) => a.position - b.position);
  if (!stops.length) return undefined;
  const color = (at: number) => {
    let low = 0,
      high = stops.length;
    while (low < high) {
      const mid = Math.floor((low + high) / 2);
      if (stops[mid]!.position <= at) low = mid + 1;
      else high = mid;
    }
    const i = Math.max(0, low - 1);
    const a = stops[i]!,
      b = stops[Math.min(i + 1, stops.length - 1)]!;
    const t = Math.max(0, Math.min(1, (at - a.position) / (b.position - a.position || 1)));
    const rgba = (raw: string) => {
      const hex = raw.replace(/^#/u, '');
      return [0, 2, 4, 6].map((offset) =>
        offset === 6 && hex.length === 6 ? 255 : parseInt(hex.slice(offset, offset + 2), 16),
      );
    };
    const ac = rgba(a.color),
      bc = rgba(b.color);
    return `#${ac
      .map((value, j) =>
        Math.round(value + (bc[j]! - value) * t)
          .toString(16)
          .padStart(2, '0'),
      )
      .join('')}`;
  };
  const rect = fill.fillToRect;
  const l = rect?.l ?? 0,
    t = rect?.t ?? 0;
  const fw = 1 - l - (rect?.r ?? 1),
    fh = 1 - t - (rect?.b ?? 1);
  const bounds = drawingMLGeometryBounds({
    ...plan,
    transform: { rotationDeg: 0, flipH: false, flipV: false },
  });
  return {
    ...(bounds
      ? { bounds: { xPt: bounds.x, yPt: bounds.y, widthPt: bounds.w, heightPt: bounds.h } }
      : {}),
    focus: { l, t, width: fw, height: fh },
    rings: Array.from({ length: 256 }, (_, i) => ({
      outside: 1 - i / 256,
      inside: 1 - (i + 1) / 256,
      color: color(1 - i / 256),
    })),
    centerColor: color(0),
  };
}

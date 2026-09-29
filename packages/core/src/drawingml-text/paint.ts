import type { DrawingMlLineSegment } from './break.js';

/**
 * Last DrawingML phase: walk already measured and aligned segments. Host
 * adapters own the actual glyph/object painter and any overlay metadata;
 * advance always comes from the measured line so paint cannot rewrap text.
 */
export function paintDrawingMlLine<T>(
  segments: readonly DrawingMlLineSegment<T>[],
  x: number,
  y: number,
  paint: (segment: DrawingMlLineSegment<T>, x: number, y: number) => void,
): number {
  let pen = x;
  for (const segment of segments) {
    paint(segment, pen, y);
    pen += segment.width;
  }
  return pen;
}

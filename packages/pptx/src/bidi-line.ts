// DrawingML segment ordering delegates to the shared UAX#9 kernel. A tab
// receives Bidi_Class S so independently aligned cells mirror under RTL.
import {
  computeSegmentLineVisualOrder, segmentLineHasRtl,
  type SegmentLineVisualOrder,
} from '@silurus/ooxml-core/internal/bidi-line';

const options = { isTab: (segment: unknown) => 'isTab' in (segment as object) };
export type LineVisualOrder = SegmentLineVisualOrder;
export const segmentsHaveRtl = (segments: readonly unknown[]): boolean =>
  segmentLineHasRtl(segments);
export const computeLineVisualOrder = (
  segments: readonly unknown[], baseRtl: boolean,
): LineVisualOrder => computeSegmentLineVisualOrder(segments, baseRtl, options);

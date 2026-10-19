type SelectionAutoScrollMode = 'cells' | 'rows' | 'cols' | 'all';

/** Width of the viewport-edge activation band in CSS pixels. */
const SELECTION_AUTO_SCROLL_EDGE_PX = 40;
/** Maximum logical viewport speed while drag-selecting outside an edge. */
const SELECTION_AUTO_SCROLL_MAX_PX_PER_SECOND = 900;

export interface SelectionAutoScrollPoint {
  readonly x: number;
  readonly y: number;
}

export interface SelectionAutoScrollViewport {
  readonly width: number;
  readonly height: number;
}

function axisVelocity(position: number, extent: number): number {
  if (extent <= 0) return 0;
  const edge = Math.min(SELECTION_AUTO_SCROLL_EDGE_PX, extent / 2);
  if (position < edge) {
    const strength = Math.min(1, Math.max(0, (edge - position) / edge));
    return -SELECTION_AUTO_SCROLL_MAX_PX_PER_SECOND * strength;
  }
  if (position > extent - edge) {
    const strength = Math.min(1, Math.max(0, (position - (extent - edge)) / edge));
    return SELECTION_AUTO_SCROLL_MAX_PX_PER_SECOND * strength;
  }
  return 0;
}

/**
 * Return a logical-LTR viewport velocity for one drag-selection pointer.
 * Physical left/right are inverted for an RTL sheet because logical x=0 is
 * always the column-A edge. Row/column header selections suppress the axis
 * that cannot extend that selection.
 */
export function selectionAutoScrollVelocity(
  point: SelectionAutoScrollPoint,
  viewport: SelectionAutoScrollViewport,
  rtl: boolean,
  mode: SelectionAutoScrollMode,
): SelectionAutoScrollPoint {
  if (mode === 'all') return { x: 0, y: 0 };
  const physicalX = mode === 'rows' ? 0 : axisVelocity(point.x, viewport.width);
  const y = mode === 'cols' ? 0 : axisVelocity(point.y, viewport.height);
  return { x: rtl ? -physicalX : physicalX, y };
}

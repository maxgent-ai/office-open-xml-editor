import type { ViewportRange, Worksheet, WorksheetCellRange } from '../types.js';

export type ViewportPreviewBlocker =
  | 'conditional-format-range'
  | 'merge-range'
  | 'drawing-dependency'
  | 'sparkline-dependency';

function intersects(range: WorksheetCellRange, viewport: ViewportRange): boolean {
  return range.top <= viewport.row + viewport.rows - 1 &&
    range.bottom >= viewport.row &&
    range.left <= viewport.col + viewport.cols - 1 &&
    range.right >= viewport.col;
}

function anchorMayReachViewport(
  anchor: { fromRow: number; fromCol: number; fromRowOff: number; fromColOff: number },
  viewport: ViewportRange,
): boolean {
  // DrawingML markers are zero-based (§20.5.2.33). An anchor is provably
  // outside only when its start is below or right of the viewport with a
  // nonnegative offset. End markers and native extents can exceed their cells,
  // so an anchor starting above/left may still reach the viewport.
  const below = anchor.fromRow + 1 > viewport.row + viewport.rows - 1 &&
    Number.isFinite(anchor.fromRowOff) && anchor.fromRowOff >= 0;
  const right = anchor.fromCol + 1 > viewport.col + viewport.cols - 1 &&
    Number.isFinite(anchor.fromColOff) && anchor.fromColOff >= 0;
  return !below && !right;
}

/** Conditional-format dependencies of the rows `top..bottom`. A statistical
 * rule that intersects them needs all rows in its own sqref, and a formula
 * operand may reach rows that have not arrived. */
function conditionalFormatBlocker(
  worksheet: Worksheet,
  top: number,
  bottom: number,
  coveredThrough: number,
): 'conditional-format-range' | null {
  for (const format of worksheet.conditionalFormats ?? []) {
    const ranges = format.sqref.filter((range) => range.top <= bottom && range.bottom >= top);
    if (ranges.length === 0) continue;
    for (const rule of format.rules) {
      if (rule.type === 'other') continue;
      if (rule.type === 'cellIs') {
        // Literal operands are row-local. A cell reference can reach unloaded
        // rows, so defer it until the complete cell graph is available.
        if (rule.formulas.every((formula) => /^\s*[-+]?(?:\d+(?:\.\d*)?|\.\d+)\s*$/.test(formula))) continue;
      }
      // Compilation gathers samples from the entire sqref, including ranges
      // outside the viewport. Any unloaded member can change the statistic.
      if (format.sqref.some((range) => range.bottom > coveredThrough)) {
        return 'conditional-format-range';
      }
      if ('activeFormula' in rule && rule.activeFormula &&
          !/^\s*[-+]?(?:\d+(?:\.\d*)?|\.\d+)\s*$/.test(rule.activeFormula)) {
        return 'conditional-format-range';
      }
      if (rule.type === 'cellIs' || rule.type === 'expression') {
        // A formula may refer outside its formatted range. The current
        // evaluator resolves cached cell references from the worksheet graph.
        return 'conditional-format-range';
      }
    }
  }
  return null;
}

/** Eligible cells have all their row-local values and every visual dependency.
 * A rule outside the viewport is irrelevant; a statistical rule that intersects
 * it needs all rows in its own sqref, not merely the first visible chunk. */
export function viewportPreviewBlocker(
  worksheet: Worksheet,
  viewport: ViewportRange,
  coveredThrough: number,
): ViewportPreviewBlocker | null {
  // A merged perimeter can take a border from its bottom/right cells even
  // when only the anchor is visible. Metadata supplies the range, but the
  // cell styles are row data and must all have arrived before painting.
  // Formula spill values are cached per cell and are not recalculated by this
  // renderer; their visible cells arrive with their own complete rows.
  if ((worksheet.mergeCells ?? []).some((merge) =>
    intersects(merge, viewport) && merge.bottom > coveredThrough)) {
    return 'merge-range';
  }
  // An unwrapped value in any painted row can overflow horizontally into
  // the viewport from an off-screen column. Its CF can alter text metrics,
  // so column-only culling here would accept an incomplete first frame.
  const formatBlocker = conditionalFormatBlocker(
    worksheet, viewport.row, viewport.row + viewport.rows - 1, coveredThrough);
  if (formatBlocker) return formatBlocker;
  if ((worksheet.charts ?? []).some((anchor) => anchorMayReachViewport(anchor, viewport)) ||
      (worksheet.shapeGroups ?? []).some((anchor) => anchorMayReachViewport(anchor, viewport)) ||
      (worksheet.slicers ?? []).some((anchor) => anchorMayReachViewport(anchor, viewport))) {
    return 'drawing-dependency';
  }
  if ((worksheet.sparklineGroups ?? []).some((group) => group.sparklines.some((sparkline) =>
    sparkline.row >= viewport.row && sparkline.row < viewport.row + viewport.rows &&
    sparkline.col >= viewport.col && sparkline.col < viewport.col + viewport.cols))) {
    return 'sparkline-dependency';
  }
  return null;
}

/** #1713 partial-preview gate for the prepared-initial anchor baseline.
 * Rows `1..lastRow` determine every eligible anchor's initial rect, so their
 * automatic heights must be final before capture. Row data has arrived through
 * `coveredThrough`; a conditional format whose rules can depend on unloaded
 * rows may still change fonts and therefore heights. That case reports
 * `drawing-dependency` and the viewer falls back to complete materialization
 * (capability boundary: the first frame of such a sheet waits for all rows;
 * sheets without eligible anchors never reach this gate). Merge exclusions
 * come from metadata; charts and sparklines do not affect row heights. */
export function anchorBaselinePreviewBlocker(
  worksheet: Worksheet,
  lastRow: number,
  coveredThrough: number,
): 'drawing-dependency' | null {
  if (lastRow < 1) return null;
  return conditionalFormatBlocker(worksheet, 1, lastRow, coveredThrough) ? 'drawing-dependency' : null;
}

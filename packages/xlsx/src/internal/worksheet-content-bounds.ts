import type { Worksheet } from '../types.js';
import { MAX_WORKSHEET_COL, MAX_WORKSHEET_ROW } from './grid-geometry.js';

type CellAnchoredObject = {
  readonly fromCol: number;
  readonly fromRow: number;
  readonly toCol: number;
  readonly toRow: number;
};

const previewBounds = new WeakMap<Worksheet, { maxRow: number; maxCol: number }>();

/** The metadata pass derives these from every MCE-visible row and cell, so the
 * scrollbars have their final extent while values are still arriving. */
export function setWorksheetPreviewBounds(
  worksheet: Worksheet, bounds: { maxRow: number; maxCol: number },
): void {
  previewBounds.set(worksheet, bounds);
}

export function inheritWorksheetPreviewBounds(source: Worksheet, target: Worksheet): void {
  const bounds = previewBounds.get(source);
  if (bounds) previewBounds.set(target, bounds);
}

/** Return the worksheet grid bounds needed to expose both cell content and
 * DrawingML objects. Two-cell anchors use zero-based markers (§20.5.2.33), so
 * an object ending at marker row 113 requires row 114 in the scroll extent.
 * Excel permits drawings well beyond the `<dimension>` / populated cell range;
 * ignoring them makes those authored objects unreachable in a viewer. */
export function worksheetContentBounds(ws: Worksheet): { maxRow: number; maxCol: number } {
  const preview = previewBounds.get(ws);
  let maxRow = Math.max(50, ws.freezeRows ?? 0, preview?.maxRow ?? 0);
  let maxCol = Math.max(26, ws.freezeCols ?? 0, preview?.maxCol ?? 0);
  for (const row of ws.rows) {
    if (row.index > maxRow) maxRow = row.index;
    for (const cell of row.cells) {
      if (cell.col > maxCol) maxCol = cell.col;
    }
  }
  const anchored: readonly CellAnchoredObject[] = [
    ...ws.charts,
    ...ws.images,
    ...(ws.shapeGroups ?? []),
    ...(ws.slicers ?? []),
  ];
  for (const anchor of anchored) {
    const fromRow = Number.isSafeInteger(anchor.fromRow) ? anchor.fromRow + 1 : 0;
    const toRow = Number.isSafeInteger(anchor.toRow) ? anchor.toRow + 1 : 0;
    const fromCol = Number.isSafeInteger(anchor.fromCol) ? anchor.fromCol + 1 : 0;
    const toCol = Number.isSafeInteger(anchor.toCol) ? anchor.toCol + 1 : 0;
    maxRow = Math.max(maxRow, fromRow, toRow);
    maxCol = Math.max(maxCol, fromCol, toCol);
  }
  return {
    maxRow: Math.min(MAX_WORKSHEET_ROW, maxRow),
    maxCol: Math.min(MAX_WORKSHEET_COL, maxCol),
  };
}

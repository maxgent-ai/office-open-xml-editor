import type { SheetMeta, SheetVisibility } from './types';

/**
 * Pure core of {@link XlsxWorkbook.sheetVisibility}: the visibility of the sheet
 * at `sheetIndex` (0-based, absolute). Like pptx's `isHidden`/`getNotes`,
 * the index is NOT clamped — out-of-range / non-integer ⇒ `'visible'`
 * ("no sheet here, so treat as visible").
 */
export function selectSheetVisibility(
  sheets: readonly SheetMeta[],
  sheetIndex: number,
): SheetVisibility {
  if (!Number.isInteger(sheetIndex) || sheetIndex < 0 || sheetIndex >= sheets.length) {
    return 'visible';
  }
  return sheets[sheetIndex].visibility ?? 'visible';
}

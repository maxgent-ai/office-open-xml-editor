/**
 * The style a cell paints with: the one rule the renderer and font-source
 * demand (font-source-load.ts) share, so demand charges every painted text to
 * the face that actually paints it.
 *
 * A cell's own `c/@s` (ECMA-376 §18.3.1.4) wins; a cell without one takes its
 * column's `<col style>` (§18.3.1.13, `colStyleRanges`, last matching range);
 * otherwise Normal (0). The worksheet model carries no row style
 * (`<row s customFormat>`, §18.3.1.73), so a row default does not apply.
 */
import type { Cell, CellXf, Styles, Worksheet } from '../types.js';

export function columnStyleIndex(worksheet: Pick<Worksheet, 'colStyleRanges'>, col: number): number {
  const ranges = worksheet.colStyleRanges ?? [];
  for (let index = ranges.length - 1; index >= 0; index--) {
    const range = ranges[index];
    if (col >= range.min && col <= range.max) return range.styleIndex;
  }
  return 0;
}

export function effectiveCellStyleIndex(
  worksheet: Pick<Worksheet, 'colStyleRanges'>,
  cell: Pick<Cell, 'styleIndex'> | undefined,
  col: number,
): number {
  return cell?.styleIndex ?? columnStyleIndex(worksheet, col);
}

/** The cellXfs entry of a style index; an out-of-range index paints as the
 * first entry (the renderer's resolveXf). */
export function cellXfAt(styles: Pick<Styles, 'cellXfs'>, styleIndex: number): CellXf | undefined {
  return styles.cellXfs?.[styleIndex] ?? styles.cellXfs?.[0];
}

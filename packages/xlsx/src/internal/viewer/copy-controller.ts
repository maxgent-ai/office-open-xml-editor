import type { Worksheet } from '../../types.js';
import type { XlsxWorkbook } from '../../workbook.js';
import type { XlsxSelectionState } from '../../selection.js';
import type { XlsxCopyResult } from '../../viewer.js';

// Keep clipboard materialization within the same hard cell-count envelope as a
// worksheet. A sparse range can span billions of coordinates even when only a
// handful of cells are populated, so its rectangular TSV must never be built.
const MAX_CLIPBOARD_CELLS = 250_000;
// Bound the retained TSV and its final joined copy. This is a resource-safety
// contract, not a worksheet semantic limit; callers can handle `too-large`
// without the viewer attempting an unbounded JavaScript string allocation.
const MAX_CLIPBOARD_UTF16_CODE_UNITS = 8 * 1_024 * 1_024;

function encodeTsvFieldWithin(value: string, remaining: number): string | null {
  let quoteCount = 0;
  let needsQuotes = false;
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    if (code === 34) { quoteCount++; needsQuotes = true; }
    else if (code === 9 || code === 10 || code === 13) needsQuotes = true;
  }
  const length = value.length + (needsQuotes ? quoteCount + 2 : 0);
  if (length > remaining) return null;
  return needsQuotes ? `"${value.replace(/"/g, '""')}"` : value;
}

/** Displayed worksheet, selection and clipboard the copy command reads. */
export interface CopyControllerHost {
  worksheet(): Worksheet | null;
  selection(): XlsxSelectionState | null;
  workbook(): XlsxWorkbook | null;
  /** Resolved only once a bounded TSV is ready to write. */
  clipboard(): Clipboard | undefined;
}

/**
 * Copies the selected area as bounded TSV. The same limits apply regardless of
 * whether pointer, keyboard, or API created the selection.
 */
export class CopyController {
  constructor(private readonly host: CopyControllerHost) {}

  async copy(): Promise<XlsxCopyResult> {
    const ws = this.host.worksheet();
    const state = this.host.selection();
    if (!ws || !state) return { status: 'empty-selection' };
    if (state.areas.length !== 1) return { status: 'unsupported-multiple-areas' };
    const area = state.areas[0];
    const wb = this.host.workbook();

    // Whole-row/column/sheet selections are unbounded Excel concepts. Copying
    // narrows them to used cells without changing the logical selection.
    let maxRow = 1, maxCol = 1;
    for (const row of ws.rows) {
      if (row.index > maxRow) maxRow = row.index;
      for (const cell of row.cells) {
        if (cell.col > maxCol) maxCol = cell.col;
      }
    }

    const { r1, r2, c1, c2 } = area.kind === 'sheet'
      ? { r1: 1, r2: maxRow, c1: 1, c2: maxCol }
      : area.kind === 'rows'
        ? { r1: area.firstRow, r2: area.lastRow, c1: 1, c2: maxCol }
        : area.kind === 'columns'
          ? { r1: 1, r2: maxRow, c1: area.firstColumn, c2: area.lastColumn }
          : { r1: area.top, r2: area.bottom, c1: area.left, c2: area.right };

    const rowCount = r2 - r1 + 1;
    const colCount = c2 - c1 + 1;
    if (rowCount > Math.floor(MAX_CLIPBOARD_CELLS / colCount)) {
      return { status: 'too-large', limit: 'cells' };
    }
    const cellCount = rowCount * colCount;

    let utf16CodeUnits = Math.max(0, rowCount - 1) + rowCount * Math.max(0, colCount - 1);
    if (utf16CodeUnits > MAX_CLIPBOARD_UTF16_CODE_UNITS) {
      return { status: 'too-large', limit: 'text' };
    }
    const cellMap = new Map<number, Map<number, string>>();
    for (const row of ws.rows) {
      if (row.index < r1 || row.index > r2) continue;
      for (const cell of row.cells) {
        if (cell.col < c1 || cell.col > c2) continue;
        const v = cell.value;
        let text = wb?.cellText(ws, cell) ?? '';
        if (!wb) {
          if (v.type === 'text') text = v.runs ? v.runs.map((r) => r.text).join('') : v.text;
          else if (v.type === 'number') text = String(v.number);
          else if (v.type === 'bool') text = v.bool ? 'TRUE' : 'FALSE';
          else if (v.type === 'error') text = v.error;
        }
        if (text) {
          const encoded = encodeTsvFieldWithin(
            text,
            MAX_CLIPBOARD_UTF16_CODE_UNITS - utf16CodeUnits,
          );
          if (encoded === null) return { status: 'too-large', limit: 'text' };
          utf16CodeUnits += encoded.length;
          let values = cellMap.get(row.index);
          if (!values) { values = new Map(); cellMap.set(row.index, values); }
          values.set(cell.col, encoded);
        }
      }
    }

    const lines: string[] = [];
    for (let r = r1; r <= r2; r++) {
      const cols: string[] = [];
      const values = cellMap.get(r);
      for (let c = c1; c <= c2; c++) {
        const value = values?.get(c) ?? '';
        cols.push(value);
      }
      lines.push(cols.join('\t'));
    }
    const clipboard = this.host.clipboard();
    if (!clipboard) return { status: 'clipboard-unavailable' };
    try {
      await clipboard.writeText(lines.join('\n'));
      return { status: 'copied', cellCount, utf16CodeUnits };
    } catch {
      return { status: 'clipboard-denied' };
    }
  }
}

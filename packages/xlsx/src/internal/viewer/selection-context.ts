import type { Cell, Row, Worksheet, XlsxComment } from '../../types.js';
import type { XlsxWorkbook } from '../../workbook.js';
import type {
  XlsxSelectionContext,
  XlsxSelectionContextCell,
  XlsxSelectionContextOptions,
  XlsxSelectionState,
} from '../../selection.js';
import {
  MAX_SELECTION_CONTEXT_CELLS,
  MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
  selectionCoordinateCountUpperBound,
} from '../../selection.js';
import { MAX_WORKSHEET_COL, MAX_WORKSHEET_ROW } from '../grid-geometry.js';

type SelectionInterval = Readonly<{ first: number; last: number }>;

function mergeSelectionIntervals(intervals: readonly SelectionInterval[]): SelectionInterval[] {
  const sorted = [...intervals].sort((a, b) => a.first - b.first || a.last - b.last);
  const merged: SelectionInterval[] = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || interval.first > previous.last + 1) {
      merged.push({ ...interval });
    } else if (interval.last > previous.last) {
      merged[merged.length - 1] = { first: previous.first, last: interval.last };
    }
  }
  return merged;
}

function intervalContains(intervals: readonly SelectionInterval[], value: number): boolean {
  let low = 0;
  let high = intervals.length - 1;
  while (low <= high) {
    const middle = (low + high) >>> 1;
    const interval = intervals[middle];
    if (value < interval.first) high = middle - 1;
    else if (value > interval.last) low = middle + 1;
    else return true;
  }
  return false;
}

function lowerBoundBy<T>(items: readonly T[], value: number, key: (item: T) => number): number {
  let low = 0;
  let high = items.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (key(items[middle]) < value) low = middle + 1;
    else high = middle;
  }
  return low;
}

function orderedBy<T>(items: readonly T[], key: (item: T) => number): readonly T[] {
  for (let index = 1; index < items.length; index++) {
    if (key(items[index - 1]) > key(items[index])) {
      return [...items].sort((left, right) => key(left) - key(right));
    }
  }
  return items;
}

const DEFAULT_SELECTION_CONTEXT_TEXT_CHARACTERS = 1 * 1_024 * 1_024;
const MAX_SELECTION_CONTEXT_FIELD_CHARACTERS = 65_536;

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, Math.max(0, maxCodeUnits));
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

/**
 * Builds the serializable, bounded range context for a cell selection: the
 * populated cells (value, display text, formula and comment thread) covered
 * by every selected area, with cell-count and text budgets. Scans only the
 * selected row and column intervals of the parsed model.
 */
export class SelectionContextReader {
  // SpreadsheetML permits explicit row/cell references to appear out of
  // coordinate order. Cache a canonical view once per immutable parsed model
  // so range extraction can use binary search without silently skipping such
  // cells on every subsequent context read.
  private readonly rows = new WeakMap<Worksheet, readonly Row[]>();
  private readonly cells = new WeakMap<Row, readonly Cell[]>();

  /** Range context for `selection` on `worksheet` (sheet `sheetIndex`).
   *  `comments` is the authored `"row:col"` comment index and `workbook`
   *  formats display text. */
  read(
    worksheet: Worksheet,
    sheetIndex: number,
    selection: XlsxSelectionState,
    comments: ReadonlyMap<string, XlsxComment>,
    workbook: XlsxWorkbook | null,
    options: XlsxSelectionContextOptions,
  ): XlsxSelectionContext {
    const requestedMax = options.maxCells ?? 1_000;
    if (!Number.isFinite(requestedMax) || requestedMax < 0) {
      throw new RangeError('maxCells must be a finite non-negative number.');
    }
    const maxCells = Math.min(MAX_SELECTION_CONTEXT_CELLS, Math.floor(requestedMax));
    const requestedTextMax = options.maxTextCharacters ?? DEFAULT_SELECTION_CONTEXT_TEXT_CHARACTERS;
    if (!Number.isFinite(requestedTextMax) || requestedTextMax < 0) {
      throw new RangeError('maxTextCharacters must be a finite non-negative number.');
    }
    const maxTextCharacters = Math.min(
      MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
      Math.floor(requestedTextMax),
    );
    let textCharacters = 0;
    let textTruncated = false;
    const boundedField = (input: string | readonly Readonly<{ text: string }>[]): string => {
      const parts: readonly (string | Readonly<{ text: string }>)[] =
        typeof input === 'string' ? [input] : input;
      const chunks: string[] = [];
      let fieldCharacters = 0;
      for (let index = 0; index < parts.length; index++) {
        const sourcePart = parts[index];
        const part = typeof sourcePart === 'string' ? sourcePart : sourcePart.text;
        const allowed = Math.max(0, Math.min(
          MAX_SELECTION_CONTEXT_FIELD_CHARACTERS - fieldCharacters,
          maxTextCharacters - textCharacters,
        ));
        const chunk = safeUtf16Prefix(part, allowed);
        chunks.push(chunk);
        fieldCharacters += chunk.length;
        textCharacters += chunk.length;
        if (chunk.length < part.length || index + 1 < parts.length && allowed === 0) {
          textTruncated = true;
          break;
        }
      }
      return chunks.join('');
    };
    const sheetSelected = selection.areas.some((area) => area.kind === 'sheet');
    const rowIntervals = mergeSelectionIntervals(selection.areas.flatMap((area) =>
      area.kind === 'rows' ? [{ first: area.firstRow, last: area.lastRow }] : []));
    const columnIntervals = mergeSelectionIntervals(selection.areas.flatMap((area) =>
      area.kind === 'columns'
        ? [{ first: area.firstColumn, last: area.lastColumn }]
        : []));
    const rectangles = selection.areas.flatMap((area) => area.kind === 'cells' ? [area] : []);
    const events = rectangles.flatMap((area, index) => [
      { row: area.top, index, active: true },
      { row: area.bottom + 1, index, active: false },
    ]).sort((a, b) => a.row - b.row || Number(a.active) - Number(b.active));
    const activeRectangles = new Set<number>();
    let eventIndex = 0;
    let activeColumnIntervals: SelectionInterval[] = [];
    const cells: XlsxSelectionContextCell[] = [];
    let cellsTruncated = false;
    const selectedRowIntervals = sheetSelected || columnIntervals.length > 0
      ? [{ first: 1, last: MAX_WORKSHEET_ROW }]
      : mergeSelectionIntervals([
          ...rowIntervals,
          ...rectangles.map((area) => ({ first: area.top, last: area.bottom })),
        ]);
    let rows = this.rows.get(worksheet);
    if (!rows) {
      rows = orderedBy(worksheet.rows, (row) => row.index);
      this.rows.set(worksheet, rows);
    }

    cellScan: for (const selectedRows of selectedRowIntervals) {
      let rowIndex = lowerBoundBy(rows, selectedRows.first, (row) => row.index);
      while (rowIndex < rows.length) {
        const row = rows[rowIndex++];
        if (row.index > selectedRows.last) break;
        let changed = false;
        while (eventIndex < events.length && events[eventIndex].row <= row.index) {
          const event = events[eventIndex++];
          if (event.active) activeRectangles.add(event.index);
          else activeRectangles.delete(event.index);
          changed = true;
        }
        if (changed) {
          activeColumnIntervals = mergeSelectionIntervals([...activeRectangles].map((index) => ({
            first: rectangles[index].left,
            last: rectangles[index].right,
          })));
        }
        const wholeRow = sheetSelected || intervalContains(rowIntervals, row.index);
        const selectedColumns = wholeRow
          ? [{ first: 1, last: MAX_WORKSHEET_COL }]
          : mergeSelectionIntervals([...columnIntervals, ...activeColumnIntervals]);
        for (const selectedColumnsInterval of selectedColumns) {
          let rowCells = this.cells.get(row);
          if (!rowCells) {
            rowCells = orderedBy(row.cells, (cell) => cell.col);
            this.cells.set(row, rowCells);
          }
          let cellIndex = lowerBoundBy(rowCells, selectedColumnsInterval.first, (cell) => cell.col);
          while (cellIndex < rowCells.length) {
            const cell = rowCells[cellIndex++];
            if (cell.col > selectedColumnsInterval.last) break;
            const raw = cell.value;
            const sourceComment = comments.get(`${cell.row}:${cell.col}`);
            if (raw.type === 'empty' && cell.formula === undefined && !sourceComment) continue;
            if (cells.length >= maxCells) { cellsTruncated = true; break cellScan; }
            const displayText = boundedField(workbook?.cellText(worksheet, cell) ?? '');
            const value = raw.type === 'text'
              ? boundedField(raw.runs ?? raw.text)
              : raw.type === 'number'
                ? raw.number
                : raw.type === 'bool'
                  ? raw.bool
                  : raw.type === 'error'
                    ? boundedField(raw.error)
                  : null;
            const comment = sourceComment ? {
              root: {
                id: sourceComment.id,
                author: sourceComment.author,
                date: sourceComment.date,
                text: boundedField(sourceComment.rootText ?? sourceComment.text),
                status: sourceComment.resolved ? 'resolved' as const : 'active' as const,
              },
              replies: (sourceComment.replies ?? []).map((reply) => ({
                id: reply.id,
                author: reply.author,
                date: reply.date,
                text: boundedField(reply.text),
                status: reply.resolved ? 'resolved' as const : 'active' as const,
              })),
            } : undefined;
            cells.push({
              address: { row: cell.row, col: cell.col },
              displayText,
              valueType: raw.type,
              value,
              ...(cell.formula === undefined ? {} : { formula: boundedField(cell.formula) }),
              ...(comment === undefined ? {} : { comment }),
            });
            if (textTruncated) break cellScan;
          }
        }
      }
    }
    const truncationReasons: Array<'cells' | 'text'> = [];
    if (cellsTruncated) truncationReasons.push('cells');
    if (textTruncated) truncationReasons.push('text');
    return {
      format: 'xlsx',
      kind: 'range',
      sheetIndex,
      sheetName: worksheet.name,
      selection,
      coordinateCountUpperBound: selectionCoordinateCountUpperBound(selection),
      cells,
      truncated: truncationReasons.length > 0,
      truncationReasons,
      maxCells,
      textCharacters,
      maxTextCharacters,
    };
  }
}

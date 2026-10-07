import type { SourceRef } from './types.js';
import type { BodyTableContinuationCursor } from './body-layout-kernel.js';
import type { TableFragmentCursor } from './table-pagination.js';
import { stableFingerprint } from './fingerprint.js';

export function sourceKey(source: SourceRef): string {
  return `${source.story}:${encodeURIComponent(source.storyInstance)}:${source.path.join('.')}`;
}

export function bodyOccurrenceKey(
  source: SourceRef,
  flowDomainId: string,
  fragmentStartKey: string,
): string {
  if (flowDomainId.length === 0 || fragmentStartKey.length === 0) {
    throw new RangeError('Body occurrence identity requires a flow domain and fragment start');
  }
  return [
    'body-occurrence',
    encodeURIComponent(sourceKey(source)),
    encodeURIComponent(flowDomainId),
    encodeURIComponent(fragmentStartKey),
  ].join('/');
}

function tableCursorKey(cursor: TableFragmentCursor): readonly unknown[] {
  return [
    cursor.rowIndex,
    cursor.rowFragmentIndex,
    cursor.cells.map((cell) => [
      cell.blockIndex,
      cell.paragraphLineStart,
      cell.nestedFragmentIndex,
      cell.nestedCursor === null ? null : tableCursorKey(cell.nestedCursor),
    ]),
  ];
}

/** Fragment-start component of a body table occurrence key. The paginator
 * binds accepted occurrences with it; owner-run measurement derives the same
 * key to keep its own occurrence out of the registry it resolves against. */
export function tableFragmentStartKey(cursor: BodyTableContinuationCursor | undefined): string {
  if (cursor === undefined) return 'root';
  if (cursor.kind === 'table') return `table:${JSON.stringify(tableCursorKey(cursor.cursor))}`;
  const tableCursor = cursor.cursor.tableCursor;
  return `adjacent-table:${cursor.cursor.tableIndex}:${cursor.cursor.sourceRowIndex}:${JSON.stringify(
    tableCursor === undefined ? null : tableCursorKey(tableCursor),
  )}`;
}

/** Acquisition identity shared by the page prescan and the accepted root table. */
export function bodyRootFloatingTablePlacementKey(
  source: SourceRef,
  pageIndex: number,
  rowIndex: number,
  rowFragmentIndex: number,
): string {
  return `table:${source.path.join('.')}:root:${pageIndex}:${rowIndex}:${rowFragmentIndex}`;
}

export function imageResourceKey(source: SourceRef, partPath: string): string {
  return `image:${sourceKey(source)}:${encodeURIComponent(partPath)}`;
}

export function mathResourceKey(source: SourceRef, localName: string): string {
  return `math:${sourceKey(source)}:${encodeURIComponent(localName)}`;
}

export function chartResourceKey(source: SourceRef): string {
  return stableFingerprint('chart-resource', source);
}

export function anchorOccurrenceKey(source: SourceRef, parserLocalId: string): string {
  return `anchor:${sourceKey(source)}:${encodeURIComponent(parserLocalId)}`;
}

import type { BodyElement } from '../types.js';

export type TableBodyElement = Extract<BodyElement, { type: 'table' }>;

/** Parser-owned ECMA-376 Part 1 §17.4.37 logical-table membership facts for one
 * body table. `table === null` when the parser preserved no logical-table
 * identity (e.g. a hand-built public `DocTable` with no acquisition wire); such
 * a table can never join a logical sequence. Membership, style-identity
 * validity, and the effective-floating barrier are all decided by the parser
 * and surface here only as the assigned `logicalSequenceId`; layout does not
 * re-derive §17.4.37 grouping from style ids or positioning. */
export interface AdjacentTableSequenceInput {
  readonly element: BodyElement;
  readonly table: Readonly<{
    readonly logicalSequenceId: string;
    readonly logicalRowOffset: number;
    readonly logicalTotalRows: number;
    readonly rowCount: number;
  }> | null;
}

export type NormalizedBodySequenceEntry =
  | Readonly<{ kind: 'body-element'; element: BodyElement }>
  | Readonly<{
    kind: 'adjacent-table-group';
    logicalSequenceId: string;
    tables: readonly TableBodyElement[];
  }>;

type PendingMember = Readonly<{
  element: TableBodyElement;
  logicalRowOffset: number;
  logicalTotalRows: number;
  rowCount: number;
}>;

/** Guard the parser-owned row sequence the layer is about to trust: within one
 * logical table the authored member offsets must be contiguous from zero and
 * agree on the shared total. Layout never recomputes these; a violation means
 * the parser boundary is inconsistent, which must fail loudly rather than
 * silently produce a mismatched union grid. */
function assertParserOwnedSequence(sequenceId: string, members: readonly PendingMember[]): void {
  const totalRows = members[0]!.logicalTotalRows;
  let expectedOffset = 0;
  for (const member of members) {
    if (
      member.logicalTotalRows !== totalRows
      || !Number.isInteger(member.rowCount)
      || member.rowCount < 0
      || member.logicalRowOffset !== expectedOffset
    ) {
      throw new Error(`Parser-owned adjacent table sequence ${sequenceId} is inconsistent`);
    }
    expectedOffset += member.rowCount;
  }
  if (expectedOffset !== totalRows) {
    throw new Error(`Parser-owned adjacent table sequence ${sequenceId} is incomplete`);
  }
}

/**
 * Group the body sequence into ECMA-376 Part 1 §17.4.37 logical tables using
 * only the parser-owned `logicalSequenceId`. The parser has already applied the
 * normative adjacency rule together with the compatibility-owned positioned-
 * table exclusion and table-style-identity validation, so directly adjacent
 * source tables carry the same id iff they form one logical table. This layer
 * only coalesces the maximal contiguous run sharing an id and validates the
 * parser-owned row totals; it does not compare derived geometry or style ids.
 */
export function* normalizeAdjacentTables(
  body: Iterable<AdjacentTableSequenceInput>,
): Generator<NormalizedBodySequenceEntry> {
  let pendingSequenceId: string | null = null;
  let pendingMembers: PendingMember[] = [];

  function* flush(): Generator<NormalizedBodySequenceEntry> {
    if (pendingMembers.length > 0) {
      assertParserOwnedSequence(pendingSequenceId!, pendingMembers);
    }
    if (pendingMembers.length === 1) {
      yield Object.freeze({ kind: 'body-element', element: pendingMembers[0]!.element });
    } else if (pendingMembers.length > 1) {
      yield Object.freeze({
        kind: 'adjacent-table-group',
        logicalSequenceId: pendingSequenceId!,
        tables: Object.freeze(pendingMembers.map((member) => member.element)),
      });
    }
    pendingSequenceId = null;
    pendingMembers = [];
  }

  for (const { element, table } of body) {
    if (element.type === 'table' && table !== null) {
      if (pendingMembers.length > 0 && pendingSequenceId !== table.logicalSequenceId) yield* flush();
      pendingSequenceId = table.logicalSequenceId;
      pendingMembers.push(Object.freeze({
        element: element as TableBodyElement,
        logicalRowOffset: table.logicalRowOffset,
        logicalTotalRows: table.logicalTotalRows,
        rowCount: table.rowCount,
      }));
      continue;
    }
    yield* flush();
    yield Object.freeze({ kind: 'body-element', element });
  }
  yield* flush();
}

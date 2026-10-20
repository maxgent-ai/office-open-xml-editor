import type {
  SourceRef,
  TableEdgeInputs,
  TableLayoutInput,
  TableRowLayoutInput,
} from './types.js';
import { firstAuthoredTableBorder } from './table-border-layer.js';
import {
  addExactLengthKeys,
  compareExactLengthKeys,
  divideExactLengthKey,
  exactLengthKeyFromNumber,
  exactLengthKeyToNumber,
  subtractExactLengthKeys,
  type ExactLengthKey,
} from './exact-length.js';

/** One authored source row inside a §17.4.37 logical group. It extends only
 * TableRowLayoutInput; the group-specific facts are additive and REQUIRED so a
 * consumer cannot mistake a group row for a full-grid row. The row's own
 * `exceptionBorders` is fixed null because the folded source-table border layer
 * is retained separately (and unambiguously) in `sourceTableEdges`. */
export interface AdjacentGroupRowFacts extends TableRowLayoutInput {
  readonly exceptionBorders: null;
  /** Union-grid interval [start, end) owned by this row's authored source tbl.
   * Retaining it keeps each source table authoritative for its own outer
   * border and spacing ownership even after the grids are unioned. */
  readonly sourceOuterColumnStart: number;
  readonly sourceOuterColumnEnd: number;
  /** The source table's tblBorders/tblPrEx layer, folded (and bidi-oriented into
   * the group frame) for this row: outer edges apply only at the source table's
   * true edges, every source seam resolves through insideH. */
  readonly sourceTableEdges: TableEdgeInputs;
}

/** The transient union grid produced from a §17.4.37 adjacent-table group. This
 * is a DISCRIMINATED intermediate (`kind: 'adjacent-table-group-grid'`) that is
 * intentionally NOT assignable to TableLayoutInput: it has no `ordinaryFlow` and
 * no whole-table `borders`, so a paint/pagination consumer must handle the group
 * grid explicitly rather than silently treating it as an ordinary table. */
export interface AdjacentTableGroupLayoutInput {
  readonly kind: 'adjacent-table-group-grid';
  readonly id: string;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly alignment: 'left' | 'center' | 'right';
  readonly indentPt: number;
  readonly bidiVisual: boolean;
  readonly columnWidthsPt: readonly number[];
  readonly columnWidthKeys: readonly (ExactLengthKey | null)[];
  readonly rows: readonly AdjacentGroupRowFacts[];
}

function physicalEdges(
  edges: TableEdgeInputs,
  sourceBidiVisual: boolean,
  groupBidiVisual: boolean,
): TableEdgeInputs {
  if (sourceBidiVisual === groupBidiVisual) return edges;
  return Object.freeze({ ...edges, left: edges.right, right: edges.left });
}

/** Fold a source table's whole-table borders and this row's tblPrEx exception
 * borders into one edge set oriented to the group frame. The compatibility
 * border layer owns the none/nil distinction. */
function foldSourceTableEdges(
  input: TableLayoutInput,
  row: TableRowLayoutInput,
  groupBidiVisual: boolean,
): TableEdgeInputs {
  const table = physicalEdges(input.borders, input.bidiVisual, groupBidiVisual);
  const exception = row.exceptionBorders == null
    ? null
    : physicalEdges(row.exceptionBorders, input.bidiVisual, groupBidiVisual);
  if (!exception) return table;
  return Object.freeze({
    top: firstAuthoredTableBorder(exception.top, table.top),
    right: firstAuthoredTableBorder(exception.right, table.right),
    bottom: firstAuthoredTableBorder(exception.bottom, table.bottom),
    left: firstAuthoredTableBorder(exception.left, table.left),
    insideH: firstAuthoredTableBorder(exception.insideH, table.insideH),
    insideV: firstAuthoredTableBorder(exception.insideV, table.insideV),
  });
}

interface GroupBoundary {
  readonly position: ExactLengthKey;
  readonly sym: number;
  readonly identity: string;
}

type SymbolNode =
  | Readonly<{ kind: 'zero' }>
  | Readonly<{ kind: 'token' }>
  | Readonly<{ kind: 'add' | 'sub'; left: number; right: number }>
  | Readonly<{ kind: 'div'; value: number; divisor: bigint }>;

/** Combine-call-local symbolic expressions. Nodes are interned by collision-free
 * structural keys containing only immediate integer child ids. The fixed O(1)
 * rewrites below are deliberately not an algebraic normalizer: no expression is
 * expanded, commuted, or reassociated. Consequently the resource bound is
 * O(unknown tracks + total row boundaries), with O(1) work and storage per
 * requested symbolic operation. */
class SymbolDag {
  readonly nodes: SymbolNode[] = [Object.freeze({ kind: 'zero' })];
  private readonly interned = new Map<string, number>([['Z', 0]]);

  private intern(key: string, node: SymbolNode): number {
    const existing = this.interned.get(key);
    if (existing !== undefined) return existing;
    const id = this.nodes.length;
    this.nodes.push(Object.freeze(node));
    this.interned.set(key, id);
    return id;
  }

  token(member: number, track: number): number {
    return this.intern(`T:${member}:${track}`, { kind: 'token' });
  }

  add(left: number, right: number): number {
    if (left === 0) return right;
    if (right === 0) return left;
    const leftNode = this.nodes[left]!;
    const rightNode = this.nodes[right]!;
    if (leftNode.kind === 'sub' && leftNode.right === right) return leftNode.left;
    if (rightNode.kind === 'sub' && rightNode.right === left) return rightNode.left;
    return this.intern(`A:${left}:${right}`, { kind: 'add', left, right });
  }

  subtract(left: number, right: number): number {
    if (left === right) return 0;
    if (right === 0) return left;
    const rightNode = this.nodes[right]!;
    if (rightNode.kind === 'sub' && rightNode.left === left) return rightNode.right;
    return this.intern(`S:${left}:${right}`, { kind: 'sub', left, right });
  }

  divide(value: number, divisor: bigint): number {
    if (value === 0) return 0;
    return this.intern(`D:${value}:${divisor}`, { kind: 'div', value, divisor });
  }
}

function boundary(position: ExactLengthKey, sym = 0): GroupBoundary {
  return Object.freeze({ position, sym, identity: `${position}|${sym}` });
}

function addBoundaries(
  dag: SymbolDag,
  left: GroupBoundary,
  right: GroupBoundary,
): GroupBoundary {
  return boundary(addExactLengthKeys(left.position, right.position), dag.add(left.sym, right.sym));
}

function subtractBoundaries(
  dag: SymbolDag,
  left: GroupBoundary,
  right: GroupBoundary,
): GroupBoundary {
  return boundary(
    subtractExactLengthKeys(left.position, right.position),
    dag.subtract(left.sym, right.sym),
  );
}

function divideBoundary(dag: SymbolDag, value: GroupBoundary, divisor: bigint): GroupBoundary {
  return boundary(divideExactLengthKey(value.position, divisor), dag.divide(value.sym, divisor));
}

function cumulative(
  dag: SymbolDag,
  input: TableLayoutInput,
  memberIndex: number,
): readonly GroupBoundary[] {
  const result: GroupBoundary[] = [boundary('0/1')];
  input.columnWidthsPt.forEach((width, trackIndex) => {
    const authoredKey = input.columnWidthKeys?.[trackIndex];
    const numericKey = exactLengthKeyFromNumber(width) ?? '0/1';
    const track = authoredKey === null
      ? boundary(numericKey, dag.token(memberIndex, trackIndex))
      : boundary(authoredKey ?? numericKey);
    result.push(addBoundaries(dag, result.at(-1)!, track));
  });
  return Object.freeze(result);
}

function rowPhysicalShift(
  dag: SymbolDag,
  row: TableRowLayoutInput,
  sourceWidth: GroupBoundary,
  groupWidth: GroupBoundary,
): GroupBoundary {
  const difference = subtractBoundaries(dag, groupWidth, sourceWidth);
  if (row.alignment === 'right') return difference;
  if (row.alignment === 'center') return divideBoundary(dag, difference, 2n);
  return boundary('0/1');
}

function physicalBoundary(
  dag: SymbolDag,
  logicalBoundary: GroupBoundary,
  sourceWidth: GroupBoundary,
  bidiVisual: boolean,
  shift: GroupBoundary,
): GroupBoundary {
  return addBoundaries(
    dag,
    shift,
    bidiVisual ? subtractBoundaries(dag, sourceWidth, logicalBoundary) : logicalBoundary,
  );
}

interface RowPlan {
  readonly input: TableLayoutInput;
  readonly row: TableRowLayoutInput;
  /** Group-logical position of every source boundary index (0..source columns),
   * in SOURCE order (ascending source index). */
  readonly groupBoundaries: readonly GroupBoundary[];
  /** True when the source table's bidi differs from the group frame, so its
   * source-order boundaries descend along the group-logical axis. Derived from
   * the bidi flags, never inferred from coordinates (which cannot distinguish a
   * mirrored table from an ordinary one at coincident zero-width tracks). */
  readonly descending: boolean;
}

/** ECMA-376 Part 1 §17.4.37 makes adjacent same-style ordinary-flow tbl
 * siblings one logical table. Build that transient union grid before border
 * resolution and pagination so source seams receive insideH semantics and one
 * continuation cursor spans every authored row. SourceRef and node ids remain
 * untouched.
 *
 * WHY (flag for final architecture review): the available specifications do not define how
 * independently authored tblGrid/alignment/bidiVisual payloads are reconciled
 * into a single physical grid. This union is a deterministic internal layout
 * policy that builds the transient column topology for seam-border and
 * pagination geometry, while each source table stays authoritative for its own
 * first/last-column, vertical banding, and corner membership through the
 * retained AdjacentGroupRowFacts outer interval. Column boundaries are a MULTISET
 * union (max multiplicity per exact coordinate) so a source table's zero-width
 * tracks survive rather than collapsing as they would in a set. */
export function combineAdjacentTableLayoutInputs(
  groupId: string,
  inputs: readonly TableLayoutInput[],
): AdjacentTableGroupLayoutInput {
  if (groupId.length === 0) throw new RangeError('Adjacent table group id must not be empty');
  if (inputs.length === 0) throw new RangeError('Adjacent table group requires at least one table');
  // `word-positioned-table-adjacency-exclusion`
  if (inputs.some((input) => !input.ordinaryFlow)) {
    throw new Error('An absolutely positioned table cannot join an adjacent table group');
  }
  const first = inputs[0]!;
  const groupBidi = first.bidiVisual;
  const dag = new SymbolDag();
  const zeroBoundary = boundary('0/1');
  const sourceBoundaries = inputs.map((input, member) => cumulative(dag, input, member));
  const sourceWidths = sourceBoundaries.map((boundaries) => boundaries.at(-1) ?? zeroBoundary);
  const groupWidth = sourceWidths.reduce((maximum, width) => (
    compareExactLengthKeys(width.position, maximum.position) > 0 ? width : maximum
  ), zeroBoundary);

  const groupLogical = (
    sourceLogicalBoundary: GroupBoundary,
    sourceWidth: GroupBoundary,
    sourceBidi: boolean,
    shift: GroupBoundary,
  ): GroupBoundary => {
    const physical = physicalBoundary(dag, sourceLogicalBoundary, sourceWidth, sourceBidi, shift);
    return groupBidi ? subtractBoundaries(dag, groupWidth, physical) : physical;
  };

  const rowPlans: RowPlan[] = [];
  inputs.forEach((input, inputIndex) => {
    const boundaries = sourceBoundaries[inputIndex]!;
    const sourceWidth = sourceWidths[inputIndex]!;
    const descending = input.bidiVisual !== groupBidi;
    input.rows.forEach((row) => {
      const shift = rowPhysicalShift(dag, row, sourceWidth, groupWidth);
      const groupBoundaries = boundaries.map((boundary) => (
        groupLogical(boundary, sourceWidth, input.bidiVisual, shift)
      ));
      rowPlans.push({ input, row, groupBoundaries, descending });
    });
  });

  // Multiset union: the max multiplicity of each exact coordinate across rows,
  // seeded with the group frame [0, groupWidth]. Two boundaries a source row
  // places at the same coordinate (a zero-width track) survive because their
  // multiplicity, not just their coordinate, is retained.
  const maxCount = new Map<string, { boundary: GroupBoundary; count: number }>();
  for (const frame of [zeroBoundary, groupWidth]) {
    maxCount.set(frame.identity, { boundary: frame, count: 1 });
  }
  for (const plan of rowPlans) {
    const counts = new Map<string, { boundary: GroupBoundary; count: number }>();
    for (const position of plan.groupBoundaries) {
      const current = counts.get(position.identity);
      counts.set(position.identity, { boundary: position, count: (current?.count ?? 0) + 1 });
    }
    for (const [identity, entry] of counts) {
      const previous = maxCount.get(identity);
      if (entry.count > (previous?.count ?? 0)) maxCount.set(identity, entry);
    }
  }

  interface PositionBucket {
    readonly position: ExactLengthKey;
    readonly identities: Map<string, { boundary: GroupBoundary; count: number }>;
    readonly edges: Map<string, Set<string>>;
    readonly firstSeen: Map<string, number>;
  }
  const buckets = new Map<ExactLengthKey, PositionBucket>();
  const bucketFor = (position: ExactLengthKey): PositionBucket => {
    let bucket = buckets.get(position);
    if (!bucket) {
      bucket = {
        position,
        identities: new Map(),
        edges: new Map(),
        firstSeen: new Map(),
      };
      buckets.set(position, bucket);
    }
    return bucket;
  };
  for (const [identity, entry] of maxCount) {
    bucketFor(entry.boundary.position).identities.set(identity, entry);
  }

  // Equal numeric positions cannot be ordered by opaque symbolic ids. Each row
  // instead contributes its directed group-axis order. Consecutive duplicates
  // are one identity occurrence for ordering, while max multiplicity remains a
  // separate multiset property. A cycle is unreachable for supported inputs:
  // every row is a monotone projection onto the same physical group axis, and
  // the fixed DAG rewrites preserve equal symbolic boundaries across mirror,
  // right-align, and center transforms. A cycle therefore signals a violated
  // invariant and must fail rather than be resolved by a heuristic.
  let firstSeenRank = 0;
  for (const plan of rowPlans) {
    const axisOrder = plan.descending ? [...plan.groupBoundaries].reverse() : plan.groupBoundaries;
    let runPosition: ExactLengthKey | null = null;
    let previousIdentity: string | null = null;
    for (const current of axisOrder) {
      const bucket = bucketFor(current.position);
      if (!bucket.firstSeen.has(current.identity)) {
        bucket.firstSeen.set(current.identity, firstSeenRank++);
      }
      if (runPosition !== current.position) {
        runPosition = current.position;
        previousIdentity = null;
      }
      if (previousIdentity !== null && previousIdentity !== current.identity) {
        let outgoing = bucket.edges.get(previousIdentity);
        if (!outgoing) {
          outgoing = new Set();
          bucket.edges.set(previousIdentity, outgoing);
        }
        outgoing.add(current.identity);
      }
      previousIdentity = current.identity;
    }
  }
  for (const bucket of buckets.values()) {
    for (const identity of bucket.identities.keys()) {
      if (!bucket.firstSeen.has(identity)) bucket.firstSeen.set(identity, firstSeenRank++);
    }
  }

  const orderedBuckets = [...buckets.values()].sort((left, right) => (
    compareExactLengthKeys(left.position, right.position)
  ));
  const unionSeq: GroupBoundary[] = [];
  const unionIndexAt = new Map<string, number>();
  for (const bucket of orderedBuckets) {
    const indegree = new Map([...bucket.identities.keys()].map((identity) => [identity, 0]));
    for (const outgoing of bucket.edges.values()) {
      for (const target of outgoing) indegree.set(target, (indegree.get(target) ?? 0) + 1);
    }
    const ready: string[] = [];
    const pushReady = (identity: string): void => {
      ready.push(identity);
      let child = ready.length - 1;
      while (child > 0) {
        const parent = Math.floor((child - 1) / 2);
        if (bucket.firstSeen.get(ready[parent]!)! <= bucket.firstSeen.get(ready[child]!)!) break;
        [ready[parent], ready[child]] = [ready[child]!, ready[parent]!];
        child = parent;
      }
    };
    const popReady = (): string => {
      const first = ready[0]!;
      const last = ready.pop()!;
      if (ready.length > 0) {
        ready[0] = last;
        let parent = 0;
        while (true) {
          const left = parent * 2 + 1;
          const right = left + 1;
          if (left >= ready.length) break;
          let child = left;
          if (right < ready.length
            && bucket.firstSeen.get(ready[right]!)! < bucket.firstSeen.get(ready[left]!)!) {
            child = right;
          }
          if (bucket.firstSeen.get(ready[parent]!)! <= bucket.firstSeen.get(ready[child]!)!) break;
          [ready[parent], ready[child]] = [ready[child]!, ready[parent]!];
          parent = child;
        }
      }
      return first;
    };
    for (const identity of bucket.identities.keys()) {
      if (indegree.get(identity) === 0) pushReady(identity);
    }
    const ordered: string[] = [];
    while (ready.length > 0) {
      const identity = popReady();
      ordered.push(identity);
      for (const target of bucket.edges.get(identity) ?? []) {
        const remaining = indegree.get(target)! - 1;
        indegree.set(target, remaining);
        if (remaining === 0) pushReady(target);
      }
    }
    if (ordered.length !== bucket.identities.size) {
      throw new Error(`Adjacent table symbolic boundary ordering cycle at ${bucket.position}`);
    }
    for (const identity of ordered) {
      const { boundary: position, count } = bucket.identities.get(identity)!;
      unionIndexAt.set(identity, unionSeq.length);
      for (let occurrence = 0; occurrence < count; occurrence += 1) unionSeq.push(position);
    }
  }
  const columnWidthKeys = unionSeq.slice(1).map((right, index) => {
    const left = unionSeq[index]!;
    return right.sym === left.sym ? subtractExactLengthKeys(right.position, left.position) : null;
  });
  const columnWidthsPt = unionSeq.slice(1).map((right, index) => exactLengthKeyToNumber(
    subtractExactLengthKeys(right.position, unionSeq[index]!.position),
  ));

  // Map each source boundary index to a union boundary index by occurrence
  // identity, scanning boundaries in SOURCE order and counting local occurrences
  // of each coordinate. An ascending source takes the group's low occurrences
  // (base + local); a descending (bidi-mismatched) source takes the high ones
  // (base + multiplicity - 1 - local) so its mirrored zero-width tracks tile the
  // union in the correct physical order.
  const unionIndicesFor = (
    groupBoundaries: readonly GroupBoundary[],
    descending: boolean,
  ): number[] => {
    const local = new Map<string, number>();
    const indices = new Array<number>(groupBoundaries.length);
    groupBoundaries.forEach((position, sourceIndex) => {
      const localOrdinal = local.get(position.identity) ?? 0;
      local.set(position.identity, localOrdinal + 1);
      const base = unionIndexAt.get(position.identity)!;
      const multiplicity = maxCount.get(position.identity)!.count;
      indices[sourceIndex] = descending
        ? base + (multiplicity - 1 - localOrdinal)
        : base + localOrdinal;
    });
    return indices;
  };

  let logicalRowIndex = 0;
  const rows = rowPlans.map((plan): AdjacentGroupRowFacts => {
    const { input, row, groupBoundaries, descending } = plan;
    const unionIndexOf = unionIndicesFor(groupBoundaries, descending);
    const outerStart = unionIndexOf[0]!;
    const outerEnd = unionIndexOf[groupBoundaries.length - 1]!;
    const cells = row.cells.map((cell) => {
      const startIndex = unionIndexOf[cell.columnStart];
      const endIndex = unionIndexOf[cell.columnStart + cell.columnSpan];
      if (startIndex == null || endIndex == null) {
        throw new RangeError(`Table cell ${cell.id} exceeds its authored grid`);
      }
      const columnStart = Math.min(startIndex, endIndex);
      const columnEnd = Math.max(startIndex, endIndex);
      if (columnEnd <= columnStart) {
        throw new Error(`Table cell ${cell.id} cannot be mapped into the logical group grid`);
      }
      const borders = physicalEdges(cell.borders, input.bidiVisual, groupBidi);
      return Object.freeze({
        ...cell,
        columnStart,
        columnSpan: columnEnd - columnStart,
        borders,
      });
    });
    return Object.freeze({
      ...row,
      logicalRowIndex: logicalRowIndex++,
      // The group grid is not a TableLayoutInput; the folded source-table border
      // layer lives in `sourceTableEdges`, so the row's own exceptionBorders slot
      // is fixed null to avoid a second, ambiguous border source.
      exceptionBorders: null,
      sourceTableEdges: foldSourceTableEdges(input, row, groupBidi),
      // Re-orient the source's leading-edge indent into the group frame when the
      // source and group bidi differ; the group frame owns final placement.
      indentPt: input.bidiVisual === groupBidi ? row.indentPt : -row.indentPt,
      sourceOuterColumnStart: Math.min(outerStart, outerEnd),
      sourceOuterColumnEnd: Math.max(outerStart, outerEnd),
      cells: Object.freeze(cells),
    });
  });

  return Object.freeze({
    kind: 'adjacent-table-group-grid',
    id: groupId,
    source: first.source,
    flowDomainId: `${first.flowDomainId}:adjacent-group:${groupId}`,
    alignment: first.alignment,
    indentPt: first.indentPt,
    bidiVisual: groupBidi,
    columnWidthsPt: Object.freeze(columnWidthsPt),
    columnWidthKeys: Object.freeze(columnWidthKeys),
    rows: Object.freeze(rows),
  });
}

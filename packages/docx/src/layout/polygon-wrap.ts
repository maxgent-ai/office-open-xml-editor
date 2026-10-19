import {
  compareExactRational,
  decodeBinary64,
  exactRationalKey,
  exactRationalToNumber,
  exactRationalToNumberUp,
  midpointExactRational,
  normalizeExactRational,
  subtractExactRational,
  type ExactRational,
} from './exact-geometry.js';

export type PolygonPoint = Readonly<{ xPt: number; yPt: number }>;

type CompiledEdge = Readonly<{
  from: PolygonPoint;
  to: PolygonPoint;
  minYPt: number;
  maxYPt: number;
  slopeXPerY: number;
  interceptX: number;
}>;

export type AffineX = Readonly<{ slope: number; intercept: number }>;

interface ContourYSpan {
  readonly yTopPt: number;
  readonly yBottomPt: number;
}

interface CompiledContourSpan extends ContourYSpan {
  readonly left: CompiledEdge;
  readonly right: CompiledEdge;
}

interface ContourSpanIndex<T extends ContourYSpan = CompiledContourSpan> {
  readonly centerYPt: number;
  readonly crossingByTop: readonly T[];
  readonly crossingByBottom: readonly T[];
  readonly below: ContourSpanIndex<T> | null;
  readonly above: ContourSpanIndex<T> | null;
}

export interface PolygonAffineInterval {
  readonly left: AffineX;
  readonly right: AffineX;
}

export interface CompiledPolygonWrap {
  readonly kind: 'tight' | 'through';
  readonly edges: readonly CompiledEdge[];
  readonly eventYPts: readonly number[];
  readonly contourSpans: readonly CompiledContourSpan[];
  readonly contourSpanIndex: ContourSpanIndex | null;
  readonly intersectionCount: number;
  readonly compileOrderComparisonCount: number;
  readonly compilePairMembershipVisitCount: number;
  readonly polygonLeftPt: number;
  readonly polygonRightPt: number;
  readonly polygonTopPt: number;
  readonly polygonBottomPt: number;
  readonly padLeftPt: number;
  readonly padRightPt: number;
  readonly padTopPt: number;
  readonly padBottomPt: number;
}

export interface PolygonCompileInput {
  readonly kind: 'tight' | 'through';
  readonly imageKey: string;
  readonly points: readonly PolygonPoint[] | undefined;
  readonly xLeftPt: number;
  readonly xRightPt: number;
  readonly yTopPt: number;
  readonly yBottomPt: number;
}

export function assertValidPolygonCompileInput(
  input: PolygonCompileInput,
): asserts input is PolygonCompileInput & { readonly points: readonly PolygonPoint[] } {
  if (!input.points || input.points.length < 3
    || input.points.some((point) =>
      !Number.isFinite(point.xPt) || !Number.isFinite(point.yPt))) {
    throw new Error(`Invalid ${input.kind} wrapPolygon for ${input.imageKey}`);
  }
  if (![input.xLeftPt, input.xRightPt, input.yTopPt, input.yBottomPt].every(Number.isFinite)
    || input.xRightPt < input.xLeftPt || input.yBottomPt < input.yTopPt) {
    throw new Error(`Invalid finite wrap bounds for ${input.imageKey}`);
  }
}

export interface PolygonInterval {
  l: number;
  r: number;
}

interface ExactPoint {
  readonly x: bigint;
  readonly y: bigint;
}

interface ExactEdge {
  readonly index: number;
  readonly from: ExactPoint;
  readonly to: ExactPoint;
  readonly minY: bigint;
  readonly maxY: bigint;
  readonly dx: bigint;
  readonly dy: bigint;
  readonly c: bigint;
}

interface ExactIntersection {
  readonly y: ExactRational;
  readonly contact: 'shared-endpoint' | 'horizontal' | 'active-crossing';
}

interface ExactContourSpan {
  readonly yTop: ExactRational;
  readonly yBottom: ExactRational;
  readonly leftEdge: number;
  readonly rightEdge: number;
}

interface ExactContourSpanIndex {
  readonly centerY: ExactRational;
  readonly crossingByTop: readonly ExactContourSpan[];
  readonly crossingByBottom: readonly ExactContourSpan[];
  readonly below: ExactContourSpanIndex | null;
  readonly above: ExactContourSpanIndex | null;
}

interface ExactCompiledPolygon {
  readonly scaleExponent: number;
  readonly edges: readonly ExactEdge[];
  readonly eventYs: readonly ExactRational[];
  readonly spans: readonly ExactContourSpan[];
  readonly spanIndex: ExactContourSpanIndex | null;
  readonly polygonLeft: ExactRational;
  readonly polygonRight: ExactRational;
  readonly polygonTop: ExactRational;
  readonly polygonBottom: ExactRational;
  readonly padLeft: ExactRational;
  readonly padRight: ExactRational;
  readonly padTop: ExactRational;
  readonly padBottom: ExactRational;
}

const exactCompiledPolygons = new WeakMap<CompiledPolygonWrap, ExactCompiledPolygon>();

export interface ExactAffineX {
  readonly slope: ExactRational;
  readonly intercept: ExactRational;
}

export interface ExactPolygonAffineInterval {
  readonly left: ExactAffineX;
  readonly right: ExactAffineX;
}

function sameExactPoint(left: ExactPoint, right: ExactPoint): boolean {
  return left.x === right.x && left.y === right.y;
}

function cross(
  leftX: bigint,
  leftY: bigint,
  rightX: bigint,
  rightY: bigint,
): bigint {
  return leftX * rightY - leftY * rightX;
}

function parameterIsOnSegment(numerator: bigint, denominator: bigint): boolean {
  return denominator > 0n
    ? numerator >= 0n && numerator <= denominator
    : numerator <= 0n && numerator >= denominator;
}

function exactSegmentIntersection(
  left: ExactEdge,
  right: ExactEdge,
): ExactIntersection | null {
  const shared = sameExactPoint(left.from, right.from) || sameExactPoint(left.from, right.to)
    ? left.from
    : sameExactPoint(left.to, right.from) || sameExactPoint(left.to, right.to)
      ? left.to
      : null;
  if (shared) {
    return Object.freeze({
      y: normalizeExactRational(shared.y, 1n),
      contact: 'shared-endpoint',
    });
  }
  const leftDx = left.to.x - left.from.x;
  const leftDy = left.to.y - left.from.y;
  const rightDx = right.to.x - right.from.x;
  const rightDy = right.to.y - right.from.y;
  const denominator = cross(leftDx, leftDy, rightDx, rightDy);
  if (denominator === 0n) return null;
  const betweenX = right.from.x - left.from.x;
  const betweenY = right.from.y - left.from.y;
  const leftParameter = cross(betweenX, betweenY, rightDx, rightDy);
  const rightParameter = cross(betweenX, betweenY, leftDx, leftDy);
  if (!parameterIsOnSegment(leftParameter, denominator)
    || !parameterIsOnSegment(rightParameter, denominator)) return null;
  return Object.freeze({
    y: normalizeExactRational(
      left.from.y * denominator + leftDy * leftParameter,
      denominator,
    ),
    contact: leftDy === 0n || rightDy === 0n ? 'horizontal' : 'active-crossing',
  });
}

interface EdgeIntersection {
  readonly y: ExactRational;
  readonly contact: ExactIntersection['contact'];
  readonly leftEdge: number;
  readonly rightEdge: number;
}

interface CompiledContourSweep {
  readonly spans: readonly CompiledContourSpan[];
  readonly exactSpans: readonly ExactContourSpan[];
  readonly orderComparisonCount: number;
  readonly pairMembershipVisitCount: number;
}

function pairKey(leftEdge: number, rightEdge: number): string {
  return `${leftEdge}:${rightEdge}`;
}

function pairsAtIndices(
  order: readonly number[],
  pairIndices: ReadonlySet<number>,
  visitPair: () => void,
): Set<string> {
  const pairs = new Set<string>();
  for (const pairIndex of pairIndices) {
    if (pairIndex < 0 || pairIndex >= Math.floor(order.length / 2)) continue;
    visitPair();
    const leftEdge = order[pairIndex * 2]!;
    const rightEdge = order[pairIndex * 2 + 1]!;
    pairs.add(pairKey(leftEdge, rightEdge));
  }
  return pairs;
}

function buildContourSpanIndex<T extends ContourYSpan>(
  spans: readonly T[],
): ContourSpanIndex<T> | null {
  const build = (ordered: readonly T[]): ContourSpanIndex<T> | null => {
    if (ordered.length === 0) return null;
    const centerYPt = ordered[Math.floor(ordered.length / 2)]!.yTopPt;
    const below: T[] = [];
    const above: T[] = [];
    const crossing: T[] = [];
    for (const span of ordered) {
      if (span.yBottomPt <= centerYPt) below.push(span);
      else if (span.yTopPt > centerYPt) above.push(span);
      else crossing.push(span);
    }
    // Median-top authority plus strict half-open classification confines each
    // child to one side of the median even when interval midpoints round alike.
    return Object.freeze({
      centerYPt,
      crossingByTop: Object.freeze(crossing),
      crossingByBottom: Object.freeze(
        crossing.slice().sort((left, right) => right.yBottomPt - left.yBottomPt),
      ),
      below: build(below),
      above: build(above),
    });
  };
  return build(spans.slice().sort((left, right) =>
    left.yTopPt - right.yTopPt || left.yBottomPt - right.yBottomPt));
}

function queryContourSpans<T extends ContourYSpan>(
  node: ContourSpanIndex<T> | null,
  queryTop: number,
  queryBottom: number,
  result: T[],
  diagnostics?: { nodeVisitCount: number },
): void {
  if (!node || queryBottom <= queryTop) return;
  if (diagnostics) diagnostics.nodeVisitCount += 1;
  if (queryBottom <= node.centerYPt) {
    for (const span of node.crossingByTop) {
      if (span.yTopPt >= queryBottom) break;
      result.push(span);
    }
    queryContourSpans(node.below, queryTop, queryBottom, result, diagnostics);
    return;
  }
  if (queryTop >= node.centerYPt) {
    for (const span of node.crossingByBottom) {
      if (span.yBottomPt <= queryTop) break;
      result.push(span);
    }
    queryContourSpans(node.above, queryTop, queryBottom, result, diagnostics);
    return;
  }
  result.push(...node.crossingByTop);
  queryContourSpans(node.below, queryTop, queryBottom, result, diagnostics);
  queryContourSpans(node.above, queryTop, queryBottom, result, diagnostics);
}

/** Internal module test seam; it is intentionally not re-exported by the package root. */
export function __test_contourSpanIndexDiagnostics(
  spans: readonly Readonly<ContourYSpan>[],
  queryTop: number,
  queryBottom: number,
): Readonly<{
  height: number;
  matchingSpanCount: number;
  queryNodeVisitCount: number;
}> {
  const index = buildContourSpanIndex(spans);
  const heightOf = (node: ContourSpanIndex<Readonly<ContourYSpan>> | null): number =>
    node ? 1 + Math.max(heightOf(node.below), heightOf(node.above)) : 0;
  const matches: Readonly<ContourYSpan>[] = [];
  const queryDiagnostics = { nodeVisitCount: 0 };
  queryContourSpans(index, queryTop, queryBottom, matches, queryDiagnostics);
  return Object.freeze({
    height: heightOf(index),
    matchingSpanCount: matches.length,
    queryNodeVisitCount: queryDiagnostics.nodeVisitCount,
  });
}

function exactLatticeToNumber(value: ExactRational, scaleExponent: number): number {
  return exactRationalToNumber(scaleExponent >= 0
    ? {
        numerator: value.numerator << BigInt(scaleExponent),
        denominator: value.denominator,
      }
    : {
        numerator: value.numerator,
        denominator: value.denominator << BigInt(-scaleExponent),
      });
}

function exactLatticeToNumberUp(value: ExactRational, scaleExponent: number): number {
  return exactRationalToNumberUp(scaleExponent >= 0
    ? {
        numerator: value.numerator << BigInt(scaleExponent),
        denominator: value.denominator,
      }
    : {
        numerator: value.numerator,
        denominator: value.denominator << BigInt(-scaleExponent),
      });
}

function exactLatticeValue(
  value: number,
  scaleExponent: number,
): ExactRational {
  const decoded = decodeBinary64(value);
  const exponent = decoded.exponent - scaleExponent;
  return exponent >= 0
    ? {
        numerator: decoded.coefficient << BigInt(exponent),
        denominator: 1n,
      }
    : {
        numerator: decoded.coefficient,
        denominator: 1n << BigInt(-exponent),
      };
}

function compareExactEdgeAtY(
  left: ExactEdge,
  right: ExactEdge,
  y: ExactRational,
): number {
  const leftNumerator = left.dx * y.numerator - left.c * y.denominator;
  const rightNumerator = right.dx * y.numerator - right.c * y.denominator;
  const delta = leftNumerator * right.dy - rightNumerator * left.dy;
  return delta < 0n ? -1 : delta > 0n ? 1 : 0;
}

function buildExactContourSpanIndex(
  spans: readonly ExactContourSpan[],
): ExactContourSpanIndex | null {
  const build = (ordered: readonly ExactContourSpan[]): ExactContourSpanIndex | null => {
    if (ordered.length === 0) return null;
    const centerY = ordered[Math.floor(ordered.length / 2)]!.yTop;
    const below: ExactContourSpan[] = [];
    const above: ExactContourSpan[] = [];
    const crossing: ExactContourSpan[] = [];
    for (const span of ordered) {
      if (compareExactRational(span.yBottom, centerY) <= 0) below.push(span);
      else if (compareExactRational(span.yTop, centerY) > 0) above.push(span);
      else crossing.push(span);
    }
    return Object.freeze({
      centerY,
      crossingByTop: Object.freeze(crossing),
      crossingByBottom: Object.freeze(crossing.slice().sort((left, right) =>
        compareExactRational(right.yBottom, left.yBottom))),
      below: build(below),
      above: build(above),
    });
  };
  return build(spans.slice().sort((left, right) =>
    compareExactRational(left.yTop, right.yTop)
      || compareExactRational(left.yBottom, right.yBottom)));
}

function queryExactContourSpans(
  node: ExactContourSpanIndex | null,
  queryTop: ExactRational,
  queryBottom: ExactRational,
  result: ExactContourSpan[],
): void {
  if (!node || compareExactRational(queryBottom, queryTop) <= 0) return;
  if (compareExactRational(queryBottom, node.centerY) <= 0) {
    for (const span of node.crossingByTop) {
      if (compareExactRational(span.yTop, queryBottom) >= 0) break;
      result.push(span);
    }
    queryExactContourSpans(node.below, queryTop, queryBottom, result);
    return;
  }
  if (compareExactRational(queryTop, node.centerY) >= 0) {
    for (const span of node.crossingByBottom) {
      if (compareExactRational(span.yBottom, queryTop) <= 0) break;
      result.push(span);
    }
    queryExactContourSpans(node.above, queryTop, queryBottom, result);
    return;
  }
  result.push(...node.crossingByTop);
  queryExactContourSpans(node.below, queryTop, queryBottom, result);
  queryExactContourSpans(node.above, queryTop, queryBottom, result);
}

function compileContourSpans(
  edges: readonly CompiledEdge[],
  exactEdges: readonly ExactEdge[],
  eventYs: readonly ExactRational[],
  scaleExponent: number,
  intersections: readonly EdgeIntersection[],
): CompiledContourSweep {
  const starts = new Map<string, number[]>();
  const ends = new Map<string, number[]>();
  exactEdges.forEach((edge, index) => {
    if (edge.minY === edge.maxY) return;
    const startKey = exactRationalKey(normalizeExactRational(edge.minY, 1n));
    const endKey = exactRationalKey(normalizeExactRational(edge.maxY, 1n));
    const starting = starts.get(startKey);
    if (starting) starting.push(index);
    else starts.set(startKey, [index]);
    const ending = ends.get(endKey);
    if (ending) ending.push(index);
    else ends.set(endKey, [index]);
  });
  const crossingsByY = new Map<string, Set<number>>();
  for (const intersection of intersections) {
    if (intersection.contact !== 'active-crossing') continue;
    const key = exactRationalKey(intersection.y);
    let edgeSet = crossingsByY.get(key);
    if (!edgeSet) crossingsByY.set(key, edgeSet = new Set());
    edgeSet.add(intersection.leftEdge);
    edgeSet.add(intersection.rightEdge);
  }
  const activeOrder: number[] = [];
  const noCrossingEdges: ReadonlySet<number> = new Set();
  const openPairs = new Map<string, {
    leftEdge: number;
    rightEdge: number;
    yTop: ExactRational;
  }>();
  const exactSpans: ExactContourSpan[] = [];
  const spans: CompiledContourSpan[] = [];
  let orderComparisonCount = 0;
  let pairMembershipVisitCount = 0;
  const syncPairs = (
    before: Set<string>,
    after: Set<string>,
    y: ExactRational,
  ): void => {
    for (const key of before) {
      if (after.has(key)) continue;
      const open = openPairs.get(key);
      if (open && compareExactRational(y, open.yTop) > 0) {
        const exactSpan = Object.freeze({
          yTop: open.yTop,
          yBottom: y,
          leftEdge: open.leftEdge,
          rightEdge: open.rightEdge,
        });
        exactSpans.push(exactSpan);
        const yTopPt = exactLatticeToNumber(open.yTop, scaleExponent);
        const yBottomPt = exactLatticeToNumber(y, scaleExponent);
        if (yBottomPt > yTopPt) {
          spans.push(Object.freeze({
            yTopPt,
            yBottomPt,
            left: edges[open.leftEdge]!,
            right: edges[open.rightEdge]!,
          }));
        }
      }
      openPairs.delete(key);
    }
    for (const key of after) {
      if (before.has(key)) continue;
      const separator = key.indexOf(':');
      openPairs.set(key, {
        leftEdge: Number(key.slice(0, separator)),
        rightEdge: Number(key.slice(separator + 1)),
        yTop: y,
      });
    }
  };
  const activePositions = new Map<number, number>();
  const appendAdjacentPairIndices = (indices: Set<number>, position: number): void => {
    indices.add(Math.floor((position - 1) / 2));
    indices.add(Math.floor(position / 2));
    indices.add(Math.floor((position + 1) / 2));
  };
  const localPairs = (indices: ReadonlySet<number>): Set<string> =>
    pairsAtIndices(activeOrder, indices, () => {
      pairMembershipVisitCount += 2;
    });
  const completePairs = (): Set<string> => {
    const pairs = new Set<string>();
    for (let position = 0; position + 1 < activeOrder.length; position += 2) {
      pairMembershipVisitCount += 2;
      pairs.add(pairKey(activeOrder[position]!, activeOrder[position + 1]!));
    }
    return pairs;
  };
  for (let eventIndex = 0; eventIndex < eventYs.length; eventIndex += 1) {
    const y = eventYs[eventIndex]!;
    const nextY = eventYs[eventIndex + 1];
    const slabY = nextY ? midpointExactRational(y, nextY) : y;
    const compareAbove = (leftEdge: number, rightEdge: number): number => {
      orderComparisonCount += 1;
      return compareExactEdgeAtY(exactEdges[leftEdge]!, exactEdges[rightEdge]!, slabY)
        || leftEdge - rightEdge;
    };
    const insertActive = (edgeIndex: number): void => {
      let low = 0;
      let high = activeOrder.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        if (compareAbove(activeOrder[middle]!, edgeIndex) <= 0) low = middle + 1;
        else high = middle;
      }
      activeOrder.splice(low, 0, edgeIndex);
      for (let index = low; index < activeOrder.length; index += 1) {
        activePositions.set(activeOrder[index]!, index);
      }
    };
    const removeActive = (edgeIndex: number): void => {
      const position = activePositions.get(edgeIndex);
      if (position === undefined) return;
      activeOrder.splice(position, 1);
      activePositions.delete(edgeIndex);
      for (let index = position; index < activeOrder.length; index += 1) {
        activePositions.set(activeOrder[index]!, index);
      }
    };
    const reorderCrossingEdges = (crossingEdges: ReadonlySet<number>): void => {
      const activeCrossingEdges = [...crossingEdges].filter((edgeIndex) =>
        activePositions.has(edgeIndex)
        && compareExactRational(
          normalizeExactRational(exactEdges[edgeIndex]!.minY, 1n),
          y,
        ) <= 0
        && compareExactRational(
          y,
          normalizeExactRational(exactEdges[edgeIndex]!.maxY, 1n),
        ) < 0);
      const positions = activeCrossingEdges
        .map((edgeIndex) => activePositions.get(edgeIndex)!)
        .sort((left, right) => left - right);
      // An exact segment intersection is only a candidate boundary. The exact
      // open-slab edge order decides whether paired membership actually changes.
      activeCrossingEdges.sort(compareAbove);
      for (let index = 0; index < positions.length; index += 1) {
        const position = positions[index]!;
        const edgeIndex = activeCrossingEdges[index]!;
        activeOrder[position] = edgeIndex;
        activePositions.set(edgeIndex, position);
      }
    };
    const eventKey = exactRationalKey(y);
    const ending = ends.get(eventKey) ?? [];
    const starting = starts.get(eventKey) ?? [];
    const crossingEdges = crossingsByY.get(eventKey) ?? noCrossingEdges;
    if (ending.length === 0 && starting.length === 0
      && (crossingEdges.size === 0 || nextY === undefined)) continue;
    const hasVertexChanges = ending.length > 0 || starting.length > 0;
    const beforePairIndices = new Set<number>();
    if (!hasVertexChanges) {
      for (const edgeIndex of crossingEdges) {
        const position = activePositions.get(edgeIndex);
        if (position !== undefined) appendAdjacentPairIndices(beforePairIndices, position);
      }
    }
    // Vertex insertions/removals can shift the parity of every later slot, so
    // only pure intersection events are safe to synchronize participant-locally.
    const before = hasVertexChanges ? completePairs() : localPairs(beforePairIndices);
    for (const edgeIndex of ending) removeActive(edgeIndex);
    for (const edgeIndex of starting) insertActive(edgeIndex);
    if (crossingEdges.size > 0 && nextY !== undefined) reorderCrossingEdges(crossingEdges);
    if (activeOrder.length % 2 !== 0) {
      throw new Error('Compiled wrapPolygon produced an odd open-slab crossing count');
    }
    if (hasVertexChanges) {
      syncPairs(before, completePairs(), y);
    } else {
      const afterPairIndices = new Set<number>();
      for (const edgeIndex of crossingEdges) {
        const position = activePositions.get(edgeIndex);
        if (position !== undefined) appendAdjacentPairIndices(afterPairIndices, position);
      }
      syncPairs(before, localPairs(afterPairIndices), y);
    }
  }
  return Object.freeze({
    spans: Object.freeze(spans),
    exactSpans: Object.freeze(exactSpans),
    orderComparisonCount,
    pairMembershipVisitCount,
  });
}

/**
 * Compile the CT_WrapPath at the float-acquisition boundary.
 *
 * ECMA-376 Part 1 §20.4.2.16 requires one start and at least two lineTo
 * points; when the last point differs from the start, closure is inferred. The
 * modulo edge below is that inferred segment. The schema does not require a
 * non-zero signed area, and signed shoelace area cancels for valid bow-ties, so
 * compilation validates the path contract rather than imposing an uncited
 * simple-polygon restriction. Binary64 coordinates are decoded to an exact
 * power-of-two lattice, so every finite-segment event and edge order has one
 * rational authority. The active-edge sweep pairs only
 * neighboring crossings in each open slab; no line query reconstructs or
 * compares inactive edge pairs. Vertex events synchronize the complete paired
 * order in O(V) each, hence O(V²) total; pure intersection events inspect only
 * participant-adjacent slots, hence O(K) pair-membership work.
 */
export function compilePolygonWrap(input: PolygonCompileInput): CompiledPolygonWrap {
  assertValidPolygonCompileInput(input);
  const points = input.points;
  const retainedPoints = Object.freeze(points.map((point) => Object.freeze({ ...point })));
  const latticeValues = [
    ...retainedPoints.flatMap((point) => [point.xPt, point.yPt]),
    input.xLeftPt,
    input.xRightPt,
    input.yTopPt,
    input.yBottomPt,
  ];
  const decodedValues = latticeValues.map(decodeBinary64);
  const nonzeroValues = decodedValues.filter(({ coefficient }) => coefficient !== 0n);
  const scaleExponent = nonzeroValues.length === 0
    ? 0
    : Math.min(...nonzeroValues.map(({ exponent }) => exponent));
  const latticeInteger = (value: number): bigint => {
    const decoded = decodeBinary64(value);
    if (decoded.coefficient === 0n) return 0n;
    return decoded.coefficient << BigInt(decoded.exponent - scaleExponent);
  };
  const exactPoints = retainedPoints.map((point): ExactPoint => Object.freeze({
    x: latticeInteger(point.xPt),
    y: latticeInteger(point.yPt),
  }));
  const edges: CompiledEdge[] = retainedPoints.map((from, index) => {
    const to = retainedPoints[(index + 1) % retainedPoints.length]!;
    const deltaY = to.yPt - from.yPt;
    const slopeXPerY = deltaY === 0 ? 0 : (to.xPt - from.xPt) / deltaY;
    return Object.freeze({
      from,
      to,
      minYPt: Math.min(from.yPt, to.yPt),
      maxYPt: Math.max(from.yPt, to.yPt),
      slopeXPerY,
      interceptX: deltaY === 0 ? from.xPt : from.xPt - slopeXPerY * from.yPt,
    });
  });
  const exactEdges: ExactEdge[] = exactPoints.map((from, index) => {
    const to = exactPoints[(index + 1) % exactPoints.length]!;
    const top = from.y <= to.y ? from : to;
    const bottom = from.y <= to.y ? to : from;
    const dx = bottom.x - top.x;
    const dy = bottom.y - top.y;
    return Object.freeze({
      index,
      from,
      to,
      minY: top.y,
      maxY: bottom.y,
      dx,
      dy,
      c: dx * top.y - dy * top.x,
    });
  });
  const intersections: EdgeIntersection[] = [];
  for (let left = 0; left < exactEdges.length; left += 1) {
    for (let right = left + 1; right < exactEdges.length; right += 1) {
      const point = exactSegmentIntersection(exactEdges[left]!, exactEdges[right]!);
      if (point) intersections.push(Object.freeze({
        y: point.y,
        contact: point.contact,
        leftEdge: left,
        rightEdge: right,
      }));
    }
  }
  let polygonLeftPt = Number.POSITIVE_INFINITY;
  let polygonRightPt = Number.NEGATIVE_INFINITY;
  let polygonTopPt = Number.POSITIVE_INFINITY;
  let polygonBottomPt = Number.NEGATIVE_INFINITY;
  for (const point of retainedPoints) {
    polygonLeftPt = Math.min(polygonLeftPt, point.xPt);
    polygonRightPt = Math.max(polygonRightPt, point.xPt);
    polygonTopPt = Math.min(polygonTopPt, point.yPt);
    polygonBottomPt = Math.max(polygonBottomPt, point.yPt);
  }
  const exactEventByKey = new Map<string, ExactRational>();
  for (const point of exactPoints) {
    const y = normalizeExactRational(point.y, 1n);
    exactEventByKey.set(exactRationalKey(y), y);
  }
  for (const intersection of intersections) {
    exactEventByKey.set(exactRationalKey(intersection.y), intersection.y);
  }
  const exactEventYs = Object.freeze(
    [...exactEventByKey.values()].sort(compareExactRational),
  );
  const eventYPts = Object.freeze([
    ...new Set(exactEventYs.map((eventY) =>
      exactLatticeToNumber(eventY, scaleExponent))),
  ].sort((left, right) => left - right));
  const contourSweep = compileContourSpans(
    edges,
    exactEdges,
    exactEventYs,
    scaleExponent,
    intersections,
  );
  const exactPolygonLeft = normalizeExactRational(
    exactPoints.reduce((minimum, point) => point.x < minimum ? point.x : minimum, exactPoints[0]!.x),
    1n,
  );
  const exactPolygonRight = normalizeExactRational(
    exactPoints.reduce((maximum, point) => point.x > maximum ? point.x : maximum, exactPoints[0]!.x),
    1n,
  );
  const exactPolygonTop = normalizeExactRational(
    exactPoints.reduce((minimum, point) => point.y < minimum ? point.y : minimum, exactPoints[0]!.y),
    1n,
  );
  const exactPolygonBottom = normalizeExactRational(
    exactPoints.reduce((maximum, point) => point.y > maximum ? point.y : maximum, exactPoints[0]!.y),
    1n,
  );
  const zero = normalizeExactRational(0n, 1n);
  const positiveDifference = (left: ExactRational, right: ExactRational): ExactRational => {
    const difference = subtractExactRational(left, right);
    return compareExactRational(difference, zero) > 0 ? difference : zero;
  };
  const exactCompiled = Object.freeze({
    scaleExponent,
    edges: Object.freeze(exactEdges),
    eventYs: exactEventYs,
    spans: contourSweep.exactSpans,
    spanIndex: buildExactContourSpanIndex(contourSweep.exactSpans),
    polygonLeft: exactPolygonLeft,
    polygonRight: exactPolygonRight,
    polygonTop: exactPolygonTop,
    polygonBottom: exactPolygonBottom,
    padLeft: positiveDifference(
      exactPolygonLeft,
      normalizeExactRational(latticeInteger(input.xLeftPt), 1n),
    ),
    padRight: positiveDifference(
      normalizeExactRational(latticeInteger(input.xRightPt), 1n),
      exactPolygonRight,
    ),
    padTop: positiveDifference(
      exactPolygonTop,
      normalizeExactRational(latticeInteger(input.yTopPt), 1n),
    ),
    padBottom: positiveDifference(
      normalizeExactRational(latticeInteger(input.yBottomPt), 1n),
      exactPolygonBottom,
    ),
  });
  const compiled = Object.freeze({
    kind: input.kind,
    edges: Object.freeze(edges),
    eventYPts,
    contourSpans: contourSweep.spans,
    contourSpanIndex: buildContourSpanIndex(contourSweep.spans),
    intersectionCount: intersections.length,
    compileOrderComparisonCount: contourSweep.orderComparisonCount,
    compilePairMembershipVisitCount: contourSweep.pairMembershipVisitCount,
    polygonLeftPt,
    polygonRightPt,
    polygonTopPt,
    polygonBottomPt,
    padLeftPt: Math.max(0, polygonLeftPt - input.xLeftPt),
    padRightPt: Math.max(0, input.xRightPt - polygonRightPt),
    padTopPt: Math.max(0, polygonTopPt - input.yTopPt),
    padBottomPt: Math.max(0, input.yBottomPt - polygonBottomPt),
  });
  exactCompiledPolygons.set(compiled, exactCompiled);
  return compiled;
}

interface ExactInterval {
  readonly l: ExactRational;
  readonly r: ExactRational;
}

function exactEdgeXAt(edge: ExactEdge, y: ExactRational): ExactRational {
  return {
    numerator: edge.dx * y.numerator - edge.c * y.denominator,
    denominator: edge.dy * y.denominator,
  };
}

function exactEdgesAreCoincident(left: ExactEdge, right: ExactEdge): boolean {
  return left.dx * right.dy === right.dx * left.dy
    && left.c * right.dy === right.c * left.dy;
}

function mergeExactIntervals(intervals: readonly ExactInterval[]): ExactInterval[] {
  const sorted = intervals
    .filter((interval) => compareExactRational(interval.r, interval.l) > 0)
    .slice()
    .sort((left, right) =>
      compareExactRational(left.l, right.l) || compareExactRational(left.r, right.r));
  const merged: ExactInterval[] = [];
  for (const interval of sorted) {
    const previous = merged.at(-1);
    if (!previous || compareExactRational(interval.l, previous.r) > 0) {
      merged.push({ ...interval });
    } else if (compareExactRational(interval.r, previous.r) > 0) {
      merged[merged.length - 1] = { l: previous.l, r: interval.r };
    }
  }
  return merged;
}

function addUnreducedExactRational(
  left: ExactRational,
  right: ExactRational,
): ExactRational {
  return {
    numerator: left.numerator * right.denominator
      + right.numerator * left.denominator,
    denominator: left.denominator * right.denominator,
  };
}

function subtractUnreducedExactRational(
  left: ExactRational,
  right: ExactRational,
): ExactRational {
  return {
    numerator: left.numerator * right.denominator
      - right.numerator * left.denominator,
    denominator: left.denominator * right.denominator,
  };
}

function projectPolygonLatticeIntervals(
  polygon: CompiledPolygonWrap,
  lineTop: ExactRational,
  lineBottom: ExactRational,
): ExactInterval[] {
  const exact = exactCompiledPolygons.get(polygon);
  if (!exact) throw new Error('Compiled polygon omitted its exact geometry authority');
  const movingTop = subtractUnreducedExactRational(lineTop, exact.padBottom);
  const movingBottom = addUnreducedExactRational(lineBottom, exact.padTop);
  const queryTop = compareExactRational(exact.polygonTop, movingTop) >= 0
    ? exact.polygonTop
    : movingTop;
  const queryBottom = compareExactRational(exact.polygonBottom, movingBottom) <= 0
    ? exact.polygonBottom
    : movingBottom;
  if (compareExactRational(queryBottom, queryTop) <= 0) return [];
  const projected: ExactInterval[] = [];
  const spans: ExactContourSpan[] = [];
  queryExactContourSpans(exact.spanIndex, queryTop, queryBottom, spans);
  for (const span of spans) {
    const slabTop = compareExactRational(queryTop, span.yTop) >= 0
      ? queryTop
      : span.yTop;
    const slabBottom = compareExactRational(queryBottom, span.yBottom) <= 0
      ? queryBottom
      : span.yBottom;
    if (compareExactRational(slabBottom, slabTop) <= 0) continue;
    const left = exact.edges[span.leftEdge]!;
    const right = exact.edges[span.rightEdge]!;
    if (exactEdgesAreCoincident(left, right)) continue;
    const leftTop = exactEdgeXAt(left, slabTop);
    const leftBottom = exactEdgeXAt(left, slabBottom);
    const rightTop = exactEdgeXAt(right, slabTop);
    const rightBottom = exactEdgeXAt(right, slabBottom);
    projected.push({
      l: subtractUnreducedExactRational(
        compareExactRational(leftTop, leftBottom) <= 0 ? leftTop : leftBottom,
        exact.padLeft,
      ),
      r: addUnreducedExactRational(
        compareExactRational(rightTop, rightBottom) >= 0 ? rightTop : rightBottom,
        exact.padRight,
      ),
    });
  }
  // Horizontal segments and a line-band boundary have zero overlapping area.
  // Omitting them is what preserves a through opening that begins exactly at
  // the line top while still letting the adjacent open slab contribute area.
  const expanded = mergeExactIntervals(projected);
  // §20.4.2.18 permits through text inside the polygon's maximum left/right
  // extents, so its disconnected filled intervals must stay disconnected.
  // §20.4.2.19 forbids that placement for tight wrapping, so tight takes the
  // outer hull while retaining the same contour-derived extrema.
  const selected = polygon.kind === 'through' || expanded.length === 0
    ? expanded
    : [{ l: expanded[0]!.l, r: expanded.at(-1)!.r }];
  return selected;
}

export interface ExactPolygonInterval {
  readonly l: ExactRational;
  readonly r: ExactRational;
}

/** Exact projection authority for the half-open line band [top,bottom). */
export function projectPolygonExactIntervals(
  polygon: CompiledPolygonWrap,
  lineTopPt: number,
  lineBottomPt: number,
): readonly ExactPolygonInterval[] {
  const exact = exactCompiledPolygons.get(polygon);
  if (!exact) throw new Error('Compiled polygon omitted its exact geometry authority');
  const toPhysical = (value: ExactRational): ExactRational =>
    exact.scaleExponent >= 0
      ? {
          numerator: value.numerator << BigInt(exact.scaleExponent),
          denominator: value.denominator,
        }
      : {
          numerator: value.numerator,
          denominator: value.denominator << BigInt(-exact.scaleExponent),
        };
  return Object.freeze(
    projectPolygonLatticeIntervals(
      polygon,
      exactLatticeValue(lineTopPt, exact.scaleExponent),
      exactLatticeValue(lineBottomPt, exact.scaleExponent),
    )
      .map((interval) => Object.freeze({
        l: toPhysical(interval.l),
        r: toPhysical(interval.r),
      })),
  );
}

/** Exact projection for a binary64 top plus a mathematically exact binary64 height. */
export function projectPolygonExactLineIntervals(
  polygon: CompiledPolygonWrap,
  lineTopPt: number,
  probeHeightPt: number,
): readonly ExactPolygonInterval[] {
  const exact = exactCompiledPolygons.get(polygon);
  if (!exact) throw new Error('Compiled polygon omitted its exact geometry authority');
  const lineTop = exactLatticeValue(lineTopPt, exact.scaleExponent);
  const lineBottom = addUnreducedExactRational(
    lineTop,
    exactLatticeValue(probeHeightPt, exact.scaleExponent),
  );
  const toPhysical = (value: ExactRational): ExactRational =>
    exact.scaleExponent >= 0
      ? {
          numerator: value.numerator << BigInt(exact.scaleExponent),
          denominator: value.denominator,
        }
      : {
          numerator: value.numerator,
          denominator: value.denominator << BigInt(-exact.scaleExponent),
        };
  return Object.freeze(
    projectPolygonLatticeIntervals(polygon, lineTop, lineBottom)
      .map((interval) => Object.freeze({
        l: toPhysical(interval.l),
        r: toPhysical(interval.r),
      })),
  );
}

/**
 * Number adapter for diagnostics and direct callers. Production line-window
 * composition consumes `projectPolygonExactLineIntervals` instead.
 */
export function projectPolygonIntervals(
  polygon: CompiledPolygonWrap,
  lineTopPt: number,
  lineBottomPt: number,
): PolygonInterval[] {
  return projectPolygonExactIntervals(polygon, lineTopPt, lineBottomPt)
    .map((interval) => ({
      l: exactRationalToNumber(interval.l),
      r: exactRationalToNumber(interval.r),
    }));
}

/** Exact structural line-top events for the finite line-window sweep. */
export function polygonLineTopEventYPts(
  polygon: CompiledPolygonWrap,
  probeHeightPt: number,
): readonly number[] {
  const exact = exactCompiledPolygons.get(polygon);
  if (!exact) throw new Error('Compiled polygon omitted its exact geometry authority');
  const probeHeight = exactLatticeValue(probeHeightPt, exact.scaleExponent);
  const events = new Set<number>();
  for (const eventY of exact.eventYs) {
    events.add(exactLatticeToNumberUp(
      addUnreducedExactRational(eventY, exact.padBottom),
      exact.scaleExponent,
    ));
    events.add(exactLatticeToNumberUp(
      subtractUnreducedExactRational(
        subtractUnreducedExactRational(eventY, probeHeight),
        exact.padTop,
      ),
      exact.scaleExponent,
    ));
  }
  return Object.freeze([...events].filter(Number.isFinite).sort((left, right) => left - right));
}

/**
 * Return only contour pairs intersected throughout one structural top-Y slab.
 * The caller supplies a slab bounded by shifted polygon events, so which static
 * contour slabs overlap the moving half-open line band cannot change inside it.
 * Each returned endpoint is affine in line top; later composition therefore
 * needs roots only between current envelopes and adjacent free boundaries.
 */
export function polygonBandExactIntervalFunctions(
  polygon: CompiledPolygonWrap,
  probeHeightPt: number,
  intervalTopY: number,
  intervalBottomY: number,
): readonly ExactPolygonAffineInterval[] {
  const exact = exactCompiledPolygons.get(polygon);
  if (!exact) throw new Error('Compiled polygon omitted its exact geometry authority');
  const intervalTop = exactLatticeValue(intervalTopY, exact.scaleExponent);
  const intervalBottom = exactLatticeValue(intervalBottomY, exact.scaleExponent);
  const midpoint = midpointExactRational(intervalTop, intervalBottom);
  const probeHeight = exactLatticeValue(probeHeightPt, exact.scaleExponent);
  const movingTopAtMidpoint = subtractUnreducedExactRational(
    midpoint,
    exact.padBottom,
  );
  const movingBottomOffset = addUnreducedExactRational(probeHeight, exact.padTop);
  const movingBottomAtMidpoint = addUnreducedExactRational(
    midpoint,
    movingBottomOffset,
  );
  const queryTopAtMidpoint =
    compareExactRational(exact.polygonTop, movingTopAtMidpoint) >= 0
      ? exact.polygonTop
      : movingTopAtMidpoint;
  const queryBottomAtMidpoint =
    compareExactRational(exact.polygonBottom, movingBottomAtMidpoint) <= 0
      ? exact.polygonBottom
      : movingBottomAtMidpoint;
  const intervals: ExactPolygonAffineInterval[] = [];
  const spans: ExactContourSpan[] = [];
  queryExactContourSpans(
    exact.spanIndex,
    queryTopAtMidpoint,
    queryBottomAtMidpoint,
    spans,
  );
  for (const span of spans) {
    const overlapTop =
      compareExactRational(queryTopAtMidpoint, span.yTop) >= 0
        ? queryTopAtMidpoint
        : span.yTop;
    const overlapBottom =
      compareExactRational(queryBottomAtMidpoint, span.yBottom) <= 0
        ? queryBottomAtMidpoint
        : span.yBottom;
    if (compareExactRational(overlapBottom, overlapTop) <= 0) continue;
    const leftEdge = exact.edges[span.leftEdge]!;
    const rightEdge = exact.edges[span.rightEdge]!;
    if (exactEdgesAreCoincident(leftEdge, rightEdge)) continue;
    const movingTop = compareExactRational(movingTopAtMidpoint, span.yTop) > 0;
    const movingBottom =
      compareExactRational(movingBottomAtMidpoint, span.yBottom) < 0;
    const endpointFunction = (
      edge: ExactEdge,
      useMoving: boolean,
      movingOffset: ExactRational,
      fixedY: ExactRational,
    ): ExactAffineX => {
      const yOffset = useMoving ? movingOffset : fixedY;
      const latticeIntercept = {
        numerator: edge.dx * yOffset.numerator
          - edge.c * yOffset.denominator,
        denominator: edge.dy * yOffset.denominator,
      };
      return {
        slope: useMoving
          ? { numerator: edge.dx, denominator: edge.dy }
          : { numerator: 0n, denominator: 1n },
        intercept: exact.scaleExponent >= 0
          ? {
              numerator: latticeIntercept.numerator << BigInt(exact.scaleExponent),
              denominator: latticeIntercept.denominator,
            }
          : {
              numerator: latticeIntercept.numerator,
              denominator:
                latticeIntercept.denominator << BigInt(-exact.scaleExponent),
            },
      };
    };
    const topLeft = endpointFunction(
      leftEdge,
      movingTop,
      { numerator: -exact.padBottom.numerator, denominator: exact.padBottom.denominator },
      span.yTop,
    );
    const bottomLeft = endpointFunction(
      leftEdge,
      movingBottom,
      movingBottomOffset,
      span.yBottom,
    );
    const topRight = endpointFunction(
      rightEdge,
      movingTop,
      { numerator: -exact.padBottom.numerator, denominator: exact.padBottom.denominator },
      span.yTop,
    );
    const bottomRight = endpointFunction(
      rightEdge,
      movingBottom,
      movingBottomOffset,
      span.yBottom,
    );
    const leftFunction = leftEdge.dx >= 0n ? topLeft : bottomLeft;
    const rightFunction = rightEdge.dx >= 0n ? bottomRight : topRight;
    const exactInterval = Object.freeze({
      left: Object.freeze({
        slope: leftFunction.slope,
        intercept: subtractUnreducedExactRational(
          leftFunction.intercept,
          exact.scaleExponent >= 0
            ? {
                numerator: exact.padLeft.numerator << BigInt(exact.scaleExponent),
                denominator: exact.padLeft.denominator,
              }
            : {
                numerator: exact.padLeft.numerator,
                denominator: exact.padLeft.denominator << BigInt(-exact.scaleExponent),
              },
        ),
      }),
      right: Object.freeze({
        slope: rightFunction.slope,
        intercept: addUnreducedExactRational(
          rightFunction.intercept,
          exact.scaleExponent >= 0
            ? {
                numerator: exact.padRight.numerator << BigInt(exact.scaleExponent),
                denominator: exact.padRight.denominator,
              }
            : {
                numerator: exact.padRight.numerator,
                denominator: exact.padRight.denominator << BigInt(-exact.scaleExponent),
              },
        ),
      }),
    });
    intervals.push(exactInterval);
  }
  return Object.freeze(intervals);
}

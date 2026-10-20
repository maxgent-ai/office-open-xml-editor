export interface ColumnBalanceFragment {
  readonly extentPt: number;
  readonly breakAfter: 'allowed' | 'forbidden' | 'forced';
}

export interface ExactColumnBalanceInput {
  readonly columnCount: number;
  readonly fragments: readonly ColumnBalanceFragment[];
}

export interface ExactColumnBalanceResult {
  readonly targetPt: number;
  /** Source-order fragment indexes ending each occupied column, including the final fragment. */
  readonly cutIndexes: readonly number[];
  /** Diagnostic proof that the frontier walk remains linear in columns × boundaries. */
  readonly transitionExpansions: number;
}

/**
 * Solve the exact contiguous minimax partition over legal fragment boundaries.
 *
 * For one column-count layer, the preceding optimum is non-decreasing by source
 * boundary while the final-column extent is non-increasing by candidate cut.
 * Their maximum is therefore unimodal. Retaining the crossing pointer as the
 * destination boundary advances visits every legal transition at most a
 * constant number of times instead of constructing a quadratic edge set.
 */
export function solveExactColumnBalance(
  input: ExactColumnBalanceInput,
): ExactColumnBalanceResult {
  // ECMA-376 §17.18.4: a column break that ends the balanced content still
  // advances to the next column. Balance into one fewer column so the break
  // lands on this page; when interior breaks leave no such partition, the
  // break follows the last column to the next page (Windows Word 16.0.20430,
  // #1896 control C5).
  const last = input.fragments.at(-1);
  if (last?.breakAfter === 'forced' && input.columnCount > 1) {
    const reserved = solveColumnBalance({ ...input, columnCount: input.columnCount - 1 });
    if (reserved) return reserved;
  }
  const result = solveColumnBalance(input);
  if (!result) throw new Error('Authored column breaks exceed the available column frontier');
  return result;
}

function solveColumnBalance(
  input: ExactColumnBalanceInput,
): ExactColumnBalanceResult | null {
  if (!Number.isInteger(input.columnCount) || input.columnCount <= 0) {
    throw new RangeError('Column count must be a positive integer');
  }
  for (const fragment of input.fragments) {
    if (!Number.isFinite(fragment.extentPt) || fragment.extentPt < 0) {
      throw new RangeError('Column balance fragment extents must be finite and non-negative');
    }
  }
  if (input.fragments.length === 0) {
    return Object.freeze({
      targetPt: 0,
      cutIndexes: Object.freeze([]),
      transitionExpansions: 0,
    });
  }

  const prefixPt = [0];
  const endpoints = [0];
  const forcedEndpoints: number[] = [];
  input.fragments.forEach((fragment, index) => {
    prefixPt.push(prefixPt[index]! + fragment.extentPt);
    const endpoint = index + 1;
    if (fragment.breakAfter !== 'forbidden' || endpoint === input.fragments.length) {
      endpoints.push(endpoint);
    }
    if (fragment.breakAfter === 'forced' && endpoint < input.fragments.length) {
      forcedEndpoints.push(endpoint);
    }
  });

  const layerCount = Math.min(input.columnCount, endpoints.length - 1);
  const costs: number[][] = Array.from(
    { length: layerCount + 1 },
    () => Array(endpoints.length).fill(Number.POSITIVE_INFINITY),
  );
  const predecessors: number[][] = Array.from(
    { length: layerCount + 1 },
    () => Array(endpoints.length).fill(-1),
  );
  costs[0]![0] = 0;
  let transitionExpansions = 0;

  for (let columns = 1; columns <= layerCount; columns += 1) {
    const previous = costs[columns - 1]!;
    const current = costs[columns]!;
    const candidatePositions = previous.flatMap((cost, position) => (
      Number.isFinite(cost) ? [position] : []
    ));
    let crossingIndex = 0;
    let forcedIndex = 0;
    let lastForcedBefore = 0;
    let minimumStartPosition = 0;
    for (let endPosition = 1; endPosition < endpoints.length; endPosition += 1) {
      const end = endpoints[endPosition]!;
      while (
        forcedIndex < forcedEndpoints.length
        && forcedEndpoints[forcedIndex]! < end
      ) {
        lastForcedBefore = forcedEndpoints[forcedIndex]!;
        forcedIndex += 1;
      }
      while (endpoints[minimumStartPosition]! < lastForcedBefore) {
        minimumStartPosition += 1;
      }
      while (
        crossingIndex < candidatePositions.length
        && candidatePositions[crossingIndex]! < minimumStartPosition
      ) crossingIndex += 1;
      const crossing = candidatePositions[crossingIndex];
      if (crossing === undefined || crossing >= endPosition) continue;

      const costAt = (startPosition: number): number => {
        transitionExpansions += 1;
        const start = endpoints[startPosition]!;
        return Math.max(
          previous[startPosition]!,
          prefixPt[end]! - prefixPt[start]!,
        );
      };
      let selectedPosition = crossing;
      let selectedCost = costAt(selectedPosition);
      while (crossingIndex + 1 < candidatePositions.length) {
        const next = candidatePositions[crossingIndex + 1]!;
        if (next >= endPosition) break;
        const nextCost = costAt(next);
        if (nextCost > selectedCost) break;
        crossingIndex += 1;
        selectedPosition = next;
        selectedCost = nextCost;
      }
      current[endPosition] = selectedCost;
      predecessors[columns]![endPosition] = selectedPosition;
    }
  }

  const finalPosition = endpoints.length - 1;
  let selectedLayer = -1;
  let targetPt = Number.POSITIVE_INFINITY;
  for (let columns = 1; columns <= layerCount; columns += 1) {
    const candidate = costs[columns]![finalPosition]!;
    if (candidate <= targetPt) {
      targetPt = candidate;
      selectedLayer = columns;
    }
  }
  if (selectedLayer < 0 || !Number.isFinite(targetPt)) return null;

  const cuts: number[] = [];
  let endPosition = finalPosition;
  for (let columns = selectedLayer; columns > 0; columns -= 1) {
    cuts.push(endpoints[endPosition]!);
    endPosition = predecessors[columns]![endPosition]!;
    if (endPosition < 0) {
      throw new Error('Column balance frontier omitted a predecessor');
    }
  }
  cuts.reverse();
  return Object.freeze({
    targetPt,
    cutIndexes: Object.freeze(cuts),
    transitionExpansions,
  });
}

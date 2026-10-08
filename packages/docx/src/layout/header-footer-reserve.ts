import { convergeLayoutSteps, type LayoutIteration } from './convergence.js';
import { stableFingerprint } from './fingerprint.js';
import type { StoryLayout } from './types.js';

export interface HeaderFooterReserve {
  readonly top: number;
  readonly bottom: number;
  /**
   * Page-final flow top of each footnote, by note id, whose story is band
   * dependent (StoryLayout.bandDependent: it holds §17.4.57 positioned
   * tables that only a page position places).
   * Footnotes are stacked only once their page closes, so the composed
   * position one pass observes is the band the next pass lays the note out
   * with; like the reserves, it is converged exactly (body-paginator.ts
   * footnoteBandPlan). A continued tail's entry is its composed top moved by
   * its source cut: the page position of its acquisition's flow top.
   */
  readonly footnoteTopsPt?: Readonly<Record<string, number>>;
}

export interface ReservedBodyInterval {
  readonly blockStartPt: number;
  readonly blockEndPt: number;
}

export interface HeaderFooterStorySlots<T> {
  readonly default: T | null;
  readonly first: T | null;
  readonly even: T | null;
}

/** §17.10.1/.2/.5/.6: resolved slots remain distinct; an absent selected
 * first/even slot is a blank story, not permission to borrow the default slot. */
export function selectedHeaderFooterStory<T>(
  stories: HeaderFooterStorySlots<T>,
  selection: Readonly<{
    titlePage: boolean;
    firstPageOfSection: boolean;
    evenAndOddHeaders: boolean;
    displayPageNumber: number;
  }>,
): T | null {
  if (selection.titlePage && selection.firstPageOfSection) return stories.first;
  if (selection.evenAndOddHeaders && selection.displayPageNumber % 2 === 0) {
    return stories.even;
  }
  return stories.default;
}

export function reservedBodyInterval(
  geometry: Readonly<{
    pageHeight: number;
    marginTop: number;
    marginBottom: number;
  }>,
  reserve: HeaderFooterReserve,
): ReservedBodyInterval {
  if (![geometry.pageHeight, geometry.marginTop, geometry.marginBottom, reserve.top, reserve.bottom]
    .every(Number.isFinite)) {
    throw new RangeError('Reserved body interval inputs must be finite');
  }
  if (geometry.pageHeight <= 0 || reserve.top < 0 || reserve.bottom < 0) {
    throw new RangeError('Reserved body interval requires a positive page and non-negative reserves');
  }
  const blockStartPt = Math.min(
    geometry.pageHeight,
    Math.abs(geometry.marginTop) + reserve.top,
  );
  const unboundedEndPt = geometry.pageHeight - Math.abs(geometry.marginBottom) - reserve.bottom;
  return Object.freeze({
    blockStartPt,
    blockEndPt: Math.max(blockStartPt, Math.min(geometry.pageHeight, unboundedEndPt)),
  });
}

/** §17.6.11: a negative signed margin permits story overlap; otherwise only
 * the extent beyond the margin-to-distance allowance reduces the body band. */
export function headerFooterOverflowReservePt(
  storyExtentPt: number,
  marginPt: number,
  distancePt: number,
): number {
  if (![storyExtentPt, marginPt, distancePt].every(Number.isFinite)) {
    throw new RangeError('Header/footer reserve inputs must be finite');
  }
  if (storyExtentPt < 0) throw new RangeError('Story extent must be non-negative');
  // A selected but empty story has no occupied interval. Its anchor distance
  // alone cannot overlap the body or reduce the canonical body-flow domain.
  if (storyExtentPt === 0) return 0;
  return marginPt < 0 ? 0 : Math.max(0, storyExtentPt - (marginPt - distancePt));
}

/**
 * ECMA-376 §17.6.11 bases a non-negative top margin on the extent of header
 * text. Controlled Word output showed that an undecorated paragraph containing
 * only U+0020 contributes no body overflow even at 36 pt, while a visible glyph
 * at that size does. A paragraph border also contributes. This intentionally
 * recognizes only that measured non-painting class; tabs, other whitespace,
 * fields, objects, and decorations keep their acquired extent.
 */
export function headerStoryBodyReserveExtentPt(story: StoryLayout): number {
  // The Office controls covered one paragraph. A longer blank header can
  // carry its own vertical extent, so do not infer the same suppression there.
  const onlyUndecoratedSpaces = story.blocks.length <= 1 && story.blocks.every((block) =>
    block.kind === 'paragraph'
    && block.borders.length === 0
    && block.shading === undefined
    && block.resources.length === 0
    && block.drawings.length === 0
    && block.textBoxes.length === 0
    && block.events.length === 0
    && block.exclusions.length === 0
    && (block.lineNumbers?.length ?? 0) === 0
    && (block.anchorFrames?.length ?? 0) === 0
    && block.lines.every((line) =>
      (line.barTabRules?.length ?? 0) === 0
      && line.placements.every((placement) =>
        placement.kind === 'text'
        && /^ *$/.test(placement.text)
        && placement.role !== 'field-result'
        && placement.dependency === undefined
        && placement.hyperlink === undefined
        && placement.paintOps.every((op) => /^ *$/.test(op.text))
        && placement.decorations.length === 0
        && placement.highlight === undefined
        && (placement.highlightFragments?.length ?? 0) === 0
        && placement.background === undefined
        && placement.runBorder === undefined
        && (placement.runBorderFragments?.length ?? 0) === 0
        && placement.ruby === undefined
        && placement.emphasis === undefined
        && placement.emphasisMark === undefined
        && placement.noteReference === undefined,
      ),
    ),
  );
  return onlyUndecoratedSpaces ? 0 : story.advancePt;
}

export interface HeaderFooterReserveIteration<T> extends LayoutIteration {
  readonly result: T;
  readonly reserves: readonly HeaderFooterReserve[];
}

export function convergeHeaderFooterReserves<T>(input: Readonly<{
  seed: T;
  measure: (result: T) => readonly HeaderFooterReserve[];
  repaginate: (reserves: readonly HeaderFooterReserve[], current: T) => T;
  identity: (result: T) => unknown;
  requiresConvergence?: boolean;
  limit?: number;
}>): HeaderFooterReserveIteration<T> {
  const { seed, repaginate, ...rest } = input;
  const steps = convergeHeaderFooterReserveSteps<T, never, T>(seed, {
    ...rest,
    carry: (result) => result,
    repaginate: function* generatorRepaginate(reserves, current) {
      return repaginate(reserves, current);
    },
  });
  let next = steps.next();
  while (!next.done) next = steps.next();
  return next.value;
}

/**
 * {@link convergeHeaderFooterReserves} whose repagination is suspendable.
 *
 * The seed pass is supplied already-computed by the caller (which is itself
 * suspendable), so only the repagination needs to delegate here.
 *
 * `carry` projects a pass result to what the next repagination reads besides
 * the measured reserves (for body pagination: page-field contexts and the
 * accepted page-anchor plan). The seed is taken as its own argument, not as a
 * field of `input`, so that once its reserves, fingerprint and carried data are
 * derived nothing here references it: while a repagination builds, neither the
 * seed nor any superseded pass result is kept alive by this frame.
 */
export function* convergeHeaderFooterReserveSteps<T, Y, C>(
  seed: T,
  input: Readonly<{
    measure: (result: T) => readonly HeaderFooterReserve[];
    carry: (result: T) => C;
    repaginate: (reserves: readonly HeaderFooterReserve[], carried: C) => Generator<Y, T, void>;
    identity: (result: T) => unknown;
    requiresConvergence?: boolean;
    limit?: number;
  }>,
): Generator<Y, HeaderFooterReserveIteration<T>, void> {
  type Carried = Readonly<{ reserves: readonly HeaderFooterReserve[]; carried: C }>;
  const iteration = (result: T): HeaderFooterReserveIteration<T> => {
    const reserves = Object.freeze(input.measure(result).map((reserve) => Object.freeze({ ...reserve })));
    return Object.freeze({
      result,
      reserves,
      pageCount: reserves.length,
      fingerprint: stableFingerprint('header-footer-reserve-v1', {
        identity: input.identity(result), reserves,
      }),
    });
  };
  let initial: HeaderFooterReserveIteration<T> | null = iteration(seed);
  seed = undefined as unknown as T;
  if (!input.requiresConvergence && initial.reserves.every(
    (reserve) => reserve.top === 0 && reserve.bottom === 0 && reserve.footnoteTopsPt === undefined,
  )) return initial;
  const steps = convergeLayoutSteps<HeaderFooterReserveIteration<T>, Y, Carried>(
    initial,
    function* reservePass(current) {
      return iteration(yield* input.repaginate(current.reserves, current.carried));
    },
    input.limit ?? 16,
    (current) => Object.freeze({
      reserves: current.reserves,
      carried: input.carry(current.result),
    }),
  );
  // The convergence generator derives the seed's state and carried data on
  // its first step; this frame must not keep the seed pass alive meanwhile.
  initial = null;
  return yield* steps;
}

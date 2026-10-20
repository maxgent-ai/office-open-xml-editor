import {
  axisAlignedRectsOverlap,
  resolveAxisAlignedOverlap,
  type AxisAlignedRect,
} from './axis-aligned-overlap.js';
import type { FloatRegistryEntryPt, LayoutRect } from './types.js';
import type { FloatRect } from './float-wrap.js';

/** Numerical tolerances retained by callers while C1b removes the temporary
 * scalar registry. Point-space table transactions also used these exact values
 * through that adapter before C1a and must preserve them explicitly. */
export const FLOAT_OVERLAP_EPS = 0.01;
export const FLOAT_PAGE_RIGHT_SLACK = 0.5;

/** Immutable object and text-exclusion geometry in one caller-selected
 * coordinate space. Normative object collisions use `bounds`; compatibility-only
 * text-wrap displacement uses `exclusionBounds`. */
interface FloatPlacementParticipantCore {
  readonly occurrenceId: string;
  readonly paragraphId: number;
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
}

export type FloatPlacementParticipant =
  | Readonly<FloatPlacementParticipantCore & {
      readonly kind: 'table';
      /** ECMA-376 Part 1 §17.4.56. Required so blocker-side `never`
       * cannot silently become overlap-permitted. */
      readonly tableOverlap: 'never' | 'overlap';
    }>
  | Readonly<FloatPlacementParticipantCore & {
      readonly kind: 'drawingml';
    }>
  | Readonly<FloatPlacementParticipantCore & {
      readonly kind: 'frame';
      /** §17.18.104 framePr wrap projected by `frameWrapExclusionMode`, as on
       * FloatRegistryEntryPt; absent projects square. */
      readonly exclusionMode?: 'square' | 'topAndBottom';
    }>;

export type FloatAvoidance =
  | Readonly<{ kind: 'drawingml-normative' }>
  | Readonly<{ kind: 'none' }>;

/** ECMA-376 §17.4.56: a floating table avoids another floating table only
 * when either one sets tblOverlap="never" (resolved in
 * `resolveFloatPlacement`); an overlap-permitted table keeps its position
 * beside floats of other paragraphs (Windows Word 16.0.20430, #1896). */
export function floatingTableAvoidance(): FloatAvoidance {
  return Object.freeze({ kind: 'none' });
}

/** ECMA-376 §20.4.2.3: an allowOverlap=true DrawingML object keeps its
 * resolved position. Word controls from issue #1623 confirm that no
 * different-paragraph displacement applies to such pictures. */
export function drawingMLAvoidance(allowOverlap: boolean): FloatAvoidance {
  return allowOverlap
    ? Object.freeze({ kind: 'none' })
    : Object.freeze({ kind: 'drawingml-normative' });
}

/** A `w:framePr` frame keeps its resolved position over other floats
 * (Windows Word 16.0.20430 #1896 controls: frames overlapping a frame or a
 * floating table of another paragraph are not displaced). */
export function frameAvoidance(): FloatAvoidance {
  return Object.freeze({ kind: 'none' });
}

export function floatRegistryParticipant(
  entry: FloatRegistryEntryPt,
): FloatPlacementParticipant {
  const core = {
    occurrenceId: entry.occurrenceId,
    paragraphId: entry.paragraphId,
    bounds: entry.bounds,
    exclusionBounds: entry.exclusionBounds,
  };
  if (entry.kind === 'table') {
    return {
      ...core,
      kind: 'table',
      tableOverlap: entry.overlap,
    };
  }
  if (entry.kind === 'frame') {
    return {
      ...core,
      kind: 'frame',
      ...(entry.exclusionMode ? { exclusionMode: entry.exclusionMode } : {}),
    };
  }
  return { ...core, kind: 'drawingml' };
}

export function floatRectParticipant(
  float: FloatRect,
  index: number,
): FloatPlacementParticipant {
  const imageX = float.imageX;
  const imageY = float.imageY;
  const imageW = float.imageW;
  const imageH = float.imageH;
  const xLeft = float.xLeft;
  const xRight = float.xRight;
  const yTop = float.yTop;
  const yBottom = float.yBottom;
  const core = {
    occurrenceId: float.anchorOccurrenceId
      ?? float.acquisitionOccurrenceId
      ?? `display-float:${index}`,
    paragraphId: float.paraId,
    bounds: {
      xPt: imageX,
      yPt: imageY,
      widthPt: imageW,
      heightPt: imageH,
    },
    exclusionBounds: {
      xPt: xLeft,
      yPt: yTop,
      widthPt: xRight - xLeft,
      heightPt: yBottom - yTop,
    },
  };
  if (float.kind === 'table') {
    return {
      ...core,
      kind: 'table',
      tableOverlap: float.tableOverlap,
    };
  }
  // A frame's FloatRect `mode` is its frameWrapExclusionMode projection.
  return float.kind === 'frame'
    ? { ...core, kind: 'frame', exclusionMode: float.mode }
    : { ...core, kind: 'drawingml' };
}

export interface FloatPlacement {
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
  readonly displacement: Readonly<{ xPt: number; yPt: number }>;
  readonly appliedCompatibilityRuleIds: readonly string[];
}

export interface ResolveFloatPlacementInput {
  readonly moving: FloatPlacementParticipant;
  readonly blockers: readonly FloatPlacementParticipant[];
  readonly avoidance: FloatAvoidance;
  readonly rightBoundaryPt: number;
  /** Callers explicitly supply the numerical policy established for their
   * coordinate path; exact anchor reflow deliberately omits both values. */
  readonly overlapEpsilonPt?: number;
  readonly rightBoundarySlackPt?: number;
  /** The page's vertical extent. When supplied, a table that must avoid
   * another (tblOverlap="never") only takes a position on the page. */
  readonly pageVerticalBoundsPt?: Readonly<{ topPt: number; bottomPt: number }>;
}

function axisRect(rect: LayoutRect): AxisAlignedRect {
  const xPt = rect.xPt;
  const yPt = rect.yPt;
  const widthPt = rect.widthPt;
  const heightPt = rect.heightPt;
  return {
    left: xPt,
    right: xPt + widthPt,
    top: yPt,
    bottom: yPt + heightPt,
  };
}

function translateRect(
  rect: LayoutRect,
  xPt: number,
  yPt: number,
): LayoutRect {
  if (xPt === 0 && yPt === 0) return rect;
  return Object.freeze({
    xPt: rect.xPt + xPt,
    yPt: rect.yPt + yPt,
    widthPt: rect.widthPt,
    heightPt: rect.heightPt,
  });
}

function placement(
  moving: FloatPlacementParticipant,
  xPt: number,
  yPt: number,
  ruleIds: readonly string[],
): FloatPlacement {
  return Object.freeze({
    bounds: translateRect(moving.bounds, xPt, yPt),
    exclusionBounds: translateRect(moving.exclusionBounds, xPt, yPt),
    displacement: Object.freeze({ xPt, yPt }),
    appliedCompatibilityRuleIds: Object.freeze([...ruleIds]),
  });
}

function compatibilityBlockerInObjectSpace(
  moving: FloatPlacementParticipant,
  blocker: FloatPlacementParticipant,
): AxisAlignedRect {
  const movingLeftPadding = moving.bounds.xPt - moving.exclusionBounds.xPt;
  const movingTopPadding = moving.bounds.yPt - moving.exclusionBounds.yPt;
  const movingRightPadding = moving.exclusionBounds.xPt
    + moving.exclusionBounds.widthPt
    - moving.bounds.xPt
    - moving.bounds.widthPt;
  const movingBottomPadding = moving.exclusionBounds.yPt
    + moving.exclusionBounds.heightPt
    - moving.bounds.yPt
    - moving.bounds.heightPt;
  const blockerExclusion = axisRect(blocker.exclusionBounds);
  // Rigid-translation identity:
  // overlap(movingExclusion + d, blockerExclusion)
  //   === overlap(movingObject + d, this inflated blocker).
  return {
    left: blockerExclusion.left - movingRightPadding,
    right: blockerExclusion.right + movingLeftPadding,
    top: blockerExclusion.top - movingBottomPadding,
    bottom: blockerExclusion.bottom + movingTopPadding,
  };
}

function resolveObjectPosition(
  input: ResolveFloatPlacementInput,
  blockers: readonly AxisAlignedRect[],
  rightBoundaryPt = input.rightBoundaryPt,
): Readonly<{ left: number; top: number }> {
  const movingRect = axisRect(input.moving.bounds);
  if (blockers.length === 0) {
    return Object.freeze({ left: movingRect.left, top: movingRect.top });
  }
  return resolveAxisAlignedOverlap(movingRect, blockers, {
    overlapEpsilon: input.overlapEpsilonPt ?? 0,
    rightBoundary: rightBoundaryPt,
    rightBoundarySlack: input.rightBoundarySlackPt ?? 0,
  });
}

/**
 * Resolve only float-to-float displacement policy.
 *
 * Anchor axis/size/wrap geometry remains owned by `anchor-frame.ts`; text-line
 * exclusion remains owned by `float-wrap.ts`; page/column admission remains
 * owned by the paginator. ECMA-376 chooses the eligible object class but does
 * not choose a valid displacement direction, so the shared deterministic
 * right-then-down kernel preserves the established direction.
 */
export function resolveFloatPlacement(
  input: ResolveFloatPlacementInput,
): FloatPlacement {
  const { moving, avoidance } = input;
  const normativeBlockers = input.blockers.flatMap((blocker): AxisAlignedRect[] => {
    if (moving.kind === 'table'
      && blocker.kind === 'table'
      && (moving.tableOverlap === 'never' || blocker.tableOverlap === 'never')) {
      return [axisRect(blocker.bounds)];
    }
    if (avoidance.kind === 'drawingml-normative' && blocker.kind === 'drawingml') {
      return [axisRect(blocker.bounds)];
    }
    return [];
  });
  const resolved = resolveObjectPosition(input, normativeBlockers);
  const page = input.pageVerticalBoundsPt;
  if (moving.kind === 'table' && page && normativeBlockers.length > 0
    && resolved.top + moving.bounds.heightPt > page.bottomPt + (input.overlapEpsilonPt ?? 0)) {
    // `word-never-table-page-displacement`: when neither moving right nor
    // moving down keeps the table on the page, set it above the tables it
    // meets; when that leaves the page too, keep its authored position.
    // (Leaving the text area is a pagination decision, made by the caller.)
    const above = resolveAbove(axisRect(moving.bounds), normativeBlockers, input.overlapEpsilonPt ?? 0);
    const top = above !== null && above >= page.topPt - (input.overlapEpsilonPt ?? 0)
      ? above
      : moving.bounds.yPt;
    return placement(moving, 0, top - moving.bounds.yPt, []);
  }
  return placement(
    moving,
    resolved.left - moving.bounds.xPt,
    resolved.top - moving.bounds.yPt,
    [],
  );
}

/** Move `moving` straight up until it meets none of `blockers`; null when
 * that does not settle. Each step clears every blocker it meets, so it
 * settles within `blockers.length` steps. */
function resolveAbove(
  moving: AxisAlignedRect,
  blockers: readonly AxisAlignedRect[],
  overlapEpsilon: number,
): number | null {
  const height = moving.bottom - moving.top;
  let top = moving.top;
  for (let step = 0; step <= blockers.length; step += 1) {
    const current = { ...moving, top, bottom: top + height };
    const intersecting = blockers.filter((blocker) =>
      axisAlignedRectsOverlap(current, blocker, overlapEpsilon));
    if (intersecting.length === 0) return top;
    top = Math.min(...intersecting.map((blocker) => blocker.top)) - height;
  }
  return null;
}

export interface ResolveBlockFlowAdmissionInput {
  /** The block's own inline extent. */
  readonly inlineStartPt: number;
  readonly inlineEndPt: number;
  /** The inline band of the flow lines the block occupies (its column or
   * story band), in the same coordinate space. */
  readonly flowBandStartPt: number;
  readonly flowBandEndPt: number;
  readonly blockStartPt: number;
  readonly blockExtentPt: number;
  readonly blockers: readonly FloatPlacementParticipant[];
  readonly overlapEpsilonPt: number;
}

/** The inline interval over which `blocker` excludes an ordinary-flow block,
 * or null when it never does. */
function blockFlowExclusionInterval(
  blocker: FloatPlacementParticipant,
  input: ResolveBlockFlowAdmissionInput,
): Readonly<{ startPt: number; endPt: number }> | null {
  const own = { startPt: input.inlineStartPt, endPt: input.inlineEndPt };
  switch (blocker.kind) {
    // §17.4.57 floating-table text exclusion (existing library policy).
    case 'table': return own;
    // ECMA-376 §17.18.104: none/notBeside (topAndBottom) allow no flow
    // content beside the frame, so every line of the band that crosses it
    // resumes below it; around (square) leaves the remaining line space,
    // which a block cannot share (ordinary-block policy, as for tables), so
    // only a block that itself overlaps the frame is moved.
    case 'frame': return (blocker.exclusionMode ?? 'square') === 'topAndBottom'
      ? {
          startPt: Math.min(input.flowBandStartPt, own.startPt),
          endPt: Math.max(input.flowBandEndPt, own.endPt),
        }
      : own;
    // DrawingML objects keep their established paragraph-line wrap only.
    case 'drawingml': return null;
  }
}

/**
 * Admit one ordinary-flow block below the text exclusions it may not sit
 * beside: floating tables and `w:framePr` frames ({@link
 * blockFlowExclusionInterval}).
 *
 * This is not §17.4.56 float-to-float placement: the moving object is flow
 * content, so exclusion bounds apply. Each move adopts the bottom of an
 * intersecting blocker. The start is monotone and every cleared blocker can
 * never intersect again, giving a hard bound of `eligible.length` moves.
 */
export function resolveBlockFlowAdmission(
  input: ResolveBlockFlowAdmissionInput,
): Readonly<{ blockStartPt: number }> {
  if (input.inlineEndPt < input.inlineStartPt
    || input.flowBandEndPt < input.flowBandStartPt
    || input.blockExtentPt < 0) {
    throw new RangeError('Block-flow admission received a negative extent');
  }
  const eligible = input.blockers.filter((blocker) => {
    const interval = blockFlowExclusionInterval(blocker, input);
    if (!interval) return false;
    const bounds = blocker.exclusionBounds;
    return interval.endPt - bounds.xPt > input.overlapEpsilonPt
      && bounds.xPt + bounds.widthPt - interval.startPt > input.overlapEpsilonPt;
  });
  let blockStartPt = input.blockStartPt;
  for (let moveCount = 0; moveCount <= eligible.length; moveCount += 1) {
    const intersecting = eligible.filter((blocker) => {
      const bounds = blocker.exclusionBounds;
      return blockStartPt + input.blockExtentPt - bounds.yPt > input.overlapEpsilonPt
        && bounds.yPt + bounds.heightPt - blockStartPt > input.overlapEpsilonPt;
    });
    if (intersecting.length === 0) return Object.freeze({ blockStartPt });
    if (moveCount === eligible.length) {
      throw new Error('Block-flow float admission did not converge');
    }
    blockStartPt = Math.max(...intersecting.map((blocker) =>
      blocker.exclusionBounds.yPt + blocker.exclusionBounds.heightPt));
  }
  throw new Error('Block-flow float admission did not converge');
}

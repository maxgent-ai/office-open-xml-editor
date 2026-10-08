import {
  WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT,
  WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL,
} from './compatibility.js';
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
  | Readonly<{ kind: 'word-different-paragraph'; paragraphId: number }>
  | Readonly<{ kind: 'none' }>;

export function floatingTableAvoidance(
  tableOverlap: 'never' | 'overlap',
  paragraphId: number,
): FloatAvoidance {
  return tableOverlap === 'overlap'
    ? Object.freeze({ kind: 'word-different-paragraph', paragraphId })
    : Object.freeze({ kind: 'none' });
}

/** ECMA-376 §20.4.2.3: an allowOverlap=true DrawingML object keeps its
 * resolved position. Word controls from issue #1623 confirm that no
 * different-paragraph displacement applies to such pictures. */
export function drawingMLAvoidance(allowOverlap: boolean): FloatAvoidance {
  return allowOverlap
    ? Object.freeze({ kind: 'none' })
    : Object.freeze({ kind: 'drawingml-normative' });
}

/** Legacy `w:framePr` frames keep the established different-paragraph
 * displacement; no Word control has measured frame overlap. */
export function frameAvoidance(paragraphId: number): FloatAvoidance {
  return Object.freeze({ kind: 'word-different-paragraph', paragraphId });
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
  const compatibilityBlockers = avoidance.kind === 'word-different-paragraph'
    ? input.blockers.flatMap((blocker): AxisAlignedRect[] =>
        blocker.paragraphId === avoidance.paragraphId
          ? []
          : [compatibilityBlockerInObjectSpace(moving, blocker)])
    : [];
  const movingRightPaddingPt = moving.exclusionBounds.xPt
    + moving.exclusionBounds.widthPt
    - moving.bounds.xPt
    - moving.bounds.widthPt;
  // Compatibility displacement historically resolved the moving exclusion
  // rectangle. Its right edge, not only the object edge, must fit within the
  // caller's boundary and slack after converting the collision into object
  // space.
  const rightBoundaryPt = avoidance.kind === 'word-different-paragraph'
    ? input.rightBoundaryPt - movingRightPaddingPt
    : input.rightBoundaryPt;
  const normative = resolveObjectPosition(
    input,
    normativeBlockers,
    rightBoundaryPt,
  );
  const resolved = compatibilityBlockers.length === 0
    ? normative
    : resolveObjectPosition(
        input,
        [...normativeBlockers, ...compatibilityBlockers],
        rightBoundaryPt,
      );
  const xPt = resolved.left - moving.bounds.xPt;
  const yPt = resolved.top - moving.bounds.yPt;
  const compatibilityChangedPlacement = resolved.left !== normative.left
    || resolved.top !== normative.top;
  return placement(moving, xPt, yPt, compatibilityChangedPlacement
    ? [WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT.id]
    : []);
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

export interface ResolvePageAnchoredTableDeferralInput {
  readonly bounds: LayoutRect;
  readonly blockers: readonly FloatPlacementParticipant[];
  readonly overlapEpsilonPt: number;
}

/** Established Word pagination behavior, deliberately separate from normative
 * §17.4.56 raw table placement. The authored object band is admitted against
 * already-reserved §17.4.57 floating-table text-exclusion bands. */
export function resolvePageAnchoredTableDeferral(
  input: ResolvePageAnchoredTableDeferralInput,
): Readonly<{
  defer: boolean;
  appliedCompatibilityRuleIds: readonly string[];
}> {
  const moving = axisRect(input.bounds);
  const defer = input.blockers.some((blocker) =>
    blocker.kind === 'table'
      && axisAlignedRectsOverlap(
        moving,
        axisRect(blocker.exclusionBounds),
        input.overlapEpsilonPt,
      ));
  return Object.freeze({
    defer,
    appliedCompatibilityRuleIds: defer
      ? Object.freeze([WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL.id])
      : Object.freeze([]),
  });
}

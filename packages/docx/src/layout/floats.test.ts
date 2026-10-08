import { describe, expect, it } from 'vitest';
import {
  resolveBlockFlowAdmission,
  resolveFloatPlacement,
  resolvePageAnchoredTableDeferral,
  type FloatPlacementParticipant,
} from './floats.js';

const participant = (
  occurrenceId: string,
  kind: FloatPlacementParticipant['kind'],
  xPt: number,
  paragraphId: number,
  exclusionPaddingPt = 0,
  tableOverlap: 'never' | 'overlap' = 'overlap',
): FloatPlacementParticipant => Object.freeze({
  occurrenceId,
  kind,
  paragraphId,
  bounds: Object.freeze({ xPt, yPt: 0, widthPt: 10, heightPt: 10 }),
  exclusionBounds: Object.freeze({
    xPt: xPt - exclusionPaddingPt,
    yPt: -exclusionPaddingPt,
    widthPt: 10 + exclusionPaddingPt * 2,
    heightPt: 10 + exclusionPaddingPt * 2,
  }),
  ...(kind === 'table' ? { tableOverlap } : {}),
}) as FloatPlacementParticipant;

describe('float displacement policy', () => {
  it('applies DrawingML allowOverlap=false only against DrawingML object bounds', () => {
    const moving = participant('moving', 'drawingml', 0, 3, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [
        participant('table', 'table', 10, 0),
        participant('frame', 'frame', 20, 1),
        participant('drawing', 'drawingml', 0, 2),
      ],
      avoidance: { kind: 'drawingml-normative' },
      rightBoundaryPt: 100,
    });

    expect(result.bounds.xPt).toBe(10);
    expect(result.exclusionBounds.xPt).toBe(8);
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('applies tblOverlap=never to raw floating-table extents, not text padding', () => {
    const moving = participant('moving', 'table', 0, 2, 3, 'never');
    const blocker = participant('blocker', 'table', 0, 1, 4, 'overlap');
    const result = resolveFloatPlacement({
      moving,
      blockers: [blocker],
      avoidance: { kind: 'none' },
      rightBoundaryPt: 100,
    });

    expect(result.bounds.xPt).toBe(10);
    expect(result.exclusionBounds.xPt).toBe(7);
  });

  it('enforces blocker-side tblOverlap=never even within the same paragraph', () => {
    const moving = participant('moving', 'table', 0, 2, 3, 'overlap');
    const blocker = participant('blocker', 'table', 0, 2, 4, 'never');
    const result = resolveFloatPlacement({
      moving,
      blockers: [blocker],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.bounds.xPt).toBe(10);
    expect(result.exclusionBounds.xPt).toBe(7);
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('does not attribute pairwise table displacement to the Word compatibility rule', () => {
    const moving = participant('moving', 'table', 0, 2, 0, 'overlap');
    const result = resolveFloatPlacement({
      moving,
      blockers: [participant('blocker', 'table', 0, 1, 0, 'never')],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.bounds.xPt).toBe(10);
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('enforces tblOverlap=never in both source orders on raw bounds', () => {
    const place = (
      movingOverlap: 'never' | 'overlap',
      blockerOverlap: 'never' | 'overlap',
    ) => resolveFloatPlacement({
      moving: participant('moving', 'table', 0, 2, 3, movingOverlap),
      blockers: [participant('blocker', 'table', 0, 1, 4, blockerOverlap)],
      avoidance: { kind: 'none' } as const,
      rightBoundaryPt: 100,
    });

    expect(place('never', 'overlap').bounds.xPt).toBe(10);
    expect(place('overlap', 'never').bounds.xPt).toBe(10);
  });

  it('uses the supplied page or cell right boundary before moving down', () => {
    const moving = participant('moving', 'drawingml', 0, 2);
    const blocker = participant('blocker', 'drawingml', 0, 1);
    const result = resolveFloatPlacement({
      moving,
      blockers: [blocker],
      avoidance: { kind: 'drawingml-normative' },
      rightBoundaryPt: 10,
    });

    expect(result.bounds).toMatchObject({ xPt: 0, yPt: 10 });
  });

  it('is independent of blocker input order', () => {
    const moving = participant('moving', 'drawingml', 0, 3);
    const left = participant('left', 'drawingml', 0, 0);
    const right = participant('right', 'drawingml', 10, 1);
    const place = (blockers: readonly FloatPlacementParticipant[]) =>
      resolveFloatPlacement({
        moving,
        blockers,
        avoidance: { kind: 'drawingml-normative' },
        rightBoundaryPt: 30,
      });

    expect(place([left, right])).toEqual(place([right, left]));
    expect(place([left, right]).bounds.xPt).toBe(20);
  });

  it('keeps observed different-paragraph displacement on exclusion bounds', () => {
    const moving = participant('moving', 'table', 0, 2, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [
        participant('same-paragraph', 'drawingml', 0, 2, 8),
        participant('other-paragraph', 'frame', 0, 1, 4),
      ],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.exclusionBounds.xPt).toBe(14);
    expect(result.appliedCompatibilityRuleIds).toEqual([
      'word-float-different-paragraph-displacement',
    ]);
  });

  it('keeps the compatibility exclusion band within the right boundary and slack', () => {
    const moving: FloatPlacementParticipant = Object.freeze({
      occurrenceId: 'moving',
      kind: 'table',
      tableOverlap: 'overlap',
      paragraphId: 2,
      bounds: Object.freeze({
        xPt: 20,
        yPt: 10,
        widthPt: 45,
        heightPt: 10,
      }),
      exclusionBounds: Object.freeze({
        xPt: 20,
        yPt: 10,
        widthPt: 53,
        heightPt: 10,
      }),
    });
    const blocker: FloatPlacementParticipant = Object.freeze({
      occurrenceId: 'blocker',
      kind: 'frame',
      paragraphId: 1,
      bounds: Object.freeze({
        xPt: 0,
        yPt: 0,
        widthPt: 50,
        heightPt: 50,
      }),
      exclusionBounds: Object.freeze({
        xPt: 0,
        yPt: 0,
        widthPt: 50,
        heightPt: 50,
      }),
    });
    const place = (rightBoundaryPt: number) => resolveFloatPlacement({
      moving,
      blockers: [blocker],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt,
      overlapEpsilonPt: 0.01,
      rightBoundarySlackPt: 0.5,
    });

    const constrained = place(100);
    expect(constrained.bounds).toMatchObject({ xPt: 20, yPt: 50 });
    expect(constrained.exclusionBounds.xPt + constrained.exclusionBounds.widthPt)
      .toBeLessThanOrEqual(100.5);
    expect(constrained.appliedCompatibilityRuleIds).toEqual([
      'word-float-different-paragraph-displacement',
    ]);

    expect(place(110).bounds).toMatchObject({ xPt: 50, yPt: 10 });
    expect(place(102.5).bounds).toMatchObject({ xPt: 50, yPt: 10 });
    expect(place(102.48).bounds).toMatchObject({ xPt: 20, yPt: 50 });
  });

  it('reports a Word compatibility rule only when it changes placement', () => {
    const moving = participant('moving', 'table', 0, 2, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [participant('other-paragraph', 'frame', 50, 1, 4)],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.displacement).toEqual({ xPt: 0, yPt: 0 });
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('does not report a Word compatibility rule when no blocker is eligible', () => {
    const moving = participant('moving', 'table', 0, 2, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [participant('same-paragraph', 'frame', 0, 2, 4)],
      avoidance: {
        kind: 'word-different-paragraph',
        paragraphId: moving.paragraphId,
      },
      rightBoundaryPt: 100,
    });

    expect(result.displacement).toEqual({ xPt: 0, yPt: 0 });
    expect(result.appliedCompatibilityRuleIds).toEqual([]);
  });

  it('preserves authored overlap when no avoidance policy applies', () => {
    const moving = participant('moving', 'drawingml', 0, 1, 2);
    const result = resolveFloatPlacement({
      moving,
      blockers: [participant('blocker', 'drawingml', 0, 0, 2)],
      avoidance: { kind: 'none' },
      rightBoundaryPt: 100,
    });

    expect(result.bounds).toBe(moving.bounds);
    expect(result.exclusionBounds).toBe(moving.exclusionBounds);
    expect(result.displacement).toEqual({ xPt: 0, yPt: 0 });
  });
});

describe('block flow admission around floating-table exclusions', () => {
  const table = (
    occurrenceId: string,
    yPt: number,
    heightPt: number,
  ): FloatPlacementParticipant => Object.freeze({
    occurrenceId,
    kind: 'table',
    tableOverlap: 'overlap',
    paragraphId: 0,
    bounds: Object.freeze({ xPt: 0, yPt, widthPt: 20, heightPt }),
    exclusionBounds: Object.freeze({ xPt: 0, yPt, widthPt: 20, heightPt }),
  });

  const admit = (blockers: readonly FloatPlacementParticipant[]) =>
    resolveBlockFlowAdmission({
      inlineStartPt: 0,
      inlineEndPt: 20,
      flowBandStartPt: 0,
      flowBandEndPt: 20,
      blockStartPt: 0,
      blockExtentPt: 10,
      blockers,
      overlapEpsilonPt: 0.01,
    }).blockStartPt;

  it('follows a finite chain of newly intersecting exclusion bands', () => {
    expect(admit([
      table('first', 0, 10),
      table('second', 12, 18),
    ])).toBe(30);
  });

  it('is independent of blocker order and does not jump to an unreachable band', () => {
    const first = table('first', 0, 10);
    const second = table('second', 12, 18);
    const unreachable = table('unreachable', 100, 100);

    expect(admit([first, second, unreachable])).toBe(30);
    expect(admit([unreachable, second, first])).toBe(30);
  });

  it('treats an edge overlap within the supplied epsilon as clear', () => {
    expect(admit([table('touching', -10.005, 10)])).toBe(0);
  });

  it('clears a frame by its §17.18.104 wrap: none/notBeside over the band, around over the block', () => {
    // A 20pt block at the start of a 100pt band; the frame sits beside it,
    // inside the band, vertically across the block.
    const frame = (exclusionMode?: 'square' | 'topAndBottom'): FloatPlacementParticipant => Object.freeze({
      occurrenceId: `frame-${exclusionMode ?? 'absent'}`,
      kind: 'frame',
      paragraphId: 0,
      ...(exclusionMode ? { exclusionMode } : {}),
      bounds: Object.freeze({ xPt: 50, yPt: 5, widthPt: 30, heightPt: 15 }),
      exclusionBounds: Object.freeze({ xPt: 50, yPt: 5, widthPt: 30, heightPt: 15 }),
    });
    const admitInBand = (blocker: FloatPlacementParticipant, flowBandEndPt = 100) =>
      resolveBlockFlowAdmission({
        inlineStartPt: 0,
        inlineEndPt: 20,
        flowBandStartPt: 0,
        flowBandEndPt,
        blockStartPt: 0,
        blockExtentPt: 10,
        blockers: [blocker],
        overlapEpsilonPt: 0.01,
      }).blockStartPt;

    expect(admitInBand(frame('topAndBottom'))).toBe(20);
    expect(admitInBand(frame('square'))).toBe(0);
    // Absent projects square, as on the registry entry.
    expect(admitInBand(frame())).toBe(0);
    // A frame outside the block's band (another column) never blocks it.
    expect(admitInBand(frame('topAndBottom'), 40)).toBe(0);
  });
});

describe('page-anchored floating-table compatibility admission', () => {
  it('defers on the existing table text-exclusion band and reports the rule', () => {
    const blocker = Object.freeze({
      ...participant('blocker', 'table', 20, 1, 5, 'overlap'),
      bounds: Object.freeze({ xPt: 20, yPt: 0, widthPt: 10, heightPt: 10 }),
      exclusionBounds: Object.freeze({ xPt: 15, yPt: -5, widthPt: 20, heightPt: 20 }),
    }) as FloatPlacementParticipant;
    const result = resolvePageAnchoredTableDeferral({
      // Raw table extents do not overlap; only the blocker exclusion does.
      bounds: { xPt: 10, yPt: 0, widthPt: 6, heightPt: 10 },
      blockers: [blocker],
      overlapEpsilonPt: 0.01,
    });

    expect(result).toEqual({
      defer: true,
      appliedCompatibilityRuleIds: [
        'word-page-anchored-table-collision-deferral',
      ],
    });
  });
});

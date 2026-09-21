import { describe, expect, it } from 'vitest';
import type { BorderSegment, TextDecorationLayout } from '../layout/types.js';
import type {
  CanvasPaintContext,
  PaintCanvas2D,
} from './types.js';
import { paintStrokeSegment } from './canvas-border.js';

function recordingContext(operations: unknown[]): PaintCanvas2D {
  return {
    globalAlpha: 1,
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    letterSpacing: '0px',
    fontKerning: 'auto',
    fillRect(x, y, width, height) { operations.push(['fillRect', x, y, width, height]); },
    strokeRect() {},
    setLineDash(pattern) { operations.push(['dash', pattern]); },
    fillText() {},
    translate() {},
    rotate() {},
    scale() {},
    transform() {},
    drawImage() {},
    save() { operations.push(['save']); },
    restore() { operations.push(['restore']); },
    beginPath() {},
    rect() {},
    clip() {},
    moveTo() {},
    lineTo() {},
    stroke() { operations.push(['stroke', this.lineWidth]); },
    fill() {},
  };
}

const dashedSegment = (widthPt: number, dashPatternPt: readonly number[]): BorderSegment => ({
  edge: 'top',
  from: { xPt: 0, yPt: 4 },
  to: { xPt: 100, yPt: 4 },
  color: '#000000',
  widthPt,
  authoredStyle: 'dashed',
  style: 'dashed',
  dashPatternPt,
});

describe('retained Canvas border paint', () => {
  it('recomputes authored dash cadence when a minimum CSS width clamps a hairline', () => {
    const operations: unknown[] = [];
    const context: CanvasPaintContext = {
      ctx: recordingContext(operations),
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };

    paintStrokeSegment(dashedSegment(0.25, [0.75, 0.5]), context, 0.5);

    expect(operations).toContainEqual(['dash', [1.5, 1]]);
    expect(operations).toContainEqual(['stroke', 0.5]);
  });

  it('keeps the retained dash cadence when the authored width exceeds the minimum', () => {
    const operations: unknown[] = [];
    const context: CanvasPaintContext = {
      ctx: recordingContext(operations),
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };

    paintStrokeSegment(dashedSegment(1, [3, 2]), context, 0.5);

    expect(operations).toContainEqual(['dash', [3, 2]]);
    expect(operations).toContainEqual(['stroke', 1]);
  });

  it('bounds retained wavy paths with a bevel join instead of an unretained miter', () => {
    const operations: unknown[] = [];
    const ctx = recordingContext(operations);
    Object.defineProperty(ctx, 'lineJoin', {
      configurable: true,
      get: () => 'miter',
      set: (value) => operations.push(['lineJoin', value]),
    });
    const context: CanvasPaintContext = {
      ctx,
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };
    const wavy: TextDecorationLayout = {
      kind: 'underline',
      from: { xPt: 0, yPt: 10 },
      to: { xPt: 8, yPt: 10 },
      color: '#000000',
      widthPt: 2,
      style: 'wavy',
      path: [
        { xPt: 0, yPt: 9 },
        { xPt: 2, yPt: 11 },
        { xPt: 4, yPt: 9 },
      ],
    };

    paintStrokeSegment(wavy, context);

    expect(operations).toContainEqual(['lineJoin', 'bevel']);
    expect(operations).toContainEqual(['save']);
    expect(operations).toContainEqual(['stroke', 2]);
    expect(operations).toContainEqual(['restore']);
    const operationIndex = (expected: unknown): number => operations.findIndex(
      (operation) => JSON.stringify(operation) === JSON.stringify(expected),
    );
    expect(operationIndex(['save']))
      .toBeLessThan(operationIndex(['lineJoin', 'bevel']));
    expect(operationIndex(['lineJoin', 'bevel']))
      .toBeLessThan(operationIndex(['stroke', 2]));
    expect(operationIndex(['stroke', 2]))
      .toBeLessThan(operationIndex(['restore']));
  });

  it('does not override the join for a two-point wavy segment', () => {
    const operations: unknown[] = [];
    const ctx = recordingContext(operations);
    Object.defineProperty(ctx, 'lineJoin', {
      configurable: true,
      get: () => 'miter',
      set: (value) => operations.push(['lineJoin', value]),
    });
    const context: CanvasPaintContext = {
      ctx,
      scale: 1,
      dpr: 1,
      resources: { paint() {} },
    };

    paintStrokeSegment({
      kind: 'underline',
      from: { xPt: 0, yPt: 10 },
      to: { xPt: 8, yPt: 10 },
      color: '#000000',
      widthPt: 2,
      style: 'wavy',
      path: [{ xPt: 0, yPt: 9 }, { xPt: 8, yPt: 11 }],
    }, context);

    expect(operations).not.toContainEqual(['lineJoin', 'bevel']);
    expect(operations).not.toContainEqual(['save']);
    expect(operations).not.toContainEqual(['restore']);
  });

  it('paints thin-thick compound borders as distinct unequal rails', () => {
    const operations: unknown[] = [];
    const context: CanvasPaintContext = {
      ctx: recordingContext(operations),
      scale: 1,
      dpr: 4,
      resources: { paint() {} },
    };

    paintStrokeSegment({
      edge: 'top', from: { xPt: 0, yPt: 10 }, to: { xPt: 100, yPt: 10 },
      color: '#000000', widthPt: 3, authoredStyle: 'thinThickSmallGap',
      style: 'compound', dashPatternPt: [],
    }, context);

    const fills = operations.filter((operation) =>
      Array.isArray(operation) && operation[0] === 'fillRect') as unknown[][];
    expect(fills).toHaveLength(2);
    expect(fills[0]?.[4]).toBeLessThan(fills[1]?.[4] as number);
    expect(operations).not.toContainEqual(['stroke', 3]);
  });
});

import { describe, expect, it } from 'vitest';
import { EM_BOX_CENTER_FALLBACK_EM, emBoxCenterAboveBaseline } from './em-box-center';

function ctxWith(metrics: Partial<TextMetrics>): CanvasRenderingContext2D {
  return {
    textBaseline: 'middle',
    measureText: () => metrics as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

describe('emBoxCenterAboveBaseline', () => {
  it('uses the font bounding box when the engine reports it and restores the baseline', () => {
    const ctx = ctxWith({ fontBoundingBoxAscent: 40, fontBoundingBoxDescent: 10 });
    expect(emBoxCenterAboveBaseline(ctx, 'A', 48)).toBe(15);
    expect(ctx.textBaseline).toBe('middle');
  });

  it('falls back to the em fraction only without font bounding box metrics', () => {
    expect(emBoxCenterAboveBaseline(ctxWith({}), 'A', 50)).toBe(EM_BOX_CENTER_FALLBACK_EM * 50);
  });
});

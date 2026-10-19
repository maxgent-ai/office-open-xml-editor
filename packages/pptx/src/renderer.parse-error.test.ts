import { describe, it, expect } from 'vitest';
import { renderSlide } from './renderer';
import type { Slide } from './types';

/**
 * RB7: a slide carrying `parseError` renders a visible placeholder instead of
 * (empty) content. This drives {@link renderSlide} against a recording 2D
 * context and asserts the placeholder is painted — the error message reaches
 * `fillText`, and no per-element render is attempted.
 */

interface DrawCall {
  op: string;
  args: unknown[];
}

/** A minimal recording 2D context that logs the draw ops we assert on. */
function recordingCtx(): { ctx: CanvasRenderingContext2D; calls: DrawCall[] } {
  const calls: DrawCall[] = [];
  const rec =
    (op: string) =>
    (...args: unknown[]) => {
      calls.push({ op, args });
    };
  const ctx = {
    // state
    save: rec('save'),
    restore: rec('restore'),
    scale: rec('scale'),
    setTransform: rec('setTransform'),
    translate: rec('translate'),
    // style props (assigned, not called)
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 1,
    font: '',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    // paint
    fillRect: rec('fillRect'),
    strokeRect: rec('strokeRect'),
    fillText: rec('fillText'),
    setLineDash: rec('setLineDash'),
    beginPath: rec('beginPath'),
    measureText: (t: string) => ({ width: t.length * 6 }),
  } as unknown as CanvasRenderingContext2D;
  return { ctx, calls };
}

/** A canvas stub whose getContext returns the recording ctx. */
function stubCanvas(ctx: CanvasRenderingContext2D): HTMLCanvasElement {
  return {
    width: 0,
    height: 0,
    style: {} as CSSStyleDeclaration,
    offsetWidth: 960,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
}

function brokenSlide(): Slide {
  return {
    index: 2,
    slideNumber: 3,
    background: null,
    elements: [],
    parseError: 'ppt/slides/slide3.xml: unexpected end of stream',
  };
}

describe('RB7 renderSlide placeholder', () => {
  it('paints a placeholder carrying the parseError message for a broken slide', async () => {
    const { ctx, calls } = recordingCtx();
    const canvas = stubCanvas(ctx);
    await renderSlide(canvas, brokenSlide(), 9_144_000, 6_858_000, { width: 960, dpr: 1 });

    const texts = calls
      .filter((c) => c.op === 'fillText')
      .map((c) => String(c.args[0]));
    // The heading names the slide and the detail includes the part-tagged error.
    expect(texts.some((t) => t.includes('Slide 3'))).toBe(true);
    expect(texts.join(' ')).toContain('slide3.xml');
    // A filled card + at least one glyph/heading were drawn.
    expect(calls.some((c) => c.op === 'fillRect')).toBe(true);
  });

  it('a healthy slide (no parseError) does NOT draw the placeholder heading', async () => {
    const { ctx, calls } = recordingCtx();
    const canvas = stubCanvas(ctx);
    const healthy: Slide = {
      index: 0,
      slideNumber: 1,
      background: null,
      elements: [],
    };
    await renderSlide(canvas, healthy, 9_144_000, 6_858_000, { width: 960, dpr: 1 });
    const texts = calls
      .filter((c) => c.op === 'fillText')
      .map((c) => String(c.args[0]));
    expect(texts.some((t) => t.includes('could not be displayed'))).toBe(false);
  });
});

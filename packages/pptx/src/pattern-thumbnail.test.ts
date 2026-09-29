import { describe, expect, it, vi } from 'vitest';
import { renderSlide } from './renderer.js';
import type { Slide } from './types.js';

function recordingCanvas() {
  const matrices: DOMMatrix2DInit[] = [];
  const state: Record<string, unknown> = { fillStyle: '', globalAlpha: 1, font: '11px Arial' };
  const ctx = new Proxy(state, {
    get(target, prop) {
      if (prop in target) return target[prop as string];
      if (prop === 'createPattern') return () => ({ setTransform: (matrix: DOMMatrix2DInit) => matrices.push(matrix) });
      if (prop === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      if (prop === 'measureText') return (text: string) => ({ width: text.length * 6 });
      if (prop === 'getLineDash') return () => [];
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') return () => ({ addColorStop() {} });
      return () => {};
    },
    set(target, prop, value) { target[prop as string] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx } as unknown as HTMLCanvasElement;
  state.canvas = canvas;
  return { canvas, matrices };
}

describe('PPTX thumbnail preset pattern scale', () => {
  it('uses the same point cells in an actual chart and shape render', async () => {
    const prior = globalThis.OffscreenCanvas;
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(public width: number, public height: number) {}
      getContext() { return { fillStyle: '', fillRect() {} }; }
    });
    try {
      const shapeFill = { fillType: 'pattern', preset: 'pct50', fg: 'D21D54', bg: '12CED4' } as const;
      const chartFill = { fillType: 'pattern', preset: 'horz', fg: 'D21D54', bg: '12CED4' } as const;
      const slide = {
        index: 0, slideNumber: 1, background: null,
        elements: [
          {
            type: 'shape', x: 0, y: 0, width: 2_000_000, height: 1_000_000,
            geometry: 'rect', rotation: 0, flipH: false, flipV: false,
            fill: shapeFill, stroke: null,
          },
          {
            type: 'chart', x: 2_200_000, y: 0, width: 3_000_000, height: 2_000_000,
            rotation: 0, flipH: false, flipV: false,
            chart: {
              chartType: 'clusteredBar', categories: [], series: [],
              authoredWithoutSeries: true, chartFill,
              showLegend: false, showDataLabels: false,
              catAxisHidden: true, valAxisHidden: true,
            },
          },
        ],
      } as unknown as Slide;
      const { canvas, matrices } = recordingCanvas();
      await renderSlide(canvas, slide, 9_144_000, 6_858_000, { width: 240, dpr: 1 });
      expect(matrices).toHaveLength(2);
      expect(matrices[0].a).toBeCloseTo((240 / 9_144_000) * 12_700 / 8);
      expect(matrices[1].a).toBeCloseTo(matrices[0].a as number);
    } finally {
      vi.stubGlobal('OffscreenCanvas', prior);
    }
  });
});

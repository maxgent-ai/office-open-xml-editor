import { describe, expect, it, vi } from 'vitest';
import { renderChart } from './renderer.js';
import { resolveFill } from '../shape/paint.js';
import type { ChartModel } from '../types/chart.js';

const pattern = (preset: string) => ({ fillType: 'pattern', preset, fg: '000000', bg: 'FFFFFF' }) as const;

function recordingContext() {
  const matrices: DOMMatrix2DInit[] = [];
  const state: Record<string, unknown> = {
    canvas: { width: 960, height: 720 },
    fillStyle: '', strokeStyle: '', font: '11px Arial', globalAlpha: 1,
  };
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
  return { ctx, matrices };
}

const chart = {
  chartType: 'clusteredBar', categories: ['A', 'B'],
  series: [{ name: 'S', values: [1, 2], color: '888888', fillPattern: pattern('pct20') }],
  chartFill: pattern('pct5'), plotAreaFill: pattern('pct10'), legendFill: pattern('pct25'),
  showLegend: true, legendPos: 'r', showDataLabels: false,
  catAxisHidden: true, valAxisHidden: true,
} as ChartModel;

describe('chart preset pattern point units', () => {
  it.each([1, 1 / 3, 4 / 3])('uses the host scale %s in actual chart frame and mark painters', (ptToPx) => {
    const prior = globalThis.OffscreenCanvas;
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(public width: number, public height: number) {}
      getContext() { return { fillStyle: '', fillRect() {} }; }
    });
    try {
      const { ctx, matrices } = recordingContext();
      renderChart(ctx, chart, { x: 20, y: 30, w: 400, h: 300 }, ptToPx);
      // Chart space, plot area, series and legend are separate render paths.
      expect(matrices.length).toBeGreaterThanOrEqual(4);
      expect(matrices.every(matrix => matrix.a === ptToPx / 8 && matrix.d === ptToPx / 8)).toBe(true);
      expect(matrices.every(matrix => (matrix.e ?? 0) === 0 && (matrix.f ?? 0) === 0)).toBe(true);
      resolveFill(pattern('pct30'), ctx, 0, 0, 100, 100);
      expect(matrices.at(-1)?.a).toBeCloseTo((4 / 3) / 8);
    } finally {
      vi.stubGlobal('OffscreenCanvas', prior);
    }
  });

});

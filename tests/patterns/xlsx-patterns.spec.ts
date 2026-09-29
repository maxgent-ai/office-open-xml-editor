import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const rendererEntry = fileURLToPath(new URL('../../packages/xlsx/src/renderer.ts', import.meta.url));

test.beforeEach(async ({ page }) => {
  const bundle = await build({
    input: rendererEntry,
    output: { format: 'iife', name: 'xlsxRenderer' },
    platform: 'browser',
  });
  await page.addScriptTag({ content: bundle.output[0].code });
});

const baseWorksheet = {
  name: 'Patterns', rows: [], colWidths: {}, rowHeights: {},
  defaultColWidth: 8.43, defaultRowHeight: 15,
  mergeCells: [], freezeRows: 0, freezeCols: 0,
  conditionalFormats: [], images: [], charts: [], shapeGroups: [],
  defaultFontFamily: 'Arial', defaultFontSize: 11,
};

const styles = {
  fonts: [{ size: 11 }], fills: [], borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 }],
  numFmts: [], dxfs: [],
};

const fill = { fillType: 'pattern', preset: 'vert', fg: 'D21D54', bg: '12CED4' };

test('worksheet shape pattern scales with zoom through renderViewport', async ({ page }) => {
  const samples = await page.evaluate(({ worksheet, styles, fill }) => {
    const output: number[] = [];
    const renderer = (globalThis as typeof globalThis & {
      xlsxRenderer: typeof import('../../packages/xlsx/src/renderer.js');
    }).xlsxRenderer;
    for (const scale of [1, 1.5, 2]) {
      const canvas = document.createElement('canvas');
      canvas.width = 1400; canvas.height = 500;
      const ctx = canvas.getContext('2d')!;
      renderer.renderViewport(ctx, {
        ...worksheet,
        shapeGroups: [{
          fromCol: 1, fromColOff: 0, fromRow: 1, fromRowOff: 0,
          toCol: 5, toColOff: 0, toRow: 9, toRowOff: 0,
          nativeExtCx: 0, nativeExtCy: 0,
          shapes: [{
            x: 0, y: 0, w: 1, h: 1, rot: 0, flipH: false, flipV: false,
            geom: { type: 'preset', name: 'rect', adj: [] },
            fill, strokeColor: null, strokeWidth: 0, paragraphs: [],
          }],
        }],
      } as never, styles as never, { row: 1, col: 1, rows: 20, cols: 20 }, { cellScale: scale });
      const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
      let bestY = 0; let bestCount = 0;
      for (let y = 0; y < canvas.height; y++) {
        let count = 0;
        for (let x = 0; x < canvas.width; x++) {
          const p = (y * canvas.width + x) * 4;
          if (data[p] > 180 && data[p + 1] < 70 && data[p + 2] < 130) count++;
        }
        if (count > bestCount) { bestCount = count; bestY = y; }
      }
      const starts: number[] = [];
      let priorRed = false;
      for (let x = 0; x < canvas.width; x++) {
        const p = (bestY * canvas.width + x) * 4;
        const red = data[p] > 180 && data[p + 1] < 70 && data[p + 2] < 130;
        if (red && !priorRed) starts.push(x);
        priorRed = red;
      }
      output.push((starts[11] - starts[1]) / 10);
    }
    return output;
  }, { worksheet: baseWorksheet, styles, fill });
  expect(samples[0]).toBeCloseTo(8 * 4 / 3, 0);
  expect(samples[1]).toBeCloseTo(16, 0);
  expect(samples[2]).toBeCloseTo(8 * 8 / 3, 0);
});

test('chart pattern keeps the fixed viewport grid while scrolling', async ({ page }) => {
  const samples = await page.evaluate(({ worksheet, styles, fill }) => {
    const renderer = (globalThis as typeof globalThis & {
      xlsxRenderer: typeof import('../../packages/xlsx/src/renderer.js');
    }).xlsxRenderer;
    const chart = {
      chartType: 'clusteredBar', categories: [], series: [],
      authoredWithoutSeries: true, chartFill: fill,
      showLegend: false, showDataLabels: false,
      catAxisHidden: true, valAxisHidden: true,
    };
    const outputs: Array<{ scale: number; period: number; frames: Array<{ pixels: number[] }> }> = [];
    for (const cellScale of [1, 2]) {
    const frames: Array<{ minX: number; minY: number; data: Uint8ClampedArray; width: number }> = [];
    for (const scrollOffsetX of [0, 3]) {
      const canvas = document.createElement('canvas');
      canvas.width = 2200; canvas.height = 1000;
      const ctx = canvas.getContext('2d')!;
      renderer.renderViewport(ctx, {
        ...worksheet,
        charts: [{
          fromCol: 2, fromColOff: 0, fromRow: 2, fromRowOff: 0,
          toCol: 8, toColOff: 0, toRow: 12, toRowOff: 0,
          chart,
        }],
      } as never, styles as never, { row: 1, col: 1, rows: 25, cols: 15 }, { scrollOffsetX, cellScale });
      const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
      let minX = canvas.width, minY = canvas.height;
      for (let y = 0; y < canvas.height; y++) for (let x = 0; x < canvas.width; x++) {
        const p = (y * canvas.width + x) * 4;
        if (data[p] > 180 && data[p + 1] < 70 && data[p + 2] < 130) {
          minX = Math.min(minX, x); minY = Math.min(minY, y);
        }
      }
      frames.push({ minX, minY, data, width: canvas.width });
    }
    const x0 = Math.max(...frames.map(frame => frame.minX)) + 25;
    const y0 = Math.max(...frames.map(frame => frame.minY)) + 25;
    const first = frames[0];
    const starts: number[] = [];
    const stripeY = first.minY + 5 * cellScale;
    let priorRed = false;
    for (let x = first.minX; x < first.minX + 150 * cellScale; x++) {
      const p = (stripeY * first.width + x) * 4;
      const red = first.data[p] > 180 && first.data[p + 1] < 70 && first.data[p + 2] < 130;
      if (red && !priorRed) starts.push(x);
      priorRed = red;
    }
    outputs.push({ scale: cellScale, period: starts.length > 5 ? (starts[5] - starts[1]) / 4 : 0, frames: frames.map(frame => {
      const pixels: number[] = [];
      for (let y = y0; y < y0 + 30; y++) for (let x = x0; x < x0 + 30; x++) {
        const p = (y * frame.width + x) * 4;
        pixels.push(frame.data[p], frame.data[p + 1], frame.data[p + 2]);
      }
      return { pixels };
    }) });
    }
    return outputs;
  }, { worksheet: baseWorksheet, styles, fill });
  // A fixed grid can move the first red stripe by a whole cell when the chart
  // clip moves three pixels, so compare a shared viewport rectangle instead.
  for (const sample of samples) expect(sample.frames[0].pixels).toEqual(sample.frames[1].pixels);
  expect(samples[0].period).toBeCloseTo(8 * 4 / 3, 0);
  expect(samples[1].period).toBeCloseTo(16 * 4 / 3, 0);
});

test('shape rotation and reflection transform the geometry while preserving the pattern grid', async ({ page }) => {
  const comparisons = await page.evaluate(({ worksheet, styles }) => {
    const renderer = (globalThis as typeof globalThis & {
      xlsxRenderer: typeof import('../../packages/xlsx/src/renderer.js');
    }).xlsxRenderer;
    const cases = [
      { rot: 0, flipH: false, flipV: false },
      { rot: 45, flipH: false, flipV: false },
      { rot: 0, flipH: true, flipV: false },
      { rot: 0, flipH: false, flipV: true },
      { rot: 45, flipH: true, flipV: false },
    ];
    return cases.map(transform => {
      const canvas = document.createElement('canvas');
      canvas.width = 1000; canvas.height = 600;
      const ctx = canvas.getContext('2d')!;
      renderer.renderViewport(ctx, {
        ...worksheet,
        shapeGroups: [{
          fromCol: 2, fromColOff: 0, fromRow: 3, fromRowOff: 0,
          toCol: 10, toColOff: 0, toRow: 15, toRowOff: 0,
          nativeExtCx: 0, nativeExtCy: 0,
          shapes: [{
            x: 0, y: 0, w: 1, h: 1, ...transform,
            geom: { type: 'preset', name: 'rect', adj: [] },
            fill: { fillType: 'pattern', preset: 'dnDiag', fg: 'D21D54', bg: '12CED4' },
            strokeColor: null, strokeWidth: 0, paragraphs: [],
          }],
        }],
      } as never, styles as never, { row: 1, col: 1, rows: 25, cols: 15 });
      // The central 64×64 region stays inside every rotated rectangle.
      const data = ctx.getImageData(475, 170, 64, 64).data;
      return Array.from(data);
    });
  }, { worksheet: baseWorksheet, styles });
  expect(comparisons[0].some((value, index) => index % 4 === 0 && value === 210)).toBe(true);
  expect(comparisons[0].some((value, index) => index % 4 === 0 && value === 18)).toBe(true);
  for (const transformed of comparisons.slice(1)) {
    expect(transformed).toEqual(comparisons[0]);
  }
});

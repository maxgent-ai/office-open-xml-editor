import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const paintEntry = fileURLToPath(new URL('../../packages/core/src/shape/paint.ts', import.meta.url));
const chartEntry = fileURLToPath(new URL('../../packages/core/src/chart/renderer.ts', import.meta.url));
const bitmapEntry = fileURLToPath(new URL('../../packages/core/src/shape/pattern-bitmaps.ts', import.meta.url));

test('translucent diagonal tile stores independent foreground, midpoint and background alpha', async ({ page }) => {
  const bundle = await build({ input: bitmapEntry, output: { format: 'iife', name: 'patternBitmap' }, platform: 'browser' });
  await page.addScriptTag({ content: bundle.output[0].code });
  const cells = await page.evaluate(() => {
    const bitmap = (globalThis as typeof globalThis & {
      patternBitmap: typeof import('../../packages/core/src/shape/pattern-bitmaps.js');
    }).patternBitmap;
    return [
      ['D21D5480', '12CED480'],
      ['D21D5440', '12CED4BF'],
      ['D21D5400', '12CED480'],
      ['D21D54FF', '12CED400'],
    ].map(([fg, bg]) => {
      const tile = bitmap.buildPatternBitmap('dnDiag', fg, bg)!;
      const ctx = tile.getContext('2d')!;
      return [4, 12, 20].map(x => Array.from(ctx.getImageData(x, 4, 1, 1).data));
    });
  });
  const expected = [
    [[210, 29, 84, 128], [114, 118, 148, 128], [18, 206, 212, 128]],
    [[210, 29, 84, 64], [114, 118, 148, 128], [18, 206, 212, 191]],
    [[0, 0, 0, 0], [114, 118, 148, 64], [18, 206, 212, 128]],
    [[210, 29, 84, 255], [114, 118, 148, 128], [0, 0, 0, 0]],
  ];
  for (let caseIndex = 0; caseIndex < expected.length; caseIndex++) {
    for (let cellIndex = 0; cellIndex < 3; cellIndex++) {
      for (let channel = 0; channel < 4; channel++) {
        expect(Math.abs(cells[caseIndex][cellIndex][channel] - expected[caseIndex][cellIndex][channel])).toBeLessThanOrEqual(2);
      }
    }
  }
});

test('preset cells stay uniformly coloured in Chrome at several point scales', async ({ page }) => {
  const bundle = await build({
    input: paintEntry,
    output: { format: 'iife', name: 'patternPaint' },
    platform: 'browser',
  });
  await page.addScriptTag({ content: bundle.output[0].code });

  for (const pointScale of [1, 4 / 3, 2, 6]) {
    const reds = await page.evaluate((scale) => {
      const canvas = document.createElement('canvas');
      canvas.width = 80;
      canvas.height = 8 * scale;
      const ctx = canvas.getContext('2d')!;
      const paint = (globalThis as typeof globalThis & {
        patternPaint: typeof import('../../packages/core/src/shape/paint.js');
      }).patternPaint;
      ctx.fillStyle = paint.resolveFill(
        { fillType: 'pattern', preset: 'horz', fg: '000000', bg: 'FFFFFF' },
        ctx, 0, 0, 80, canvas.height, 0, scale,
      )!;
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      const pixels = ctx.getImageData(3, 0, 1, canvas.height).data;
      return Array.from({ length: canvas.height }, (_, y) => pixels[y * 4]);
    }, pointScale);
    // The office PDF's horizontal preset has one foreground row, then seven
    // background rows. Any intermediate red value is unwanted interpolation.
    expect(reds.every(red => red === 0 || red === 255), `${pointScale} px/pt`).toBe(true);
    expect(reds.filter(red => red === 0).length).toBe(Math.round(pointScale));
  }
});

test('chart effect surfaces inherit point scale and fixed page phase', async ({ page }) => {
  const bundle = await build({
    input: chartEntry,
    output: { format: 'iife', name: 'patternChart' },
    platform: 'browser',
  });
  await page.addScriptTag({ content: bundle.output[0].code });
  const matrices = await page.evaluate(() => {
    const recorded: number[][] = [];
    const original = CanvasPattern.prototype.setTransform;
    CanvasPattern.prototype.setTransform = function (matrix) {
      recorded.push([matrix.a, matrix.d, matrix.e, matrix.f]);
      return original.call(this, matrix);
    };
    try {
      const canvas = document.createElement('canvas');
      canvas.width = 500; canvas.height = 400;
      const ctx = canvas.getContext('2d')!;
      const chart = {
        chartType: 'clusteredBar', categories: [], series: [],
        authoredWithoutSeries: true,
        chartFill: { fillType: 'pattern', preset: 'horz', fg: '000000', bg: 'FFFFFF' },
        chartAreaStyle: {
          shadows: [{ color: '112233', alpha: 0.5, blur: 12700, dist: 25400, dir: 0 }],
          effectAuthored: true,
        },
        showLegend: false, showDataLabels: false,
        catAxisHidden: true, valAxisHidden: true,
      };
      (globalThis as typeof globalThis & {
        patternChart: typeof import('../../packages/core/src/chart/renderer.js');
      }).patternChart.renderChart(ctx, chart as never, { x: 20, y: 30, w: 400, h: 300 }, 1,
        0, undefined, undefined, undefined, undefined);
      return recorded;
    } finally {
      CanvasPattern.prototype.setTransform = original;
    }
  });
  expect(matrices.length).toBeGreaterThanOrEqual(2);
  expect(matrices.every(([a, d, e, f]) => a === 1 / 8 && d === 1 / 8 && e === 0 && f === 0)).toBe(true);
});

test('a scaled and rotated group keeps the slide pattern grid', async ({ page }) => {
  const bundle = await build({
    input: paintEntry,
    output: { format: 'iife', name: 'patternPaint' },
    platform: 'browser',
  });
  await page.addScriptTag({ content: bundle.output[0].code });
  const rows = await page.evaluate(() => {
    const paint = (globalThis as typeof globalThis & {
      patternPaint: typeof import('../../packages/core/src/shape/paint.js');
    }).patternPaint;
    const canvas = document.createElement('canvas');
    canvas.width = 320; canvas.height = 240;
    const ctx = canvas.getContext('2d')!;
    const slideFrame = ctx.getTransform();
    ctx.translate(160, 120);
    ctx.rotate(Math.PI / 4);
    ctx.scale(1 / 8, 1 / 8);
    ctx.translate(-800, -400);
    paint.withPatternCoordinateSpace(ctx, slideFrame, () => {
      ctx.fillStyle = paint.resolveFill(
        { fillType: 'pattern', preset: 'horz', fg: '000000', bg: 'FFFFFF' },
        ctx, 0, 0, 1600, 800, 0, 1,
      )!;
      ctx.fillRect(0, 0, 1600, 800);
    });
    return [120, 121, 124, 128].map(y => [140, 160, 180].map(x =>
      ctx.getImageData(x, y, 1, 1).data[0]));
  });
  expect(rows).toEqual([[0, 0, 0], [255, 255, 255], [255, 255, 255], [0, 0, 0]]);
});

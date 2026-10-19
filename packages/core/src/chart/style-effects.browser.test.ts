import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { build } from 'vite';
import { existsSync } from 'node:fs';

// Exercise the production chart adapter with real stroke/blur rasterization.
describe.skipIf(!existsSync(chromium.executablePath()))('chart effect coverage pixels', () => {
  let browser: Browser;
  let page: Page;
  beforeAll(async () => {
    const root = new URL('../', import.meta.url).pathname;
    const entry = '/__coverage-test.ts';
    const result = await build({
      configFile: false, logLevel: 'error',
      plugins: [{ name: 'coverage-test', resolveId: id => id === entry ? id : null,
        load: id => id === entry ? [
          `export { renderChart } from ${JSON.stringify(`${root}chart/renderer.ts`)};`,
          `export * from ${JSON.stringify(`${root}chart/style-effects.ts`)};`,
          `export { applySoftEdge, applyInnerShadow } from ${JSON.stringify(`${root}shape/effects.ts`)};`,
        ].join('\n') : null }],
      build: { write: false, minify: false,
        lib: { entry, formats: ['iife'], name: 'coverageTest' } },
    });
    const outputs = Array.isArray(result) ? result : [result];
    const chunk = outputs.flatMap(output => 'output' in output ? output.output : [])
      .find(output => output.type === 'chunk');
    if (!chunk || chunk.type !== 'chunk') throw new Error('Chart bundle missing');
    browser = await chromium.launch({ channel: 'chromium' });
    page = await browser.newPage();
    await page.setContent('<!doctype html><canvas width="400" height="400"></canvas>');
    await page.addScriptTag({ content: chunk.code });
  }, 60_000);
  afterAll(async () => { await browser?.close(); });

  it('preserves a single-slice pie with a 20pt outline and 10pt soft edge', async () => {
    const pixels = await page.evaluate(() => {
      const api = (globalThis as unknown as { coverageTest: typeof import('./renderer') }).coverageTest;
      const canvas = document.querySelector('canvas') as HTMLCanvasElement;
      const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, 400, 400);
      api.renderChart(ctx, {
        chartType: 'pie', title: null, categories: ['A'],
        series: [{ name: 'S', values: [1], color: 'FF0000',
          chartexStyle: { lineColors: ['FF0000'], lineWidthEmu: 20 * 12700,
            softEdges: [{ radius: 10 * 12700 }], effectAuthored: true } }],
        showLegend: false, showDataLabels: false,
      } as import('../types/chart').ChartModel, { x: 0, y: 0, w: 400, h: 400 }, 1);
      return [30, 35, 40].map(x => [...ctx.getImageData(x, 200, 1, 1).data]);
    });
    // The previous renderer retains the red outline here. A geometry-only
    // coverage raster turns these outside-the-circle stroke pixels nearly white.
    for (const [i, pixel] of pixels.entries()) {
      expect(pixel[0]).toBe(255);
      expect(pixel[1]).toBeLessThanOrEqual([17, 0, 0][i]);
      expect(pixel[2]).toBeLessThanOrEqual([17, 0, 0][i]);
      expect(pixel[3]).toBe(255);
    }
  });

  for (const effect of ['softEdge', 'innerShadow'] as const) {
    it(`retains miter joins and square caps through ${effect}`, async () => {
      const mismatch = await page.evaluate(effect => {
        const api = (globalThis as unknown as { coverageTest:
          typeof import('./style-effects') & typeof import('../shape/effects') }).coverageTest;
        const make = () => {
          const canvas = document.createElement('canvas');
          canvas.width = canvas.height = 400;
          return canvas.getContext('2d') as CanvasRenderingContext2D;
        };
        const actual = make(), reference = make();
        const bbox = { x: 100.25, y: 100.25, w: 100, h: 100 };
        const softEdge = { radius: 30 * 12700 };
        const innerShadow = { color: '000000', alpha: 0.5, blur: 10 * 12700,
          dist: 5 * 12700, dir: 45 };
        const paint = (ctx: import('../canvas/aux-canvas').AuxContext) => {
          ctx.setTransform({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
          ctx.strokeStyle = '#ff0000';
          ctx.lineWidth = 20;
          ctx.lineCap = 'square';
          ctx.lineJoin = 'miter';
          ctx.beginPath();
          ctx.moveTo(100.25, 200.25);
          ctx.lineTo(150.25, 100.25);
          ctx.lineTo(200.25, 200.25);
          ctx.stroke();
        };
        api.withChartEffectBudget(actual, () => api.paintChartStyleEffects(
          actual, { effectAuthored: true,
            ...(effect === 'softEdge' ? { softEdges: [softEdge] } : { innerShadows: [innerShadow] }),
          }, undefined, 0, bbox, 1, paint,
        ));
        // The old adapter let the effect consumer rasterize the entire body
        // directly. This independent reference uses the same public primitive,
        // with an opaque vector silhouette authored from the stroked path.
        if (effect === 'softEdge') api.applySoftEdge(reference, paint, bbox,
          softEdge, 1 / 12700, 400, 400, paint);
        else {
          paint(reference);
          api.applyInnerShadow(reference, paint, bbox, innerShadow, 1 / 12700, 400, 400);
        }
        const a = actual.getImageData(0, 0, 400, 400).data;
        const b = reference.getImageData(0, 0, 400, 400).data;
        let missing = 0;
        let shadowed = 0;
        // Inspect fully covered cap/miter interiors; AA differs intentionally
        // under the established nonzero-support coverage policy.
        for (const [x, y] of [[150, 85], [97, 205], [203, 205]]) {
          const p = (y * 400 + x) * 4;
          if (b[p + 3] <= 20) throw new Error('Reference probe missed the painted outline');
          if (effect === 'innerShadow') {
            if (b[p] < 250) shadowed++;
            if (Math.abs(a[p] - b[p]) > 3) missing++;
          } else if (a[p + 3] < b[p + 3] - 5) missing++;
        }
        if (effect === 'innerShadow' && !shadowed) throw new Error('Reference probes missed the inset shadow');
        return missing;
      }, effect);
      expect(mismatch).toBe(0);
    });
  }
});

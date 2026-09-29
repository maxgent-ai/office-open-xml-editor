import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const entry = fileURLToPath(new URL('../../packages/pptx/src/renderer.ts', import.meta.url));

test('PowerPoint bevel body keeps the slide pattern phase in its interior', async ({ page }) => {
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  await page.addScriptTag({ content: bundle.output[0].code });
  const agreements = await page.evaluate(async () => {
    const renderer = (globalThis as typeof globalThis & {
      pptxRenderer: typeof import('../../packages/pptx/src/renderer.js');
    }).pptxRenderer;
    const render = async (bevel: boolean, rotation: number, flipH: boolean) => {
      const canvas = document.createElement('canvas');
      const shape = {
        type: 'shape', x: 83 * 12700, y: 73 * 12700,
        width: 180 * 12700, height: 100 * 12700,
        geometry: 'rect', rotation, flipH, flipV: false,
        fill: { fillType: 'pattern', preset: 'horz', fg: 'D21D54', bg: '12CED4' },
        stroke: null,
        ...(bevel ? {
          scene3d: { camera: { prst: 'orthographicFront' }, lightRig: { rig: 'threePt', dir: 't' } },
          sp3d: { bevelT: { w: 127000, h: 76200, prst: 'circle' } },
        } : {}),
      };
      await renderer.renderSlide(canvas, { index: 0, slideNumber: 1, background: null, elements: [shape] } as never,
        960 * 12700, 540 * 12700, { width: 960, dpr: 1 });
      return canvas.getContext('2d')!.getImageData(145, 105, 60, 40).data;
    };
    const results = [];
    for (const { rotation, flipH } of [
      { rotation: 0, flipH: false },
      { rotation: 45, flipH: false },
      { rotation: 45, flipH: true },
    ]) {
      const flat = await render(false, rotation, flipH);
      const beveled = await render(true, rotation, flipH);
      let same = 0;
      for (let i = 0; i < flat.length; i += 4) {
        if ((flat[i] > flat[i + 2]) === (beveled[i] > beveled[i + 2])) same++;
      }
      results.push(same / (flat.length / 4));
    }
    return results;
  });
  // Rotated bevels resample the auxiliary raster once more than flat shapes.
  // Their cell classes still agree across the central region despite edge aliasing.
  for (const agreement of agreements) expect(agreement).toBeGreaterThan(0.95);
});

test('PowerPoint pattern cells keep slide phase through rotation and reflection', async ({ page }) => {
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  await page.addScriptTag({ content: bundle.output[0].code });
  const colors = await page.evaluate(async () => {
    const renderer = (globalThis as typeof globalThis & {
      pptxRenderer: typeof import('../../packages/pptx/src/renderer.js');
    }).pptxRenderer;
    const cases = [
      { rotation: 0, flipH: false, flipV: false, preset: 'horz', samples: [[173, 120], [173, 124]] },
      { rotation: 45, flipH: false, flipV: false, preset: 'horz', samples: [[153, 120], [173, 120], [193, 120], [173, 124]] },
      { rotation: 0, flipH: true, flipV: false, preset: 'vert', samples: [[168, 123], [172, 123]] },
      { rotation: 0, flipH: false, flipV: true, preset: 'horz', samples: [[173, 120], [173, 124]] },
      { rotation: 45, flipH: true, flipV: false, preset: 'vert', samples: [[168, 123], [172, 123]] },
    ];
    const result: number[][][] = [];
    for (const dpr of [1, 2]) {
      const rows: number[][] = [];
      for (const c of cases) {
        const canvas = document.createElement('canvas');
        const slide = {
          index: 0, slideNumber: 1, background: null,
          elements: [{
            type: 'shape', x: 83 * 12700, y: 73 * 12700,
            width: 180 * 12700, height: 100 * 12700,
            geometry: 'rect', rotation: c.rotation, flipH: c.flipH, flipV: c.flipV,
            fill: { fillType: 'pattern', preset: c.preset, fg: '000000', bg: 'FFFFFF' },
            stroke: null,
          }],
        };
        await renderer.renderSlide(canvas, slide as never, 960 * 12700, 540 * 12700, { width: 960, dpr });
        const ctx = canvas.getContext('2d')!;
        rows.push(c.samples.map(([x, y]) => ctx.getImageData(x * dpr + Math.floor(dpr / 2), y * dpr + Math.floor(dpr / 2), 1, 1).data[0]));
      }
      result.push(rows);
    }
    return result;
  });
  for (const dprCases of colors) {
    expect(dprCases).toEqual([
      [0, 255],
      [0, 0, 0, 255],
      [0, 255],
      [0, 255],
      [0, 255],
    ]);
  }
});

test('transparent picture backing keeps the slide pattern grid through rotation and flip', async ({ page }) => {
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  await page.addScriptTag({ content: bundle.output[0].code });
  const regions = await page.evaluate(async () => {
    const renderer = (globalThis as typeof globalThis & {
      pptxRenderer: typeof import('../../packages/pptx/src/renderer.js');
    }).pptxRenderer;
    const source = document.createElement('canvas');
    source.width = source.height = 2;
    const blob = await new Promise<Blob>((resolve, reject) =>
      source.toBlob(value => value ? resolve(value) : reject(new Error('PNG encode failed')), 'image/png'));
    const result: number[][] = [];
    for (const transform of [
      { rotation: 0, flipH: false },
      { rotation: 315, flipH: false },
      { rotation: 0, flipH: true },
    ]) {
      const canvas = document.createElement('canvas');
      const slide = {
        index: 0, slideNumber: 1, background: null,
        elements: [{
          type: 'picture', x: 100 * 12700, y: 100 * 12700,
          width: 160 * 12700, height: 160 * 12700,
          ...transform, flipV: false, imagePath: 'transparent.png', mimeType: 'image/png', stroke: null,
          fill: { fillType: 'pattern', preset: 'horz', fg: 'D21D54', bg: '12CED4' },
        }],
      };
      await renderer.renderSlide(canvas, slide as never, 960 * 12700, 540 * 12700,
        { width: 960, dpr: 1, fetchImage: async () => blob });
      result.push(Array.from(canvas.getContext('2d')!.getImageData(150, 150, 60, 60).data));
    }
    return result;
  });
  expect(regions[0].some((channel, i) => i % 4 === 0 && channel === 210)).toBe(true);
  expect(regions[1]).toEqual(regions[0]);
  expect(regions[2]).toEqual(regions[0]);
});

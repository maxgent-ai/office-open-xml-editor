import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const entry = fileURLToPath(new URL('../../packages/docx/src/paint/canvas-page.ts', import.meta.url));

test('header, body and footer DrawingML patterns retain the page grid through rotation', async ({ page }) => {
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'docxPage' }, platform: 'browser' });
  await page.addScriptTag({ content: bundle.output[0].code });
  const samples = await page.evaluate(async () => {
    const { paintLayoutPage } = (globalThis as typeof globalThis & {
      docxPage: typeof import('../../packages/docx/src/paint/canvas-page.js');
    }).docxPage;
    const results: number[][][] = [];
    for (const scale of [1, 2]) {
      const cases: number[][] = [];
      for (const [story, rotationDeg, size] of [
        ['header', 0, 160], ['header', 315, 160], ['header', 315, 280],
        ['body', 0, 160], ['body', 315, 160],
        ['footer', 0, 160], ['footer', 315, 160],
      ] as const) {
        const canvas = document.createElement('canvas');
        // Deliberately offset the shape from the 8-point page grid.
        const x = (320 - size) / 2 + 3;
        const y = (320 - size) / 2 + 5;
        const bounds = { xPt: x, yPt: y, widthPt: size, heightPt: size };
        const drawing = {
          kind: 'drawing', id: `${story}-pattern`,
          source: { story, storyInstance: 'default', path: [0] },
          flowDomainId: story, flowBounds: bounds, inkBounds: bounds,
          advancePt: 0, ordinaryFlow: false,
          commands: [{ kind: 'drawingml-shape', plan: {
            rect: { x, y, w: size, h: size },
            geometry: { kind: 'preset', name: 'rect', adjustments: [] },
            fill: { fillType: 'pattern', preset: 'pct50', fg: 'D21D54', bg: '12CED4' },
            stroke: null,
            transform: { rotationDeg, flipH: false, flipV: false },
          } }],
        };
        const layout = { pages: [{
          geometry: { widthPt: 320, heightPt: 320 },
          flowDomains: [], sectionRegions: [], columnSeparators: [], pageBorder: null,
          layers: { paintOrder: [{
            kind: 'drawing', layer: 'behindText', sourceLayer: story,
            rootNodeId: `${story}-pattern`, coordinateSpace: 'section-logical',
            flowDomainId: story, node: drawing, textBoxes: [], frames: [],
            layoutTranslationPt: { xPt: 0, yPt: 0 },
          }] },
        }], diagnostics: [] };
        await paintLayoutPage(layout as never, 0, canvas, { scale, dpr: 1 });
        const ctx = canvas.getContext('2d')!;
        cases.push([[152, 152], [153, 152], [152, 153], [153, 153], [160, 160], [161, 160]]
          .map(([px, py]) => ctx.getImageData(Math.round((px + 0.5) * scale), Math.round((py + 0.5) * scale), 1, 1).data[0]));
      }
      results.push(cases);
    }
    return results;
  });
  for (const scaleCases of samples) {
    expect(scaleCases).toEqual(Array.from({ length: 7 }, () => [210, 18, 18, 210, 210, 18]));
  }
});

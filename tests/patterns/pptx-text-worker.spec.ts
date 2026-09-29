import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const entry = fileURLToPath(new URL('../../packages/pptx/src/renderer.ts', import.meta.url));
const pattern = { fillType: 'pattern', preset: 'horz', fg: 'D21D54', bg: '12CED4' };
const slide = {
  index: 0, slideNumber: 1, background: null,
  elements: [
    {
      type: 'shape', x: 914400, y: 914400, width: 5486400, height: 1828800,
      rotation: 0, flipH: false, flipV: false, geometry: 'rect',
      fill: null, stroke: null,
      textBody: {
        verticalAnchor: 't',
        paragraphs: [{
          alignment: 'l', marL: 0, marR: 0, indent: 0,
          spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
          bullet: { type: 'none' }, defFontSize: null, defColor: null,
          defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true,
          runs: [{ type: 'text', text: 'HHHHHH', bold: true, italic: false,
            underline: false, strikethrough: false, fontSize: 72,
            color: null, patternFill: pattern, fontFamily: 'Arial' }],
        }],
        defaultFontSize: null, defaultBold: null, defaultItalic: null,
        lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
      },
    },
    {
      type: 'chart', x: 914400, y: 3200400, width: 5486400, height: 1828800,
      rotation: 0, flipH: false, flipV: false,
      chart: {
        chartType: 'clusteredBar', categories: [], series: [],
        authoredWithoutSeries: true, chartFill: pattern,
        showLegend: false, showDataLabels: false,
        catAxisHidden: true, valAxisHidden: true,
      },
    },
  ],
};

function countColors(data: Uint8ClampedArray) {
  let red = 0, cyan = 0;
  for (let i = 0; i < data.length; i += 4) {
    if (data[i + 3] < 220) continue;
    if (data[i] > 170 && data[i + 1] < 80 && data[i + 2] < 150) red++;
    if (data[i] < 80 && data[i + 1] > 150 && data[i + 2] > 150) cyan++;
  }
  return { red, cyan };
}

test('plain PPTX glyph pattern renders on main canvas and OffscreenCanvas worker', async ({ page }) => {
  test.setTimeout(90_000);
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  const script = bundle.output[0].code;
  await page.addScriptTag({ content: script });
  const main = await page.evaluate(async (s) => {
    const canvas = document.createElement('canvas');
    const renderer = (globalThis as typeof globalThis & { pptxRenderer: typeof import('../../packages/pptx/src/renderer.js') }).pptxRenderer;
    await renderer.renderSlide(canvas, s as never, 9144000, 6858000, { width: 960, dpr: 1 });
    return Array.from(canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data);
  }, slide);
  const mainColors = countColors(new Uint8ClampedArray(main));
  expect(mainColors.red).toBeGreaterThan(50);
  expect(mainColors.cyan).toBeGreaterThan(50);

  const worker = await page.evaluate(async ({ script, slide }) => {
    const source = `${script}\nself.onmessage = async (event) => { const canvas = new OffscreenCanvas(1, 1); await pptxRenderer.renderSlide(canvas, event.data, 9144000, 6858000, { width: 960, dpr: 1 }); self.postMessage(Array.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data)); };`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    try {
      const w = new Worker(url);
      const result = await new Promise<number[]>((resolve, reject) => {
        w.onmessage = event => resolve(event.data);
        w.onerror = event => reject(new Error(event.message));
        w.postMessage(slide);
      });
      w.terminate();
      return result;
    } finally { URL.revokeObjectURL(url); }
  }, { script, slide });
  expect(countColors(new Uint8ClampedArray(worker))).toEqual(mainColors);
});

test('PPTX pattern glyphs retain their fill with a run reflection on main and worker canvases', async ({ page }) => {
  test.setTimeout(90_000);
  const reflectedSlide = structuredClone(slide);
  reflectedSlide.elements = [reflectedSlide.elements[0]];
  const run = reflectedSlide.elements[0].textBody.paragraphs[0].runs[0] as Record<string, unknown>;
  run.reflection = { blur: 0, dist: 25400, dir: 90, stA: 1, stPos: 0, endA: 0, endPos: 1, sx: 1, sy: -1 };
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  const script = bundle.output[0].code;
  await page.addScriptTag({ content: script });
  const colors = await page.evaluate(async ({ script, slide }) => {
    const renderer = (globalThis as typeof globalThis & { pptxRenderer: typeof import('../../packages/pptx/src/renderer.js') }).pptxRenderer;
    const render = async (canvas: HTMLCanvasElement | OffscreenCanvas) => {
      await renderer.renderSlide(canvas as HTMLCanvasElement, slide as never, 9144000, 6858000, { width: 960, dpr: 1 });
      const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      return Array.from(data);
    };
    const main = await render(document.createElement('canvas'));
    const source = `${script}\nself.onmessage = async (event) => { const canvas = new OffscreenCanvas(1, 1); await pptxRenderer.renderSlide(canvas, event.data, 9144000, 6858000, { width: 960, dpr: 1 }); self.postMessage(Array.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data)); };`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    try {
      const worker = new Worker(url);
      const offscreen = await new Promise<number[]>((resolve, reject) => {
        worker.onmessage = event => resolve(event.data);
        worker.onerror = event => reject(new Error(event.message));
        worker.postMessage(slide);
      });
      worker.terminate();
      return { main, offscreen };
    } finally { URL.revokeObjectURL(url); }
  }, { script, slide: reflectedSlide });
  const mainColors = countColors(new Uint8ClampedArray(colors.main));
  expect(mainColors.red).toBeGreaterThan(50);
  expect(mainColors.cyan).toBeGreaterThan(50);
  expect(countColors(new Uint8ClampedArray(colors.offscreen))).toEqual(mainColors);
});

test('PPTX underline follows patterned glyphs or its explicit uFill on main and worker', async ({ page }) => {
  test.setTimeout(90_000);
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  const script = bundle.output[0].code;
  await page.addScriptTag({ content: script });
  const results = await page.evaluate(async ({ script, base }) => {
    const renderer = (globalThis as typeof globalThis & {
      pptxRenderer: typeof import('../../packages/pptx/src/renderer.js');
    }).pptxRenderer;
    const colorCounts = (bytes: Uint8ClampedArray, width: number) => {
      let bestY = 0; let bestCount = -1;
      for (let y = 80; y < 220; y++) {
        let count = 0;
        for (let x = 100; x < 720; x++) {
          const i = (y * width + x) * 4;
          if (bytes[i] < 245 || bytes[i + 1] < 245 || bytes[i + 2] < 245) count++;
        }
        if (count > bestCount) { bestCount = count; bestY = y; }
      }
      const counts = { red: 0, cyan: 0, green: 0, orange: 0 };
      for (let x = 100; x < 720; x++) {
        const i = (bestY * width + x) * 4;
        const [r, g, b] = [bytes[i], bytes[i + 1], bytes[i + 2]];
        if (r > 160 && g < 100 && b < 160) counts.red++;
        if (r < 100 && g > 150 && b > 150) counts.cyan++;
        if (r < 90 && g > 120 && b < 130) counts.green++;
        if (r > 170 && g > 100 && b < 100) counts.orange++;
      }
      return counts;
    };
    const outputs = [];
    for (const underlineFill of [null,
      { fillType: 'pattern', preset: 'dnDiag', fg: '00A650', bg: 'F5A623' },
      { fillType: 'solid', color: '00A650' }]) {
      const s = structuredClone(base);
      s.elements = [s.elements[0]];
      const run = s.elements[0].textBody.paragraphs[0].runs[0] as Record<string, unknown>;
      run.text = 'HHHHHHHH'; run.underline = true;
      run.patternFill = { fillType: 'pattern', preset: 'pct50', fg: 'D21D54', bg: '12CED4' };
      if (underlineFill) run.underlineFill = underlineFill;
      const canvas = document.createElement('canvas');
      await renderer.renderSlide(canvas, s as never, 9144000, 6858000, { width: 960, dpr: 1 });
      const main = colorCounts(canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data, canvas.width);
      const source = `${script}\nself.onmessage = async (event) => { const canvas = new OffscreenCanvas(1, 1); await pptxRenderer.renderSlide(canvas, event.data, 9144000, 6858000, { width: 960, dpr: 1 }); self.postMessage({ width: canvas.width, bytes: Array.from(canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data) }); };`;
      const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
      try {
        const worker = new Worker(url);
        const offscreen = await new Promise<{ width: number; bytes: number[] }>((resolve, reject) => {
          worker.onmessage = event => resolve(event.data);
          worker.onerror = event => reject(new Error(event.message));
          worker.postMessage(s);
        });
        worker.terminate();
        outputs.push({ main, worker: colorCounts(new Uint8ClampedArray(offscreen.bytes), offscreen.width) });
      } finally { URL.revokeObjectURL(url); }
    }
    return outputs;
  }, { script, base: slide });
  for (const output of results) expect(output.worker).toEqual(output.main);
  expect(results[0].main.red).toBeGreaterThan(20);
  expect(results[0].main.cyan).toBeGreaterThan(20);
  expect(results[1].main.green).toBeGreaterThan(20);
  expect(results[1].main.orange).toBeGreaterThan(20);
  expect(results[2].main.green).toBeGreaterThan(200);
  expect(results[2].main.orange).toBe(0);
});

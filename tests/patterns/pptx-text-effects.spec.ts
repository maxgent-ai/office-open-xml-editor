import { expect, test } from '@playwright/test';
import { build } from 'rolldown';
import { fileURLToPath } from 'node:url';

const entry = fileURLToPath(new URL('../../packages/pptx/src/renderer.ts', import.meta.url));
const pattern = { fillType: 'pattern', preset: 'pct50', fg: 'D21D54', bg: '12CED4' };

function textSlide(effect: string) {
  const run: Record<string, unknown> = {
    type: 'text', text: 'HHHHHHHH', bold: true, italic: false,
    underline: false, strikethrough: false, fontSize: 72,
    color: null, patternFill: pattern, fontFamily: 'Arial',
  };
  const shape: Record<string, unknown> = {
    type: 'shape', x: 914400, y: 914400, width: 5486400, height: 1828800,
    rotation: 0, flipH: false, flipV: false, geometry: 'rect', fill: null, stroke: null,
  };
  const body: Record<string, unknown> = {
    verticalAnchor: 't', defaultFontSize: null, defaultBold: null, defaultItalic: null,
    lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
  };
  const paragraph: Record<string, unknown> = {
    alignment: 'l', marL: 0, marR: 0, indent: 0,
    spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
    bullet: { type: 'none' }, defFontSize: null, defColor: null,
    defBold: null, defItalic: null, defFontFamily: null, tabStops: [], eaLnBrk: true,
    runs: [run],
  };
  switch (effect) {
    case 'highlight': run.highlight = 'FFEB3B'; break;
    case 'highlight underline reflection':
      run.highlight = 'FFEB3B'; run.underline = true;
      run.reflection = { blur: 0, stPos: 0, endPos: 1, stA: 0.7, endA: 0, dist: 25400, dir: 90, sx: 1, sy: -1 };
      break;
    case 'underline': run.underline = true; break;
    case 'strike': run.strikethrough = true; break;
    case 'double strike': run.strikethrough = true; run.strikeDouble = true; break;
    case 'outline': run.outline = { width: 12700, color: null }; break;
    case 'outline solid': run.outline = { width: 12700, color: '00A650' }; break;
    case 'shadow': run.shadow = { color: '00A650', alpha: 0.65, blur: 0, dist: 25400, dir: 45 }; break;
    case 'reflection': run.reflection = { blur: 0, stPos: 0, endPos: 1, stA: 0.7, endA: 0, dist: 25400, dir: 90, sx: 1, sy: -1 }; break;
    case 'glow': shape.glow = { color: '00A650', alpha: 0.7, radius: 38100 }; break;
    case 'soft edge': shape.softEdge = { radius: 38100 }; break;
    case 'emboss': shape.sp3d = { prstMaterial: 'warmMatte', bevelT: { w: 38100, h: 38100, prst: 'circle' } }; break;
    case 'all caps': run.caps = 'all'; run.text = 'hhhhhhhh'; break;
    case 'small caps': run.caps = 'small'; run.text = 'hhhhhhhh'; break;
    case 'baseline': run.baseline = 30000; break;
    case 'hyperlink theme': run.hyperlink = 'https://example.test/'; run.color = '0000FF'; break;
    case 'hyperlink text fill': run.hyperlink = 'https://example.test/'; run.hyperlinkUsesTextFill = true; break;
    case 'field': run.fieldType = 'datetime'; break;
    case 'vertical': body.vert = 'vert'; break;
    case 'warp arch': body.textWarp = { preset: 'textArchUp', adj: [] }; break;
    case 'warp inflate': body.textWarp = { preset: 'textInflate', adj: [] }; break;
    case 'bullet':
    case 'bullet solid':
      paragraph.bullet = { type: 'char', char: '•', color: effect === 'bullet solid' ? '00A650' : null, fontFamily: null, sizePct: 100 };
      paragraph.marL = 457200; paragraph.indent = -228600; break;
  }
  body.paragraphs = [paragraph];
  shape.textBody = body;
  return { index: 0, slideNumber: 1, background: null, elements: [shape] };
}

const effects = [
  'highlight', 'highlight underline reflection', 'underline', 'strike', 'double strike',
  'outline', 'outline solid', 'shadow', 'reflection', 'glow', 'soft edge', 'emboss',
  'all caps', 'small caps', 'baseline',
  'hyperlink theme', 'hyperlink text fill', 'field', 'vertical', 'warp arch', 'warp inflate', 'bullet', 'bullet solid',
];

test('pattern text survives every painted effect on main and worker canvases', async ({ page }) => {
  test.setTimeout(120_000);
  const bundle = await build({ input: entry, output: { format: 'iife', name: 'pptxRenderer' }, platform: 'browser' });
  const script = bundle.output[0].code;
  await page.addScriptTag({ content: script });
  const results = await page.evaluate(async ({ script, cases }) => {
    const renderer = (globalThis as typeof globalThis & {
      pptxRenderer: typeof import('../../packages/pptx/src/renderer.js');
    }).pptxRenderer;
    const source = `${script}\nself.onmessage = async event => { const c = new OffscreenCanvas(1,1); await pptxRenderer.renderSlide(c,event.data,9144000,6858000,{width:960,dpr:1}); const d=c.getContext('2d').getImageData(0,0,c.width,c.height).data; self.postMessage({width:c.width,buffer:d.buffer},[d.buffer]); };`;
    const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
    const worker = new Worker(url);
    const count = (data: Uint8ClampedArray) => {
      let red = 0; let cyan = 0; let black = 0; let blue = 0;
      for (let i = 0; i < data.length; i += 4) {
        if (data[i + 3] < 220) continue;
        if (data[i] > 170 && data[i + 1] < 80 && data[i + 2] < 150) red++;
        if (data[i] < 80 && data[i + 1] > 150 && data[i + 2] > 150) cyan++;
        if (data[i] < 40 && data[i + 1] < 40 && data[i + 2] < 40) black++;
        if (data[i] < 70 && data[i + 1] < 70 && data[i + 2] > 170) blue++;
      }
      return { red, cyan, black, blue };
    };
    const markerCount = (data: Uint8ClampedArray, width: number) => {
      let red = 0; let cyan = 0; let green = 0;
      for (let y = 90; y < 250; y++) for (let x = 100; x < 135; x++) {
        const i = (y * width + x) * 4;
        if (data[i] > 170 && data[i + 1] < 80 && data[i + 2] < 150) red++;
        if (data[i] < 80 && data[i + 1] > 150 && data[i + 2] > 150) cyan++;
        if (data[i] < 80 && data[i + 1] > 115 && data[i + 2] < 135) green++;
      }
      return { red, cyan, green };
    };
    try {
      const output = [];
      for (const { effect, slide } of cases) {
        const canvas = document.createElement('canvas');
        await renderer.renderSlide(canvas, slide as never, 9144000, 6858000, { width: 960, dpr: 1 });
        const mainBytes = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
        const main = count(mainBytes);
        const bytes = await new Promise<Uint8ClampedArray>((resolve, reject) => {
          worker.onmessage = event => resolve(new Uint8ClampedArray(event.data.buffer));
          worker.onerror = event => reject(new Error(event.message));
          worker.postMessage(slide);
        });
        output.push({
          effect, main, worker: count(bytes),
          marker: effect.startsWith('bullet')
            ? { main: markerCount(mainBytes, canvas.width), worker: markerCount(bytes, canvas.width) }
            : null,
        });
      }
      return output;
    } finally { worker.terminate(); URL.revokeObjectURL(url); }
  }, { script, cases: effects.map(effect => ({ effect, slide: textSlide(effect) })) });
  for (const { effect, main, worker, marker } of results) {
    expect(main, effect).toEqual(worker);
    if (effect === 'hyperlink theme') {
      expect(main.blue, effect).toBeGreaterThan(100);
      expect(main.red, effect).toBe(0);
      expect(main.cyan, effect).toBe(0);
    } else {
      expect(main.red, effect).toBeGreaterThan(100);
      expect(main.cyan, effect).toBeGreaterThan(100);
    }
    expect(main.black, effect).toBeLessThan(100);
    if (marker) {
      expect(marker.main, effect).toEqual(marker.worker);
      if (effect === 'bullet') {
        expect(marker.main.red).toBeGreaterThan(3);
        expect(marker.main.cyan).toBeGreaterThan(3);
      } else {
        expect(marker.main.green).toBeGreaterThan(3);
        expect(marker.main.red).toBe(0);
        expect(marker.main.cyan).toBe(0);
      }
    }
  }
});

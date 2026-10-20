import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { build } from 'vite';
import { existsSync } from 'node:fs';
import { PNG } from 'pngjs';
import type { PictureElement, ShapeElement } from './types.js';

function orthographicBeveledShape(): ShapeElement {
  return {
    type: 'shape',
    x: 914_400,
    y: 914_400,
    width: 914_400,
    height: 914_400,
    rotation: 0,
    flipH: false,
    flipV: false,
    geometry: 'roundRect',
    fill: { fillType: 'solid', color: '4472C4' },
    stroke: null,
    textBody: null,
    defaultTextColor: null,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
    shadow: null,
    scene3d: {
      camera: { prst: 'orthographicFront' },
      lightRig: { rig: 'chilly', dir: 't' },
    },
    sp3d: {
      prstMaterial: 'translucentPowder',
      bevelT: { w: 127_000, h: 25_400, prst: 'softRound' },
    },
  };
}

function projectedBeveledShape(): ShapeElement {
  const shape = orthographicBeveledShape();
  return {
    ...shape,
    scene3d: {
      camera: { prst: 'isometricLeftUp' },
      lightRig: { rig: 'chilly', dir: 't' },
    },
  };
}


// Real browser pixels cover material ownership and camera routing; allocation
// counts cannot distinguish an extra warp buffer from a second material pass.
describe.skipIf(!existsSync(chromium.executablePath()))('shape bevel and camera pixels', () => {
  let browser: Browser;
  let page: Page;
  beforeAll(async () => {
    const result = await build({
      configFile: false, logLevel: 'error',
      define: { __OOXML_MODEL_SOURCES__: 'true' },
      plugins: [{ name: 'text-projection-test', resolveId: id => id === '/__text-projection-test.ts' ? id : null,
        load: id => id === '/__text-projection-test.ts' ?
          `export * from ${JSON.stringify(new URL('./renderer.ts', import.meta.url).pathname)};
           export { computeScene3dQuad } from ${JSON.stringify(new URL('../../core/src/shape/scene3d-camera.ts', import.meta.url).pathname)};
           export { drawProjected, projectQuadPoint } from ${JSON.stringify(new URL('../../core/src/shape/scene3d-draw.ts', import.meta.url).pathname)};` : null }],
      // Preserve core module initialization in this isolated renderer bundle.
      build: { write: false, minify: false, rolldownOptions: { treeshake: false },
        lib: { entry: '/__text-projection-test.ts',
          formats: ['iife'], name: 'pptxRenderer' } },
    });
    const outputs = Array.isArray(result) ? result : [result];
    const chunk = outputs.flatMap(output => 'output' in output ? output.output : [])
      .find(output => output.type === 'chunk');
    if (!chunk || chunk.type !== 'chunk') throw new Error('Renderer bundle missing');
    browser = await chromium.launch({ channel: 'chromium' });
    page = await browser.newPage();
    await page.setContent('<!doctype html><canvas></canvas>');
    await page.addScriptTag({ content: chunk.code });
  }, 60_000);
  afterAll(async () => { await browser?.close(); });

  async function render(shape: ShapeElement | PictureElement, height = 400, sourceAlpha = 1, width = 400, dpr = 1) {
    const png = await page.evaluate(async ({ shape, height, sourceAlpha, width, dpr }) => {
      const canvas = document.querySelector('canvas') as HTMLCanvasElement;
      const renderer = (globalThis as unknown as { pptxRenderer: typeof import('./renderer') }).pptxRenderer;
      await renderer.renderSlide(canvas,
        { index: 0, slideNumber: 1, background: null, elements: [shape] },
        400 * 9525, height * 9525, { width, dpr, fetchImage: async () => {
          const source = document.createElement('canvas');
          source.width = source.height = 96;
          const ctx = source.getContext('2d') as CanvasRenderingContext2D;
          ctx.fillStyle = `rgba(68,114,196,${sourceAlpha})`;
          ctx.fillRect(0, 0, 96, 96);
          return new Promise<Blob>(resolve => source.toBlob(blob => resolve(blob as Blob)));
        } });
      return canvas.toDataURL('image/png').split(',')[1];
    }, { shape, height, sourceAlpha, width, dpr });
    return PNG.sync.read(Buffer.from(png, 'base64')).data;
  }
  const pixel = (data: Uint8Array, x: number, y: number) => [...data.subarray((y * 400 + x) * 4, (y * 400 + x) * 4 + 3)];

  it('includes the complete mitered outline in the inner-shadow coverage', async () => {
    const shape: ShapeElement = {
      ...orthographicBeveledShape(), geometry: 'triangle', scene3d: undefined, sp3d: undefined,
      stroke: { color: '4472C4', width: 20 * 9525,
        lineJoin: 'miter', lineCap: 'square' },
    };
    const plain = await render(shape);
    const shadowed = await render({ ...shape,
      innerShadow: { color: '000000', alpha: .5, blur: 0, dist: 400 * 9525, dir: 0 },
    });
    let opaqueStrokePixels = 0;
    for (let p = 0; p < plain.length; p += 4) {
      if (plain[p] !== 68 || plain[p + 1] !== 114 || plain[p + 2] !== 196) continue;
      opaqueStrokePixels++;
      // The offset is disjoint: every opaque outline pixel, including the
      // miter outside nominal geometry, receives 50% black.
      for (const [channel, expected] of [34, 57, 98].entries())
        expect(Math.abs(shadowed[p + channel] - expected)).toBeLessThanOrEqual(1);
    }
    expect(opaqueStrokePixels).toBeGreaterThan(100);
  });

  for (const projected of [false, true]) {
    it(`applies translucent material once (${projected ? 'projected' : 'identity'} camera)`, async () => {
      const shape = projected ? projectedBeveledShape() : orthographicBeveledShape();
      shape.fill = { fillType: 'solid', color: '50648C' };
      shape.scene3d!.lightRig = { rig: 'contrasting', dir: 't' };
      const data = await render(shape);
      // Contrasting has ambient=1 and no diffuse. The flat material has alpha
      // .7, so (80,100,140) over white must be (133,147,175), within rounding.
      const center = pixel(data, 144, 144);
      for (const [i, expected] of [133, 147, 175].entries())
        expect(Math.abs(center[i] - expected)).toBeLessThanOrEqual(1);
    });
  }

  it('keeps projected text unshaded when an invisible shadow selects the effect path', async () => {
    const shape = projectedBeveledShape();
    shape.geometry = 'rect';
    shape.width *= 2;
    shape.textBody = {
      verticalAnchor: 'ctr', defaultFontSize: 32, defaultBold: false, defaultItalic: false,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'none', vert: 'horz', autoFit: 'none',
      paragraphs: [{ alignment: 'ctr', marL: 0, marR: 0, indent: 0,
        spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
        bullet: { type: 'none' }, defFontSize: 32, defColor: 'FF0000',
        defBold: false, defItalic: false, defFontFamily: 'Arial', tabStops: [], eaLnBrk: true,
        runs: [{ type: 'text', text: 'MMMM', bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 32, color: 'FF0000', fontFamily: 'Arial' }] }],
    };
    const plain = await render(shape);
    const shadow = await render({ ...shape, shadow: { color: '000000', alpha: 0, blur: 0, dist: 0, dir: 0 } });
    const redPixels = (data: Uint8Array) => Array.from(data).reduce<number[]>((indices, r, i) => {
      if (i % 4 === 0 && r === 255 && data[i + 1] === 0 && data[i + 2] === 0) indices.push(i);
      return indices;
    }, []);
    expect(redPixels(plain).length).toBeGreaterThan(50);
    expect(redPixels(plain)).toEqual(redPixels(shadow));
  });

  for (const mode of ['horz', 'insets', 'vert', 'vert270', 'eaVert', 'wordArtVert', 'rotation', 'perspective'] as const) {
    it(`preserves overflowing projected text (${mode})`, async () => {
      const shape = projectedBeveledShape();
      shape.geometry = 'rect';
      shape.fill = null;
      shape.x = shape.y = 152 * 9525;
      shape.rotation = mode === 'rotation' ? 25 : 0;
      shape.scene3d!.camera = mode === 'perspective'
        ? { prst: 'perspectiveFront', rot: { lat: 15, lon: 30, rev: 0 } }
        : { prst: 'orthographicFront', rot: { lat: 0, lon: 30, rev: 0 } };
      shape.sp3d = { prstMaterial: 'matte', extrusionH: 80 * 9525,
        bevelT: { w: 127000, h: 127000, prst: 'angle' } };
      shape.textBody = {
        verticalAnchor: 'ctr', defaultFontSize: 32, defaultBold: false, defaultItalic: false,
        lIns: mode === 'insets' ? -24 * 9525 : 0, rIns: 0,
        tIns: mode === 'insets' ? 32 * 9525 : 0, bIns: 0,
        wrap: 'none', vert: ['insets', 'rotation', 'perspective'].includes(mode) ? 'horz' : mode, autoFit: 'none',
        paragraphs: [{ alignment: 'ctr', marL: 0, marR: 0, indent: 0,
          spaceBefore: null, spaceAfter: null, spaceLine: null, lvl: 0,
          bullet: { type: 'none' }, defFontSize: 32, defColor: 'FF0000',
          defBold: false, defItalic: false, defFontFamily: 'Arial', tabStops: [], eaLnBrk: true,
          runs: [{ type: 'text', text: mode === 'eaVert' ? '漢MMMM漢' : 'MMMMMMMM', bold: false, italic: false, underline: false,
            strikethrough: false, fontSize: 32, color: 'FF0000', fontFamily: 'Arial' }] }],
      };
      const expected = await page.evaluate(async shape => {
        const renderer = (globalThis as unknown as { pptxRenderer: typeof import('./renderer') & typeof import('../../core/src/shape/scene3d-draw') & typeof import('../../core/src/shape/scene3d-camera') }).pptxRenderer;
        const source = document.createElement('canvas');
        // Paint all flat text on a full slide before projection, independently
        // of the production text-raster sizing. The affine case uses analytic
        // corners; perspective reuses only the authored camera homography.
        await renderer.renderSlide(source, { index: 0, slideNumber: 1, background: null,
          elements: [{ ...shape, rotation: 0, fill: null, stroke: null, scene3d: undefined, sp3d: undefined }] },
          400 * 9525, 400 * 9525, { width: 400, dpr: 1 });
        const canvas = document.createElement('canvas');
        canvas.width = canvas.height = 400;
        const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 400, 400);
        ctx.translate(200, 200);
        ctx.rotate(shape.rotation * Math.PI / 180);
        ctx.translate(-200, -200);
        const left = 200 - 200 * Math.cos(Math.PI / 6);
        const right = 200 + 200 * Math.cos(Math.PI / 6);
        let corners = [{ x: left, y: 0 }, { x: right, y: 0 }, { x: right, y: 400 }, { x: left, y: 400 }];
        if (shape.scene3d?.camera.prst === 'perspectiveFront') {
          const face = renderer.computeScene3dQuad(shape.scene3d.camera, 96, 96).corners;
          corners = [[0, 0], [400, 0], [400, 400], [0, 400]].map(([x, y]) => {
            const point = renderer.projectQuadPoint(face, (x - 152) / 96, (y - 152) / 96) as { x: number; y: number };
            return { x: point.x + 152, y: point.y + 152 };
          });
        }
        renderer.drawProjected(source, ctx, 400, 400,
          corners as Parameters<typeof renderer.drawProjected>[4]);
        return canvas.toDataURL('image/png').split(',')[1];
      }, shape);
      const reference = PNG.sync.read(Buffer.from(expected, 'base64')).data;
      const ink = (data: Uint8Array) => {
        let count = 0, left = 400, right = 0, top = 400, bottom = 0;
        for (let i = 0; i < data.length; i += 4) {
          if (data[i] < 240 || data[i + 1] > 100 || data[i + 2] > 100) continue;
          const x = i / 4 % 400, y = Math.floor(i / 4 / 400);
          count++; left = Math.min(left, x); right = Math.max(right, x);
          top = Math.min(top, y); bottom = Math.max(bottom, y);
        }
        return { count, bounds: [left, right, top, bottom] };
      };
      const wanted = ink(reference);
      expect(wanted.count).toBeGreaterThan(500);
      for (const effects of [{}, { shadow: { color: '000000', alpha: 0, blur: 0, dist: 0, dir: 0 } }]) {
        const actual = ink(await render({ ...shape, ...effects }));
        expect(actual.count).toBeGreaterThan(wanted.count * .9);
        expect(actual.count).toBeLessThan(wanted.count * 1.1);
        actual.bounds.forEach((value, i) => expect(Math.abs(value - wanted.bounds[i])).toBeLessThanOrEqual(2));
      }
      if (mode === 'horz') {
        // Ordinary filled faces exercise the same overflow contract. Count
        // only bounds here: blue beneath the text changes edge-pixel RGB.
        const filled = ink(await render({ ...shape, fill: { fillType: 'solid', color: '4472C4' } }));
        filled.bounds.forEach((value, i) => expect(Math.abs(value - wanted.bounds[i])).toBeLessThanOrEqual(2));
      }
    });
  }

  it('projects a 100 by 1000 shape under a one-degree perspective tilt', async () => {
    const shape = orthographicBeveledShape();
    Object.assign(shape, { x: 50 * 9525, y: 80 * 9525, width: 100 * 9525,
      height: 1000 * 9525, geometry: 'rect', sp3d: undefined,
      scene3d: { camera: { prst: 'perspectiveFront', rot: { lat: 1, lon: 0, rev: 0 } } } });
    const data = await render(shape, 1200);
    // The tilted near edge extends above y=80 by roughly four pixels.
    expect(pixel(data, 100, 78)).toEqual([68, 114, 196]);
  });

  it('shows a back bevel only when the camera exposes the back face', async () => {
    const shape = orthographicBeveledShape();
    shape.sp3d = { prstMaterial: 'matte', bevelB: { w: 127000, h: 127000, prst: 'angle' } };
    const front = await render(shape);
    expect(pixel(front, 144, 144)).toEqual([68, 114, 196]);
    expect(pixel(front, 144, 105)).toEqual([68, 114, 196]);
    shape.scene3d!.camera = { prst: 'orthographicFront', rot: { lat: 0, lon: 180, rev: 0 } };
    const back = await render(shape);
    expect(pixel(back, 144, 144)).not.toEqual(pixel(await render({ ...shape, sp3d: undefined }), 144, 144));
    expect(pixel(back, 144, 105)).not.toEqual(pixel(back, 144, 144));
    shape.scene3d!.camera.rot!.lon = 150;
    shape.sp3d.extrusionH = 24 * 9525;
    const extruded = await render(shape);
    // The back plane moves 12 px right (24 * sin(150 degrees)). Its bevel
    // remains visible near x=190, with the wall extending toward the front.
    expect(pixel(extruded, 190, 144)).not.toEqual([255, 255, 255]);
    expect(pixel(extruded, 190, 144)).not.toEqual(pixel(extruded, 156, 144));
  });

  for (const adapter of ['shape', 'picture'] as const) {
    function element(shape: ShapeElement): ShapeElement | PictureElement {
      if (adapter === 'shape') return shape;
      return {
        type: 'picture', x: shape.x, y: shape.y, width: shape.width, height: shape.height,
        rotation: shape.rotation, flipH: shape.flipH, flipV: shape.flipV,
        imagePath: 'test/camera.png', mimeType: 'image/png', stroke: null,
        scene3d: shape.scene3d, sp3d: shape.sp3d, shadow: shape.shadow ?? undefined, softEdge: shape.softEdge,
      };
    }

    it(`preserves physical camera framing at different render and device scales (${adapter})`, async () => {
      const shape = orthographicBeveledShape();
      Object.assign(shape, { x: 48 * 9525, y: 96 * 9525, width: 288 * 9525,
        geometry: 'rect', sp3d: undefined,
        scene3d: { camera: { prst: 'perspectiveFront', fov: 90,
          rot: { lat: 0, lon: 30, rev: 0 } } } });
      // Independent pinhole bounds in 96dpi coordinates: fixed 6.28-inch
      // distance, 3-inch face, 30-degree longitude, no renderer helper oracle.
      const d = 6.28 * 96;
      const halfProjectedWidth = 144 * Math.cos(Math.PI / 6);
      const expected = [192 - halfProjectedWidth / (1 + 72 / d),
        192 + halfProjectedWidth / (1 - 72 / d)];
      for (const [width, dpr] of [[300, 1], [600, 2]]) {
        const data = await render(element(shape), 400, 1, width, dpr);
        const side = width * dpr;
        const bounds = [side, 0];
        for (let i = 0; i < data.length; i += 4) {
          if (data[i] > 200 || data[i + 2] < 150) continue;
          const x = i / 4 % side;
          bounds[0] = Math.min(bounds[0], x); bounds[1] = Math.max(bounds[1], x);
        }
        // Pixel-centre ownership and the existing warp mesh can move the
        // thresholded edge by one device pixel.
        bounds.forEach((v, i) => expect(Math.abs(v + .5 - expected[i] * side / 400), `edge ${i}, width ${width}, dpr ${dpr}`)
          .toBeLessThanOrEqual(1));
      }
    });

    it(`keeps a face-on inset extrusion on the unchanged front paint path (${adapter})`, async () => {
      const shape = orthographicBeveledShape();
      shape.width = 288 * 9525; shape.height = 300 * 9525; shape.x = 48 * 9525;
      shape.scene3d!.camera = { prst: 'perspectiveFront', fov: 90 };
      const plain = await render(element(shape));
      const deep = await render(element({ ...shape,
        sp3d: { ...shape.sp3d, prstMaterial: 'translucentPowder', extrusionH: 2 * 12700 } }));
      expect(deep).toEqual(plain);
    });

    it(`keeps camera-crossing geometry on the same side at CSS and device scales (${adapter})`, async () => {
      const shape = orthographicBeveledShape();
      Object.assign(shape, { x: (200 - 768) * 9525, y: 8 * 9525,
        width: 1536 * 9525, height: 384 * 9525, geometry: 'rect', sp3d: undefined,
        scene3d: { camera: { prst: 'perspectiveFront', fov: 90, zoom: .01,
          rot: { lat: 0, lon: 60, rev: 0 } } } });
      // Independent view-Z fallback for a 16in-wide solid: nearest=8*sin60,
      // gap=.02*(16*sin60). Positive depth puts the near edge to the RIGHT;
      // a signed corner clamp projects it left and folds the texture.
      const near = 768 * Math.sin(Math.PI / 3), distance = near * 1.04;
      const expected = [200 - 384 * .01 / (1 + near / distance),
        200 + 384 * .01 / (1 - near / distance)];
      for (const [width, dpr] of [[300, 1], [600, 2]]) {
        const data = await render(element(shape), 400, 1, width, dpr);
        const side = width * dpr, bounds = [side, -1];
        for (let i = 0; i < data.length; i += 4) {
          if (data[i] > 200 || data[i + 2] < 150) continue;
          const x = i / 4 % side;
          bounds[0] = Math.min(bounds[0], x); bounds[1] = Math.max(bounds[1], x);
        }
        expect(Math.abs(bounds[0] + .5 - expected[0] * side / 400)).toBeLessThanOrEqual(1);
        // Test an interior ray rather than the near silhouette: the existing
        // triangle warp can truncate that extremely magnified raster edge.
        const sampleX = Math.floor((200 + expected[1]) / 2 * side / 400);
        const sampleY = Math.floor(200 * side / 400);
        const sample = (sampleY * side + sampleX) * 4;
        expect([...data.subarray(sample, sample + 3)]).toEqual([68, 114, 196]);
      }
    });

    it(`binds the full extrusion to the front camera before device-space wall painting (${adapter})`, async () => {
      const shape = orthographicBeveledShape();
      Object.assign(shape, { x: (200 - 384) * 9525, y: 152 * 9525,
        width: 768 * 9525, height: 96 * 9525, geometry: 'rect',
        sp3d: { extrusionH: 768 * 9525, extrusionClr: '404040' },
        scene3d: { camera: { prst: 'perspectiveFront', fov: 90, zoom: .3,
          rot: { lat: 0, lon: 120, rev: 0 } } } });
      const extent = 384 * Math.sin(Math.PI / 3);
      const distance = extent + 384 + .02 * (2 * extent + 384);
      const expected = [200 - 192 * .3 / (1 - extent / distance),
        200 + 192 * .3 / (1 + extent / distance)];
      // Keep the existing pixel-by-pixel extrusion sweep small; physical
      // document dimensions and the camera boundary are identical at any DPI.
      for (const [width, dpr] of [[40, 1], [80, 2]]) {
        const data = await render(element(shape), 400, 1, width, dpr);
        const side = width * dpr, bounds = [side, -1];
        for (let i = 0; i < data.length; i += 4) {
          // The blue front is distinct from the neutral extrusion walls.
          if (data[i] > 150 || data[i + 2] < 150) continue;
          const x = i / 4 % side;
          bounds[0] = Math.min(bounds[0], x); bounds[1] = Math.max(bounds[1], x);
        }
        bounds.forEach((v, i) => expect(Math.abs(v + .5 - expected[i] * side / 400),
          `edge ${i}, width ${width}, dpr ${dpr}`).toBeLessThanOrEqual(1));
      }
    });

    it(`uses geometry and shadow color independently of body opacity (${adapter})`, async () => {
      const shape = projectedBeveledShape();
      shape.geometry = 'rect';
      shape.sp3d = undefined;
      shape.fill = { fillType: 'solid', color: '4472C480' };
      const body = element(shape);
      if (body.type === 'picture') body.imagePath = 'test/translucent-shadow.png';
      const plain = await render(body, 400, .5);
      const invisible = await render({ ...body,
        innerShadow: { color: 'FF0000', alpha: 0, blur: 0, dist: 200 * 9525, dir: 0 } }, 400, .5);
      // A fully offset mask exposes the whole shadow, independently of inset
      // placement. Alpha zero must leave a 50%-opaque body unchanged.
      expect(pixel(plain, 144, 144)).toEqual([161, 184, 225]);
      expect(pixel(invisible, 144, 144)).toEqual(pixel(plain, 144, 144));
      const red = await render({ ...body,
        innerShadow: { color: 'FF0000', alpha: .5, blur: 0, dist: 200 * 9525, dir: 0 } }, 400, .5);
      for (const [i, expected] of [208, 92, 112].entries())
        expect(Math.abs(pixel(red, 144, 144)[i] - expected)).toBeLessThanOrEqual(1);
      const softened = await render({ ...body, softEdge: { radius: 9525 } }, 400, .5);
      // Soft edges change the perimeter, not source opacity in the interior.
      expect(pixel(softened, 144, 144)).toEqual(pixel(plain, 144, 144));
    });

    it(`keeps translucent wireframe topology across 49% and 50% source opacity (${adapter})`, async () => {
      const shape = orthographicBeveledShape();
      shape.geometry = 'rect';
      shape.scene3d!.camera = { prst: 'obliqueTop' };
      shape.sp3d = { prstMaterial: 'legacyWireframe', extrusionH: 60 * 9525,
        bevelT: { w: 127000, h: 127000, prst: 'angle' } };
      const support: number[][] = [];
      for (const alpha of [.49, .5]) {
        shape.fill = { fillType: 'solid', color: `4472C4${Math.round(alpha * 255).toString(16)}` };
        const body = element(shape);
        if (body.type === 'picture') body.imagePath = `test/wire-${alpha}.png`;
        const data = await render(body, 400, alpha);
        expect(pixel(data, 144, 144)).toEqual([255, 255, 255]);
        expect(pixel(data, 144, 79)).not.toEqual([255, 255, 255]); // back rim
        expect(pixel(data, 109, 144)).not.toEqual([255, 255, 255]); // profile rim
        support.push(Array.from(data).flatMap((value, i) => i % 4 === 0 && value < 250 ? [i] : []));
      }
      expect(support[0].length).toBeGreaterThan(100);
      expect(support[0]).toEqual(support[1]);
    });

    it(`keeps an extruded back face intact with an invisible shadow (${adapter})`, async () => {
      const shape = orthographicBeveledShape();
      shape.geometry = 'rect';
      shape.x = 180 * 9525;
      shape.scene3d!.camera = { prst: 'orthographicFront', rot: { lat: 0, lon: 110, rev: 0 } };
      shape.sp3d = { prstMaterial: 'matte', extrusionH: 24 * 9525,
        bevelB: { w: 127000, h: 127000, prst: 'angle' } };
      const plain = await render(element(shape));
      const shadow = await render(element({ ...shape,
        shadow: { color: '000000', alpha: 0, blur: 0, dist: 0, dir: 0 } }));
      expect(pixel(plain, 240, 144)).not.toEqual([255, 255, 255]);
      // An intermediate transparent cache can round premultiplied channels
      // by one; it must preserve the whole back face, not clip its far edge.
      const maxDifference = plain.reduce((max, value, i) => Math.max(max, Math.abs(value - shadow[i])), 0);
      expect(maxDifference).toBeLessThanOrEqual(1);
      if (adapter === 'picture') {
        const softened = await render(element({ ...shape, softEdge: { radius: 9525 } }));
        // This interior lies beyond the front-plane mask, but well inside the
        // back plane: a soft edge must not erase it using a stale front mask.
        expect(pixel(softened, 250, 144)).not.toEqual([255, 255, 255]);
      }
    });

    for (const wireframe of [false, true]) {
      it(`retains default oblique ${wireframe ? 'wireframe' : 'solid'} depth (${adapter})`, async () => {
        const shape = orthographicBeveledShape();
        shape.geometry = 'rect';
        shape.scene3d!.camera = { prst: 'obliqueTop' };
        shape.sp3d = { prstMaterial: wireframe ? 'legacyWireframe' : 'matte', extrusionH: 60 * 9525 };
        const deep = await render(element(shape));
        const flat = await render(element({ ...shape, sp3d: { ...shape.sp3d, extrusionH: 0 } }));
        // Default obliqueTop leaves the front unchanged and pushes the back
        // 18 px upward. Both its solid wall and wire rim extend above y=96.
        expect(pixel(deep, 144, 79)).not.toEqual([255, 255, 255]);
        expect(pixel(flat, 144, 79)).toEqual([255, 255, 255]);
      });

      it(`softens the complete oblique ${wireframe ? 'wireframe' : 'solid'} geometry (${adapter})`, async () => {
        const shape = orthographicBeveledShape();
        shape.geometry = 'rect';
        shape.scene3d!.camera = { prst: 'obliqueTop' };
        shape.sp3d = { prstMaterial: wireframe ? 'legacyWireframe' : 'matte', extrusionH: 60 * 9525 };
        const plain = await render(element(shape));
        const softened = await render(element({ ...shape, softEdge: { radius: 9525 } }));
        // The 18 px extrusion lies above the front face's y=96 boundary.
        // A face-only mask erases the wall interior and displaced wire rim.
        if (wireframe) {
          expect(pixel(plain, 144, 79)).not.toEqual([255, 255, 255]);
          expect(pixel(softened, 144, 79)).not.toEqual([255, 255, 255]);
          expect(pixel(softened, 144, 144)).toEqual([255, 255, 255]);
        } else {
          for (const y of [85, 90]) {
            expect(pixel(plain, 144, y)).toEqual([64, 64, 64]);
            expect(pixel(softened, 144, y)).toEqual(pixel(plain, 144, y));
          }
        }
      });
    }
  }

  for (const camera of ['orthographicFront', 'perspectiveRelaxed'] as const) {
    it(`shades a picture's circle bevel as a convex rim (${camera})`, async () => {
      // Office renders circle as a quarter round: steepest, hence darkest on a
      // side edge under a top light, at the silhouette, brightening inward into
      // the face. Reading the §2.1.1316 points as (height, inward) instead
      // produced an unlit rim plus a dark crease at the inner band edge.
      const band = 24;
      const picture: PictureElement = {
        type: 'picture', x: 48 * 9525, y: 48 * 9525, width: 288 * 9525, height: 288 * 9525,
        rotation: 0, flipH: false, flipV: false, imagePath: `test/bevel-${camera}.png`,
        mimeType: 'image/png', stroke: null, prstGeom: 'roundRect', prstAdjust: [16_667],
        scene3d: {
          camera: camera === 'orthographicFront' ? { prst: camera }
            : { prst: camera, rot: { lat: 330, lon: 20, rev: 347 } },
          lightRig: { rig: 'threePt', dir: 't' },
        },
        sp3d: { prstMaterial: 'matte', contourW: 6350, contourClr: '969696',
          bevelT: { w: band * 9525, h: band * 9525, prst: 'circle' } },
      };
      const data = await render(picture);
      const y = 192;
      let edge = 399;
      while (edge > 0 && pixel(data, edge, y).every(v => v >= 250)) edge--;
      expect(edge).toBeGreaterThan(240);
      // Skip the contour/antialiased silhouette pixels, then walk inward.
      const lum = Array.from({ length: band + 8 }, (_, k) => {
        const [r, g, b] = pixel(data, edge - 3 - k, y);
        return 0.299 * r + 0.587 * g + 0.114 * b;
      });
      const darkest = lum.indexOf(Math.min(...lum));
      expect(darkest).toBeLessThan(band / 4);
      for (let k = 1; k < lum.length; k++) expect(lum[k]).toBeGreaterThanOrEqual(lum[k - 1] - 2);
      expect(lum[lum.length - 1] - lum[0]).toBeGreaterThan(20);
    });
  }

  it('renders wireframe rims with a transparent interior', async () => {
    const shape = projectedBeveledShape();
    shape.sp3d!.prstMaterial = 'legacyWireframe';
    const data = await render(shape);
    expect(pixel(data, 144, 144)).toEqual([255, 255, 255]);
    const ink = data.filter((v, i) => i % 4 === 0 && v < 250).length;
    expect(ink).toBeGreaterThan(100);
    expect(ink).toBeLessThan(3000);
  });
});

import { describe, expect, it, vi } from 'vitest';
import { dropBitmapCacheByPath, type BlipEffect } from '@silurus/ooxml-core';
import type { Slide } from '@silurus/ooxml-pptx';
import { syntheticDocxModel } from '../../docx/src/testing/synthetic-document.js';
import { installImageBitmapShim, installOffscreenCanvasShim, type NodeCanvasFactory } from './render.js';
import { loadDocxRendererForTests, loadSkiaForTests, importForTests } from './test-imports.js';

const skia = await loadSkiaForTests();
const docx = await loadDocxRendererForTests();
const pptx = await importForTests(() => import('../../pptx/src/renderer.js'), 'PPTX renderer');
const stages = [
  'base fetch', 'base decode', 'dependency fetch', 'dependency decode',
  'base SVG', 'dependency SVG', 'surface allocation', 'readback draw', 'readback pixels',
  'nested fill draw', 'nested fill readback', 'pixel write', 'bitmap bake', 'nested bitmap bake', 'nested duotone readback', 'nested duotone bake',
] as const;
type Stage = typeof stages[number];

describe.skipIf(!skia || !docx || !pptx)('strict blip native failures at host picture boundaries', () => {
  it.each(['DOCX', 'PPTX'] as const)('renders the valid nested-fill control through %s', async format => {
    expect(await paint(format)).toEqual([255, 0, 0, 255]);
  });
  it.each((['DOCX', 'PPTX'] as const).flatMap(format => stages.map(stage => ({ format, stage }))))(
    '$format rejects $stage instead of successfully painting an empty picture', async ({ format, stage }) => {
      await paint(format, stage);
    });
});

async function paint(format: 'DOCX' | 'PPTX', stage?: Stage): Promise<number[]> {
  const source = new skia!.Canvas(4, 4);
  source.getContext('2d').fillStyle = '#FF0000';
  source.getContext('2d').fillRect(0, 0, 4, 4);
  const bytes = new Uint8Array(await source.toBuffer('png'));
  let injected = 0, enabled = true, surfaces = 0;
  const fail = (at: Stage) => {
    if (enabled && stage === at) {
      injected++;
      throw new DOMException(`Injected ${at}`, 'InvalidStateError');
    }
  };
  const factory: NodeCanvasFactory = {
    loadImage: data => skia!.loadImage(Buffer.from(data as Uint8Array)),
    createCanvas: (width, height) => {
      fail('surface allocation');
      const surface = new skia!.Canvas(width, height);
      const ctx = surface.getContext('2d');
      const index = ++surfaces;
      const draw = ctx.drawImage.bind(ctx), read = ctx.getImageData.bind(ctx), write = ctx.putImageData.bind(ctx);
      ctx.drawImage = ((...args: Parameters<typeof draw>) => {
        fail(index === 1 ? 'readback draw' : 'nested fill draw'); return draw(...args);
      }) as typeof ctx.drawImage;
      ctx.getImageData = ((...args: Parameters<typeof read>) => {
        if (index === 1) fail('nested duotone readback');
        fail(index === 1 ? 'readback pixels' : 'nested fill readback'); return read(...args);
      }) as typeof ctx.getImageData;
      ctx.putImageData = ((...args: Parameters<typeof write>) => { fail('pixel write'); return write(...args); }) as typeof ctx.putImageData;
      return surface as unknown as ReturnType<NodeCanvasFactory['createCanvas']>;
    },
  };
  const paths = new WeakMap<Blob, string>();
  const fetchImage = async (path: string, mime: string) => {
    fail(path === 'base.png' ? 'base fetch' : 'dependency fetch');
    const blob = new Blob([bytes], { type: mime }); paths.set(blob, path); return blob;
  };
  // The configured SVG decoder is the host's rasterization boundary. Its
  // rejection must follow the same strict-chain policy as native bitmap decode.
  const svgDecoder = async (blob: Blob) => {
    fail(paths.get(blob) === 'base.png' ? 'base SVG' : 'dependency SVG');
    return skia!.loadImage(Buffer.from(bytes)) as unknown as ImageBitmap;
  };
  const effects: BlipEffect[] = [{ type: 'fillOverlay', blend: 'over', fill: {
    fillType: 'image', imagePath: 'nested.png', mimeType: stage === 'dependency SVG' ? 'image/svg+xml' : 'image/png',
    stretch: true, duotone: stage?.startsWith('nested duotone') ? { clr1: '000000', clr2: 'FF0000' } : undefined,
    blipEffects: stage === 'nested bitmap bake' ? [{ type: 'grayscale' }] : undefined,
  } }];
  const image = { imagePath: 'base.png', mimeType: stage === 'base SVG' ? 'image/svg+xml' : 'image/png', blipEffects: effects };
  const canvas = new skia!.Canvas(100, 100);
  const render = async () => {
    surfaces = 0;
    if (format === 'PPTX') {
      const slide = { index: 0, slideNumber: 1, background: null, elements: [{ type: 'picture', ...image,
        x: 0, y: 0, width: 100 * 9525, height: 100 * 9525, rotation: 0, flipH: false, flipV: false,
      }] } as Slide;
      await pptx!.renderSlideWithEmbeddedFonts(canvas as unknown as OffscreenCanvas, slide, 100 * 9525, 100 * 9525,
        { width: 100, dpr: 1, fetchImage, svgDecoder });
    } else {
      const doc = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 });
      const paragraph = doc.body[0];
      if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
      paragraph.runs = [{ type: 'image', ...image, widthPt: 32, heightPt: 32 }];
      await docx!.renderDocumentToCanvas(doc, canvas as unknown as OffscreenCanvas, 0,
        { width: 100, dpr: 1, fetchImage, svgDecoder });
    }
  };
  const restoreBitmap = installImageBitmapShim(factory), restoreOffscreen = installOffscreenCanvasShim(factory);
  const decode = globalThis.createImageBitmap;
  vi.stubGlobal('createImageBitmap', async (input: ImageBitmapSource, ...args: unknown[]) => {
    if (input instanceof Blob) fail(paths.get(input) === 'base.png' ? 'base decode' : 'dependency decode');
    else { fail('bitmap bake'); fail('nested bitmap bake'); fail('nested duotone bake'); }
    return Reflect.apply(decode, globalThis, [input, ...args]);
  });
  try {
    if (stage) {
      await expect(render()).rejects.toMatchObject({ code: 'unsupported-image-effect' });
      expect(injected).toBeGreaterThan(0);
      // A rejected preparation/bake must evict its promise and release leases:
      // the same owner must be able to retry and paint the picture.
      enabled = false;
    }
    await render();
    const pixels = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
    let painted = 0;
    for (let i = 4; i < pixels.length; i += 4) if (pixels[i + 1] < pixels[painted + 1]) painted = i;
    expect(pixels[painted + 1]).toBeLessThan(250);
    return [...pixels.slice(painted, painted + 4)];
  } finally {
    dropBitmapCacheByPath(fetchImage);
    vi.unstubAllGlobals(); restoreOffscreen(); restoreBitmap();
  }
}

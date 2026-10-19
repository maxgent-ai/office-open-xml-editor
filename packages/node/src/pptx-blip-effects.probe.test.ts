import { describe, expect, it, vi } from 'vitest';
import type { BlipEffect } from '@silurus/ooxml-core';
import type { Presentation, Slide } from '@silurus/ooxml-pptx';
import { installImageBitmapShim, renderSlideNode, type NodeCanvasFactory } from './render.js';
import { loadSkiaForTests } from './test-imports.js';

const skia = await loadSkiaForTests();
const factory: NodeCanvasFactory = {
  createCanvas: (w, h) => new skia!.Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
  loadImage: (bytes) => skia!.loadImage(Buffer.from(bytes as Uint8Array)),
};

async function render(effects: BlipEffect[], site: 'picture' | 'background' | 'shape-fill', color = '#000000', inset = 0, sample = { x: 16, y: 16 }, rotation = 0, fillRotates?: boolean) {
  const source = new skia!.Canvas(2, 2);
  const ctx = source.getContext('2d');
  ctx.fillStyle = color; ctx.fillRect(0, 0, 2, 2);
  const png = await source.toBuffer('png');
  const frame = { x: inset * 9525, y: inset * 9525, width: (32 - 2 * inset) * 9525, height: (32 - 2 * inset) * 9525,
    rotation, flipH: false, flipV: false };
  const image = { imagePath: 'ppt/media/picture.png', mimeType: 'image/png', blipEffects: effects };
  const fill = { fillType: 'image' as const, ...image, stretch: true, rotWithShape: fillRotates };
  const slide = { index: 0, slideNumber: 1,
    background: site === 'background' ? fill : null,
    elements: site === 'background' ? [] : site === 'shape-fill'
      ? [{ type: 'shape', ...frame, geometry: 'rect', fill, textBody: null }]
      : [{ type: 'picture', ...frame, ...image, prstGeom: 'rect' }],
  } as Slide;
  const presentation: Presentation = { slideWidth: 32 * 9525, slideHeight: 32 * 9525,
    slides: [slide], defaultTextColor: null, majorFont: null, minorFont: null };
  const canvas = new skia!.Canvas(32, 32);
  await renderSlideNode(canvas as unknown as Parameters<typeof renderSlideNode>[0], presentation, 0, {
    width: 32, dpr: 1, factory,
    fetchImage: async () => new Blob([new Uint8Array(png)]),
  });
  return [...canvas.getContext('2d').getImageData(sample.x, sample.y, 1, 1).data];
}

describe.skipIf(!skia)('PPTX ordered blip paint', () => {
  // At 41e557377 the PPTX parser skipped alphaMod/blur/fillOverlay/hsl/tint/
  // alphaInv, and the duotone kernel rejected RGBA endpoints as a no-op. These
  // decks already rendered; correcting their pixels must not reject loading.
  it.each([
    [[{ type: 'alphaMod', container: { kind: 'tree', effects: [{ type: 'alphaInv' }] } }], '#FF0000', [255, 255, 255, 255]],
    [[{ type: 'fillOverlay', blend: 'over', fill: { fillType: 'solid', color: '00FF00' } }], '#FF0000', [0, 255, 0, 255]],
    [[{ type: 'hsl', hue: 120, sat: 0, lum: 0 }], '#FF0000', [0, 255, 0, 255]],
    [[{ type: 'tint', hue: 120, amt: 1 }], '#FF0000', [0, 255, 0, 255]],
    [[{ type: 'alphaInv', color: '00FF00' }], 'rgba(255,0,0,.5)', [128, 255, 128, 255]],
    [[{ type: 'duotone', clr1: 'FF000080', clr2: '0000FF80' }], '#000000', [255, 127, 127, 255]],
  ] satisfies [BlipEffect[], string, number[]][])('renders the previously ignored picture chain %j', async (effects, color, expected) => {
    expect(await render(effects, 'picture', color)).toEqual(expected);
  });
  it.each(['picture', 'shape-fill'] as const)('honors page-fixed recursive and preset shadows on a rotated %s', async site => {
    const shadow = { color: '000000', alpha: 1, blur: 0, dist: 8 * 9525, dir: 0 };
    const paint = (rotWithShape: boolean, preset = false) => render([{ type: 'container', container: { kind: 'tree', effects: [
      preset ? { type: 'presetShadow', shadows: [{ ...shadow, rotWithShape }] }
        : { type: 'shadow', inner: false, shadow: { ...shadow, rotWithShape } },
    ] } }], site, '#FF0000', 8, { x: 20, y: 12 }, 90);
    expect(await paint(true)).toEqual([255, 255, 255, 255]);
    expect(await paint(false)).toEqual([0, 0, 0, 255]);
    expect(await paint(true, true)).toEqual([0, 0, 0, 255]);
  });
  it('uses the orientation of a nonrotating image-fill brush for its recursive shadow', async () => {
    const shadow = { color: '000000', alpha: 1, blur: 0, dist: 8 * 9525, dir: 0, rotWithShape: false };
    expect(await render([{ type: 'container', container: { kind: 'tree', effects: [
      { type: 'shadow', inner: false, shadow },
    ] } }], 'shape-fill', '#FF0000', 8, { x: 20, y: 20 }, 90, false)).toEqual([0, 0, 0, 255]);
  });
  it('samples effects in a nested tiled image at the tile DPI rather than the outer picture size', async () => {
    const pngs = await Promise.all([32, 2].map(async size => {
      const source = new skia!.Canvas(size, size);
      source.getContext('2d').fillStyle = '#FF0000';
      source.getContext('2d').fillRect(0, 0, size, size);
      return new Uint8Array(await source.toBuffer('png'));
    }));
    const effects: BlipEffect[] = [{ type: 'fill', fill: { fillType: 'image', imagePath: 'tile.png', mimeType: 'image/png',
      dpi: 96, tile: {}, blipEffects: [{ type: 'blur', radius: 9525, grow: false }],
    } }];
    const presentation = { slideWidth: 32 * 9525, slideHeight: 32 * 9525,
      slides: [{ index: 0, slideNumber: 1, background: null, elements: [{ type: 'picture',
        x: 0, y: 0, width: 32 * 9525, height: 32 * 9525, rotation: 0, flipH: false, flipV: false,
        imagePath: 'base.png', mimeType: 'image/png', blipEffects: effects,
      }] }], defaultTextColor: null, majorFont: null, minorFont: null } as Presentation;
    const canvas = new skia!.Canvas(32, 32);
    await renderSlideNode(canvas as unknown as Parameters<typeof renderSlideNode>[0], presentation, 0, {
      width: 32, dpr: 1, factory, fetchImage: async path => new Blob([pngs[path === 'tile.png' ? 1 : 0]]),
    });
    // A 2x2 constant-red tile blurred by one source pixel retains alpha 249
    // with the documented separable kernel. Over white, its green/blue are 6.
    expect([...canvas.getContext('2d').getImageData(16, 16, 1, 1).data]).toEqual([255, 6, 6, 255]);
  });
  it('prepares recursively nested blip fills without holding decoder slots', async () => {
    const inner = { fillType: 'image' as const, imagePath: 'ppt/media/inner.png', mimeType: 'image/png', stretch: true,
      blipEffects: [{ type: 'colorReplace' as const, color: '0000FF' }] };
    const outer = { ...inner, imagePath: 'ppt/media/outer.png',
      blipEffects: [{ type: 'fillOverlay' as const, blend: 'over' as const, fill: inner }] };
    expect(await render([{ type: 'fillOverlay', blend: 'over', fill: outer }], 'picture', '#FF0000'))
      .toEqual([0, 0, 255, 255]);
  });
  it('paints gradient and pattern overlay fills through the shared fill painter', async () => {
    const gradient = await render([{ type: 'fillOverlay', blend: 'over', fill: {
      fillType: 'gradient', gradType: 'linear', angle: 0,
      stops: [{ position: 0, color: 'FF0000' }, { position: 1, color: '0000FF' }],
    } }], 'picture');
    expect(gradient[0]).toBeGreaterThan(0); expect(gradient[2]).toBeGreaterThan(0);
    expect(gradient[1]).toBe(0);
    const pattern = await render([{ type: 'fillOverlay', blend: 'over', fill: {
      fillType: 'pattern', preset: 'horz', fg: 'FF0000', bg: '0000FF',
    } }], 'picture');
    expect(pattern[0] + pattern[2]).toBeGreaterThan(0);
    expect(pattern[1]).toBe(0);
  });
  it('renders blur instead of rejecting a deck whose blur was previously ignored', async () => {
    const pixel = await render([{ type: 'blur', radius: 16 * 9525, grow: true }], 'picture', '#FF0000');
    expect(pixel[0]).toBe(255);
    expect(pixel[1]).toBeGreaterThan(0);
    expect(pixel[1]).toBeLessThan(255);
  });
  it('honors grow outside an explicit rectangular picture preset', async () => {
    const grown = await render([{ type: 'blur', radius: 8 * 9525, grow: true }], 'picture', '#FF0000', 8, { x: 4, y: 16 });
    const clipped = await render([{ type: 'blur', radius: 8 * 9525, grow: false }], 'picture', '#FF0000', 8, { x: 4, y: 16 });
    expect(grown[0]).toBe(255); expect(grown[1]).toBeLessThan(255);
    expect(clipped).toEqual([255, 255, 255, 255]);
  });
  it.each(['picture', 'background', 'shape-fill'] as const)('applies alpha-sensitive replacement exactly once on %s', async (site) => {
    expect(await render([
      { type: 'alphaModFix', amt: .5 },
      { type: 'colorChange', from: '000000', fromAlpha: 128 / 255,
        to: 'FF0000', toAlpha: 1, useAlpha: true },
    ], site)).toEqual([255, 0, 0, 255]);
  });
  it.each(['picture', 'background', 'shape-fill'] as const)('propagates malformed effect diagnostics on %s', async (site) => {
    await expect(render([
      { type: 'luminance', bright: .9, contrast: 0 },
      { type: 'unsupported', effect: 'malformed picture effect' },
    ], site)).rejects.toMatchObject({ code: 'unsupported-image-effect' });
  });
});


describe.skipIf(!skia)('PPTX physical blip identity and native failure propagation', () => {
  async function renderCopies(site: 'picture' | 'shape-fill', sizes: number[]) {
    const source = new skia!.Canvas(16, 16);
    const ctx = source.getContext('2d');
    ctx.fillStyle = '#ff0000'; ctx.fillRect(0, 0, 16, 16);
    const png = await source.toBuffer('png');
    const effects: BlipEffect[] = [{ type: 'blur', radius: 12700, grow: false }];
    const elements = sizes.map((size, index) => {
      const frame = { x: (index ? 20 : 0) * 12700, y: 0, width: size * 12700, height: size * 12700,
        rotation: 0, flipH: false, flipV: false };
      const image = { imagePath: 'ppt/media/shared.png', mimeType: 'image/png', blipEffects: effects };
      return site === 'picture' ? { type: 'picture', ...frame, ...image, prstGeom: 'rect' }
        : { type: 'shape', ...frame, geometry: 'rect', textBody: null,
          fill: { fillType: 'image', ...image, stretch: true } };
    });
    const presentation = { slideWidth: 110 * 12700, slideHeight: 90 * 12700,
      slides: [{ index: 0, slideNumber: 1, background: null, elements }],
      defaultTextColor: null, majorFont: null, minorFont: null } as Presentation;
    const canvas = new skia!.Canvas(110, 90);
    await renderSlideNode(canvas as unknown as Parameters<typeof renderSlideNode>[0], presentation, 0,
      { width: 110, dpr: 1, factory, fetchImage: async () => new Blob([new Uint8Array(png)]) });
    return [...canvas.getContext('2d').getImageData(0, 0, 8, 8).data];
  }
  it.each(['picture', 'shape-fill'] as const)('keeps an 8pt %s unchanged when an 80pt copy shares its PNG and 1pt blur', async site => {
    const alone = await renderCopies(site, [8]);
    expect(alone).toEqual(await renderCopies(site, [8, 80]));
    expect(alone.slice(0, 4)).not.toEqual([255, 0, 0, 255]);
  });
  it.each(['picture', 'shape-fill'] as const)('rejects native InvalidStateError during the %s effect bake', async site => {
    const restore = installImageBitmapShim(factory);
    const decode = globalThis.createImageBitmap;
    vi.stubGlobal('createImageBitmap', async (source: ImageBitmapSource, ...args: unknown[]) => {
      if (!(source instanceof Blob)) throw new DOMException('Injected native bake failure', 'InvalidStateError');
      return Reflect.apply(decode, globalThis, [source, ...args]);
    });
    try {
      await expect(renderCopies(site, [8])).rejects.toMatchObject({ code: 'unsupported-image-effect' });
    } finally { vi.unstubAllGlobals(); restore(); }
  });
});

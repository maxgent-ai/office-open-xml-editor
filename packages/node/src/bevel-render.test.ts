import { describe, it, expect } from 'vitest';
import { renderSlideNode } from './render';
import type { NodeCanvasFactory } from './render';
import type { Presentation, Slide, PictureElement } from '@silurus/ooxml-pptx';
import { loadSkiaForTests } from './test-imports';

// skia-canvas ships a native binding via a devDependency, so `pnpm install`
// provides it in CI as well as locally. Load it through the shared test helper:
// absent → skip cleanly (local), but under OOXML_REQUIRE_SKIA=1 (CI) a load
// failure becomes a hard error instead of a silent skip.
const skia = await loadSkiaForTests();
// Non-null aliases for use inside the (skia-gated) test bodies. When `skia` is
// null the whole suite is skipped via `describe.skipIf`, so these are never
// dereferenced; the cast keeps the helpers strongly typed without `as any`.
type Skia = typeof import('skia-canvas');
const { Canvas, loadImage } = (skia ?? {}) as Skia;

/**
 * Regression test for the Node OffscreenCanvas shim.
 *
 * `packages/core/src/shape/effects.ts`'s `createAuxCanvas` needs an
 * `OffscreenCanvas` (or DOM `document`) to allocate the auxiliary canvas that
 * the pptx renderer's `paintBeveledFlat` uses to bake sp3d bevel-lip shading.
 * Under Node neither exists, so without the shim the renderer silently skips
 * the bevel and blits the flat image. This test renders the SAME synthetic
 * slide (a solid picture carrying an sp3d `bevelT`) twice — once WITH the shim
 * (factory passed → bevel shading active) and once WITHOUT (no factory → flat
 * fallback) — and asserts the rim band differs. If the shim regresses, both
 * renders match and this fails.
 */

const factory: NodeCanvasFactory | null =
  Canvas && loadImage
    ? {
        createCanvas: (w, h) =>
          new Canvas(w, h) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
        loadImage: (buf) =>
          loadImage(buf as Buffer) as unknown as ReturnType<NodeCanvasFactory['loadImage']>,
      }
    : null;

const g = globalThis as unknown as { createImageBitmap?: unknown; OffscreenCanvas?: unknown };

/** Build a small flat-colour PNG as raw bytes using skia-canvas. The lazy
 *  image pipeline fetches bytes by path, so the test carries a path on the
 *  element and serves these bytes through `fetchImage`. */
async function flatPngBytes(w: number, h: number, color: string): Promise<Uint8Array> {
  const c = new Canvas(w, h);
  const ctx = c.getContext('2d');
  ctx.fillStyle = color;
  ctx.fillRect(0, 0, w, h);
  return new Uint8Array(await c.toBuffer('png'));
}

const EMU_PER_PX = 9525; // 96 dpi
const PIC_IMAGE_PATH = 'ppt/media/image1.png';

function buildBeveledSlide(): Presentation {
  // A 300x300 px picture centred on a 400x400 px slide, with a chunky top
  // bevel so the lit/shadowed rim is unmistakable.
  const px = (n: number) => Math.round(n * EMU_PER_PX);
  const pic: PictureElement = {
    type: 'picture',
    x: px(50),
    y: px(50),
    width: px(300),
    height: px(300),
    rotation: 0,
    flipH: false,
    flipV: false,
    imagePath: PIC_IMAGE_PATH,
    mimeType: 'image/png',
    stroke: null,
    sp3d: {
      prstMaterial: 'warmMatte',
      // 360000 EMU ≈ 38 px bevel at 1:1; circle lip → strong gradient.
      bevelT: { w: 360000, h: 360000, prst: 'circle' },
    },
  };
  const slide: Slide = {
    index: 0,
    slideNumber: 1,
    background: null,
    elements: [pic],
  };
  return {
    slideWidth: px(400),
    slideHeight: px(400),
    slides: [slide],
    defaultTextColor: null,
    majorFont: null,
    minorFont: null,
  };
}

async function renderToRgba(
  presentation: Presentation,
  pngBytes: Uint8Array,
  withShim: boolean,
): Promise<{ data: Uint8ClampedArray; width: number; height: number }> {
  const width = 400;
  const dpr = 1;
  const canvas = new Canvas(width * dpr, 400 * dpr);
  // createImageBitmap is always needed to decode the picture's bytes.
  const g2 = globalThis as unknown as { createImageBitmap?: unknown };
  const prevBitmap = g2.createImageBitmap;
  g2.createImageBitmap = async (source: Blob | ArrayBuffer | Uint8Array) => {
    let buf: ArrayBuffer | Uint8Array;
    if (source instanceof Uint8Array || source instanceof ArrayBuffer) buf = source;
    else buf = await (source as Blob).arrayBuffer();
    return loadImage(Buffer.from(buf as ArrayBuffer));
  };
  try {
    await renderSlideNode(
      canvas as unknown as Parameters<typeof renderSlideNode>[0],
      presentation,
      0,
      // `factory` is non-null whenever this suite runs (it is gated by
      // describe.skipIf(!skia)); the `&& factory` narrows the nullable type
      // away so the spread matches renderSlideNode's `factory?` option.
      {
        width,
        dpr,
        // Serve the picture's bytes by path (lazy image pipeline).
        fetchImage: async (_path: string, mime: string) =>
          new Blob([pngBytes as BlobPart], { type: mime }),
        ...(withShim && factory ? { factory } : {}),
      },
    );
  } finally {
    g2.createImageBitmap = prevBitmap as typeof globalThis.createImageBitmap;
  }
  const ctx = canvas.getContext('2d');
  const img = ctx.getImageData(0, 0, width * dpr, 400 * dpr);
  return { data: img.data as unknown as Uint8ClampedArray, width: width * dpr, height: 400 * dpr };
}

describe.skipIf(!skia)('node bevel rendering (OffscreenCanvas shim)', () => {
  it('shim activates sp3d bevel shading — rim pixels differ from the flat fallback', async () => {
    // Guard: no ambient OffscreenCanvas in the test runtime, so the only way the
    // bevel can appear is via the shim installed by renderSlideNode(factory).
    expect(typeof g.OffscreenCanvas).toBe('undefined');

    const pngBytes = await flatPngBytes(8, 8, '#808080'); // mid-grey face
    const presentation = buildBeveledSlide();

    const flat = await renderToRgba(presentation, pngBytes, /* withShim */ false);
    const beveled = await renderToRgba(presentation, pngBytes, /* withShim */ true);

    // Shim must not leak out of renderSlideNode.
    expect(typeof g.OffscreenCanvas).toBe('undefined');

    // The picture occupies px (50,50)..(350,350) at dpr=1. The top bevel lip
    // spans the rows from the top edge (y=50) inward by the bevel width,
    // bevelT w = 360000 EMU / 9525 EMU per px ≈ 37.8 px (ECMA-376 §20.1.5.3).
    // Where in that band the swing peaks depends on the preset's profile and
    // the light. The convex circle is near-vertical at the silhouette, where
    // its diffuse term under the top light is close to the face's; the swing
    // peaks mid-band, where the normal leans most toward the light. So scan
    // every lip row (excluding the antialiased edge row) and take the
    // strongest. Columns stay clear of the side lips.
    const { data: fd, width } = flat;
    const { data: bd } = beveled;
    const top = 50;
    const bandPx = 360000 / EMU_PER_PX;
    let diffCount = 0;
    let maxDelta = 0;
    for (let y = top + 1; y < top + Math.floor(bandPx); y++) {
      let rowCount = 0;
      let rowMax = 0;
      for (let x = top + Math.ceil(bandPx); x < 350 - Math.ceil(bandPx); x++) {
        const i = (y * width + x) * 4;
        const delta = Math.max(
          Math.abs(fd[i] - bd[i]),
          Math.abs(fd[i + 1] - bd[i + 1]),
          Math.abs(fd[i + 2] - bd[i + 2]),
        );
        if (delta > 8) rowCount++;
        if (delta > rowMax) rowMax = delta;
      }
      diffCount = Math.max(diffCount, rowCount);
      maxDelta = Math.max(maxDelta, rowMax);
    }

    // Without the shim the flat render fills this band with the flat grey; with
    // the shim the bevel lip lightens/darkens it. Expect some lip row with a
    // substantial run of differing pixels, and some lip row with a clear
    // luminance swing (each maximized over the band independently).
    expect(diffCount).toBeGreaterThan(40);
    expect(maxDelta).toBeGreaterThan(20);
  });

  it('without a factory the render is the flat fallback (no aux canvas, no throw)', async () => {
    const pngBytes = await flatPngBytes(8, 8, '#808080');
    const presentation = buildBeveledSlide();
    // Two flat renders must be identical (deterministic, no bevel).
    const a = await renderToRgba(presentation, pngBytes, false);
    const b = await renderToRgba(presentation, pngBytes, false);
    let diff = 0;
    for (let i = 0; i < a.data.length; i++) if (a.data[i] !== b.data[i]) diff++;
    expect(diff).toBe(0);
  });
});

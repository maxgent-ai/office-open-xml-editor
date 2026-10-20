import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  applyOuterShadow,
  applySoftEdge,
  applyReflection,
  createAuxCanvas,
  createRasterCoveragePaint,
  rasterEffectCoverageBounds,
} from './effects';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';
import type { Shadow, SoftEdge, Reflection } from '../types/common';

/**
 * Minimal recording 2D context. Records the ordered sequence of state mutations
 * and draw calls so we can assert the compositing pipeline each effect builds,
 * without a real canvas. Effects are pure compositing logic, so the op sequence
 * is the contract under test (ECMA-376 §20.1.8.40 / .50 / .53).
 */
class RecordingCtx {
  ops: Array<{ op: string; args?: unknown[] }> = [];
  fillStyle: unknown = '#000';
  // Every composite-op / filter assignment is recorded, not just the last value,
  // so a multi-pass pipeline (e.g. set destination-in + blur, then reset) can be
  // asserted by history rather than the final field value.
  usedCompositeOps: string[] = [];
  usedFilters: string[] = [];

  #filter = 'none';
  #gco = 'source-over';

  get filter() { return this.#filter; }
  set filter(v: string) { this.#filter = v; this.usedFilters.push(v); }
  get globalCompositeOperation() { return this.#gco; }
  set globalCompositeOperation(v: string) { this.#gco = v; this.usedCompositeOps.push(v); }

  /** Filter strings that were a blur(...) (i.e. excluding the 'none' resets). */
  get usedBlurFilters(): string[] {
    return this.usedFilters.filter(f => f.startsWith('blur('));
  }

  save() { this.ops.push({ op: 'save' }); }
  restore() { this.ops.push({ op: 'restore' }); }
  translate(x: number, y: number) { this.ops.push({ op: 'translate', args: [x, y] }); }
  scale(x: number, y: number) { this.ops.push({ op: 'scale', args: [x, y] }); }
  rotate(angle: number) { this.ops.push({ op: 'rotate', args: [angle] }); }
  transform(...a: number[]) { this.ops.push({ op: 'transform', args: a }); }
  setTransform(...a: number[]) { this.ops.push({ op: 'setTransform', args: a }); }
  fillRect(...a: number[]) { this.ops.push({ op: 'fillRect', args: a }); }
  fill(...a: unknown[]) { this.ops.push({ op: 'fill', args: a }); }
  clip(...a: unknown[]) { this.ops.push({ op: 'clip', args: a }); }
  beginPath() { this.ops.push({ op: 'beginPath' }); }
  rect(...a: number[]) { this.ops.push({ op: 'rect', args: a }); }
  drawImage(...a: unknown[]) { this.ops.push({ op: 'drawImage', args: a }); }
  createLinearGradient() {
    this.ops.push({ op: 'createLinearGradient' });
    const stops: Array<[number, string]> = [];
    return {
      stops,
      addColorStop: (o: number, c: string) => { stops.push([o, c]); },
    } as unknown as CanvasGradient;
  }
}

/** Install a fake OffscreenCanvas that hands back RecordingCtx instances. */
function installFakeCanvas(): { auxCtxs: RecordingCtx[] } {
  const auxCtxs: RecordingCtx[] = [];
  class FakeOffscreen {
    width: number;
    height: number;
    constructor(w: number, h: number) { this.width = w; this.height = h; }
    getContext(_kind: string) {
      const c = new RecordingCtx();
      auxCtxs.push(c);
      return c;
    }
  }
  vi.stubGlobal('OffscreenCanvas', FakeOffscreen);
  return { auxCtxs };
}

const BBOX = { x: 100, y: 50, w: 200, h: 80 };
const DEVICE_W = 800;
const DEVICE_H = 600;
// scale is px-per-EMU; 1 px == 9525 EMU in DrawingML, but the effect helper just
// multiplies, so pick a scale where the arithmetic is easy to read.
const SCALE = 1 / 9525; // 9525 EMU -> 1 px

describe('raster effect coverage allocation', () => {
  afterEach(() => { vi.unstubAllGlobals(); vi.restoreAllMocks(); });

  it.each([
    { width: 4096, height: 4096, bbox: { x: 10, y: 10, w: 4076, h: 4076 } },
    { width: 32767, height: 1, bbox: { x: 10, y: 0, w: 32747, h: 1 } },
  ])('keeps coverage available at the supported $width × $height boundary', ({ width, height, bbox }) => {
    const allocations: number[][] = [];
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(w: number, h: number) { allocations.push([w, h]); }
      getContext() {
        // Exercise the allocation guard without allocating a maximum-size
        // backing store or readback buffer. Pixel semantics are covered by
        // the real-canvas effect suites.
        return Object.assign(new RecordingCtx(), {
          getImageData: () => ({ data: new Uint8ClampedArray(0) }),
          putImageData: () => {},
        });
      }
    });
    const bounds = rasterEffectCoverageBounds(bbox, {
      softEdge: { radius: 8 },
      innerShadow: { color: '000000', alpha: 0.5, blur: 2, dist: 2, dir: 0 },
    }, 1, width, height);

    const coverage = createRasterCoveragePaint(() => {}, bounds);

    expect(coverage).toBeTypeOf('function');
    expect(allocations).toEqual([[width, height]]);
    const live = new RecordingCtx();
    coverage?.(live as never);
    expect(live.ops.find(op => op.op === 'drawImage')?.args?.slice(1)).toEqual([0, 0]);
  });

  it('rejects an oversized source before allocation and warns once about omitted effects', () => {
    const allocate = vi.fn();
    vi.stubGlobal('OffscreenCanvas', allocate);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const paint = vi.fn();
    const bbox = { x: 0, y: 0, w: 4097, h: 4096 };

    expect(createRasterCoveragePaint(paint, bbox)).toBeNull();
    expect(createRasterCoveragePaint(paint, bbox)).toBeNull();

    expect(allocate).not.toHaveBeenCalled();
    expect(paint).not.toHaveBeenCalled();
    expect(warn).toHaveBeenCalledTimes(1);
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('soft edge and inner shadow omitted'));
  });
});

describe('createAuxCanvas', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('returns an OffscreenCanvas when available, clamped to >=1 px', () => {
    installFakeCanvas();
    const c = createAuxCanvas(0, 0) as { width: number; height: number };
    expect(c).not.toBeNull();
    expect(c.width).toBe(1);
    expect(c.height).toBe(1);
  });

  it('returns null in a headless environment', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    expect(createAuxCanvas(10, 10)).toBeNull();
  });

  it('uses the live context canvas constructor last in a headless skia-like environment', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    class SkiaLikeCanvas {
      constructor(public width: number, public height: number) {}
    }
    const ctx = { canvas: new SkiaLikeCanvas(20, 20) } as unknown as CanvasRenderingContext2D;

    const aux = createAuxCanvasForContext(ctx, 10.1, 0) as unknown as SkiaLikeCanvas;

    expect(aux).toBeInstanceOf(SkiaLikeCanvas);
    expect([aux.width, aux.height]).toEqual([11, 1]);
  });

  it('returns null when every ctx-aware allocation strategy throws', () => {
    vi.stubGlobal('OffscreenCanvas', class { constructor() { throw new Error('offscreen'); } });
    vi.stubGlobal('document', { createElement: () => { throw new Error('dom'); } });
    class BrokenCanvas { constructor() { throw new Error('constructor'); } }
    const ctx = { canvas: Object.create(BrokenCanvas.prototype) } as unknown as CanvasRenderingContext2D;

    expect(createAuxCanvasForContext(ctx, 10, 10)).toBeNull();
  });
});

describe('applyOuterShadow (ECMA-376 §20.1.8.45)', () => {
  let fake: { auxCtxs: RecordingCtx[] };
  beforeEach(() => { fake = installFakeCanvas(); });
  afterEach(() => vi.unstubAllGlobals());

  it('resets the crop transform before recolouring, then blurs/offsets once', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => {
      const target = c as RecordingCtx;
      target.setTransform({ a: 2, b: 0, c: 0, d: 2, e: 400, f: 200 } as never);
      target.ops.push({ op: 'paintShape' });
    });
    const shadow: Shadow = {
      color: '000000', alpha: 0.38, blur: 38_100, dist: 19_050, dir: 90,
    };

    expect(applyOuterShadow(
      live as never, paint as never, BBOX, shadow, SCALE, DEVICE_W, DEVICE_H,
    )).toBe(true);

    expect(paint).toHaveBeenCalledTimes(1);
    expect(fake.auxCtxs).toHaveLength(1);
    expect(fake.auxCtxs[0].usedCompositeOps).toContain('source-in');
    const recolour = fake.auxCtxs[0].ops.findIndex(o => o.op === 'fillRect');
    expect(recolour).toBeGreaterThan(0);
    expect(fake.auxCtxs[0].ops[recolour - 1]).toMatchObject({
      op: 'setTransform',
      args: [1, 0, 0, 1, 0, 0],
    });
    expect(live.usedBlurFilters).toEqual(['blur(4px)']);
    const finalBlit = live.ops.filter(o => o.op === 'drawImage');
    expect(finalBlit).toHaveLength(1);
    expect(finalBlit[0].args?.slice(1)).toEqual([84, 34]);
    expect(live.ops.filter(o => o.op === 'translate')[0].args?.[1]).toBeCloseTo(2, 6);
  });

  it('uses one auxiliary surface and no filter for a sharp shadow', () => {
    const live = new RecordingCtx();
    const shadow: Shadow = {
      color: '000000', alpha: 0.38, blur: 0, dist: 19_050, dir: 90,
    };

    expect(applyOuterShadow(
      live as never, (() => {}) as never, BBOX, shadow, SCALE, DEVICE_W, DEVICE_H,
    )).toBe(true);

    expect(fake.auxCtxs).toHaveLength(1);
    expect(fake.auxCtxs[0].usedBlurFilters).toEqual([]);
    expect(live.usedBlurFilters).toEqual([]);
    expect(live.ops.filter(o => o.op === 'drawImage')).toHaveLength(1);
  });

  it('applies alignment, scale, skew and shape-relative rotation before the blit', () => {
    const live = new RecordingCtx();
    const shadow: Shadow = {
      color: '000000', alpha: 0.5, blur: 0, dist: 9525, dir: 0,
      sx: 0.5, sy: 1.5, kx: 10, ky: -5, algn: 'tr', rotWithShape: true,
    };

    expect(applyOuterShadow(
      live as never, (() => {}) as never, BBOX, shadow, SCALE, DEVICE_W, DEVICE_H, 90,
      [75, 75],
    )).toBe(true);

    const transforms = live.ops.filter(o =>
      o.op === 'translate' || o.op === 'rotate' || o.op === 'transform');
    expect(transforms[0].args?.[0]).toBeCloseTo(0, 6);
    expect(transforms[0].args?.[1]).toBeCloseTo(1, 6);
    expect(transforms[1]).toEqual({ op: 'translate', args: [75, 75] });
    expect(transforms[2].args?.[0]).toBeCloseTo(Math.PI / 2, 9);
    expect(transforms[3].args).toEqual([
      0.5,
      Math.tan(-5 * Math.PI / 180),
      Math.tan(10 * Math.PI / 180),
      1.5,
      0,
      0,
    ]);
    expect(transforms[4].args?.[0]).toBeCloseTo(-Math.PI / 2, 9);
    expect(transforms[5]).toEqual({ op: 'translate', args: [-75, -75] });
  });

  it('keeps direction and transform in page space when rotWithShape is false', () => {
    const live = new RecordingCtx();
    const shadow: Shadow = {
      color: '000000', alpha: 0.5, blur: 0, dist: 9525, dir: 0,
      rotWithShape: false,
    };

    expect(applyOuterShadow(
      live as never, (() => {}) as never, BBOX, shadow, SCALE, DEVICE_W, DEVICE_H, 90,
    )).toBe(true);

    const transforms = live.ops.filter(o =>
      o.op === 'translate' || o.op === 'rotate' || o.op === 'transform');
    expect(transforms[0]).toEqual({ op: 'translate', args: [1, 0] });
    expect(transforms.some(o => o.op === 'rotate')).toBe(false);
  });
});

describe('applySoftEdge (ECMA-376 §20.1.8.53)', () => {
  let fake: { auxCtxs: RecordingCtx[] };
  beforeEach(() => { fake = installFakeCanvas(); });
  afterEach(() => vi.unstubAllGlobals());

  it('paints onto the live context directly when radius is 0 (no feather)', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    const se: SoftEdge = { radius: 0 };
    applySoftEdge(live as never, paint as never, BBOX, se, SCALE, DEVICE_W, DEVICE_H);
    // No aux canvas allocated; shape painted straight onto live.
    expect(fake.auxCtxs).toHaveLength(0);
    expect(live.ops.some(o => o.op === 'paintShape')).toBe(true);
    expect(live.ops.some(o => o.op === 'drawImage')).toBe(false);
  });

  it('builds an edge-clamp colour layer, replaces alpha with a blurred silhouette, then blits once', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    const mask = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintMask' }); });
    const se: SoftEdge = { radius: 28575 }; // 3 px (radius/3 = 1px blur)
    applySoftEdge(live as never, paint as never, BBOX, se, SCALE, DEVICE_W, DEVICE_H, mask as never);

    // THREE aux canvases: the sharp image, the silhouette mask, and the compose
    // layer. PowerPoint's soft edge feathers symmetrically (outward + inward), so
    // we cannot just mask the hard-clipped image (that feathers inward only and
    // leaves a hard outer step). Build an opaque colour layer that extends past
    // the geometry, then replace its alpha with the blurred silhouette.
    expect(fake.auxCtxs).toHaveLength(3);
    const [imageAux, maskAux, composeAux] = fake.auxCtxs;

    // Sharp image: one full-shape paint. Mask layer: one silhouette paint.
    expect(imageAux.ops.filter(o => o.op === 'paintShape')).toHaveLength(1);
    expect(maskAux.ops.filter(o => o.op === 'paintMask')).toHaveLength(1);
    expect(maskAux.ops.some(o => o.op === 'paintShape')).toBe(false);

    // Compose layer: the edge-clamp STRETCH is a 9-arg drawImage (src bbox →
    // dst bbox expanded by radius on every side), followed by the sharp image
    // drawn 1:1 (3-arg) to keep the interior crisp.
    const composeDraws = composeAux.ops.filter(o => o.op === 'drawImage');
    expect(composeDraws.length).toBeGreaterThanOrEqual(3); // 2 colour-layer draws + mask draw
    const stretch = composeDraws[0];
    expect(stretch.args).toHaveLength(9);
    // The aux canvases are cropped to bbox+margin (perf: A4), so coordinates are
    // in the CROPPED canvas's LOCAL space. radius=3px → margin=ceil(3)+2=5, so
    // crop origin = (bbox.x−5, bbox.y−5) = (95, 45); bbox in local space is
    // (bx,by)=(5,5). The stretch dst is the local bbox grown by `radius` on all
    // sides = (bx−3, by−3, w+6, h+6) = (2, 2, 206, 86).
    const CROP_MARGIN = 5;            // ceil(radius)+2 = ceil(3)+2
    // When the crop is unclamped, crop.x = bbox.x − margin, so bbox in local
    // space sits exactly `margin` in from the crop edge: bx = by = CROP_MARGIN.
    const bx = CROP_MARGIN;
    const by = CROP_MARGIN;
    expect(stretch.args?.[5]).toBeCloseTo(bx - 3, 6);       // dx (local)
    expect(stretch.args?.[6]).toBeCloseTo(by - 3, 6);       // dy (local)
    expect(stretch.args?.[7]).toBeCloseTo(BBOX.w + 6, 6);   // dWidth (unchanged)
    expect(stretch.args?.[8]).toBeCloseTo(BBOX.h + 6, 6);   // dHeight (unchanged)
    const sharp = composeDraws[1];
    expect(sharp.args).toHaveLength(3);
    // The final live blit is at the crop origin (95, 45), not (0, 0).
    const liveBlit = live.ops.find(o => o.op === 'drawImage');
    expect(liveBlit?.args?.[1]).toBeCloseTo(BBOX.x - CROP_MARGIN, 6);
    expect(liveBlit?.args?.[2]).toBeCloseTo(BBOX.y - CROP_MARGIN, 6);

    // Then the alpha is replaced by the blurred silhouette: a destination-in
    // pass and a blur(...) filter were used on the compose layer.
    expect(composeAux.usedCompositeOps).toContain('destination-in');
    expect(composeAux.usedBlurFilters.length).toBeGreaterThan(0);
    // Soft-edge `rad` spans ~3σ, so the blur std-dev is radius/3 = 1px.
    expect(composeAux.usedBlurFilters[0]).toBe('blur(1px)');

    // The compose layer is blitted onto live exactly once; the image/mask auxes
    // are not blitted directly.
    expect(live.ops.filter(o => o.op === 'drawImage')).toHaveLength(1);
  });
});

describe('applyReflection (ECMA-376 §20.1.8.50)', () => {
  let fake: { auxCtxs: RecordingCtx[] };
  beforeEach(() => { fake = installFakeCanvas(); });
  afterEach(() => vi.unstubAllGlobals());

  const reflection: Reflection = {
    blur: 0, dist: 0, dir: 90,
    stA: 0.5, stPos: 0, endA: 0, endPos: 0.35,
    sx: 1, sy: -1,
  };

  it('paints shape, builds a vertical alpha gradient, mirrors via sy<0, blits', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    applyReflection(live as never, paint as never, BBOX, reflection, SCALE, DEVICE_W, DEVICE_H);

    const aux = fake.auxCtxs[0];
    expect(aux.ops.some(o => o.op === 'createLinearGradient')).toBe(true);
    expect(aux.ops.some(o => o.op === 'fillRect')).toBe(true);
    // The live blit mirrors with scale(sx, sy) where sy<0.
    const scaleOp = live.ops.find(o => o.op === 'scale');
    expect(scaleOp?.args).toEqual([1, -1]);
    expect(live.ops.filter(o => o.op === 'drawImage')).toHaveLength(1);
  });

  it('anchors the mirror at the shape bottom edge (algn="b") with dist offset', () => {
    const live = new RecordingCtx();
    const paint = vi.fn();
    const withDist: Reflection = { ...reflection, dist: 9525, dir: 90 }; // 1px down
    applyReflection(live as never, paint as never, BBOX, withDist, SCALE, DEVICE_W, DEVICE_H);
    const translates = live.ops.filter(o => o.op === 'translate');
    // First translate moves origin to (bbox.x + offX, bottom + offY).
    const bottom = BBOX.y + BBOX.h; // 130
    expect(translates[0].args?.[0]).toBeCloseTo(BBOX.x + 0, 6);
    expect(translates[0].args?.[1]).toBeCloseTo(bottom + 1, 6);
  });
});

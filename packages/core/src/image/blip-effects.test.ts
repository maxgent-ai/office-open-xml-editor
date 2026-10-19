import { ImageEffectWorkLedger } from './image-effect-work.js';
import { describe, it, expect, vi, afterEach } from 'vitest';
import {
  applyBlipPixelEffects,
  blipEffectsMayExtendCoverage,
  blipEffectsCoverageReach,
  blipPixelEffectsKey,
  blipPixelEffectWork,
  assertBlipPixelEffectsBudget,
  blipGrayLevel,
  blipLuminance,
  type BlipEffect,
  type BlipPixelEffects,
} from './blip-effects';
import {
  MAX_IMAGE_EFFECT_BASE_PIXELS,
  MAX_RASTER_PIXELS,
  MAX_IMAGE_EFFECT_PASSES,
  MAX_IMAGE_EFFECT_PIXEL_WORK,
  isOoxmlDecodedImageLimitError,
} from './pixel-budget';
import {
  getCachedDuotoneBitmapByPath,
  duotoneCacheKey,
  dropDuotoneBitmapCache,
} from './duotone-bitmap-by-path';
import { dropBitmapCacheByPath } from './bitmap-image-by-path';
import type { OffscreenFactory } from './duotone';

const buffer = (...pixels: number[][]) => ({
  data: new Uint8ClampedArray(pixels.flat()),
  width: pixels.length,
  height: 1,
});

describe('applyBlipPixelEffects (ECMA-376 §20.1.8.11/16/34)', () => {
  // §20.1.8.39 adjusts H by degrees and S/L by relative percentages.
  it('adjusts HSL in the shared DrawingML color space, preserving alpha and order', () => {
    const buf = buffer([255, 0, 0, 128]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'hsl', hue: 120, sat: 0, lum: -.5 }] });
    expect([...buf.data]).toEqual([0, 128, 0, 128]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'hsl', hue: 0, sat: -1, lum: 0 }] });
    expect([...buf.data]).toEqual([64, 64, 64, 128]);
  });

  it('tints towards and away from the authored hue with bounded RGB interpolation', () => {
    const buf = buffer([0, 0, 255, 29]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'tint', hue: 0, amt: .5 }] });
    expect([...buf.data]).toEqual([128, 0, 128, 29]);
    const away = buffer([128, 64, 64, 255]);
    applyBlipPixelEffects(away, { effects: [{ type: 'tint', hue: 0, amt: -.5 }] });
    expect([...away.data]).toEqual([65, 96, 96, 255]);
  });

  it('colored alphaInv replaces RGB and inverts source alpha before a following effect', () => {
    const buf = buffer([0, 0, 0, 64]);
    applyBlipPixelEffects(buf, { effects: [
      { type: 'alphaInv', color: 'FF0000' }, { type: 'alphaModFix', amt: .5 },
    ] });
    expect([...buf.data]).toEqual([255, 0, 0, 96]);
  });

  it('interpolates duotone endpoint opacity and multiplies it by the source alpha', () => {
    const buf = buffer([0, 0, 0, 128], [255, 255, 255, 128], [128, 128, 128, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'duotone', clr1: 'FF000000', clr2: '00FF0080' }] });
    expect([...buf.data]).toEqual([255, 0, 0, 0, 0, 255, 0, 64, 127, 128, 0, 64]);
  });

  it('alphaMod recursively evaluates ordered containers on the current alpha', () => {
    const buf = buffer([12, 34, 56, 128]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'alphaMod', container: {
      kind: 'tree', effects: [ { type: 'alphaRepl', alpha: .5 },
        { type: 'alphaMod', container: { kind: 'tree', effects: [{ type: 'alphaInv' }] } } ],
    } }] });
    expect([...buf.data]).toEqual([12, 34, 56, 32]);
  });

  it('applies sibling masks independently instead of chaining them', () => {
    const buf = buffer([1, 2, 3, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'alphaMod', container: {
      kind: 'sib', effects: [{ type: 'alphaModFix', amt: .5 }, { type: 'alphaModFix', amt: .5 }],
    } }] });
    expect(buf.data[3]).toBe(192); // 50% over 50%, not 25%
  });

  it.each<BlipEffect>([
    { type: 'alphaModFix', amt: 1 }, { type: 'luminance', bright: 0, contrast: 0 },
  ])('an identity sibling preserves opaque coverage beside alphaInv: %j', identity => {
    const buf = buffer([1, 2, 3, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'alphaMod', container: {
      kind: 'sib', effects: [identity, { type: 'alphaInv' }],
    } }] });
    expect([...buf.data]).toEqual([1, 2, 3, 255]);
  });

  it.each([
    ['darken', [64, 64, 64, 255]], ['lighten', [128, 128, 128, 255]],
    ['mult', [32, 32, 32, 255]], ['over', [128, 64, 128, 255]],
    ['screen', [160, 160, 160, 255]],
  ] as const)('blends fillOverlay using %s before the next alpha operation', (blend, expected) => {
    const buf = buffer([64, 128, 64, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'fillOverlay', fill: { fillType: 'solid', color: '804080' }, blend }] });
    expect([...buf.data]).toEqual(expected);
  });

  it('blur is a separable premultiplied convolution with clipped or grown bounds', () => {
    const source = () => buffer([255, 0, 0, 255]);
    const grown = applyBlipPixelEffects(source(), { effects: [{ type: 'blur', radius: 9525, grow: true }] });
    expect([grown.width, grown.height, grown.frame]).toEqual([3, 3, { x: 1, y: 1, width: 1, height: 1 }]);
    expect([...grown.data.slice(16, 20)]).toEqual([255, 0, 0, 244]);
    const clipped = applyBlipPixelEffects(source(), { effects: [{ type: 'blur', radius: 9525, grow: false }] });
    expect([clipped.width, clipped.height, ...clipped.data]).toEqual([1, 1, 255, 0, 0, 244]);
  });

  it('zero-radius soft edges preserve pixels exactly, including inside alphaMod', () => {
    const source = buffer([12, 34, 56, 128], [78, 90, 12, 0]);
    expect(applyBlipPixelEffects(source, { effects: [{ type: 'softEdge', radius: 0 }] }).data)
      .toEqual(new Uint8ClampedArray([12, 34, 56, 128, 78, 90, 12, 0]));
    const nested = applyBlipPixelEffects(buffer([12, 34, 56, 128]), { effects: [
      { type: 'alphaMod', container: { kind: 'tree', effects: [{ type: 'softEdge', radius: 0 }] } },
    ] });
    expect([...nested.data]).toEqual([12, 34, 56, 64]);
  });

  it('feathers coverage once independently of source opacity and retains fill colors', () => {
    const render = (alpha: number) => applyBlipPixelEffects(buffer([12, 34, 56, alpha]),
      { effects: [{ type: 'softEdge', radius: 9525 }] });
    const opaque = render(255), translucent = render(128);
    expect(opaque.data[3]).toBe(244);
    expect([...translucent.data]).toEqual([12, 34, 56, Math.round(128 * 244 / 255)]);
    expect(render(0).data[3]).toBe(0);
  });

  it('preflights blur and nested effect work before touching source pixels', () => {
    const buf = buffer([10, 20, 30, 40]);
    expect(() => applyBlipPixelEffects(buf, { effects: [
      { type: 'alphaRepl', alpha: 1 }, { type: 'blur', radius: 1e15, grow: true },
    ] })).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-work' }));
    expect([...buf.data]).toEqual([10, 20, 30, 40]);
  });

  it('orders HSL and tint relative to grayscale', () => {
    const a = buffer([255, 0, 0, 255]), b = buffer([255, 0, 0, 255]);
    const hsl: BlipEffect = { type: 'hsl', hue: 120, sat: 0, lum: 0 };
    applyBlipPixelEffects(a, { effects: [hsl, { type: 'grayscale' }] });
    applyBlipPixelEffects(b, { effects: [{ type: 'grayscale' }, hsl] });
    expect([...a.data]).toEqual([182, 182, 182, 255]);
    expect([...b.data]).toEqual([54, 54, 54, 255]);
    const c = buffer([0, 0, 255, 255]), d = buffer([0, 0, 255, 255]);
    const tint: BlipEffect = { type: 'tint', hue: 0, amt: .5 };
    applyBlipPixelEffects(c, { effects: [tint, { type: 'grayscale' }] });
    applyBlipPixelEffects(d, { effects: [{ type: 'grayscale' }, tint] });
    expect([...c.data]).toEqual([36, 36, 36, 255]);
    expect([...d.data]).toEqual([137, 9, 9, 255]);
  });
  it('orders blur, duotone and recursive masks before alpha replacement', () => {
    for (const effect of [
      { type: 'blur', radius: 9525, grow: true },
      { type: 'duotone', clr1: 'FF000080', clr2: '00FF0080' },
      { type: 'alphaMod', container: { kind: 'tree', effects: [{ type: 'alphaRepl', alpha: .5 }] } },
    ] satisfies BlipEffect[]) {
      const a = applyBlipPixelEffects(buffer([255, 0, 0, 64]), { effects: [effect, { type: 'alphaRepl', alpha: 1 }] });
      const b = applyBlipPixelEffects(buffer([255, 0, 0, 64]), { effects: [{ type: 'alphaRepl', alpha: 1 }, effect] });
      expect(a.data[3]).toBe(255);
      expect(b.data[3]).toBeLessThan(255);
    }
    const buf = buffer([0, 0, 0, 255]);
    applyBlipPixelEffects(buf, { effects: [
      { type: 'fillOverlay', blend: 'over', fill: { fillType: 'solid', color: 'FF0000' } },
      { type: 'colorChange', from: 'FF0000', fromAlpha: 1, to: '00FF00', toAlpha: 1, useAlpha: true },
    ] });
    expect([...buf.data]).toEqual([0, 255, 0, 255]);
  });
  it('resolves named masks and the original fill reference at their authored positions', () => {
    const buf = buffer([10, 20, 30, 255]);
    applyBlipPixelEffects(buf, { effects: [
      { type: 'alphaModFix', amt: .5 },
      { type: 'alphaMod', container: { kind: 'tree', effects: [
        { type: 'container', container: { name: 'mask', kind: 'tree', effects: [{ type: 'reference', reference: 'fill' }] } },
        { type: 'reference', reference: 'mask' },
      ] } },
    ] });
    expect([...buf.data]).toEqual([10, 20, 30, 128]);
  });
  it('charges recursive scratch and Gaussian kernels to the four-surface memory ceiling', () => {
    const nested: BlipEffect = { type: 'alphaMod', container: { kind: 'tree', effects: [
      { type: 'alphaMod', container: { kind: 'tree', effects: [
        { type: 'alphaMod', container: { kind: 'tree', effects: [{ type: 'alphaRepl', alpha: .5 }] } },
      ] } },
    ] } };
    expect(() => assertBlipPixelEffectsBudget({ effects: [nested] }, MAX_IMAGE_EFFECT_BASE_PIXELS,
      { width: 4096, height: 2048 })).toThrowError(expect.objectContaining({ metric: 'active-decoded-bytes' }));
    const radiusPixels = Math.floor(MAX_RASTER_PIXELS / Float64Array.BYTES_PER_ELEMENT) + 1;
    expect(() => applyBlipPixelEffects(buffer([1, 2, 3, 255]), { effects: [
      { type: 'blur', radius: radiusPixels * 9525, grow: false },
    ] })).toThrowError(expect.objectContaining({ metric: 'active-decoded-bytes' }));
  });

  it('evaluates spatial and shape-effect children as alpha masks, without recoloring the parent', () => {
    const shifted = buffer([10, 20, 30, 255], [40, 50, 60, 128]);
    applyBlipPixelEffects(shifted, { effects: [{ type: 'alphaMod', container: { kind: 'tree', effects: [
      { type: 'relativeOffset', tx: .5, ty: 0 },
    ] } }] });
    expect([...shifted.data]).toEqual([10, 20, 30, 0, 40, 50, 60, 128]);
    const shadow = { color: 'FF0000', alpha: .5, blur: 0, dist: 9525, dir: 0 };
    for (const [effect, alpha] of [
      [{ type: 'alphaOutset', radius: 9525 }, 255],
      [{ type: 'alphaOutset', radius: -9525 }, 0],
      [{ type: 'glow', glow: { color: 'FF0000', alpha: .5, radius: 0 } }, 128],
      [{ type: 'softEdge', radius: 9525 }, 244],
      [{ type: 'shadow', inner: false, shadow }, 0],
      [{ type: 'shadow', inner: true, shadow }, 128],
      [{ type: 'presetShadow', shadows: [shadow] }, 0],
    ] satisfies [BlipEffect, number][]) {
      const buf = buffer([10, 20, 30, 255]);
      applyBlipPixelEffects(buf, { effects: [{ type: 'alphaMod', container: { kind: 'tree', effects: [effect] } }] });
      expect([...buf.data], effect.type).toEqual([10, 20, 30, alpha]);
    }
  });

  it('keeps overlay fill geometry in the authored frame after a grown blur', () => {
    const result = applyBlipPixelEffects(buffer([255, 0, 0, 255]), { effects: [
      { type: 'blur', radius: 9525, grow: true },
      { type: 'fillOverlay', blend: 'over', fill: { fillType: 'solid', color: '00FF00' } },
    ] });
    expect([...result.data.slice(16, 20)]).toEqual([0, 255, 0, 255]);
    expect(result.data[3]).toBe(0);
    expect(result.data[7]).toBeLessThan(255);
  });

  it('evaluates part-local named definitions and charges their lookup work', () => {
    const effects: BlipEffect[] = [
      { type: 'alphaMod', container: { kind: 'tree', effects: [{ type: 'reference', reference: 'external' }] } },
      { type: 'definitions', work: 10, containers: [
        { name: 'external', kind: 'tree', effects: [{ type: 'alphaRepl', alpha: .5 }] },
      ] },
    ];
    const buf = buffer([1, 2, 3, 128]);
    applyBlipPixelEffects(buf, { effects });
    expect([...buf.data]).toEqual([1, 2, 3, 64]);
    expect(() => assertBlipPixelEffectsBudget({ effects: [effects[0], {
      type: 'definitions', work: MAX_IMAGE_EFFECT_PIXEL_WORK + 1, containers: [],
    }] }, 0)).toThrowError(expect.objectContaining({ metric: 'image-effect-work' }));
  });

  it('admits identity blur at the existing maximum base without fictitious kernel allocation', () => {
    expect(() => assertBlipPixelEffectsBudget({ effects: [{ type: 'blur', radius: 0, grow: true }] },
      MAX_IMAGE_EFFECT_BASE_PIXELS, { width: 4096, height: 2048 })).not.toThrow();
  });

  it('grayscl + biLevel(50%) thresholds the Rec. 709 luma PowerPoint uses', () => {
    // PowerPoint PDF evidence: (2,167,223) → white; (0,147,190), (0,126,229)
    // and (9,74,178) → black under its "Black and White" picture setting.
    const buf = buffer([2, 167, 223, 255], [0, 147, 190, 255], [0, 126, 229, 128], [9, 74, 178, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }, { type: 'biLevel', thresh: 0.5 }] });
    expect([...buf.data]).toEqual([255, 255, 255, 255, 0, 0, 0, 255, 0, 0, 0, 128, 0, 0, 0, 255]);
  });

  it('grayscl keeps alpha and writes the luma to every channel', () => {
    const buf = buffer([255, 0, 0, 7]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }] });
    const gray = Math.floor(blipLuminance(255, 0, 0) * 255);
    expect([...buf.data]).toEqual([gray, gray, gray, 7]);
  });

  it('grayscl truncates the luma to a whole level, as the PowerPoint boundary evidence shows', () => {
    // Luma 127.59 must become 127 (black under biLevel 50%), 128.02 stays 128.
    expect(blipGrayLevel(99, 131, 178)).toBe(127);
    expect(blipGrayLevel(98, 132, 177)).toBe(128);
    expect(blipGrayLevel(255, 255, 255)).toBe(255);
    const buf = buffer([99, 131, 178, 255], [98, 132, 177, 255]);
    applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }, { type: 'biLevel', thresh: 0.5 }] });
    expect([...buf.data]).toEqual([0, 0, 0, 255, 255, 255, 255, 255]);
  });

  it('clrChange replaces exact RGB matches with the target colour and alpha', () => {
    const buf = buffer([255, 255, 255, 255], [254, 255, 255, 255], [255, 255, 255, 10]);
    applyBlipPixelEffects(buf, {
      effects: [{ type: 'colorChange', from: 'FFFFFF', fromAlpha: 1, to: 'FFFFFF', toAlpha: 0, useAlpha: false }],
    });
    expect([...buf.data]).toEqual([255, 255, 255, 0, 254, 255, 255, 255, 255, 255, 255, 0]);
    const strict = buffer([255, 255, 255, 255], [255, 255, 255, 10]);
    applyBlipPixelEffects(strict, {
      effects: [{ type: 'colorChange', from: 'FFFFFF', fromAlpha: 1, to: '000000', toAlpha: 0.5, useAlpha: true }],
    });
    expect([...strict.data]).toEqual([0, 0, 0, 128, 255, 255, 255, 10]);
  });

  it('applies effects in document order, with the duotone at its marker', () => {
    const duotone = { clr1: '000000', clr2: 'FF0000' };
    // Duotone first maps white to red; grayscale after it grays the red.
    const first = buffer([255, 255, 255, 255]);
    applyBlipPixelEffects(first, { effects: [{ type: 'duotone' }, { type: 'grayscale' }], duotone });
    expect(first.data[0]).toBe(first.data[1]);
    // Grayscale first leaves white, which the later duotone maps to red.
    const second = buffer([255, 255, 255, 255]);
    applyBlipPixelEffects(second, { effects: [{ type: 'grayscale' }, { type: 'duotone' }], duotone });
    expect([...second.data]).toEqual([255, 0, 0, 255]);
    // Without a marker the duotone applies after the listed effects.
    const unmarked = buffer([255, 255, 255, 255]);
    applyBlipPixelEffects(unmarked, { effects: [{ type: 'grayscale' }], duotone });
    expect([...unmarked.data]).toEqual([255, 0, 0, 255]);
  });

  it.each([
    [{ type: 'alphaBiLevel', thresh: 128 / 255 }, [0, 0, 255, 255]],
    [{ type: 'alphaCeiling' }, [0, 255, 255, 255]],
    [{ type: 'alphaFloor' }, [0, 0, 0, 255]],
    [{ type: 'alphaInv' }, [255, 128, 127, 0]],
    [{ type: 'alphaModFix', amt: 2 }, [0, 254, 255, 255]],
    [{ type: 'alphaRepl', alpha: .5 }, [128, 128, 128, 128]],
  ] satisfies [BlipEffect, number[]][])('applies the specified alpha boundaries for %j', (effect, expected) => {
    const buf = buffer(...[0, 127, 128, 255].map(alpha => [10, 20, 30, alpha]));
    applyBlipPixelEffects(buf, { effects: [effect] });
    expect([...buf.data]).toEqual(expected.flatMap(alpha => [10, 20, 30, alpha]));
  });

  it('allows a later alpha replacement to reveal a duotone applied to transparent RGB', () => {
    const buf = buffer([0, 0, 0, 0]);
    applyBlipPixelEffects(buf, { effects: [
      { type: 'duotone', clr1: 'FF0000', clr2: 'FFFFFF' },
      { type: 'alphaRepl', alpha: 1 },
    ] });
    expect([...buf.data]).toEqual([255, 0, 0, 255]);
  });

  it('modulates alpha before an alpha-sensitive color change', () => {
    const buf = buffer([0, 0, 0, 255]);
    applyBlipPixelEffects(buf, { effects: [
      { type: 'alphaModFix', amt: .5 },
      { type: 'colorChange', from: '000000', fromAlpha: 128 / 255,
        to: 'FF0000', toAlpha: 1, useAlpha: true },
    ] });
    expect([...buf.data]).toEqual([255, 0, 0, 255]);
  });

  it('retains each duotone palette and keys different second palettes apart', () => {
    const effects: BlipEffect[] = [
      { type: 'luminance', bright: .9, contrast: 0 },
      { type: 'duotone', clr1: '000000', clr2: 'FF0000' },
      { type: 'duotone', clr1: '000000', clr2: '00FF00' },
    ];
    const buf = buffer([0, 0, 0, 255]);
    applyBlipPixelEffects(buf, { effects });
    expect(buf.data[0]).toBe(0);
    expect(buf.data[1]).toBeGreaterThan(0);
    expect(buf.data[2]).toBe(0);
    expect(blipPixelEffectsKey({ effects })).not.toBe(blipPixelEffectsKey({
      effects: [...effects.slice(0, 2), { type: 'duotone', clr1: '000000', clr2: '0000FF' }],
    }));
  });

  it('replaces RGB after luminance while retaining source alpha', () => {
    const buf = buffer([0, 0, 0, 64]);
    applyBlipPixelEffects(buf, { effects: [
      { type: 'luminance', bright: .9, contrast: 0 },
      { type: 'colorReplace', color: 'FF0000' },
    ] });
    expect([...buf.data]).toEqual([255, 0, 0, 64]);
  });

  it('rejects a malformed parser diagnostic before mutating pixels', () => {
    const buf = buffer([0, 0, 0, 255]);
    expect(() => applyBlipPixelEffects(buf, { effects: [
      { type: 'luminance', bright: .9, contrast: 0 },
      { type: 'unsupported', effect: 'malformed effect' },
    ] })).toThrow(expect.objectContaining({ code: 'unsupported-image-effect' }));
    expect([...buf.data]).toEqual([0, 0, 0, 255]);
  });

  // Levels read from PowerPoint's PDF export of a 256-step gray ramp under
  // <a:lum> (inputs 0, 32, 64, 96, 128, 160, 192, 224, 255).
  it('lum follows the brightness/contrast levels PowerPoint renders', () => {
    const inputs = [0, 32, 64, 96, 128, 160, 192, 224, 255];
    const levels = (bright: number, contrast: number) => {
      const buf = buffer(inputs.flatMap((v) => [v, v, v, 200]));
      applyBlipPixelEffects(buf, { effects: [{ type: 'luminance', bright, contrast }] });
      expect([...buf.data].filter((_, i) => i % 4 === 3).every((a) => a === 200)).toBe(true);
      return [...buf.data].filter((_, i) => i % 4 === 0);
    };
    const near = (actual: number[], office: number[]) => actual.forEach((value, i) => {
      expect(Math.abs(value - office[i])).toBeLessThanOrEqual(1);
    });
    near(levels(0, -0.7), [89, 99, 108, 118, 128, 137, 147, 156, 166]);
    near(levels(0, -0.35), [45, 65, 86, 107, 128, 149, 169, 190, 210]);
    near(levels(0.35, 0), [89, 121, 153, 185, 217, 249, 255, 255, 255]);
    near(levels(0.7, -0.7), [205, 215, 224, 234, 244, 253, 255, 255, 255]);
    expect(levels(1, 0)).toEqual(Array(9).fill(255));
    expect(levels(-1, 0)).toEqual(Array(9).fill(0));
    expect(levels(0, -1)).toEqual(Array(9).fill(128));
    expect(levels(0, 1)).toEqual([0, 0, 0, 0, 255, 255, 255, 255, 255]);
  });
});

describe('getCachedDuotoneBitmapByPath with CT_Blip effects', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('does not reuse a luminance bitmap for a colliding decimal parameter pair', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 1, height: 1, close() {} })));
    const written: number[] = [];
    const factory = ((width: number, height: number) => ({
      width, height,
      getContext: () => ({
        drawImage() {},
        getImageData: () => buffer([115, 115, 115, 255]),
        putImageData: (image: ImageData) => written.push(image.data[0]),
      }),
    })) as unknown as OffscreenFactory;
    const fetchImage = async () => new Blob([new Uint8Array([1])], { type: 'image/png' });
    const get = (bright: number, contrast: number) => getCachedDuotoneBitmapByPath(
      'lum.png', 'image/png', { effects: [{ type: 'luminance', bright, contrast }] },
      fetchImage, { offscreenFactory: factory },
    );
    try {
      // Concatenated decimal keys could alias distinct tuples; the structural
      // key keeps every in-range ST_FixedPercentage pair distinct, so each
      // recipe runs once and never reuses another recipe's pixels.
      const first = await get(0.1, 0.25);
      const second = await get(0.01, 0.25);
      expect(first).not.toBe(second);
      expect(written).toHaveLength(2);
      expect(written[0]).not.toBe(written[1]);
    } finally {
      dropDuotoneBitmapCache(fetchImage);
      dropBitmapCacheByPath(fetchImage);
    }
  });

  it('rejects ordered effects when the pixel pipeline is unavailable', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 1, height: 1, close() {} })));
    await expect(getCachedDuotoneBitmapByPath('picture.png', 'image/png', {
      effects: [{ type: 'alphaModFix', amt: .5 }],
    }, async () => new Blob(), { offscreenFactory: () => null }))
      .rejects.toMatchObject({ code: 'unsupported-image-effect' });
  });

  it('rejects non-finite physical cache dimensions before decoding', async () => {
    const fetchImage = vi.fn(async () => new Blob());
    await expect(getCachedDuotoneBitmapByPath('blur.png', 'image/png', {
      effects: [{ type: 'blur', radius: 12700, grow: false }],
    }, fetchImage, { widthPt: NaN, heightPt: 3 })).rejects.toMatchObject({ code: 'unsupported-image-effect' });
    expect(fetchImage).not.toHaveBeenCalled();
  });

  it('keeps physical blur sizes 1.2 × 3 and 1 × 2.3 independent in the real cache', async () => {
    const pixels = new Uint8ClampedArray(Array.from({ length: 9 }, () => [20, 40, 60, 255]).flat());
    const factory = ((w: number, h: number) => {
      const surface = { width: w, height: h, pixels: new Uint8ClampedArray(), getContext: () => ({
        drawImage() {},
        getImageData: () => ({ data: pixels.slice(), width: 3, height: 3 }),
        createImageData: (width: number, height: number) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) }),
        putImageData: (data: ImageData) => { surface.pixels = data.data.slice(); },
      }) };
      return surface;
    }) as unknown as OffscreenFactory;
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: { pixels?: Uint8ClampedArray }) => ({
      width: 3, height: 3, pixels: src.pixels?.slice(), close() {},
    })));
    const effects: BlipPixelEffects = { effects: [{ type: 'blur', radius: 12700, grow: false }] };
    const fetchImage = async () => new Blob();
    const independentFetch = async () => new Blob();
    const render = (owner: typeof fetchImage, widthPt: number, heightPt: number) =>
      getCachedDuotoneBitmapByPath('blur.png', 'image/png', effects, owner,
        { widthPt, heightPt, offscreenFactory: factory }) as Promise<ImageBitmap & { pixels: Uint8ClampedArray }>;
    const first = await render(fetchImage, 1.2, 3);
    const second = await render(fetchImage, 1, 2.3);
    const independent = await render(independentFetch, 1, 2.3);
    expect(second.pixels).toEqual(independent.pixels);
    expect(second.pixels).not.toEqual(first.pixels);
    expect(await render(fetchImage, 1, 2.3)).toBe(second);
    dropBitmapCacheByPath(fetchImage);
    dropBitmapCacheByPath(independentFetch);
  });

  it('distinguishes decimal numeric tuples in effect and EMU keys', () => {
    for (const [a, b] of [
      [{ effects: [{ type: 'hsl', hue: 1.2, sat: .3, lum: 0 }] },
        { effects: [{ type: 'hsl', hue: 1, sat: .23, lum: 0 }] }],
      [{ effects: [], emuPerPixelX: 1.2, emuPerPixelY: 3 },
        { effects: [], emuPerPixelX: 1, emuPerPixelY: 2.3 }],
      [{ effects: [{ type: 'luminance', bright: .1, contrast: 1 }] },
        { effects: [{ type: 'luminance', bright: 0, contrast: .11 }] }],
    ] satisfies [BlipPixelEffects, BlipPixelEffects][]) {
      expect(blipPixelEffectsKey(a)).not.toBe(blipPixelEffectsKey(b));
    }
  });

  it('runs the ordered effects once through the shared decode cache, keyed apart from the duotone', async () => {
    const base = { width: 1, height: 1, close() {} } as unknown as ImageBitmap;
    const transformed = { width: 1, height: 1, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (src: unknown) => (src instanceof Blob ? base : transformed));
    vi.stubGlobal('createImageBitmap', cib);
    const written: Uint8ClampedArray[] = [];
    const factory = ((w: number, h: number) => ({
      width: w,
      height: h,
      getContext() {
        return {
          drawImage() {},
          getImageData() {
            return { data: new Uint8ClampedArray([2, 167, 223, 255]), width: 1, height: 1 } as unknown as ImageData;
          },
          putImageData(img: ImageData) {
            written.push(img.data);
          },
        };
      },
    })) as unknown as OffscreenFactory;
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const effects: BlipPixelEffects = {
      effects: [{ type: 'grayscale' }, { type: 'biLevel', thresh: 0.5 }],
      duotone: null,
    };
    const path = 'ppt/media/blip-effects.png';
    const first = await getCachedDuotoneBitmapByPath(path, 'image/png', effects, fetchImage, { offscreenFactory: factory });
    const second = await getCachedDuotoneBitmapByPath(path, 'image/png', effects, fetchImage, { offscreenFactory: factory });
    expect(first).toBe(transformed);
    expect(second).toBe(transformed);
    expect(written).toHaveLength(1);
    expect([...written[0]]).toEqual([255, 255, 255, 255]);
    expect(duotoneCacheKey(path, effects)).not.toBe(duotoneCacheKey(path, null));
    expect(duotoneCacheKey(path, { clr1: '000000', clr2: 'FFFFFF' })).toBe(`${path}|duo:000000:FFFFFF`);
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });
});

describe('CT_Blip effect workload budget (shared resource policy)', () => {
  const grays = (count: number): BlipEffect[] => Array.from({ length: count }, () => ({ type: 'grayscale' }));
  const limitOf = (run: () => unknown) => {
    try {
      run();
    } catch (error) {
      if (!isOoxmlDecodedImageLimitError(error)) throw error;
      return { metric: error.metric, limit: error.limit, observed: error.observed };
    }
    return undefined;
  };

  it('derives the pass limit from the work ceiling at the largest effect base', () => {
    expect(MAX_IMAGE_EFFECT_PASSES).toBe(16);
    expect(MAX_IMAGE_EFFECT_PASSES * MAX_IMAGE_EFFECT_BASE_PIXELS).toBe(MAX_IMAGE_EFFECT_PIXEL_WORK);
  });

  it('admits the pass limit and rejects one more pass before touching a pixel', () => {
    const atLimit = buffer([255, 0, 0, 9]);
    applyBlipPixelEffects(atLimit, { effects: grays(MAX_IMAGE_EFFECT_PASSES) });
    expect(atLimit.data[0]).toBe(atLimit.data[1]);

    const over = buffer([255, 0, 0, 9]);
    expect(limitOf(() => applyBlipPixelEffects(over, { effects: grays(MAX_IMAGE_EFFECT_PASSES + 1) })))
      .toEqual({ metric: 'image-effect-count', limit: 16, observed: 17 });
    expect([...over.data]).toEqual([255, 0, 0, 9]);
    // A duotone without a position marker is one more pass.
    expect(limitOf(() => applyBlipPixelEffects(buffer([255, 0, 0, 9]), {
      effects: grays(MAX_IMAGE_EFFECT_PASSES),
      duotone: { clr1: '000000', clr2: 'FFFFFF' },
    }))).toEqual({ metric: 'image-effect-count', limit: 16, observed: 17 });
  });

  it('admits cumulative work at the ceiling and rejects one pixel visit more', () => {
    const two = { effects: grays(2) };
    const half = MAX_IMAGE_EFFECT_PIXEL_WORK / 2;
    expect(limitOf(() => assertBlipPixelEffectsBudget(two, half))).toBeUndefined();
    expect(limitOf(() => assertBlipPixelEffectsBudget(two, half + 1)))
      .toEqual({ metric: 'image-effect-work', limit: MAX_IMAGE_EFFECT_PIXEL_WORK, observed: 2 * (half + 1) });
    // The applied transform checks the grid it would scan before reading it:
    // this buffer has no pixel storage at all, only its length.
    const unread = { data: { length: (half + 1) * 4 } as unknown as Uint8ClampedArray, width: half + 1, height: 1 };
    expect(limitOf(() => applyBlipPixelEffects(unread, two))?.metric).toBe('image-effect-work');
  });
});

describe('getCachedDuotoneBitmapByPath effect budget', () => {
  afterEach(() => vi.unstubAllGlobals());

  const factory = ((w: number, h: number) => ({
    width: w,
    height: h,
    getContext() {
      return {
        drawImage() {},
        getImageData() {
          return { data: new Uint8ClampedArray([255, 0, 0, 255]), width: 1, height: 1 } as unknown as ImageData;
        },
        putImageData() {},
      };
    },
  })) as unknown as OffscreenFactory;
  const grays = (count: number): BlipEffect[] => Array.from({ length: count }, () => ({ type: 'grayscale' }));

  it('rejects an over-long effect list before fetching or decoding the blip', async () => {
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const rejected = getCachedDuotoneBitmapByPath(
      'ppt/media/too-many-effects.png',
      'image/png',
      { effects: grays(MAX_IMAGE_EFFECT_PASSES + 1), duotone: null },
      fetchImage,
      { offscreenFactory: factory },
    );
    await expect(rejected).rejects.toMatchObject({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-count' });
    expect(fetchImage).not.toHaveBeenCalled();
  });

  it('runs the pass limit on the largest effect base (work exactly at the ceiling)', async () => {
    // 4096 × 2048 = MAX_IMAGE_EFFECT_BASE_PIXELS; 16 passes = 2^27 visits.
    const base = { width: 4096, height: 2048, close() {} } as unknown as ImageBitmap;
    const transformed = { width: 4096, height: 2048, close() {} } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: unknown) => (src instanceof Blob ? base : transformed)));
    const fetchImage = vi.fn(async (_p: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }));
    const result = await getCachedDuotoneBitmapByPath(
      'ppt/media/largest-effect-base.png',
      'image/png',
      { effects: grays(MAX_IMAGE_EFFECT_PASSES), duotone: null },
      fetchImage,
      { offscreenFactory: factory },
    );
    expect(result).toBe(transformed);
    dropDuotoneBitmapCache(fetchImage);
    dropBitmapCacheByPath(fetchImage);
  });
});

it('keeps numeric effect parameters distinct in cache keys', () => {
  const key = (bright: number, contrast: number) => blipPixelEffectsKey({
    effects: [{ type: 'luminance', bright, contrast }],
  });
  expect(key(0.1, 0.23)).not.toBe(key(0.12, 0.3));
  expect(key(0.12, 0.3)).not.toBe(key(0.1, 0.23));
  expect(key(-0.5, 0.25)).not.toBe(key(-0.52, 0.5));
});

it('charges the reviewer’s two tiled overlays before any painting, including recursive masks', () => {
  const fill = { fillType: 'image' as const, imagePath: 'tile.png', mimeType: 'image/png',
    dpi: 65535, tile: { sx: .00003, sy: 1000 } };
  const overlay: BlipEffect = { type: 'fillOverlay', fill, blend: 'over' };
  const runtime = { images: new Map([[fill, { width: 1, height: 1 } as CanvasImageSource]]),
    offscreenFactory: vi.fn(() => { throw new Error('painting must not start'); }) };
  const physical = { emuPerPixelX: 12700, emuPerPixelY: 12700 };
  // Each fill alone fits; their 182,041,680 draw iterations together do not.
  const one = { ...physical, effects: [overlay] };
  expect(blipPixelEffectWork(one, 1, 1, runtime)).toBeGreaterThanOrEqual(91_020_840);
  expect(blipPixelEffectWork(one, 1, 1, runtime)).toBeLessThan(MAX_IMAGE_EFFECT_PIXEL_WORK);
  for (const effects of [[overlay, overlay], [overlay, { type: 'alphaMod', container: {
    kind: 'tree', effects: [overlay],
  } } as BlipEffect]]) {
    const buf = buffer([10, 20, 30, 40]);
    expect(() => applyBlipPixelEffects(buf, { ...physical, effects }, runtime))
      .toThrowError(expect.objectContaining({ metric: 'image-effect-work' }));
    expect([...buf.data]).toEqual([10, 20, 30, 40]);
  }
  expect(runtime.offscreenFactory).not.toHaveBeenCalled();
});


it('budgets both separable blur passes cumulatively through recursive alphaMod', () => {
  const blur: BlipEffect = { type: 'blur', radius: 4500 * 9525, grow: false };
  const one: BlipPixelEffects = { effects: [blur] };
  const work = blipPixelEffectWork(one, 64, 64);
  // Two tap loops and their output visits, plus kernel construction.
  expect(work).toBeGreaterThanOrEqual(64 * 64 * (18002 + 2));
  expect(work).toBeLessThan(MAX_IMAGE_EFFECT_PIXEL_WORK);
  expect(() => assertBlipPixelEffectsBudget({ effects: [blur, { type: 'alphaMod', container: {
    kind: 'tree', effects: [blur],
  } }] }, 64 * 64, { width: 64, height: 64 }))
    .toThrowError(expect.objectContaining({ metric: 'image-effect-work' }));
});

it.each([
  new DOMException('Native bake failed', 'InvalidStateError'),
  Object.assign(new Error('Native bake failed'), { code: 'ERR_NATIVE_BAKE' }),
])('normalizes non-library bake errors even when they carry a code', async failure => {
  vi.stubGlobal('createImageBitmap', vi.fn(async (source: unknown) => {
    if (!(source instanceof Blob)) throw failure;
    return { width: 1, height: 1, close() {} };
  }));
  const factory = (() => ({ width: 1, height: 1, getContext: () => ({
    drawImage() {}, getImageData: () => buffer([0, 0, 0, 255]), putImageData() {},
  }) })) as unknown as OffscreenFactory;
  await expect(getCachedDuotoneBitmapByPath('picture.png', 'image/png', {
    effects: [{ type: 'grayscale' }],
  }, async () => new Blob(), { offscreenFactory: factory }))
    .rejects.toMatchObject({ code: 'unsupported-image-effect' });
});

it.each([
  { type: 'transform', sx: 2, sy: 1, kx: 0, ky: 0, tx: 0, ty: 0 },
  { type: 'blur', radius: 9525, grow: true },
] satisfies BlipEffect[])('relOff uses the previous grid after $type, also inside alphaMod', first => {
  const source = () => buffer([255, 0, 0, 255], [0, 0, 255, 255]);
  const prior = applyBlipPixelEffects(source(), { effects: [first] });
  expect(prior.width).toBe(4);
  const moved = applyBlipPixelEffects(source(), { effects: [first, { type: 'relativeOffset', tx: .5, ty: 0 }] });
  expect(moved.frame?.x).toBe((prior.frame?.x ?? 0) - 2);
  const parent = applyBlipPixelEffects(source(), { effects: [first] });
  const masked = applyBlipPixelEffects(parent, { effects: [{ type: 'alphaMod', container: {
    kind: 'tree', effects: [{ type: 'relativeOffset', tx: .5, ty: 0 }],
  } }] });
  expect(masked.data[3]).toBe(0);
  expect(masked.data[7]).toBe(0);
});

it('counter-rotates recursive shadows and keys orientation, including forced preset shadows', () => {
  const source = () => buffer([255, 0, 0, 255], [0, 0, 255, 255]);
  const shadow = { color: '000000', alpha: 1, blur: 0, dist: 9525, dir: 0 };
  const render = (rotWithShape: boolean, preset = false) => applyBlipPixelEffects(source(), {
    hostRotation: 90, effects: [{ type: 'container', container: { kind: 'tree', effects: [
      preset ? { type: 'presetShadow', shadows: [{ ...shadow, rotWithShape }] }
        : { type: 'shadow', inner: false, shadow: { ...shadow, rotWithShape } },
    ] } }],
  });
  expect(render(true).frame).toMatchObject({ x: -1, y: 0 });
  expect(render(false).frame).toMatchObject({ x: 0, y: 1 });
  expect(render(true, true).frame).toEqual(render(false).frame);
  expect(render(true, true).data.slice(0, 8)).toEqual(render(false).data);
  const effects: BlipEffect[] = [{ type: 'shadow', inner: false, shadow: { ...shadow, rotWithShape: false } }];
  expect(blipPixelEffectsKey({ effects, hostRotation: 0 })).not.toBe(blipPixelEffectsKey({ effects, hostRotation: 90 }));
});

it('rejects cyclic fill recipes and malformed model-source effects before allocating pixels', () => {
  const effects: BlipEffect[] = [];
  effects.push({ type: 'fillOverlay', blend: 'over', fill: { fillType: 'image', imagePath: 'cycle.png', mimeType: 'image/png', stretch: true, blipEffects: effects } });
  expect(() => blipPixelEffectsKey({ effects })).toThrowError(expect.objectContaining({ code: 'unsupported-image-effect' }));
  for (const effect of [
    { type: 'relativeOffset', tx: Infinity, ty: 0 },
    { type: 'colorChange', from: 'invalid', to: '000000', fromAlpha: 1, toAlpha: 1, useAlpha: false },
    { type: 'fillOverlay', blend: 'invalid', fill: { fillType: 'solid', color: 'FF0000' } },
    { type: 'softEdge', radius: -1 },
    { type: 'fill', fill: { fillType: 'image', imagePath: 'bad-crop.png', mimeType: 'image/png', stretch: true, srcRect: { l: NaN, t: 0, r: 0, b: 0 } } },
  ] as unknown as BlipEffect[]) {
    const buf = buffer([1, 2, 3, 255]);
    expect(() => applyBlipPixelEffects(buf, { effects: [{ type: 'grayscale' }, effect] }))
      .toThrowError(expect.objectContaining({ code: 'unsupported-image-effect' }));
    expect([...buf.data]).toEqual([1, 2, 3, 255]);
  }
});

it('keeps independent nested fill definitions scoped to their image evaluator', () => {
  const fill = () => ({ fillType: 'image' as const, imagePath: 'fill.png', mimeType: 'image/png', stretch: true,
    blipEffects: [{ type: 'definitions' as const, work: 1, containers: [{ kind: 'tree' as const, name: 'same', effects: [{ type: 'grayscale' as const }] }] }] });
  expect(() => blipPixelEffectsKey({ effects: [
    { type: 'fillOverlay', blend: 'over', fill: fill() }, { type: 'fillOverlay', blend: 'over', fill: fill() },
  ] })).not.toThrow();
});

it('orients reflection transforms and fades in page axes inside a container', () => {
  const render = (rotWithShape: boolean) => applyBlipPixelEffects({ width: 2, height: 2,
    data: new Uint8ClampedArray(Array(4).fill([0, 0, 0, 255]).flat()) }, {
    hostRotation: 90, effects: [{ type: 'container', container: { kind: 'tree', effects: [{ type: 'reflection', reflection: {
      blur: 0, dist: 0, dir: 0, stA: 1, stPos: 0, endA: 0, endPos: 1, sx: 1, sy: 1, fadeDir: 0, rotWithShape,
    } }] } }],
  });
  expect([...render(true).data].filter((_, i) => i % 4 === 3)).toEqual([191, 64, 191, 64]);
  expect([...render(false).data].filter((_, i) => i % 4 === 3)).toEqual([64, 64, 191, 191]);
});

it('uses a hard reflection fade when start and end occupy the same position', () => {
  const result = applyBlipPixelEffects(buffer([1, 2, 3, 255], [4, 5, 6, 255]), { effects: [{ type: 'reflection', reflection: {
    blur: 0, dist: 0, dir: 0, stA: 1, stPos: .5, endA: 0, endPos: .5, sx: 1, sy: 1, fadeDir: 0,
  } }] });
  expect([result.data[3], result.data[7]]).toEqual([255, 0]);
});


it.each(['fill', 'fillLine'])('restores the original grid when budgeting %s references inside an alphaMod tree', reference => {
  const value: BlipPixelEffects = { effects: [{ type: 'alphaMod', container: { kind: 'tree', effects: [
    { type: 'transform', sx: .001, sy: .001, kx: 0, ky: 0, tx: 0, ty: 0 },
    { type: 'reference', reference },
    ...Array.from({ length: 4 }, (): BlipEffect => ({ type: 'blur', radius: 8 * 9525, grow: false })),
  ] } }] };
  // Four 1024-square, radius-eight convolutions alone need 155,189,520 visits.
  expect(blipPixelEffectWork(value, 1024, 1024)).toBeGreaterThanOrEqual(155_189_520);
  const data = new Uint8ClampedArray(1024 * 1024 * 4).fill(137);
  expect(() => applyBlipPixelEffects({ width: 1024, height: 1024, data }, value))
    .toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-work' }));
  expect(data.every(v => v === 137)).toBe(true);
});


it('bounds live execution work for seeded random recursive grids and fails closed at the shared ceiling', () => {
  let seed = 0x181005;
  const next = (n: number) => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return Math.floor(seed / 0x100000000 * n);
  };
  const shadow = { color: '000000', alpha: .5, blur: 9525, dist: 9525, dir: 90 };
  const effects = (depth: number): BlipEffect[] => Array.from({ length: 1 + next(3) }, (): BlipEffect => {
    switch (next(depth ? 18 : 10)) {
      case 0: return { type: 'transform', sx: [.25, 1, 1.5][next(3)], sy: [.5, 1, 1.5][next(3)],
        kx: next(2) * 10, ky: 0, tx: (next(3) - 1) * 9525, ty: next(2) * 9525 };
      case 1: return { type: 'reference', reference: ['fill', 'fillLine', 'line', 'children', 'named'][next(5)] };
      case 2: return { type: 'blur', radius: next(3) * 9525, grow: !!next(2) };
      case 3: return { type: 'relativeOffset', tx: .25, ty: -.25 };
      case 4: return { type: 'alphaOutset', radius: (next(2) ? 1 : -1) * 9525 };
      case 5: return { type: 'luminance', bright: .2, contrast: -.3 };
      case 6: return { type: 'fillOverlay', blend: 'mult', fill: { fillType: 'solid', color: 'FF0000' } };
      case 7: return { type: 'fill', fill: { fillType: 'solid', color: '00FF00' } };
      case 8: return { type: 'duotone', clr1: '000000', clr2: 'FFFFFF' };
      case 9: return { type: 'alphaInv' };
      case 10: return { type: 'glow', glow: { color: '0000FF', alpha: .5, radius: 9525 } };
      case 11: return { type: 'shadow', shadow, inner: !!next(2) };
      case 12: return { type: 'softEdge', radius: 9525 };
      case 13: return { type: 'presetShadow', shadows: [shadow, { ...shadow, dir: 180 }] };
      case 14: return { type: 'reflection', reflection: { ...shadow, sx: 1, sy: -1, stA: 1, endA: 0, stPos: 0, endPos: 1 } };
      default: {
        const container = { kind: next(2) ? 'tree' as const : 'sib' as const, effects: effects(depth - 1) };
        return next(2) ? { type: next(2) ? 'alphaMod' : 'container', container } : { type: 'blend', blend: 'over', container };
      }
    }
  });
  let evaluated = 0;
  for (let i = 0; i < 160; i++) {
    const width = 1 + next(8), height = 1 + next(8);
    const value: BlipPixelEffects = { effects: [
      { type: 'definitions', work: 7, containers: [{ name: 'named', kind: 'tree', effects: [
        { type: 'blur', radius: 9525, grow: true }, { type: 'reference', reference: 'fill' },
      ] }] }, ...effects(3),
    ], emuPerPixelX: 9525, emuPerPixelY: 12700, duotone: { clr1: '000000', clr2: 'FFFFFF' } };
    const makeBuffer = () => ({ width, height, data: new Uint8ClampedArray(width * height * 4).fill(137) });
    const label = `case ${i}, seed ${seed}: ${JSON.stringify(value)}`;
    let planned: number;
    try { planned = blipPixelEffectWork(value, width, height); }
    catch (error) {
      expect(error, label).toMatchObject({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-count' });
      expect(() => applyBlipPixelEffects(makeBuffer(), value), label)
        .toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
      continue;
    }
    const ledger = new ImageEffectWorkLedger();
    try { applyBlipPixelEffects(makeBuffer(), value, { workLedger: ledger }); }
    catch (error) {
      // Random trees may legitimately exceed the finite pass envelope after
      // compound lowering; every such refusal must be an explicit quota.
      expect(error, label).toMatchObject({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-count' });
      continue;
    }
    expect(ledger.used, label).toBeLessThanOrEqual(planned);
    expect(ledger.used, label).toBeGreaterThan(0);
    // An independent earlier nested fill has already consumed the rest of the
    // shared chain ceiling. Even an admitted plan must stop before overspend.
    const busy = new ImageEffectWorkLedger();
    busy.charge(MAX_IMAGE_EFFECT_PIXEL_WORK - ledger.used + 1);
    expect(() => applyBlipPixelEffects(makeBuffer(), value, { workLedger: busy }), label)
      .toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit', metric: 'image-effect-work' }));
    expect(busy.used, label).toBeLessThanOrEqual(MAX_IMAGE_EFFECT_PIXEL_WORK);
    evaluated++;
  }
  expect(evaluated).toBeGreaterThan(80);
});


it('budgets restored original-fill workspace while the parent and original snapshot remain live', () => {
  const width = 2048, height = 4096;
  const value: BlipPixelEffects = { effects: [{ type: 'alphaMod', container: { kind: 'tree', effects: [
    { type: 'transform', sx: .001, sy: .001, kx: 0, ky: 0, tx: 0, ty: 0 },
    { type: 'reference', reference: 'fill' }, { type: 'blur', radius: 9525, grow: false },
  ] } }] };
  expect(blipPixelEffectWork(value, width, height)).toBeLessThan(MAX_IMAGE_EFFECT_PIXEL_WORK);
  expect(() => assertBlipPixelEffectsBudget(value, width * height, { width, height }))
    .toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit', metric: 'active-decoded-bytes' }));
});

it('restores the original frame before subsequent affine sampling', () => {
  const original = { width: 8, height: 6, frame: { x: 2, y: 1, width: 4, height: 3 },
    data: new Uint8ClampedArray(8 * 6 * 4).fill(137) };
  const move: BlipEffect = { type: 'transform', sx: 1, sy: 1, kx: 0, ky: 0, tx: 9525, ty: -9525 };
  const expected = applyBlipPixelEffects({ ...original, data: original.data.slice() }, { effects: [move] });
  const actual = applyBlipPixelEffects(original, { effects: [
    { type: 'transform', sx: .001, sy: .001, kx: 0, ky: 0, tx: 0, ty: 0 },
    { type: 'reference', reference: 'fillLine' }, move,
  ] });
  expect(actual).toEqual(expected);
});

it('keeps prototype-backed native ImageData from a recursive fill effect', () => {
  // Native ImageData exposes data/width/height as prototype accessors.
  class NativeImageData {
    readonly #data = new Uint8ClampedArray([9, 8, 7, 255]);
    get data() { return this.#data; }
    get width() { return 1; }
    get height() { return 1; }
  }
  const fill = { fillType: 'image' as const, imagePath: 'fill.png', mimeType: 'image/png', stretch: true };
  const ctx = { save() {}, restore() {}, beginPath() {}, rect() {}, clip() {}, drawImage() {},
    translate() {}, rotate() {}, scale() {}, setTransform() {}, getImageData: () => new NativeImageData() };
  const runtime = { images: new Map([[fill, { width: 1, height: 1 } as CanvasImageSource]]),
    offscreenFactory: (() => ({ width: 1, height: 1, getContext: () => ctx })) as unknown as OffscreenFactory };
  const buf = buffer([1, 2, 3, 255]);
  const result = applyBlipPixelEffects(buf, { effects: [{ type: 'fill', fill }] }, runtime);
  expect([...result.data]).toEqual([9, 8, 7, 255]);
});

it.each([
  [false, -1],
  [true, 1],
] as const)('counter-rotates a rotWithShape=0 brush with the host flip parity (flipH=%s)', (flipH, sign) => {
  const fill = { fillType: 'image' as const, imagePath: 'fill.png', mimeType: 'image/png', stretch: true,
    rotWithShape: false };
  const rotate = vi.fn();
  const ctx = { save() {}, restore() {}, beginPath() {}, rect() {}, clip() {}, drawImage() {},
    translate() {}, rotate, scale() {}, getImageData: () => ({ data: new Uint8ClampedArray(4), width: 1, height: 1 }) };
  const runtime = { images: new Map([[fill, { width: 1, height: 1 } as CanvasImageSource]]),
    offscreenFactory: (() => ({ width: 1, height: 1, getContext: () => ctx })) as unknown as OffscreenFactory };
  applyBlipPixelEffects(buffer([1, 2, 3, 255]), { effects: [{ type: 'fill', fill }], hostRotation: 45, hostFlipH: flipH }, runtime);
  expect(rotate).toHaveBeenCalledWith(sign * 45 * Math.PI / 180);
});

it('classifies only chains that can paint outside their frame as spatial', () => {
  expect(blipEffectsMayExtendCoverage([{ type: 'grayscale' }, { type: 'alphaModFix', amt: 0.5 },
    { type: 'blur', radius: 12700, grow: false }])).toBe(false);
  expect(blipEffectsMayExtendCoverage([{ type: 'blur', radius: 12700, grow: true }])).toBe(true);
  expect(blipEffectsMayExtendCoverage([{ type: 'alphaMod', container: { kind: 'tree',
    effects: [{ type: 'relativeOffset', tx: 1, ty: 0 }] } }])).toBe(true);
});

it('fails closed when a recursive tiled gradient brush has no auxiliary surface', () => {
  const fill = { fillType: 'gradient' as const, gradType: 'linear', angle: 0,
    stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
    tileRect: { l: 0, t: 0, r: 0.5, b: 0 } };
  const ctx = { save() {}, restore() {}, beginPath() {}, rect() {}, clip() {}, fillRect() {},
    createLinearGradient: () => ({ addColorStop() {} }), getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }),
    getImageData: () => ({ data: new Uint8ClampedArray(4), width: 1, height: 1 }) };
  const runtime = { offscreenFactory: (() => ({ width: 1, height: 1, getContext: () => ctx })) as unknown as OffscreenFactory };
  const buf = buffer([1, 2, 3, 255]);
  expect(() => applyBlipPixelEffects(buf, { effects: [{ type: 'fill', fill }] }, runtime))
    .toThrowError(expect.objectContaining({ code: 'unsupported-image-effect' }));
  expect([...buf.data]).toEqual([1, 2, 3, 255]);
});

it('fails closed when a recursive tiled gradient pattern cannot be built after its surface allocates', () => {
  const fill = { fillType: 'gradient' as const, gradType: 'linear', angle: 0,
    stops: [{ position: 0, color: '000000' }, { position: 1, color: 'FFFFFF' }],
    tileRect: { l: 0, t: 0, r: 0.5, b: 0 } };
  const surfaceContext = () => ({ save() {}, restore() {}, fillRect() {}, translate() {}, scale() {}, drawImage() {},
    createLinearGradient: () => ({ addColorStop() {} }), getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }) });
  vi.stubGlobal('OffscreenCanvas', class { constructor(readonly width: number, readonly height: number) {}
    getContext() { return surfaceContext(); } });
  try {
    const ctx = { ...surfaceContext(), beginPath() {}, rect() {}, clip() {}, createPattern: () => null,
      getImageData: () => ({ data: new Uint8ClampedArray(4), width: 1, height: 1 }) };
    const runtime = { offscreenFactory: (() => ({ width: 1, height: 1, getContext: () => ctx })) as unknown as OffscreenFactory };
    const buf = buffer([1, 2, 3, 255]);
    expect(() => applyBlipPixelEffects(buf, { effects: [{ type: 'fill', fill }] }, runtime))
      .toThrowError(expect.objectContaining({ code: 'unsupported-image-effect' }));
    expect([...buf.data]).toEqual([1, 2, 3, 255]);
  } finally {
    vi.unstubAllGlobals();
  }
});

it('bounds effect reach conservatively and axis-independently, or not at all', () => {
  // Offsets and transforms are measured on the previous effect's grid
  // (§20.1.8.51/.61), so no static bound exists.
  expect(blipEffectsCoverageReach([{ type: 'blur', radius: 100, grow: true },
    { type: 'relativeOffset', tx: 10, ty: 0 }, { type: 'alphaRepl', alpha: 1 }], 1000, 1000)).toBeNull();
  expect(blipEffectsCoverageReach([{ type: 'transform', sx: 1, sy: 1, kx: 0, ky: 0, tx: 5, ty: 0 }], 10, 10)).toBeNull();
  // A shadow offset only along x is still reported on both axes (rotated
  // hosts), including one whole-frame source-pixel rounding margin.
  const shadow = { color: '000000', alpha: 1, blur: 0, dist: 50, dir: 0 };
  expect(blipEffectsCoverageReach([{ type: 'shadow', shadow, inner: false }], 1000, 400))
    .toEqual({ x: 1050, y: 1050 });
  expect(blipEffectsCoverageReach([{ type: 'grayscale' }], 1000, 400)).toEqual({ x: 0, y: 0 });
  // A named recipe may run several times; its reach is not static.
  expect(blipEffectsCoverageReach([{ type: 'reference', reference: 'grow' }], 10, 10)).toBeNull();
  // Skewed/scaled shadows have no cheap static bound.
  expect(blipEffectsCoverageReach([{ type: 'shadow', shadow: { ...shadow, dist: 0, kx: 80 }, inner: false }], 10, 10))
    .toBeNull();
});

it('rejects an over-wide union grid before allocating it', async () => {
  const { unionEffectGrids } = await import('./blip-spatial');
  expect(() => unionEffectGrids({ width: 1, height: 1 }, { width: 40_001, height: 1, frame: { x: 0, y: 0, width: 1, height: 1 } }))
    .toThrowError(expect.objectContaining({ metric: 'image-dimension' }));
});

it('charges the pinned original readback while replacement buffers are live', () => {
  // Two replacing blurs on a 4096×1792 base: decoded source + pinned readback
  // + first result + scratch + second result exceed the 128 MiB working set.
  const width = 4096, height = 1792;
  const blur: BlipEffect = { type: 'blur', radius: 95, grow: false };
  expect(() => assertBlipPixelEffectsBudget({ effects: [blur, blur] }, width * height, { width, height }))
    .toThrowError(expect.objectContaining({ metric: 'active-decoded-bytes' }));
  // An in-place pointwise chain mutates that readback and stays admitted.
  expect(() => assertBlipPixelEffectsBudget({ effects: [{ type: 'grayscale' }] }, width * height, { width, height }))
    .not.toThrow();
});

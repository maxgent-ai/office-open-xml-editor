import { setImageEffectFrame } from '../image/image-effect-frame';
import { markChartFillPreparationSource } from './image-fill.js';
import { describe, expect, it } from 'vitest';
import type { ImageFill } from '../types/common.js';
import {
  chartThreeDSurfacePictureSceneFace,
  paintChartThreeDSurfacePicture,
} from './three-d-surface-picture.js';

function paint(
  fill: ImageFill,
  options: {
    rejectSourceDraw?: boolean;
    imageWidth?: number;
    imageHeight?: number;
    faceWidth?: number;
    faceHeight?: number;
    pictureFormat?: 'stretch' | 'stack' | 'stackScale';
    projectXScale?: number;
    pictureStackAspect?: number;
    thicknessPercent?: number;
    faceIndex?: number;
    faceIndices?: readonly number[];
    effectFrame?: { x: number; y: number; width: number; height: number };
    sourceFor?: (widthPt: number, heightPt: number) => CanvasImageSource | null;
    preparationSource?: boolean;
  } = {},
): {
  painted: boolean;
  draws: unknown[][];
  sourceDraws: unknown[][];
  transforms: number[][];
  operations: Array<{ kind: 'translate' | 'scale'; values: number[] }>;
} {
  const transforms: number[][] = [];
  const draws: unknown[][] = [];
  const operations: Array<{ kind: 'translate' | 'scale'; values: number[] }> = [];
  const image = {
    width: options.imageWidth ?? 100,
    height: options.imageHeight ?? 100,
  } as unknown as CanvasImageSource;
  if (options.effectFrame) setImageEffectFrame(image, options.effectFrame);
  if (options.preparationSource) markChartFillPreparationSource(image);
  const makeContext = (canvas: object): CanvasRenderingContext2D => {
    const state: Record<string, unknown> = { globalAlpha: 1, canvas };
    return new Proxy(state, {
      get(_target, property: string) {
        if (property === 'getTransform') {
          return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
        }
        if (property === 'setTransform') {
          return (...args: number[]) => transforms.push(args);
        }
        if (property === 'drawImage') {
          return (...args: unknown[]) => {
            if (options.rejectSourceDraw && args[0] === image) throw new Error('unsafe grid reached paint');
            draws.push(args);
          };
        }
        if (property === 'translate' || property === 'scale') {
          return (...values: number[]) => operations.push({ kind: property, values });
        }
        if (property in state) return state[property];
        return () => undefined;
      },
      set(_target, property: string, value) {
        state[property] = value;
        return true;
      },
    }) as unknown as CanvasRenderingContext2D;
  };
  class RecordingCanvas {
    readonly width: number;
    readonly height: number;
    private readonly context: CanvasRenderingContext2D;
    constructor(width: number, height: number) {
      this.width = width;
      this.height = height;
      this.context = makeContext(this);
    }
    getContext(): CanvasRenderingContext2D { return this.context; }
  }
  const ctx = makeContext(new RecordingCanvas(1, 1));
  const faceWidth = options.faceWidth ?? 100;
  const faceHeight = options.faceHeight ?? 100;
  const faceIndex = options.faceIndex ?? 0;
  const faceIndices = options.faceIndices ?? [faceIndex];
  const thicknessPercent = options.thicknessPercent ?? 0;
  const inner = [
    { x: 0, y: faceHeight, depth: 0 },
    { x: faceWidth, y: faceHeight, depth: 0 },
    { x: faceWidth, y: 0, depth: 0 },
    { x: 0, y: 0, depth: 0 },
  ];
  const painted = paintChartThreeDSurfacePicture(
    ctx,
    fill,
    image,
    {
      thicknessPercent,
      pictureOptions: {
        applyToFront: true,
        applyToSides: true,
        applyToEnd: true,
        pictureFormat: options.pictureFormat ?? 'stretch',
      },
    },
    'backWall',
    {
      thickness: thicknessPercent > 0 ? 1 : 0,
      inner,
      outer: [],
      faces: thicknessPercent > 0
        ? Array.from({ length: 6 }, (_, index) => faceIndices.includes(index) ? inner : [])
        : [],
      pictureStackAspect: options.pictureStackAspect
        ?? faceWidth * (options.projectXScale ?? 1) / faceHeight,
      modelDepth: faceWidth,
    },
    faceIndices,
    point => ({ x: point.x * (options.projectXScale ?? 1), y: point.y }),
    10,
    options.sourceFor,
  );
  return {
    painted,
    draws,
    sourceDraws: draws.filter(call => call[0] === image),
    transforms,
    operations,
  };
}

const imageFill = {
  fillType: 'image' as const,
  imagePath: 'surface.png',
  mimeType: 'image/png',
  stretch: true,
};

describe('CT_Surface stretch source and destination rectangles', () => {
  it('rotates only the positive-thickness back-wall top face to Office texture order', () => {
    const topFace = [
      { x: 0, y: 0, depth: 0 },
      { x: 1, y: 0, depth: 0 },
      { x: 1, y: 1, depth: 0 },
      { x: 0, y: 1, depth: 0 },
    ];
    const geometry = {
      thickness: 1,
      inner: topFace,
      outer: topFace,
      faces: [[], [], [], [], topFace, []],
      pictureStackAspect: 1,
      modelDepth: 1,
    };
    const identity = (point: { x: number; y: number }) => point;

    expect(chartThreeDSurfacePictureSceneFace(geometry, 'backWall', 4, identity))
      .toEqual([topFace[1], topFace[2], topFace[3], topFace[0]]);
    expect(chartThreeDSurfacePictureSceneFace(geometry, 'sideWall', 4, identity))
      .toEqual(topFace);
  });

  it('maps the complete source into the authored projected fillRect', () => {
    const { draws, transforms } = paint({
      ...imageFill,
      fillRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.1 },
    });
    expect(draws).toHaveLength(1);
    expect(draws[0].slice(1, 5)).toEqual([0, 0, 100, 100]);
    expect(transforms).toHaveLength(1);
    expect(transforms[0][0]).toBeCloseTo(0.6, 6);
    expect(transforms[0][3]).toBeCloseTo(0.7, 6);
    expect(transforms[0][4]).toBeCloseTo(10, 6);
    expect(transforms[0][5]).toBeCloseTo(20, 6);
  });

  it('clips a negative fillRect outset at the complete face', () => {
    const { draws, transforms } = paint({
      ...imageFill,
      fillRect: { l: -0.25, t: 0, r: 0, b: 0 },
    });
    expect(draws).toHaveLength(1);
    expect(draws[0].slice(1, 5)).toEqual([0, 0, 100, 100]);
    expect(transforms[0][0]).toBeCloseTo(1.25, 6);
    expect(transforms[0][4]).toBeCloseTo(-25, 6);
  });

  it('preserves transparent destination space for a negative srcRect outset', () => {
    const { draws, transforms } = paint({
      ...imageFill,
      srcRect: { l: -0.25, t: 0, r: 0, b: 0 },
    });
    expect(draws).toHaveLength(1);
    expect(draws[0].slice(1, 5)).toEqual([0, 0, 100, 100]);
    expect(transforms[0][0]).toBeCloseTo(0.8, 6);
    expect(transforms[0][4]).toBeCloseTo(20, 6);
  });
});

describe('CT_Surface plain stacked pictures', () => {
  it('preserves image aspect, anchors at the value-axis minimum, and repeats upward', () => {
    const one = paint(imageFill, {
      imageWidth: 400, imageHeight: 100,
      faceWidth: 400, faceHeight: 100,
      pictureFormat: 'stack',
    });
    expect(one.painted).toBe(true);
    expect(one.draws).toHaveLength(1);
    expect(one.transforms[0]).toEqual([1, 0, 0, 1, 0, 0]);

    const two = paint(imageFill, {
      imageWidth: 800, imageHeight: 100,
      faceWidth: 400, faceHeight: 100,
      pictureFormat: 'stack',
    });
    expect(two.painted).toBe(true);
    expect(two.draws).toHaveLength(2);
    expect(two.transforms.map(transform => transform[5])).toEqual([50, 0]);

    const clipped = paint(imageFill, {
      imageWidth: 200, imageHeight: 100,
      faceWidth: 400, faceHeight: 100,
      pictureFormat: 'stack',
    });
    expect(clipped.painted).toBe(true);
    expect(clipped.draws).toHaveLength(1);
    expect(clipped.transforms[0][3]).toBeCloseTo(2, 6);
    expect(clipped.transforms[0][5]).toBeCloseTo(-100, 6);
  });

  it('does not let authored DPI change plain stack geometry', () => {
    const low = paint({ ...imageFill, dpi: 48 }, {
      imageWidth: 800, imageHeight: 100,
      faceWidth: 400, faceHeight: 100,
      pictureFormat: 'stack',
    });
    const high = paint({ ...imageFill, dpi: 192 }, {
      imageWidth: 800, imageHeight: 100,
      faceWidth: 400, faceHeight: 100,
      pictureFormat: 'stack',
    });
    expect(low.transforms).toEqual(high.transforms);
  });

  it('derives aspect from the projected face rather than model-space depth', () => {
    const result = paint(imageFill, {
      imageWidth: 200, imageHeight: 100,
      faceWidth: 100, faceHeight: 100,
      pictureFormat: 'stack',
      projectXScale: 2,
    });
    expect(result.painted).toBe(true);
    expect(result.draws).toHaveLength(1);
  });

  it('shares the plot reference aspect with a differently shaped target wall', () => {
    const result = paint(imageFill, {
      imageWidth: 800, imageHeight: 100,
      faceWidth: 25, faceHeight: 100,
      pictureFormat: 'stack',
      pictureStackAspect: 4,
    });
    expect(result.painted).toBe(true);
    expect(result.draws).toHaveLength(2);
  });

  it('repeats thick front/side faces but maps one complete source on end faces', () => {
    const common = {
      imageWidth: 800,
      imageHeight: 100,
      faceWidth: 400,
      faceHeight: 100,
      pictureFormat: 'stack' as const,
      pictureStackAspect: 4,
      thicknessPercent: 25,
    };
    const front = paint(imageFill, { ...common, faceIndex: 0 });
    const side = paint(imageFill, { ...common, faceIndex: 3 });
    const end = paint(imageFill, { ...common, faceIndex: 2 });
    expect(front.painted).toBe(true);
    expect(side.painted).toBe(true);
    expect(end.painted).toBe(true);
    expect(front.draws).toHaveLength(2);
    expect(side.draws).toHaveLength(2);
    expect(end.draws).toHaveLength(1);
  });

  it('fails closed before drawing when aspect repetition exceeds the image-work ceiling', () => {
    expect(() => paint(imageFill, {
      imageWidth: 4_097 * 400, imageHeight: 100,
      faceWidth: 400, faceHeight: 100,
      pictureFormat: 'stack',
    })).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
  });

  it('bounds the aggregate plain-stack work across every thick slab face', () => {
    const common = {
      imageHeight: 100,
      faceWidth: 400,
      faceHeight: 100,
      pictureFormat: 'stack' as const,
      pictureStackAspect: 4,
      thicknessPercent: 25,
      faceIndices: [0, 2, 3, 4, 5],
    };
    const withinLimit = paint(imageFill, { ...common, imageWidth: 400 * 1_364 });
    expect(withinLimit.painted).toBe(true);
    expect(withinLimit.draws).toHaveLength(4_094);

    expect(() => paint(imageFill, { ...common, imageWidth: 400 * 1_365 }))
      .toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
  });
});

describe('CT_Surface DrawingML tiles', () => {
  it('projects the back-wall tile grid with the surface instead of keeping device-space tiles', () => {
    const result = paint({
      ...imageFill,
      stretch: false,
      dpi: 96,
      tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'ctr' },
    }, {
      imageWidth: 80,
      imageHeight: 40,
      faceWidth: 1_080,
      faceHeight: 40,
      projectXScale: 0.4,
    });
    const projectedDraw = result.draws.find(call => {
      const source = call[0] as { width?: number };
      return source && source !== result.sourceDraws[0]?.[0] && source.width != null;
    });
    // Assert the source extent passed to real paint, before scratch release.
    expect(projectedDraw?.[3]).toBe(1_080);
    // 1,080 / 80 = 13.5 tiles: the final half tile leaves 27 visible
    // half-width colour blocks after the complete grid is projected.
    expect(result.sourceDraws).toHaveLength(14);
  });
});

describe('CT_Surface tiled pictures', () => {
  const tiled = {
    ...imageFill,
    stretch: false,
    dpi: 96,
    tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' },
  } satisfies ImageFill;

  it('rejects the reviewer’s valid-coordinate unsafe tile indices before painting', () => {
    expect(() => paint({ ...tiled, dpi: 65535,
      tile: { ...tiled.tile, tx: 27273042316900, sx: .0001, sy: 10000 } },
      { imageWidth: 1, imageHeight: 1, faceWidth: 1 / 9525, faceHeight: 1 / 9525, rejectSourceDraw: true }))
      .toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
  });

  it('repeats the physical source size in face-local coordinates', () => {
    const result = paint(tiled, {
      imageWidth: 50,
      imageHeight: 25,
      faceWidth: 100,
      faceHeight: 100,
    });
    expect(result.painted).toBe(true);
    expect(result.sourceDraws).toHaveLength(8);
  });

  it('applies scale and alignment before projecting the tile grid', () => {
    const full = paint(tiled, {
      imageWidth: 50,
      imageHeight: 25,
      faceWidth: 100,
      faceHeight: 100,
    });
    const half = paint({
      ...tiled,
      tile: { ...tiled.tile, sx: 0.5, sy: 0.5, algn: 'ctr' },
    }, {
      imageWidth: 50,
      imageHeight: 25,
      faceWidth: 100,
      faceHeight: 100,
    });
    expect(full.painted).toBe(true);
    expect(half.painted).toBe(true);
    expect(half.sourceDraws.length).toBeGreaterThan(full.sourceDraws.length);
    expect(half.operations).not.toEqual(full.operations);
  });

  it('mirrors alternating face-local columns and rows', () => {
    const none = paint(tiled, {
      imageWidth: 50,
      imageHeight: 25,
      faceWidth: 100,
      faceHeight: 100,
    });
    const mirrored = paint({
      ...tiled,
      tile: { ...tiled.tile, flip: 'xy' },
    }, {
      imageWidth: 50,
      imageHeight: 25,
      faceWidth: 100,
      faceHeight: 100,
    });
    expect(mirrored.painted).toBe(true);
    expect(mirrored.sourceDraws).toHaveLength(none.sourceDraws.length);
    expect(mirrored.operations).not.toEqual(none.operations);
  });

  it('applies a positive source crop independently inside each tile', () => {
    const result = paint({
      ...tiled,
      srcRect: { l: 0.25, t: 0, r: 0, b: 0 },
    }, {
      imageWidth: 80,
      imageHeight: 40,
      faceWidth: 80,
      faceHeight: 40,
    });
    expect(result.painted).toBe(true);
    expect(result.sourceDraws).toHaveLength(1);
    expect(result.sourceDraws[0].slice(1)).toEqual([20, 0, 60, 40, 0, 0, 80, 40]);
  });

  it('preserves transparent destination space for a source outset in each tile', () => {
    const result = paint({
      ...tiled,
      srcRect: { l: -0.25, t: 0, r: 0, b: 0 },
    }, {
      imageWidth: 80,
      imageHeight: 40,
      faceWidth: 80,
      faceHeight: 40,
    });
    expect(result.painted).toBe(true);
    expect(result.sourceDraws).toHaveLength(1);
    expect(result.sourceDraws[0].slice(1)).toEqual([0, 0, 80, 40, 16, 0, 64, 40]);
  });

  it('accepts the exact tile-work ceiling and rejects one additional column atomically', () => {
    const exact = paint(tiled, {
      imageWidth: 1,
      imageHeight: 1,
      faceWidth: 64,
      faceHeight: 64,
    });
    expect(exact.painted).toBe(true);
    expect(exact.sourceDraws).toHaveLength(4096);

    expect(() => paint(tiled, {
      imageWidth: 1,
      imageHeight: 1,
      faceWidth: 65,
      faceHeight: 64,
    })).toThrowError(expect.objectContaining({ code: 'ooxml-decoded-image-limit' }));
  });
});


it('maps only the authored frame of an effect-grown picture onto a stretched face', () => {
  const result = paint(imageFill, { effectFrame: { x: 25, y: 25, width: 50, height: 50 } });
  expect(result.painted).toBe(true);
  expect(result.sourceDraws.length).toBeGreaterThan(0);
  for (const draw of result.sourceDraws) {
    const [, sx, sy, sw, sh] = draw as number[];
    // drawProjected pads each source cell by half a pixel against seams.
    expect(sx).toBeGreaterThanOrEqual(24.5 - 1e-9);
    expect(sy).toBeGreaterThanOrEqual(24.5 - 1e-9);
    expect(sx + sw).toBeLessThanOrEqual(75.5 + 1e-9);
    expect(sy + sh).toBeLessThanOrEqual(75.5 + 1e-9);
  }
});

it('draws neighbouring 3-D tiles whose effect outset reaches the face', () => {
  const tiled = { ...imageFill, stretch: false, dpi: 96,
    tile: { tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none', algn: 'tl' } } satisfies ImageFill;
  const options = { imageWidth: 20, imageHeight: 10, faceWidth: 20, faceHeight: 10 };
  const plain = paint(tiled, options);
  // A 5 px outset each side of a 10 px frame reaches one more period.
  const grown = paint(tiled, { ...options, effectFrame: { x: 5, y: 0, width: 10, height: 10 } });
  expect(plain.painted && grown.painted).toBe(true);
  expect(grown.sourceDraws.length).toBeGreaterThan(plain.sourceDraws.length);
});

it('resolves a size-dependent source per stacked repetition at its mapped size', () => {
  const sizes: Array<[number, number]> = [];
  const unit = { width: 800, height: 100 } as unknown as CanvasImageSource;
  const result = paint(imageFill, {
    imageWidth: 800, imageHeight: 100, faceWidth: 400, faceHeight: 100, pictureFormat: 'stack',
    sourceFor: (widthPt, heightPt) => { sizes.push([widthPt, heightPt]); return unit; },
  });
  expect(result.painted).toBe(true);
  // Two stacked repetitions, each mapped onto half the face height.
  expect(sizes).toHaveLength(2);
  for (const [widthPt, heightPt] of sizes) {
    expect(widthPt).toBeCloseTo(400, 6);
    expect(heightPt).toBeCloseTo(50, 6);
  }
});

it('sizes a planar stretched picture as face width by face height', () => {
  const sizes: Array<[number, number]> = [];
  const result = paint(imageFill, {
    faceWidth: 400, faceHeight: 100,
    sourceFor: (widthPt, heightPt) => {
      sizes.push([widthPt, heightPt]);
      return { width: 100, height: 100 } as unknown as CanvasImageSource;
    },
  });
  expect(result.painted).toBe(true);
  expect(sizes).toHaveLength(1);
  expect(sizes[0]![0]).toBeCloseTo(400, 6);
  expect(sizes[0]![1]).toBeCloseTo(100, 6);
});

it('projects each per-mapping source with its own grid and effect frame', () => {
  // The per-size source grew a 10 px outset around its 100 px frame; the
  // initial whole-face source has none. Each draw must use the resolved
  // source's metrics, exactly as when that source is the only one.
  const grown = { width: 120, height: 120 } as unknown as CanvasImageSource;
  setImageEffectFrame(grown, { x: 10, y: 10, width: 100, height: 100 });
  const mapped = paint(imageFill, { sourceFor: () => grown });
  const direct = paint(imageFill, { imageWidth: 120, imageHeight: 120,
    effectFrame: { x: 10, y: 10, width: 100, height: 100 } });
  expect(mapped.painted && direct.painted).toBe(true);
  const strip = (calls: unknown[][]) => calls.map(call => call.slice(1));
  expect(mapped.draws.length).toBeGreaterThan(0);
  expect(mapped.draws.every(call => call[0] === grown)).toBe(true);
  expect(strip(mapped.draws)).toEqual(strip(direct.sourceDraws));
});

it('records no plain-stack repetition from a preparation placeholder', () => {
  const sizes: Array<[number, number]> = [];
  const result = paint(imageFill, {
    imageWidth: 800, imageHeight: 100, faceWidth: 400, faceHeight: 100, pictureFormat: 'stack',
    preparationSource: true,
    sourceFor: (widthPt, heightPt) => { sizes.push([widthPt, heightPt]); return null; },
  });
  expect(result.painted).toBe(false);
  expect(sizes).toEqual([]);
});

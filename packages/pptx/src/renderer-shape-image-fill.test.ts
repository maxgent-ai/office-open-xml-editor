import { beforeEach, describe, expect, it, vi } from 'vitest';
import { OoxmlDecodedImageLimitError, OoxmlResourceLimitError } from '@silurus/ooxml-core';
import type { Slide } from './types.js';

const mocks = vi.hoisted(() => ({
  image: { width: 2, height: 1, close() {} } as unknown as ImageBitmap,
  decode: vi.fn(),
}));

vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return { ...actual, getCachedDuotoneBitmapByPath: mocks.decode,
    rasterizeMathSvg: async () => ({ source: mocks.image, widthPx: 1, heightPx: 1 }),
    tintMathRaster: () => mocks.image };
});

import { paintPreparedShapeImageFill, prepareSlideMath, renderSlide } from './renderer.js';

function recordingCanvas() {
  const calls: Array<[string, ...unknown[]]> = [];
  const state: Record<string, unknown> = { globalAlpha: 1, fillStyle: '', strokeStyle: '' };
  const stack: Array<Record<string, unknown>> = [];
  const ctx = new Proxy(state, {
    get(target, property: string) {
      if (property in target) return target[property];
      if (property === 'save') return () => stack.push({ ...target });
      if (property === 'restore') return () => Object.assign(target, stack.pop());
      if (property === 'getTransform') return () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 });
      if (property === 'measureText') return (text: string) => ({ width: text.length * 8 });
      return (...args: unknown[]) => { calls.push([property, ...args]); };
    },
    set(target, property: string, value) { target[property] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  const canvas = {
    width: 0, height: 0, style: {}, offsetWidth: 960,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
  return { canvas, ctx, calls, state };
}

function pngHeader(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return bytes;
}

function shape(imagePath: string, x = 0, withText = false) {
  return {
    type: 'shape', x, y: 0, width: 914_400, height: 914_400,
    rotation: 0, flipH: false, flipV: false, geometry: 'ellipse',
    fill: {
      fillType: 'image', imagePath, mimeType: 'image/png', stretch: true,
      rotWithShape: true, alpha: 0.5,
    },
    stroke: { color: '000000', width: 12_700 },
    textBody: withText ? {
      verticalAnchor: 't', paragraphs: [{
        alignment: 'l', marL: 0, marR: 0, indent: 0, spaceBefore: null,
        spaceAfter: null, spaceLine: null, lvl: 0, bullet: { type: 'inherit' },
        defFontSize: null, defColor: null, defBold: null, defItalic: null,
        defFontFamily: null, tabStops: [], eaLnBrk: true,
        runs: [{
          type: 'text', text: 'Picture fill text', bold: null, italic: null,
          underline: false, strikethrough: false, fontSize: 18,
          color: null, fontFamily: null,
        }],
      }], defaultFontSize: 18, defaultBold: null, defaultItalic: null,
      lIns: 0, rIns: 0, tIns: 0, bIns: 0, wrap: 'square', vert: 'horz', autoFit: 'none',
    } : null,
    defaultTextColor: null, custGeom: null,
    adj: null, adj2: null, adj3: null, adj4: null,
    adj5: null, adj6: null, adj7: null, adj8: null, shadow: null,
  } as const;
}

describe('PPTX ordinary-shape image fills', () => {
  beforeEach(() => {
    mocks.decode.mockReset();
    mocks.decode.mockResolvedValue(mocks.image);
  });

  it('clips the asymmetric source to the current path and isolates alpha/crop state', () => {
    const { ctx, calls, state } = recordingCanvas();
    const painted = paintPreparedShapeImageFill(ctx, {
      fillType: 'image', imagePath: 'ppt/media/asymmetric.png', mimeType: 'image/png',
      stretch: true, alpha: 0.5,
      srcRect: { l: 0.5, t: 0, r: 0, b: 0 },
      fillRect: { l: 0.1, t: 0.2, r: 0.3, b: 0.1 },
    }, { image: mocks.image }, { x: 10, y: 20, w: 100, h: 50 }, 1, true);

    expect(painted).toBe(true);
    expect(calls).toContainEqual(['clip', 'evenodd']);
    const draw = calls.find(([name]) => name === 'drawImage');
    expect(draw?.slice(0, 8)).toEqual(['drawImage', mocks.image, 1, 0, 1, 1, 20, 30]);
    expect(draw?.[8]).toBeCloseTo(60);
    expect(draw?.[9]).toBeCloseTo(35);
    expect(state.globalAlpha).toBe(1);
  });

  // §20.1.8.14 EG_FillModeProperties: an omitted mode has no default and a
  // tile+stretch pair is schema-invalid; neither paints (nor decodes) a blip.
  it.each([
    ['omitted', {}],
    ['conflicting', { stretch: true, tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' } }],
  ])('fails closed when the fill mode is %s', async (_label, mode) => {
    const fill = {
      fillType: 'image', imagePath: 'ppt/media/mode.png', mimeType: 'image/png', dpi: 96, ...mode,
    } as const;
    const direct = recordingCanvas();
    expect(paintPreparedShapeImageFill(
      direct.ctx, fill, { image: mocks.image }, { x: 10, y: 20, w: 100, h: 50 }, 1,
    )).toBe(false);
    expect(direct.calls.some(([name]) => name === 'drawImage')).toBe(false);

    const { canvas, calls } = recordingCanvas();
    await renderSlide(canvas, {
      index: 0, slideNumber: 1, background: null,
      elements: [{ ...shape('ppt/media/mode.png'), fill }],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob([pngHeader(2, 1) as BlobPart], { type: 'image/png' })),
    });
    expect(mocks.decode).not.toHaveBeenCalled();
    expect(calls.some(([name]) => name === 'drawImage')).toBe(false);
    expect(calls.some(([name]) => name === 'stroke')).toBe(true);
  });

  // Two shapes sharing one PNG with different CT_Blip pixel effects must each
  // paint their own transformed decode, whichever comes first on the slide.
  it.each([false, true])('keeps per-shape blip effects for a shared source (reversed=%s)', async (reversed) => {
    const plain = { kind: 'plain', width: 2, height: 1, close() {} };
    const gray = { kind: 'gray', width: 2, height: 1, close() {} };
    const threshold = { kind: 'threshold', width: 2, height: 1, close() {} };
    mocks.decode.mockImplementation(async (_path: string, _mime: string, transform?: {
      effects?: Array<{ type: string }>;
    }) => {
      const kinds = transform?.effects?.map((effect) => effect.type).join(',');
      if (kinds === 'grayscale') return gray;
      if (kinds === 'biLevel,grayscale') return threshold;
      return plain;
    });
    const shared = 'ppt/media/shared-effects.png';
    const withEffects = (x: number, blipEffects?: unknown[]) => ({
      ...shape(shared, x),
      fill: { ...shape(shared).fill, alpha: undefined, ...(blipEffects ? { blipEffects } : {}) },
    });
    const elements = [
      withEffects(0),
      withEffects(1_000_000, [{ type: 'grayscale' }]),
      withEffects(2_000_000, [{ type: 'biLevel', thresh: 0.5 }, { type: 'grayscale' }]),
    ];
    if (reversed) elements.reverse();
    const { canvas, calls } = recordingCanvas();
    await renderSlide(canvas, {
      index: 0, slideNumber: 1, background: null, elements,
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob([pngHeader(2, 1) as BlobPart], { type: 'image/png' })),
    });

    const painted = calls
      .filter(([name]) => name === 'drawImage')
      .map(([, image]) => (image as { kind: string }).kind);
    const expected = ['plain', 'gray', 'threshold'];
    expect(painted).toEqual(reversed ? expected.reverse() : expected);
    expect(mocks.decode).toHaveBeenCalledTimes(3);
  });

  it('prepares a shared resource once, then paints each owning shape before its stroke', async () => {
    const { canvas, calls } = recordingCanvas();
    const slide = {
      index: 0, slideNumber: 1, background: null,
      elements: [shape('ppt/media/shared.png', 0, true), shape('ppt/media/shared.png', 1_000_000)],
    } as unknown as Slide;
    const textRuns = vi.fn();
    await renderSlide(canvas, slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['png'], { type: 'image/png' })),
    }, textRuns);

    expect(mocks.decode).toHaveBeenCalledTimes(1);
    expect(calls.filter(([name]) => name === 'drawImage')).toHaveLength(2);
    const firstImage = calls.findIndex(([name]) => name === 'drawImage');
    const firstStroke = calls.findIndex(([name]) => name === 'stroke');
    expect(firstImage).toBeGreaterThan(-1);
    expect(firstImage).toBeLessThan(firstStroke);
    expect(textRuns.mock.calls.map(([run]) => run.text).join('')).toBe('Picture fill text');
  });

  it('contains an ordinary decode failure while retaining shape stroke and text', async () => {
    mocks.decode.mockRejectedValue(new Error('bad image'));
    const { canvas, calls } = recordingCanvas();
    const textRuns = vi.fn();
    await expect(renderSlide(canvas, {
      index: 0, slideNumber: 1, background: null,
      elements: [shape('ppt/media/bad.png', 0, true)],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob(['bad'], { type: 'image/png' })),
    }, textRuns)).resolves.toBeDefined();

    expect(calls.some(([name]) => name === 'drawImage')).toBe(false);
    expect(calls.some(([name]) => name === 'stroke')).toBe(true);
    expect(textRuns.mock.calls.map(([run]) => run.text).join('')).toBe('Picture fill text');
  });

  it('samples a tiled effect chain at the fill DPI, independent of the shape size', async () => {
    mocks.decode.mockClear();
    mocks.decode.mockResolvedValue(mocks.image);
    const { canvas } = recordingCanvas();
    const effected = {
      ...shape('ppt/media/dpi-tile.png'),
      fill: {
        fillType: 'image', imagePath: 'ppt/media/dpi-tile.png', mimeType: 'image/png', dpi: 192,
        tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
        blipEffects: [{ type: 'blur', radius: 12700, grow: false }],
      },
    };
    await renderSlide(canvas, {
      index: 0, slideNumber: 1, background: null, elements: [effected],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob([pngHeader(16, 16) as BlobPart], { type: 'image/png' })),
    }).catch(() => undefined); // tile painting needs a surface; only preparation is asserted
    expect(mocks.decode).toHaveBeenCalledTimes(1);
    const [, , transform, , options] = mocks.decode.mock.calls[0]!;
    expect(transform).toMatchObject({ emuPerPixelX: 4762.5, emuPerPixelY: 4762.5 });
    expect(options).toMatchObject({ widthPt: 0, heightPt: 0 });
  });

  it('rejects a tiled effect fill whose source inspection fails natively instead of omitting it', async () => {
    mocks.decode.mockClear();
    const { canvas } = recordingCanvas();
    const effected = {
      ...shape('ppt/media/effected-tile.png'),
      fill: {
        fillType: 'image', imagePath: 'ppt/media/effected-tile.png', mimeType: 'image/png', dpi: 96,
        tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
        blipEffects: [{ type: 'grayscale' }],
      },
    };
    await expect(renderSlide(canvas, {
      index: 0, slideNumber: 1, background: null, elements: [effected],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => { throw new TypeError('native read failed'); }),
    })).rejects.toMatchObject({ code: 'unsupported-image-effect' });
    expect(mocks.decode).not.toHaveBeenCalled();
  });

  it.each([false, true])('keeps a shared tile native-sized regardless of consumer order (%s)', async (reversed) => {
    class TileCanvas {
      getContext() { return recordingCanvas().ctx; }
    }
    vi.stubGlobal('OffscreenCanvas', TileCanvas);
    const target = recordingCanvas();
    target.state.createPattern = () => ({});
    const shared = 'ppt/media/tile-and-stretch.png';
    const tiled = {
      ...shape(shared),
      fill: {
        fillType: 'image', imagePath: shared, mimeType: 'image/png', dpi: 300,
        tile: { algn: 'tl', tx: 0, ty: 0, sx: 1, sy: 1, flip: 'none' },
      },
    };
    const stretched = {
      ...shape(shared, 1_000_000),
      fill: {
        fillType: 'image', imagePath: shared, mimeType: 'image/png', stretch: true,
        srcRect: { l: 0.25, t: 0, r: 0, b: 0 },
        fillRect: { l: 0.1, t: 0, r: 0.2, b: 0 },
      },
    };
    mocks.decode.mockClear();
    await renderSlide(target.canvas, {
      index: 0, slideNumber: 1, background: null,
      elements: reversed ? [stretched, tiled] : [tiled, stretched],
    } as unknown as Slide, 9_144_000, 6_858_000, {
      width: 960, dpr: 1,
      fetchImage: vi.fn(async () => new Blob([pngHeader(300, 150) as BlobPart], { type: 'image/png' })),
    });

    expect(mocks.decode).toHaveBeenCalledTimes(1);
    expect(mocks.decode.mock.calls[0][4]).not.toHaveProperty('targetWidthPx');
    expect(mocks.decode.mock.calls[0][4]).not.toHaveProperty('targetHeightPx');
    vi.unstubAllGlobals();
  });

  it.each([
    { sx: Number.POSITIVE_INFINITY, sy: 1, flip: 'none' },
  ])('rejects an unbounded tile auxiliary and restores caller state: %o', (tile) => {
    const { ctx, calls, state } = recordingCanvas();
    expect(() => paintPreparedShapeImageFill(ctx, {
      fillType: 'image', imagePath: 'ppt/media/huge.png', mimeType: 'image/png',
      dpi: 96, tile: { algn: 'tl', tx: 0, ty: 0, ...tile },
    }, { image: mocks.image }, { x: 0, y: 0, w: 100, h: 100 }, 1)).toThrowError(expect.objectContaining({ code: 'unsupported-image-effect' }));
    expect(calls.some(([name]) => name === 'drawImage')).toBe(false);
    expect(state.globalAlpha).toBe(1);
  });

  it('preserves in-budget tile sampling and bounds only oversized repeating cells', () => {
    const allocations: Array<[number, number]> = [];
    const patternScales: number[][] = [];
    class ScratchCanvas {
      constructor(width: number, height: number) { allocations.push([width, height]); }
      getContext() { return recordingCanvas().ctx; }
    }
    vi.stubGlobal('OffscreenCanvas', ScratchCanvas);
    const paint = (sx: number, intrinsicSize?: { width: number; height: number }) => {
      const { ctx, calls } = recordingCanvas();
      ctx.createPattern = () => ({ setTransform() {} }) as unknown as CanvasPattern;
      const painted = paintPreparedShapeImageFill(ctx, {
        fillType: 'image', imagePath: 'ppt/media/boundary.png', mimeType: 'image/png',
        dpi: 96,
        tile: { algn: 'tl', tx: 0, ty: 0, sx, sy: 4096.1, flip: 'none' },
      }, { image: mocks.image, intrinsicSize }, { x: 0, y: 0, w: 100, h: 100 }, 1 / 9525);
      patternScales.push(calls.find(([name]) => name === 'scale')!.slice(1) as number[]);
      return painted;
    };

    expect(paint(2047.05)).toBe(true); // ceil(4094.1) × ceil(4096.1) fits the old budget
    expect(allocations).toEqual([[4095, 4097]]);
    expect(patternScales[0]).toEqual([1, 1]);
    expect(paint(2047.55)).toBe(true); // ceil(4095.1) × ceil(4096.1) crosses the old budget
    expect(allocations).toHaveLength(2);
    expect(allocations[1]).toEqual([2, 1]);
    expect(allocations[1][0] * allocations[1][1]).toBeLessThanOrEqual(16_777_216);
    expect(paint(1_000_000)).toBe(true);
    expect(allocations[2]).toEqual([2, 1]);
    expect(allocations[2][0] * patternScales[2][0]).toBeCloseTo(2_000_000);
    expect(allocations[2][1] * patternScales[2][1]).toBeCloseTo(4096.1);
    // A cropped/outset native source can itself cross the budget only after
    // independent rounding: 4095.1×4096.1 would ceil to 4096×4097.
    expect(paint(1, { width: 4095.1, height: 4096.1 })).toBe(true);
    expect(allocations[3][0] * allocations[3][1]).toBeLessThanOrEqual(16_777_216);
    vi.unstubAllGlobals();
  });

  it.each([false, true])('counter-rotates an image fill after clipping (flipped=%s)', (flipH) => {
    const { ctx, calls } = recordingCanvas();
    expect(paintPreparedShapeImageFill(ctx, {
      fillType: 'image', imagePath: 'ppt/media/image.png', mimeType: 'image/png',
      stretch: true, rotWithShape: false,
    }, { image: mocks.image }, { x: 10, y: 20, w: 100, h: 50 }, 1, false,
    { rotation: 90, flipH, flipV: false })).toBe(true);
    const clip = calls.findIndex(([name]) => name === 'clip');
    const rotate = calls.findIndex(([name]) => name === 'rotate');
    expect(rotate).toBeGreaterThan(clip);
    expect(calls[rotate][1]).toBeCloseTo((flipH ? 1 : -1) * Math.PI / 2);
    expect(calls).toContainEqual(['translate', 60, 45]);
    expect(calls.some(([name]) => name === 'drawImage')).toBe(true);
  });
});


it('selects the complete authored equation fallback branch before planning image decodes', async () => {
  mocks.decode.mockClear();
  const first = shape('ppt/media/first.png', 0, true);
  const equation = { ...shape('ppt/media/live.png', 0, true),
    textBody: { ...first.textBody!, paragraphs: [{ ...first.textBody!.paragraphs[0], runs: [
      { type: 'math', nodes: [{ kind: 'run', text: 'x', style: 'roman' }], display: false, fontSize: 18 },
    ] }] },
    mathFallback: { before: 1, after: 1, elements: [shape('ppt/media/preview.png', 0, true)] },
  };
  const slide = { index: 0, slideNumber: 1, background: null,
    elements: [first, equation, shape('ppt/media/last.png', 0, true)] } as unknown as Slide;
  const text = vi.fn();
  await renderSlide(recordingCanvas().canvas, slide, 9_144_000, 6_858_000, {
    width: 960, dpr: 1,
    fetchImage: async () => new Blob([pngHeader(2, 1) as BlobPart], { type: 'image/png' }),
    math: { loadMathJax: async () => {}, mathMLToSvg: async () => { throw new Error('conversion'); } },
  }, text);
  expect(mocks.decode.mock.calls.map(args => args[0])).toEqual(['ppt/media/preview.png']);
  expect(text.mock.calls.map(([run]) => run.text).join('')).toBe('Picture fill text');
  expect(slide.elements).toHaveLength(3); // rendering never rewrites caller-owned input
});

it.each([
  new OoxmlDecodedImageLimitError('active-decoded-bytes', 1, 2),
  new OoxmlResourceLimitError('equation resource limit', {
    stage: 'parsing', violation: {
      format: 'pptx', operation: 'parse-slide', resource: 'xml-tree', metric: 'depth',
      limit: 1, observed: 2, configurable: false,
      usage: { archiveEntryCount: 1, declaredInflatedBytes: 2,
        distinctInflatedBytes: 2, operationInflatedBytes: 2 },
    },
  }),
].flatMap(error => (['load', 'conversion'] as const).map(phase => ({ error, phase }))))(
  'propagates equation resource failures during $phase instead of selecting a preview ($error.name)', async ({ error, phase }) => {
    const live = shape('', 0, true);
    const slide = { elements: [{ ...live, textBody: { ...live.textBody!, paragraphs: [{
      ...live.textBody!.paragraphs[0], runs: [
        { type: 'math', nodes: [{ kind: 'run', text: 'x', style: 'roman' }], display: false },
      ],
    }] } }] } as unknown as Slide;
    await expect(prepareSlideMath(slide, {
      loadMathJax: async () => { if (phase === 'load') throw error; },
      mathMLToSvg: async () => { throw error; },
    })).rejects.toBe(error);
  },
);

function equationTable(body: unknown, mathFallback?: unknown) {
  return { type: 'table', x: 0, y: 0, width: 914_400, height: 914_400,
    rotation: 0, flipH: false, flipV: false, cols: [914_400],
    rows: [{ height: 914_400, cells: [{ textBody: body, gridSpan: 1, rowSpan: 1 }] }], mathFallback };
}

it.each(['all', 'first', 'last', 'none'])(
  'replaces an inline equation branch exactly once after %s conversion failures', async (failure) => {
    const live = shape('', 0, true);
    const paragraph = live.textBody!.paragraphs[0];
    const equations = ['x', 'y'].map(text => ({ type: 'math',
      nodes: [{ kind: 'run', text, style: 'roman' }], display: false, fontSize: 18 }));
    const body = { ...live.textBody!, paragraphs: [{ ...paragraph,
      runs: [{ ...paragraph.runs[0], text: 'before' }, ...equations,
        { ...paragraph.runs[0], text: 'after' }],
      mathFallbacks: [{ start: 1, end: 3, text: 'inner preview' },
        { start: 1, end: 3, text: 'complete preview' }],
    }] };
    const slide = { index: 0, slideNumber: 1, background: null,
      elements: [{ ...live, fill: null, stroke: null, textBody: body }] } as unknown as Slide;
    const original = JSON.stringify(slide);
    const recording = recordingCanvas();
    await renderSlide(recording.canvas, slide, 9_144_000, 6_858_000, { width: 960, dpr: 1,
      math: { loadMathJax: async () => {}, mathMLToSvg: async xml => {
        if (failure === 'all' || (failure === 'first' && xml.includes('>x<'))
          || (failure === 'last' && xml.includes('>y<'))) throw new Error('conversion');
        return { svg: '', widthEm: 1, ascentEm: 1, descentEm: 0 };
      } },
    });
    const painted = recording.calls.filter(([name]) => name === 'fillText').map(([, text]) => text);
    expect(painted).toEqual(failure === 'none' ? ['before', 'after'] : ['before', 'complete preview', 'after']);
    expect(recording.calls.filter(([name]) => name === 'drawImage')).toHaveLength(failure === 'none' ? 2 : 0);
    expect(JSON.stringify(slide)).toBe(original);
  },
);

it.each(['table', 'mixed'])('selects a complete %s equation Choice after a table-cell conversion failure', async (kind) => {
  mocks.decode.mockClear();
  const live = shape('ppt/media/live.png', 0, true);
  const body = { ...live.textBody!, paragraphs: [{ ...live.textBody!.paragraphs[0], runs: [
    { type: 'math', nodes: [{ kind: 'run', text: 'x', style: 'roman' }], display: false, fontSize: 18 },
  ] }] };
  const preview = shape('ppt/media/preview.png', 0, true);
  const fallback = { before: 0, after: kind === 'mixed' ? 1 : 0, elements: [preview] };
  const elements = kind === 'table'
    ? [equationTable(body, fallback)]
    : [{ ...live, mathFallback: fallback }, equationTable(body)];
  const text = vi.fn();
  await renderSlide(recordingCanvas().canvas, { index: 0, slideNumber: 1, background: null, elements } as unknown as Slide,
    9_144_000, 6_858_000, { width: 960, dpr: 1,
      fetchImage: async () => new Blob([pngHeader(2, 1) as BlobPart], { type: 'image/png' }),
      math: { loadMathJax: async () => {}, mathMLToSvg: async () => { throw new Error('conversion'); } },
    }, text);
  expect(mocks.decode.mock.calls.map(args => args[0])).toEqual(['ppt/media/preview.png']);
  expect(text.mock.calls.map(([run]) => run.text).join('')).toBe('Picture fill text');
});

it('preserves source locators after preview replacement across master, layout and slide origins', async () => {
  const live = shape('ppt/media/live.png', 0, true);
  const equation = { ...live, textBody: { ...live.textBody!, paragraphs: [{ ...live.textBody!.paragraphs[0], runs: [
    { type: 'math', nodes: [{ kind: 'run', text: 'x', style: 'roman' }], display: false, fontSize: 18 },
  ] }] }, mathFallback: { before: 1, after: 1, elements: [{ ...live, id: 'preview' }] } };
  const slide = { index: 0, slideNumber: 1, background: null,
    elements: [live, equation, live, { ...equationTable(live.textBody), id: 'layout' }, { ...live, id: 'slide' }],
    elementSources: ['master', 'master', 'master', 'layout', 'slide'].map(origin => ({ origin })),
  } as unknown as Slide;
  const text = vi.fn();
  await renderSlide(recordingCanvas().canvas, slide, 9_144_000, 6_858_000, { width: 960, dpr: 1,
    fetchImage: async () => new Blob([pngHeader(2, 1) as BlobPart], { type: 'image/png' }),
    math: { loadMathJax: async () => {}, mathMLToSvg: async () => { throw new Error('conversion'); } },
  }, text);
  const runs = text.mock.calls.map(([run]) => run);
  // A preview is display text outside the source element tree: no invented
  // element/shape locator. Unchanged content retains its original model index.
  expect([...new Map(runs.map(run => [JSON.stringify([run.elementIndex, run.origin, run.shapeId]),
    [run.elementIndex, run.origin, run.shapeId]])).values()]).toEqual([
    [undefined, undefined, undefined], [3, 'layout', 'layout'], [4, 'slide', 'slide'],
  ]);
  expect([...new Set(runs.map(run => run.renderElementIndex))]).toEqual([0, 1, 2]);
  expect(slide.elements).toHaveLength(5);
});

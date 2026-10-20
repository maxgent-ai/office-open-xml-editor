import { expect, it, vi } from 'vitest';
import { chartImageFillHostKey, chartImageFillKey, chartImageFillPixelTransform, paintChartImageFill, withChartImageLookup } from './image-fill.js';
import { setImageEffectFrame } from '../image/image-effect-frame.js';
import { OoxmlDecodedImageLimitError } from '../image/pixel-budget.js';
import { EMU_PER_PT, PT_TO_PX } from '../units.js';
import type { ImageFill } from '../types/common.js';
import { MAX_CHART_IMAGE_FILL_TILES } from './resource-limits.js';

/** A draw spy that turns a non-terminating tile loop into a bounded failure. */
function guardedDrawImage() {
  const spy = vi.fn(() => {
    if (spy.mock.calls.length > MAX_CHART_IMAGE_FILL_TILES) throw new Error('tile loop did not terminate');
  });
  return spy;
}

it('paints a huge finite tile translation at its remainder phase', () => {
  const translations = (tx: number) => {
    const recorded: number[][] = [];
    const ctx = new Proxy({}, {
      get: (_target, key) => key === 'translate'
        ? (x: number, y: number) => { recorded.push([x, y]); }
        : () => {},
    }) as unknown as CanvasRenderingContext2D;
    const fill: ImageFill = {
      fillType: 'image', imagePath: 'tile.png', mimeType: 'image/png', dpi: 72,
      tile: { algn: 'tl', sx: 1, sy: 1, tx, ty: 0, flip: 'x' },
    };
    const painted = withChartImageLookup(() => ({ width: 10, height: 10 }) as CanvasImageSource,
      () => paintChartImageFill(ctx, fill, 0, 0, 40, 40));
    expect(painted).toBe(true);
    return recorded;
  };
  // 10pt tiles mirrored on x repeat every 20pt. 2^30 whole periods plus 5pt
  // must keep the 5pt phase and flip parity rather than stall or be rejected.
  const huge = translations((2 ** 30 * 20 + 5) * EMU_PER_PT);
  const near = translations(5 * EMU_PER_PT);
  expect(huge).toHaveLength(near.length);
  huge.forEach(([x, y], index) => {
    expect(x).toBeCloseTo(near[index]![0]!, 3);
    expect(y).toBeCloseTo(near[index]![1]!, 3);
  });
});

it('rejects the extreme shifted tile grid instead of stalling', () => {
  // Before phase normalization x0 + step === x0 and this loop never advanced.
  // Its tiles are far narrower than the offset's rounding, so the phase is not
  // representable and the fill fails closed before any draw.
  const drawImage = guardedDrawImage();
  const ctx = new Proxy({ drawImage }, {
    get: (target, key) => key === 'drawImage' ? target.drawImage : () => {},
  }) as unknown as CanvasRenderingContext2D;
  const fill: ImageFill = {
    fillType: 'image', imagePath: 'tile.png', mimeType: 'image/png', dpi: 65535,
    tile: { algn: 'tl', sx: 0.0001, sy: 10000, tx: 27273042316900, ty: 0 },
  };
  const emu = PT_TO_PX / EMU_PER_PT;
  expect(() => withChartImageLookup(() => ({ width: 1, height: 1 }) as CanvasImageSource,
    () => paintChartImageFill(ctx, fill, 0, 0, emu, emu))).toThrow(OoxmlDecodedImageLimitError);
  expect(drawImage).not.toHaveBeenCalled();
});

it('fails closed when conversion rounding would misregister a huge mirrored tile phase', () => {
  const ctx = new Proxy({ drawImage: guardedDrawImage() }, {
    get: (target, key) => key === 'drawImage' ? target.drawImage : () => {},
  }) as unknown as CanvasRenderingContext2D;
  const fill = (tx: number): ImageFill => ({
    fillType: 'image', imagePath: 'tile.png', mimeType: 'image/png', dpi: 65535,
    tile: { algn: 'tl', sx: 0.0001, sy: 10000, tx, ty: 0, flip: 'x' },
  });
  const zoomedPtToPx = (4 / 3) * 1e7;
  const render = (tx: number) => withChartImageLookup(() => ({ width: 2, height: 1 }) as CanvasImageSource,
    () => paintChartImageFill(ctx, fill(tx), 0, 0, 1000, 1000, zoomedPtToPx));
  // These offsets differ by whole mirrored periods, but the converted pixel
  // offset of the first has lost more precision than one tile is wide.
  expect(render(508)).toBe(true);
  expect(() => render(27273042316900)).toThrow(OoxmlDecodedImageLimitError);
});

it('tiles a chart fill by the authored frame of an effect-grown source', () => {
  const recorded: number[] = [];
  const ctx = new Proxy({}, {
    get: (_target, key) => key === 'translate'
      ? (x: number) => { recorded.push(x); }
      : () => {},
  }) as unknown as CanvasRenderingContext2D;
  // A blur grow added 5 px on each side of a 10 px authored frame.
  const grown = { width: 20, height: 10 } as CanvasImageSource;
  setImageEffectFrame(grown, { x: 5, y: 0, width: 10, height: 10 });
  const fill: ImageFill = {
    fillType: 'image', imagePath: 'tile.png', mimeType: 'image/png', dpi: 96,
    tile: { algn: 'tl', sx: 1, sy: 1, tx: 0, ty: 0 },
  };
  expect(withChartImageLookup(() => grown, () => paintChartImageFill(ctx, fill, 0, 0, 40, 10))).toBe(true);
  // 96 dpi: one source pixel per CSS pixel, so the period is the 10 px frame.
  expect(recorded[1]! - recorded[0]!).toBeCloseTo(10, 6);
  // The 5 px outset reaches one neighbouring period on each horizontal side:
  // 4 + 2 boundary + 2 outset columns, 1 + 2 boundary rows.
  expect(recorded).toHaveLength(8 * 3);
});

it('evaluates a chart fill chain on the source grid at the fill DPI', () => {
  const effects = [{ type: 'blur' as const, radius: 9525, grow: false }];
  expect(chartImageFillPixelTransform({ fillType: 'image', imagePath: 'a.png', mimeType: 'image/png',
    dpi: 192, blipEffects: effects }, { rotation: 30, flipH: true })).toEqual({ effects, duotone: null,
    emuPerPixelX: 4762.5, emuPerPixelY: 4762.5, hostRotation: 30, hostFlipH: true, hostFlipV: false });
  // Sampling DPI and host orientation are part of the decoded identity.
  const fill = { fillType: 'image' as const, imagePath: 'a.png', mimeType: 'image/png', blipEffects: effects };
  expect(chartImageFillKey({ ...fill, dpi: 96 })).not.toBe(chartImageFillKey({ ...fill, dpi: 192 }));
  expect(chartImageFillHostKey(fill, { rotation: 0 })).not.toBe(chartImageFillHostKey(fill, { rotation: 90 }));
  expect(chartImageFillHostKey({ ...fill, rotWithShape: false }, { rotation: 90 }))
    .toBe(chartImageFillHostKey({ ...fill, rotWithShape: false }, { rotation: 0 }));
  const duotone = { clr1: '000000', clr2: 'FFFFFF' };
  expect(chartImageFillPixelTransform({ fillType: 'image', imagePath: 'a.png', mimeType: 'image/png', duotone }))
    .toBe(duotone);
});

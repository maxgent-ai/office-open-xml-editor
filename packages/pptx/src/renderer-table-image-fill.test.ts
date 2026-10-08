import { afterEach, expect, it, vi } from 'vitest';
import { renderSlide } from './renderer.js';
import type { Fill, Slide, TableCell, TableElement } from './types.js';

function table(background: Fill, cells: Array<Fill | null> = [null]): TableElement {
  return {
    type: 'table', x: 0, y: 0, width: 914_400, height: 914_400,
    rotation: 0, flipH: false, flipV: false, cols: [914_400], background,
    rows: cells.map(fill => ({ height: 914_400 / cells.length, cells: [{
      textBody: null, fill, borderL: null, borderR: null, borderT: null, borderB: null,
      gridSpan: 1, rowSpan: 1, hMerge: false, vMerge: false,
    } as TableCell] })),
  };
}

function recordingCanvas() {
  const paints: Array<{ name: string; args: unknown[]; alpha: unknown; style: unknown }> = [];
  const state: Record<string, unknown> = { globalAlpha: 1, fillStyle: '' };
  const stack: Array<Record<string, unknown>> = [];
  const ctx = new Proxy(state, {
    get(target, name: string) {
      if (name in target) return target[name];
      if (name === 'save') return () => stack.push({ ...target });
      if (name === 'restore') return () => Object.assign(target, stack.pop());
      return (...args: unknown[]) => paints.push({ name, args, alpha: target.globalAlpha, style: target.fillStyle });
    },
    set(target, name: string, value) { target[name] = value; return true; },
  });
  const canvas = { width: 0, height: 0, getContext: () => ctx } as unknown as OffscreenCanvas;
  return { canvas, paints };
}

function pngHeader() {
  const bytes = new Uint8Array(26);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  bytes.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
  const view = new DataView(bytes.buffer);
  view.setUint32(16, 800);
  view.setUint32(20, 600);
  return bytes;
}
const image: Fill = { fillType: 'image', imagePath: 'ppt/media/table.png', mimeType: 'image/png', stretch: true };
const slide = (element: TableElement): Slide => ({ index: 0, slideNumber: 1, background: null, elements: [element] });
afterEach(() => vi.unstubAllGlobals());

it('loads one shared table blip, paints the full background before alpha bands and clips direct cell images', async () => {
  const decode = vi.fn(async () => ({ width: 800, height: 600, close() {} }));
  vi.stubGlobal('createImageBitmap', decode);
  const fetchImage = vi.fn(async () => new Blob([pngHeader()], { type: 'image/png' }));
  const { canvas, paints } = recordingCanvas();
  await renderSlide(canvas, slide(table(image, [
    { fillType: 'solid', color: 'FF000080' },
    { ...image, alpha: 0.25, srcRect: { l: 0.5, t: 0, r: 0, b: 0 } },
  ])), 914_400, 914_400, { width: 96, dpr: 1, fetchImage });
  expect(fetchImage).toHaveBeenCalledWith('ppt/media/table.png', 'image/png');
  expect(decode).toHaveBeenCalledTimes(1);
  const fills = paints.filter(p => p.name === 'drawImage' || p.name === 'fillRect');
  expect(fills.slice(-3).map(p => p.name)).toEqual(['drawImage', 'fillRect', 'drawImage']);
  expect(fills.at(-3)?.args.slice(-4)).toEqual([0, 0, 96, 96]);
  expect(fills.at(-2)?.style).toBe('rgba(255,0,0,0.5019607843137255)');
  expect(fills.at(-1)?.alpha).toBe(0.25);
  expect(fills.at(-1)?.args.slice(1)).toEqual([400, 0, 400, 600, 0, 48, 96, 48]);
});

it('plans every table image together under the decoded-image budget', async () => {
  const decode = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => ({
    width: options?.resizeWidth ?? 800, height: options?.resizeHeight ?? 600, close() {},
  }));
  vi.stubGlobal('createImageBitmap', decode);
  const fetchImage = vi.fn(async () => new Blob([pngHeader()], { type: 'image/png' }));
  const { canvas } = recordingCanvas();
  await renderSlide(canvas, slide(table(image, [{ ...image, imagePath: 'ppt/media/cell.png' }])),
    914_400, 914_400, { width: 96, dpr: 1, fetchImage,
      imageResources: { resolution: 'display', strategy: 'adaptive', decodedByteBudget: 73_728 } });
  expect(decode).toHaveBeenCalledTimes(2);
  for (const [, options] of decode.mock.calls) {
    expect(options).toEqual({ resizeWidth: 96, resizeHeight: 96, resizeQuality: 'high' });
  }
});

it('propagates a table-only decode budget failure through renderSlide', async () => {
  vi.stubGlobal('createImageBitmap', vi.fn(async () => ({ width: 800, height: 600, close() {} })));
  await expect(renderSlide(recordingCanvas().canvas, slide(table(image)), 914_400, 914_400, {
    width: 960, dpr: 1,
    fetchImage: async () => new Blob([pngHeader()], { type: 'image/png' }),
    imageResources: { strategy: 'strict', decodedByteBudget: 65_536 },
  })).rejects.toMatchObject({ code: 'ooxml-decoded-image-limit' });
});

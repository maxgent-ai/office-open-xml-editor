import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// Hold every raster decode open until the test settles it.
vi.mock('@silurus/ooxml-core', async (importOriginal) => {
  const actual = await importOriginal<typeof import('@silurus/ooxml-core')>();
  return {
    ...actual,
    getCachedBitmapByPath: vi.fn(),
    inspectCachedRasterSource: vi.fn(async () => null),
  };
});

import { renderDocumentToCanvas } from './renderer';
import { getCachedBitmapByPath, OoxmlDecodedImageLimitError } from '@silurus/ooxml-core';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

/**
 * Chrome rasterizes a canvas recording in chunks split at task boundaries, and
 * antialiased-path coverage can depend on that chunking (issue #1612). A page
 * must therefore not be cleared or drawn until every decode it needs has
 * settled, and must then be painted without yielding to another task.
 */

/** A fresh byte source per render keeps each case's decode caches cold. */
const bytes = () => vi.fn(
  async (_path: string, mime: string) => new Blob([new Uint8Array([1, 2, 3])], { type: mime }),
);

function imageDoc(paths: readonly string[]): DocxDocumentModel {
  const para: DocParagraph = {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: paths.map((imagePath) => ({
      type: 'image', imagePath, mimeType: 'image/png', widthPt: 40, heightPt: 30, anchor: false,
    })) as unknown as DocParagraph['runs'],
    defaultFontSize: 16, defaultFontFamily: 'Arial',
    widowControl: false,
  };
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...para } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

/** Canvas whose clear (resize) and draw calls are logged with task boundaries. */
function recordingCanvas() {
  const events: string[] = [];
  let taskBoundary = false;
  const record = (name: string) => {
    if (events.length === 0) setTimeout(() => { taskBoundary = true; }, 0);
    else if (taskBoundary) {
      events.push('<task boundary>');
      taskBoundary = false;
      setTimeout(() => { taskBoundary = true; }, 0);
    }
    events.push(name);
  };
  let font = '16px serif';
  const context = new Proxy({} as Record<string | symbol, unknown>, {
    get(target, property) {
      if (property === 'font') return font;
      if (property in target) return target[property];
      if (property === 'measureText') {
        return (s: string) => {
          const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
          return {
            width: [...s].length * px,
            fontBoundingBoxAscent: px * 0.8,
            fontBoundingBoxDescent: px * 0.2,
            actualBoundingBoxAscent: px * 0.8,
            actualBoundingBoxDescent: px * 0.2,
          };
        };
      }
      if (property === 'createLinearGradient') return () => ({ addColorStop() {} });
      if (typeof property === 'string'
        && ['fillRect', 'drawImage', 'fill', 'stroke', 'fillText', 'clip'].includes(property)) {
        return () => record(property);
      }
      return () => undefined;
    },
    set(target, property, value) {
      if (property === 'font') font = value as string;
      else target[property] = value;
      return true;
    },
  });
  const canvas = {
    set width(_value: number) { record('clear'); },
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => context,
  } as unknown as HTMLCanvasElement;
  return { canvas, events };
}

function holdDecodes() {
  const pending: Array<{
    path: string;
    resolve: (bitmap: ImageBitmap) => void;
    reject: (error: unknown) => void;
  }> = [];
  vi.mocked(getCachedBitmapByPath).mockImplementation(
    (path: string) => new Promise((resolve, reject) => {
      pending.push({ path, resolve: resolve as (bitmap: ImageBitmap) => void, reject });
    }),
  );
  const bitmap = { width: 40, height: 30, close() {} } as unknown as ImageBitmap;
  return {
    pending,
    settleAll: () => { for (const { resolve } of pending) resolve(bitmap); },
  };
}

const nextTask = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

describe('DOCX page paint atomicity', () => {
  beforeEach(() => {
    vi.mocked(getCachedBitmapByPath).mockReset();
  });
  afterEach(() => vi.unstubAllGlobals());

  it('paints nothing until all decodes settle, then paints in one task', async () => {
    const decodes = holdDecodes();
    const { canvas, events } = recordingCanvas();
    const fetchImage = bytes();
    const render = renderDocumentToCanvas(
      imageDoc(['word/media/a.png', 'word/media/b.png']), canvas, 0, { dpr: 1, width: 400, fetchImage },
    );
    for (let i = 0; i < 5 && decodes.pending.length < 2; i++) await nextTask();
    expect(decodes.pending.length).toBe(2);
    await nextTask();
    expect(events).toEqual([]);

    decodes.settleAll();
    await render;
    expect(events[0]).toBe('clear');
    expect(events.filter((event) => event === 'drawImage')).toHaveLength(2);
    expect(events).not.toContain('<task boundary>');
  });

  it('stops a superseded render before it clears the newer render', async () => {
    const decodes = holdDecodes();
    const { canvas, events } = recordingCanvas();
    const doc = imageDoc(['word/media/a.png']);
    const fetchImage = bytes();
    const first = renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400, fetchImage });
    const second = renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400, fetchImage });
    let settled = false;
    const both = Promise.all([first, second]).finally(() => { settled = true; });
    // Renders sharing a byte source may decode one after another.
    while (!settled) {
      decodes.settleAll();
      await nextTask();
    }
    await both;
    expect(events.filter((event) => event === 'clear')).toHaveLength(1);
    expect(events.filter((event) => event === 'drawImage')).toHaveLength(1);
  });

  it('raises a decoded-image budget failure after the page background, before content', async () => {
    const decodes = holdDecodes();
    const { canvas, events } = recordingCanvas();
    const fetchImage = bytes();
    const render = renderDocumentToCanvas(
      imageDoc(['word/media/a.png']), canvas, 0, { dpr: 1, width: 400, fetchImage },
    );
    for (let i = 0; i < 5 && decodes.pending.length < 1; i++) await nextTask();
    expect(events).toEqual([]);
    decodes.pending[0]!.reject(
      new OoxmlDecodedImageLimitError('image-pixels', 1, 2),
    );
    await expect(render).rejects.toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(events).toEqual(['clear', 'fillRect']);
  });
});

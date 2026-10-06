import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import { projectXlsxElementContext } from './element-context.js';
import type { XlsxElementContext } from './selection.js';
import type { Styles, Worksheet } from './types.js';

describe('worksheet picture transforms', () => {
  it('wraps crop and alpha paint in the authored centre rotation and flips', () => {
    const ops: Array<{ name: string; args: unknown[]; alpha?: number }> = [];
    const state: Record<string, unknown> = { globalAlpha: 1 };
    const ctx = new Proxy(state, {
      get(target, prop) {
        if (prop === 'canvas') return { width: 400, height: 300 };
        if (prop === 'measureText') return () => ({ width: 7 });
        if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
          return () => ({ addColorStop() {} });
        }
        if (prop === 'drawImage') return (...args: unknown[]) => {
          ops.push({ name: 'drawImage', args, alpha: target.globalAlpha as number });
          if (target.throwDraw) throw new Error('synthetic draw failure');
        };
        if (typeof prop === 'string') return (...args: unknown[]) => ops.push({ name: prop, args });
      },
      set(target, prop, value) { target[String(prop)] = value; return true; },
    }) as unknown as CanvasRenderingContext2D;
    const ws = {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
      defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
      freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], shapeGroups: [],
      images: [{
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
        nativeExtCx: 0, nativeExtCy: 0,
        imagePath: 'xl/media/asymmetric.png', mimeType: 'image/png',
        rotation: 90, flipH: true, flipV: true,
        srcRect: { l: .1, t: .2, r: .3, b: .1 }, alpha: .4,
      }],
    } as Worksheet;

    renderViewport(ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
      { row: 1, col: 1, rows: 10, cols: 10 },
      { loadedImages: new Map([['xl/media/asymmetric.png', { width: 200, height: 100 } as CanvasImageSource]]) });

    expect(ops.find((op) => op.name === 'rotate')?.args[0]).toBeCloseTo(Math.PI / 2, 12);
    expect(ops).toContainEqual({ name: 'scale', args: [-1, -1] });
    const draw = ops.find((op) => op.name === 'drawImage');
    expect(draw?.args).toHaveLength(9);
    expect(draw?.args.slice(1, 5)).toEqual([20, 20, 120, 70]);
    expect(draw?.args.slice(5)).toEqual([50, 22, 134, 40]);
    expect(draw?.alpha).toBe(.4);
    const rotatesAt = ops.findIndex((op) => op.name === 'rotate');
    expect(ops.slice(rotatesAt - 1, rotatesAt + 3)).toEqual([
      { name: 'translate', args: [117, 42] },
      { name: 'rotate', args: [Math.PI / 2] },
      { name: 'scale', args: [-1, -1] },
      { name: 'translate', args: [-117, -42] },
    ]);

    ops.length = 0;
    state.throwDraw = true;
    expect(() => renderViewport(
      ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
      { row: 1, col: 1, rows: 10, cols: 10 },
      { loadedImages: new Map([['xl/media/asymmetric.png', { width: 200, height: 100 } as CanvasImageSource]]) },
    )).toThrow('synthetic draw failure');
    const failedDraw = ops.findIndex((op) => op.name === 'drawImage');
    expect(ops.slice(failedDraw + 1)).toEqual([
      { name: 'restore', args: [] }, // alpha frame
      { name: 'restore', args: [] }, // DrawingML transform frame
    ]);

    // Exercise the other reflection combination in a non-default viewport.
    // RTL changes placement only; the authored local horizontal reflection and
    // negative clockwise angle retain their DrawingML order around the new box.
    ops.length = 0;
    state.throwDraw = false;
    ws.rightToLeft = true;
    ws.images[0].rotation = -30;
    ws.images[0].flipH = true;
    ws.images[0].flipV = false;
    renderViewport(
      ctx, ws, { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles,
      { row: 1, col: 1, rows: 10, cols: 10 },
      {
        cellScale: 1.5, scrollOffsetX: 2, scrollOffsetY: 3,
        loadedImages: new Map([['xl/media/asymmetric.png', { width: 200, height: 100 } as CanvasImageSource]]),
      },
    );
    const secondDraw = ops.find((op) => op.name === 'drawImage');
    expect(secondDraw?.args.slice(1, 5)).toEqual([20, 20, 120, 70]);
    expect(secondDraw?.args.slice(5)).toEqual([126, 28.5, 202, 60]);
    const secondRotate = ops.findIndex((op) => op.name === 'rotate');
    expect(ops.slice(secondRotate - 1, secondRotate + 3)).toEqual([
      { name: 'translate', args: [227, 58.5] },
      { name: 'rotate', args: [-Math.PI / 6] },
      { name: 'scale', args: [-1, 1] },
      { name: 'translate', args: [-227, -58.5] },
    ]);
  });
});

// #1713: an XML `twoCellAnchor` saved with `editAs="oneCell"` keeps its
// from/to markers as the authoritative INITIAL display rectangle (ECMA-376
// Part 1 §20.5.2.33); only later band edits retain that size (§20.5.3.2).
// The anchor-level XML kind is the explicit optional `anchorTag` model fact.
// Untagged models keep the legacy native-extent path.
const EMU_PER_TEST_PX = 9525;
const NATIVE_100x50 = { nativeExtCx: 100 * EMU_PER_TEST_PX, nativeExtCy: 50 * EMU_PER_TEST_PX };
const MARKERS = {
  fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
  toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
};
const XML_TWO_CELL_ANCHOR = { anchorTag: 'twoCellAnchor' } as const;
const EMPTY_STYLES = { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles;
const RANGE = { row: 1, col: 1, rows: 20, cols: 20 };

/** Explicit authored bands isolate the anchor geometry from row auto-fit. */
function explicitBandSheet(extra: Record<string, unknown>): Worksheet {
  const colWidths: Record<number, number> = {};
  const rowHeights: Record<number, number> = {};
  for (let index = 0; index <= 24; index++) {
    colWidths[index] = 20;
    rowHeights[index] = 45;
  }
  return {
    name: 'Sheet1', rows: [], colWidths, rowHeights,
    defaultColWidth: 8.43, defaultRowHeight: 15, defaultRowHeightCustom: true, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], shapeGroups: [], images: [],
    ...extra,
  } as unknown as Worksheet;
}

function paintedDestinations(
  ws: Worksheet,
  loadedImages: Map<string, CanvasImageSource>,
): Map<CanvasImageSource, number[]> {
  const destinations = new Map<CanvasImageSource, number[]>();
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(_target, prop) {
      if (prop === 'canvas') return { width: 800, height: 600 };
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      if (prop === 'drawImage') return (...args: unknown[]) => {
        // Destination is always the trailing x, y, width, height quadruple.
        destinations.set(args[0] as CanvasImageSource, args.slice(-4) as number[]);
      };
      if (typeof prop === 'string') return () => undefined;
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, ws, EMPTY_STYLES, RANGE, { loadedImages });
  return destinations;
}

const ELEMENT_VIEWPORT = {
  width: 800, height: 600, cellScale: 1, viewport: RANGE,
  scrollOffsetX: 0, scrollOffsetY: 0, freezeRows: 0, freezeCols: 0,
};

function elementContext(
  elementType: 'image' | 'shape',
  elementIndex: number,
  shapeIndex?: number,
): XlsxElementContext {
  return {
    format: 'xlsx', kind: 'element', sheetIndex: 0, sheetName: 'Sheet1',
    elementType, elementIndex,
    ...(shapeIndex === undefined ? {} : { shapeIndex }),
    anchor: {
      from: { row: 1, col: 1, offsetX: 0, offsetY: 0 },
      to: { row: 3, col: 3, offsetX: 0, offsetY: 0 },
    },
    truncated: false, truncationReasons: [], textCharacters: 0, maxTextCharacters: 0,
  };
}

describe('XML twoCellAnchor editAs=oneCell initial geometry (#1713)', () => {
  it('paints the tagged picture at its twoCell sibling rectangle, not the child ext', () => {
    const sources = {
      twoCell: { width: 200, height: 100 } as CanvasImageSource,
      oneCell: { width: 200, height: 100 } as CanvasImageSource,
      legacy: { width: 200, height: 100 } as CanvasImageSource,
    };
    const picture = (path: string, extra: Record<string, unknown>) => ({
      ...MARKERS, ...NATIVE_100x50, imagePath: path, mimeType: 'image/png', ...extra,
    });
    const ws = explicitBandSheet({
      images: [
        picture('xl/media/two-cell.png', { ...XML_TWO_CELL_ANCHOR, editAs: 'twoCell' }),
        picture('xl/media/one-cell.png', { ...XML_TWO_CELL_ANCHOR, editAs: 'oneCell' }),
        // Old untagged public model: compatibility path, not the oracle.
        picture('xl/media/legacy.png', { editAs: 'oneCell' }),
      ],
    });
    const painted = paintedDestinations(ws, new Map([
      ['xl/media/two-cell.png', sources.twoCell],
      ['xl/media/one-cell.png', sources.oneCell],
      ['xl/media/legacy.png', sources.legacy],
    ]));

    const sibling = painted.get(sources.twoCell);
    expect(sibling).toHaveLength(4);
    // Fixture sanity: the from/to rectangle must be distinguishable from the
    // 100x50 px child/native ext, or the regression below would be vacuous.
    expect(Math.abs(sibling![2] - 100)).toBeGreaterThan(10);
    expect(Math.abs(sibling![3] - 50)).toBeGreaterThan(10);

    expect(painted.get(sources.oneCell)).toEqual(sibling);

    const legacy = painted.get(sources.legacy);
    expect(legacy?.[2]).toBeCloseTo(100, 0);
    expect(legacy?.[3]).toBeCloseTo(50, 0);
  });

  it('projects hit/outline geometry for tagged pictures and shape leaves from the same initial rectangle', () => {
    const leaf = { zOrder: 1, x: 0.25, y: 0.5, w: 0.5, h: 0.25, rot: 0 };
    const group = (extra: Record<string, unknown>) => ({
      ...MARKERS, ...NATIVE_100x50, ...XML_TWO_CELL_ANCHOR, shapes: [leaf], ...extra,
    });
    const ws = explicitBandSheet({
      images: [
        { ...MARKERS, ...NATIVE_100x50, ...XML_TWO_CELL_ANCHOR, editAs: 'twoCell', imagePath: 'xl/media/a.png', mimeType: 'image/png' },
        { ...MARKERS, ...NATIVE_100x50, ...XML_TWO_CELL_ANCHOR, editAs: 'oneCell', imagePath: 'xl/media/b.png', mimeType: 'image/png' },
      ],
      shapeGroups: [group({ editAs: 'twoCell' }), group({ editAs: 'oneCell' })],
    });
    const project = (context: XlsxElementContext) =>
      projectXlsxElementContext(ws, context, ELEMENT_VIEWPORT)?.rect;

    const imageSibling = project(elementContext('image', 0));
    expect(imageSibling).toBeDefined();
    expect(Math.abs(imageSibling!.width - 100)).toBeGreaterThan(10);
    expect(project(elementContext('image', 1))).toEqual(imageSibling);

    const groupSibling = project(elementContext('shape', 0));
    expect(groupSibling).toEqual(imageSibling);
    expect(project(elementContext('shape', 1))).toEqual(groupSibling);

    // Raw normalized child transforms still map into the anchor rectangle.
    const leafSibling = project(elementContext('shape', 0, 0));
    expect(leafSibling?.width).toBeCloseTo(groupSibling!.width * 0.5, 9);
    expect(project(elementContext('shape', 1, 0))).toEqual(leafSibling);
  });
});

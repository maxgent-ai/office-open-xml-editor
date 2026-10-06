import { describe, expect, it } from 'vitest';
import { getGridGeometryForWorksheet, renderViewport } from './renderer.js';
import { projectXlsxElementContext } from './element-context.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';
import {
  bindInitialAnchorSizes,
  captureInitialAnchorSizes,
  lookupInitialAnchorSize,
  releaseInitialAnchorSizeReference,
  type InitialAnchorSizeReference,
} from './internal/initial-anchor-sizes.js';
import { WorksheetViewProjectionCache, type WireSizeOverrides } from './worker-protocol.js';
import type { XlsxElementContext } from './selection.js';
import type { Styles, Worksheet } from './types.js';

// #1713 B1: a tagged XML twoCellAnchor editAs="oneCell" keeps its PREPARED
// INITIAL size (ECMA-376 Part 1 §20.5.2.33 initial from/to; §20.5.3.2 band
// edits move without resizing) on a projection bound to a captured reference.
// The default twoCell sibling keeps resizing with its markers.
const EMU = 9525;
const NATIVE_100x50 = { nativeExtCx: 100 * EMU, nativeExtCy: 50 * EMU };
const MARKERS = {
  fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
  toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
};
// Initially inverted (from offset beyond the to edge) so it is not renderable;
// widening band 1 by the edit below would make its markers positive.
const INVERTED_MARKERS = {
  fromCol: 0, fromColOff: 300 * EMU, fromRow: 0, fromRowOff: 100 * EMU,
  toCol: 1, toColOff: 0, toRow: 1, toRowOff: 0,
};
const TAGGED_ONE_CELL = { anchorTag: 'twoCellAnchor', editAs: 'oneCell' } as const;
const TAGGED_TWO_CELL = { anchorTag: 'twoCellAnchor', editAs: 'twoCell' } as const;
const EMPTY_STYLES = { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles;
const RANGE = { row: 1, col: 1, rows: 20, cols: 20 };
const SOURCES = {
  one: { width: 200, height: 100 } as CanvasImageSource,
  two: { width: 200, height: 100 } as CanvasImageSource,
  legacy: { width: 200, height: 100 } as CanvasImageSource,
  inverted: { width: 200, height: 100 } as CanvasImageSource,
};
const LOADED = new Map<string, CanvasImageSource>([
  ['xl/media/one.png', SOURCES.one],
  ['xl/media/two.png', SOURCES.two],
  ['xl/media/legacy.png', SOURCES.legacy],
  ['xl/media/inverted.png', SOURCES.inverted],
]);
const BAND_EDIT: WireSizeOverrides = { cols: { 1: 60 }, rows: { 1: 90 } };

/** Cached parsed sheet with explicit prepared bands (no row auto-fit). Its
 * anchor objects are frozen: the source model must stay immutable. */
function sourceSheet(): Worksheet {
  const colWidths: Record<number, number> = {};
  const rowHeights: Record<number, number> = {};
  for (let index = 0; index <= 24; index++) {
    colWidths[index] = 20;
    rowHeights[index] = 45;
  }
  const picture = (path: string, markers: object, extra: object) => Object.freeze({
    ...markers, ...NATIVE_100x50, imagePath: path, mimeType: 'image/png', ...extra,
  });
  // Zero-size leaves keep paint independent of shape drawing details.
  const group = (extra: object) => Object.freeze({
    ...MARKERS, ...NATIVE_100x50, shapes: [{ x: 0, y: 0, w: 0, h: 0, rot: 0 }], ...extra,
  });
  return {
    name: 'Sheet1', rows: [], colWidths, rowHeights,
    defaultColWidth: 8.43, defaultRowHeight: 15, defaultRowHeightCustom: true, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [],
    images: [
      picture('xl/media/one.png', MARKERS, TAGGED_ONE_CELL),
      picture('xl/media/two.png', MARKERS, TAGGED_TWO_CELL),
      picture('xl/media/legacy.png', MARKERS, { editAs: 'oneCell' }),
      picture('xl/media/inverted.png', INVERTED_MARKERS, TAGGED_ONE_CELL),
    ],
    shapeGroups: [group(TAGGED_ONE_CELL), group(TAGGED_TWO_CELL)],
  } as unknown as Worksheet;
}

function editBands(ws: Worksheet): void {
  ws.colWidths[1] = BAND_EDIT.cols![1] as number;
  ws.rowHeights[1] = BAND_EDIT.rows![1] as number;
  GridGeometry.invalidate(ws);
}

function prepare(ws: Worksheet): InitialAnchorSizeReference {
  const reference = captureInitialAnchorSizes(ws, getGridGeometryForWorksheet(ws));
  expect(reference).toBeDefined();
  return reference!;
}

function paint(ws: Worksheet): Map<CanvasImageSource, number[]> {
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
        destinations.set(args[0] as CanvasImageSource, args.slice(-4) as number[]);
      };
      if (typeof prop === 'string') return () => undefined;
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, ws, EMPTY_STYLES, RANGE, { loadedImages: LOADED });
  return destinations;
}

function outline(
  ws: Worksheet,
  elementType: 'image' | 'shape',
  elementIndex: number,
  cellScale = 1,
) {
  const context: XlsxElementContext = {
    format: 'xlsx', kind: 'element', sheetIndex: 0, sheetName: 'Sheet1',
    elementType, elementIndex,
    anchor: {
      from: { row: 1, col: 1, offsetX: 0, offsetY: 0 },
      to: { row: 3, col: 3, offsetX: 0, offsetY: 0 },
    },
    truncated: false, truncationReasons: [], textCharacters: 0, maxTextCharacters: 0,
  };
  return projectXlsxElementContext(ws, context, {
    width: 800, height: 600, cellScale, viewport: RANGE,
    scrollOffsetX: 0, scrollOffsetY: 0, freezeRows: 0, freezeCols: 0,
  })?.rect ?? null;
}

function expectSize(actual: readonly number[] | undefined, width: number, height: number): void {
  expect(actual).toHaveLength(4);
  expect(actual![2]).toBeCloseTo(width, 6);
  expect(actual![3]).toBeCloseTo(height, 6);
}

describe('prepared initial anchor sizes for tagged twoCellAnchor editAs=oneCell (#1713)', () => {
  it('keeps the prepared initial size on paint and outline while the twoCell sibling grows', () => {
    const source = sourceSheet();
    const view = createSheetViewModel(source);
    const reference = prepare(view);
    // Compact primitive facts only, for eligible tagged anchors only.
    expect(structuredClone(reference)).toEqual(reference);
    expect(reference.entries.map((entry) => [entry.family, entry.index])).toEqual([
      ['image', 0], ['image', 3], ['shape', 0],
    ]);
    bindInitialAnchorSizes(view, reference);

    const before = paint(view);
    const initial = before.get(SOURCES.two)!;
    expect(initial).toHaveLength(4);
    expectSize(before.get(SOURCES.one), initial[2], initial[3]);
    expect(before.get(SOURCES.one)!.slice(0, 2)).toEqual(initial.slice(0, 2));
    expect(before.has(SOURCES.inverted)).toBe(false);

    editBands(view);
    const after = paint(view);
    const grown = after.get(SOURCES.two)!;
    // Fixture sanity: the edit really resizes default twoCell anchors.
    expect(grown[2]).toBeGreaterThan(initial[2] + 10);
    expect(grown[3]).toBeGreaterThan(initial[3] + 10);
    expectSize(after.get(SOURCES.one), initial[2], initial[3]);
    // A non-renderable prepared initial rect is not revived by edits.
    expect(after.has(SOURCES.inverted)).toBe(false);

    const image = outline(view, 'image', 0);
    expect(image?.width).toBeCloseTo(initial[2], 6);
    expect(image?.height).toBeCloseTo(initial[3], 6);
    expect(outline(view, 'image', 1)?.width).toBeCloseTo(grown[2], 6);
    expect(outline(view, 'image', 3)).toBeNull();
    expect(outline(view, 'shape', 0)?.width).toBeCloseTo(initial[2], 6);
    expect(outline(view, 'shape', 0)?.height).toBeCloseTo(initial[3], 6);
    expect(outline(view, 'shape', 1)?.width).toBeCloseTo(grown[2], 6);
    // EMU storage: zoom rescales the retained size.
    expect(outline(view, 'image', 0, 2)?.width).toBeCloseTo(initial[2] * 2, 6);

    // The cached source model stays immutable and unbound.
    expect(source.colWidths[1]).toBe(20);
    expect(Object.isFrozen(source.images[0])).toBe(true);
    expect(lookupInitialAnchorSize(source, source.images[0])).toBeUndefined();
  });

  it('keeps sibling references independent, rebinds a reacquired projection, and leaves unbound renders initial-only', () => {
    const source = sourceSheet();
    const edited = createSheetViewModel(source);
    const sibling = createSheetViewModel(source);
    const unbound = createSheetViewModel(source);
    const reference = prepare(edited);
    bindInitialAnchorSizes(edited, reference);
    bindInitialAnchorSizes(sibling, prepare(sibling));
    const initial = paint(sibling).get(SOURCES.one)!;

    editBands(edited);
    editBands(unbound);
    expect(paint(sibling).get(SOURCES.one)).toEqual(initial);
    expectSize(paint(edited).get(SOURCES.one), initial[2], initial[3]);
    // Stateless public renderer: no edit history, so the markers are only the
    // initial display and the edited projection resizes (documented limit).
    const stateless = paint(unbound);
    expect(stateless.get(SOURCES.one)![2]).toBeGreaterThan(initial[2] + 10);
    expect(stateless.has(SOURCES.inverted)).toBe(true);

    // A projection rebuilt from the same workbook binds the same reference.
    const reacquired = createSheetViewModel(source);
    bindInitialAnchorSizes(reacquired, reference);
    editBands(reacquired);
    expectSize(paint(reacquired).get(SOURCES.one), initial[2], initial[3]);
    expect(paint(reacquired).has(SOURCES.inverted)).toBe(false);
  });

  it('rejects mismatched references and never revives edited markers after release', () => {
    const source = sourceSheet();
    const view = createSheetViewModel(source);
    const reference = prepare(view);
    const withImages = (images: Worksheet['images']) =>
      createSheetViewModel({ ...source, images } as Worksheet);
    expect(() => bindInitialAnchorSizes(
      withImages([{ ...source.images[0], toCol: 3 }, ...source.images.slice(1)]), reference,
    )).toThrow(/does not match/);
    expect(() => bindInitialAnchorSizes(
      withImages([{ ...source.images[0], editAs: 'twoCell' }, ...source.images.slice(1)]), reference,
    )).toThrow(/does not match/);
    expect(() => bindInitialAnchorSizes(withImages(source.images.slice(0, 3)), reference))
      .toThrow(/does not match/);
    expect(() => bindInitialAnchorSizes(view, { ...reference, entries: [] })).toThrow(/does not match/);

    for (const cx of [NaN, Infinity, -1]) {
      expect(() => bindInitialAnchorSizes(createSheetViewModel(source), {
        ...reference,
        entries: reference.entries.map((entry, index) => index === 0 ? { ...entry, cx } : entry),
      })).toThrow(/EMU/);
    }
    bindInitialAnchorSizes(view, reference);
    bindInitialAnchorSizes(view, reference);
    expect(() => bindInitialAnchorSizes(view, prepare(view))).toThrow(/already bound/);

    // Replacing acquisition arrays invalidates their association. Retained
    // objects must not fall back to edited markers on this bound projection.
    view.images = [...view.images];
    view.shapeGroups = [...(view.shapeGroups ?? [])];
    editBands(view);
    expect(paint(view).has(SOURCES.one)).toBe(false);
    expect(outline(view, 'image', 0)).toBeNull();
    expect(() => bindInitialAnchorSizes(view, reference)).toThrow(/reference/);
    releaseInitialAnchorSizeReference(reference);
    editBands(view);
    const painted = paint(view);
    expect(painted.has(SOURCES.one)).toBe(false);
    expect(painted.has(SOURCES.two)).toBe(true);
    expect(outline(view, 'image', 0)).toBeNull();
    expect(outline(view, 'shape', 0)).toBeNull();
    expect(() => bindInitialAnchorSizes(createSheetViewModel(source), reference)).toThrow(/released/);
  });

  it('binds a structured-cloned reference to the worker-local projection per revision', () => {
    const source = sourceSheet();
    const wire = structuredClone(prepare(createSheetViewModel(source)));
    const cache = new WorksheetViewProjectionCache();
    const projection = (revision: number, initialAnchorSizes = wire) =>
      ({ id: 7, revision, autoRowHeightsPrepared: true, initialAnchorSizes });

    // Without overrides the reference still gets its own shallow projection;
    // the cached sheet shared by every viewer is never bound or copied.
    const first = cache.resolve(source, 0, projection(1), undefined);
    expect(first.created).toBe(true);
    expect(first.worksheet).not.toBe(source);
    expect(first.worksheet.rowHeights).toBe(source.rowHeights);
    expect(lookupInitialAnchorSize(source, source.images[0])).toBeUndefined();
    const reused = cache.resolve(source, 0, projection(1, structuredClone(wire)), undefined);
    expect(reused.worksheet).toBe(first.worksheet);
    expect(reused.created).toBe(false);
    const initial = paint(first.worksheet).get(SOURCES.one)!;
    expect(initial).toHaveLength(4);

    // A cache hit may reuse a structured clone, but must reject a changed
    // reference before it can paint a cached rectangle for this revision.
    expect(() => cache.resolve(source, 0, projection(1, { ...wire, imageCount: 1 }), undefined))
      .toThrow(/reference/);
    const changedSize = {
      ...wire, entries: wire.entries.map((entry, index) => index === 0
        ? { ...entry, cx: entry.cx + 9525 } : entry),
    };
    expect(() => cache.resolve(source, 0, projection(1, changedSize), undefined))
      .toThrow(/reference/);

    const edited = cache.resolve(source, 0, projection(2), BAND_EDIT);
    expect(edited.created).toBe(true);
    const afterEdit = paint(edited.worksheet);
    expectSize(afterEdit.get(SOURCES.one), initial[2], initial[3]);
    expect(afterEdit.get(SOURCES.two)![2]).toBeGreaterThan(initial[2] + 10);
    expect(afterEdit.has(SOURCES.inverted)).toBe(false);
    expect(source.colWidths[1]).toBe(20);

    // Another viewer on the same cached sheet has no reference: no bleed.
    const other = cache.resolve(source, 0, { id: 8, revision: 2 }, BAND_EDIT);
    expect(paint(other.worksheet).get(SOURCES.one)![2]).toBeGreaterThan(initial[2] + 10);
    expect(cache.resolve(source, 0, { id: 9, revision: 1 }, undefined).worksheet).toBe(source);

    // Release tombstone: late work still paints the retained size but never
    // resurrects a cache entry.
    cache.release(7);
    const late = cache.resolve(source, 0, projection(2), BAND_EDIT);
    expect(late.created).toBe(true);
    expect(late.worksheet).not.toBe(edited.worksheet);
    expectSize(paint(late.worksheet).get(SOURCES.one), initial[2], initial[3]);
    expect(cache.resolve(source, 0, projection(2), BAND_EDIT).worksheet).not.toBe(late.worksheet);

    // A mismatched reference is loud, before any paint.
    expect(() => cache.resolve(source, 0, projection(1, { ...wire, imageCount: 1 }), undefined))
      .toThrow(/does not match/);
  });

  it('captures nothing for ordinary sheets', () => {
    const source = sourceSheet();
    const ordinary = createSheetViewModel({
      ...source,
      images: source.images.slice(1, 3),
      shapeGroups: (source.shapeGroups ?? []).slice(1),
    } as Worksheet);
    expect(captureInitialAnchorSizes(ordinary, getGridGeometryForWorksheet(ordinary))).toBeUndefined();
    expect(lookupInitialAnchorSize(ordinary, ordinary.images[0])).toBeUndefined();
  });
});

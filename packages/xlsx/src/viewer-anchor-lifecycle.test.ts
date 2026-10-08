import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { prepareXlsxViewerRowHeights } from './workbook.js';
import { renderViewport } from './renderer.js';
import { GridGeometry } from './internal/grid-geometry.js';
import { pxToColWidth, pxToRowHeight } from './internal/grid-metrics.js';
import { WorksheetViewProjectionCache, extractViewerRenderContext } from './worker-protocol.js';
import type { Styles, Worksheet } from './types.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';

// #1713 B2: real viewer lifecycle. A tagged twoCellAnchor editAs="oneCell"
// keeps the size of this viewer's first prepared projection (ECMA-376 Part 1
// §20.5.2.33/§20.5.3.2; host-measured automatic heights) through view edits,
// navigation, reacquisition and both render modes, while the twoCell sibling
// resizes with its markers. Archive/worker, font measurement and Canvas are
// doubles; every rectangle comes from the production renderer.

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const EMU = 9525;
const ONE = { width: 200, height: 100 } as CanvasImageSource;
const TWO = { width: 200, height: 100 } as CanvasImageSource;
const LOADED = new Map<string, CanvasImageSource>([
  ['xl/media/one.png', ONE],
  ['xl/media/two.png', TWO],
]);
const EMPTY_STYLES = { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles;
const RANGE = { row: 1, col: 1, rows: 20, cols: 20 };
const STUBBED = [
  'hideCommentPopup', 'hideValidationPanel', 'updateSelectionOverlay', 'updateTabActive',
  'buildCommentMap', 'buildHyperlinkMap', 'buildOutline', 'layoutGutters', 'updateSpacerSize',
  'resetHorizontalScroll', 'updateFindOverlay', 'emitViewportChange', 'buildTabs',
];
type RenderOpts = Parameters<typeof extractViewerRenderContext>[0];
type Frame = Map<CanvasImageSource, number[]>;
type Engine = Record<string, unknown> & {
  wb: unknown;
  canvasArea: object;
  currentWorksheet: Worksheet | null;
  showSheet(index: number): Promise<void>;
  prepareWorkbook(workbook: unknown): boolean;
  renderCurrentSheet(): Promise<void>;
  recordSizeOverride(axis: 'row' | 'col', index: number): void;
  refitAutoRowsAfterColumnResize(): void;
};

/** Immutable cached source: columns A:B 160 px and rows 1:2 60 px; both
 * pictures span A1:B2 (zero-based markers 0..2). */
function sheet(name: string): Worksheet {
  const colWidths: Record<number, number> = {};
  const rowHeights: Record<number, number> = {};
  for (let index = 0; index <= 24; index++) {
    colWidths[index] = pxToColWidth(160);
    rowHeights[index] = pxToRowHeight(60);
  }
  const picture = (imagePath: string, editAs: string) => Object.freeze({
    fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
    toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
    nativeExtCx: 100 * EMU, nativeExtCy: 50 * EMU,
    imagePath, mimeType: 'image/png', anchorTag: 'twoCellAnchor', editAs,
  });
  return {
    name, rows: [], colWidths: Object.freeze(colWidths), rowHeights: Object.freeze(rowHeights),
    defaultColWidth: 8.43, defaultRowHeight: 15, defaultRowHeightCustom: true, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], shapeGroups: [],
    images: Object.freeze([picture('xl/media/one.png', 'oneCell'), picture('xl/media/two.png', 'twoCell')]),
  } as unknown as Worksheet;
}

function paint(ws: Worksheet, cellScale = 1): Frame {
  const drawn: Frame = new Map();
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(_target, prop) {
      if (prop === 'canvas') return { width: 800, height: 600 };
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      if (prop === 'drawImage') return (...args: unknown[]) => {
        drawn.set(args[0] as CanvasImageSource, args.slice(-4) as number[]);
      };
      if (typeof prop === 'string') return () => undefined;
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, ws, EMPTY_STYLES, RANGE, { loadedImages: LOADED, cellScale });
  return drawn;
}

function expectSize(frame: Frame, source: CanvasImageSource, width: number, height: number): void {
  const rect = frame.get(source);
  expect(rect).toHaveLength(4);
  expect(rect![2]).toBeCloseTo(width, 6);
  expect(rect![3]).toBeCloseTo(height, 6);
}

/** Structural workbook double. Main paints the bound viewer projection; the
 * worker double resolves its own parsed model through the production
 * projection cache, as the render worker does. */
function createWorkbook(
  sources: readonly Worksheet[],
  frames: Frame[],
  hostRowHeight?: number,
  gate?: Promise<void>,
) {
  const workerSources = sources.map((source) => ({ ...source, colWidths: { ...source.colWidths }, rowHeights: { ...source.rowHeights } }));
  const projections = new WorksheetViewProjectionCache();
  const lease = async (index: number) => {
    if (index === 1 && gate) await gate;
    return { worksheet: sources[index], release: vi.fn() };
  };
  return {
    sheetNames: sources.map((source) => source.name),
    sheetCount: sources.length,
    getWorksheet: vi.fn(async (index: number) => (await lease(index)).worksheet),
    acquireWorksheetLease: vi.fn(lease),
    retainWorksheetReference: vi.fn(async () => vi.fn()),
    renderViewport: vi.fn(async (_canvas: unknown, _index: number, _range: unknown, opts: RenderOpts) => {
      frames.push(paint(extractViewerRenderContext(opts).worksheet!, opts.cellScale));
    }),
    renderViewportToBitmap: vi.fn(async (index: number, _range: unknown, opts: RenderOpts) => {
      const context = extractViewerRenderContext(opts);
      // The real worker commits the degraded parsed placeholder before render.
      const workerSource = context.worksheet?.parseError
        ? { ...workerSources[index], images: [], shapeGroups: [], parseError: context.worksheet.parseError }
        : workerSources[index];
      frames.push(paint(projections.resolve(
        workerSource, index, context.projection, context.opts.sizeOverrides,
      ).worksheet, opts.cellScale));
      return { width: 1, height: 1, close: vi.fn() };
    }),
    destroy: vi.fn(),
    // Host-font automatic height of row 1 (wrapped text): `hostRowHeight`
    // until column A is widened, then 30 px.
    ...(hostRowHeight === undefined ? {} : {
      [prepareXlsxViewerRowHeights](view: Worksheet): void {
        view.rowHeights[1] = pxToRowHeight(view.colWidths[1] > pxToColWidth(160) ? 30 : hostRowHeight);
        GridGeometry.invalidate(view);
      },
    }),
  };
}

function mount(
  sources: readonly Worksheet[],
  options: { worker?: boolean; hostRowHeight?: number; gate?: Promise<void>; cellScale?: number } = {},
) {
  installDom();
  const container = makeContainer();
  const viewer = new XlsxViewer(
    container as unknown as HTMLElement,
    { ...(options.worker ? { mode: 'worker' } : {}), cellScale: options.cellScale } as ConstructorParameters<typeof XlsxViewer>[1],
  );
  const engine = viewer as unknown as Engine;
  for (const method of STUBBED) engine[method] = vi.fn();
  for (const [key, value] of [['clientWidth', 800], ['clientHeight', 600]] as const) {
    Object.defineProperty(engine.canvasArea, key, { configurable: true, value });
  }
  const frames: Frame[] = [];
  const open = (hostRowHeight = options.hostRowHeight) => {
    const workbook = createWorkbook(sources, frames, hostRowHeight, options.gate);
    engine.wb = workbook;
    engine.prepareWorkbook(workbook);
  };
  open();
  const frame = async (): Promise<Frame> => {
    await engine.renderCurrentSheet();
    return frames[frames.length - 1];
  };
  return { viewer, engine, frame, open, lastFrame: () => frames[frames.length - 1] };
}

/** Commit of a column-A drag through the viewer's own record + refit path. */
function resizeColumnA(engine: Engine, width: number): void {
  const ws = engine.currentWorksheet!;
  ws.colWidths[1] = pxToColWidth(width);
  GridGeometry.invalidate(ws);
  engine.recordSizeOverride('col', 1);
  engine.refitAutoRowsAfterColumnResize();
}

describe('XlsxViewer prepared initial anchor size lifecycle (#1713)', () => {
  it('main: keeps the host-measured initial size through resize, navigation and reacquisition', async () => {
    const sources = [sheet('A'), sheet('B')];
    const a = mount(sources, { hostRowHeight: 40 });
    const b = mount(sources, { hostRowHeight: 70 });
    await a.engine.showSheet(0);
    await b.engine.showSheet(0);
    let frame = await a.frame();
    expectSize(frame, ONE, 320, 100);
    expectSize(frame, TWO, 320, 100);
    expectSize(await b.frame(), ONE, 320, 130);

    resizeColumnA(a.engine, 250);
    frame = await a.frame();
    // The initial measurement, neither authored 60 px nor the 30 px refit.
    expectSize(frame, ONE, 320, 100);
    expectSize(frame, TWO, 410, 90);

    await a.engine.showSheet(1);
    await a.engine.showSheet(0);
    frame = await a.frame();
    expectSize(frame, ONE, 320, 100);
    expectSize(frame, TWO, 410, 90);
    expectSize(await b.frame(), ONE, 320, 130);
    expect(sources[0].colWidths[1]).toBe(pxToColWidth(160));
    expect(sources[0].rowHeights[1]).toBe(pxToRowHeight(60));
  });

  it('worker: the projection carries the reference before any size edit and after reacquisition', async () => {
    const { engine, frame } = mount([sheet('A'), sheet('B')], { worker: true });
    await engine.showSheet(0);
    let current = await frame();
    expectSize(current, ONE, 320, 120);
    expectSize(current, TWO, 320, 120);
    resizeColumnA(engine, 250);
    current = await frame();
    expectSize(current, ONE, 320, 120);
    expectSize(current, TWO, 410, 120);
    await engine.showSheet(1);
    await engine.showSheet(0);
    current = await frame();
    expectSize(current, ONE, 320, 120);
    expectSize(current, TWO, 410, 120);
  });

  it('reload, a superseded acquisition and destroy release references and never install stale ones', async () => {
    let openGate!: () => void;
    const gate = new Promise<void>((resolve) => { openGate = resolve; });
    const { viewer, engine, frame, open } = mount([sheet('A'), sheet('B')], { hostRowHeight: 40, gate });
    await engine.showSheet(0);
    resizeColumnA(engine, 250);
    const previous = engine.currentWorksheet!;
    expectSize(paint(previous), ONE, 320, 100);

    // The old workbook's acquisition is still pending when a reload replaces it.
    const superseded = engine.showSheet(1);
    open(50);
    await engine.showSheet(0);
    openGate();
    await superseded;
    // A released reference stays non-renderable; it never revives edited markers.
    expect(paint(previous).has(ONE)).toBe(false);
    let current = await frame();
    expectSize(current, ONE, 320, 110);
    expectSize(current, TWO, 320, 110);
    resizeColumnA(engine, 250);
    current = await frame();
    expectSize(current, ONE, 320, 110);
    expectSize(current, TWO, 410, 90);

    const displayed = engine.currentWorksheet!;
    viewer.destroy();
    expect(paint(displayed).has(ONE)).toBe(false);
  });
});

// Owner regressions after independent B2 review: real view lifecycle contracts,
// not the test doubles' representation or pixel/OOXML unit conversion.
describe('prepared initial anchor lifecycle boundary regressions (#1713)', () => {
  it.each([false, true])('preserves first fractional-zoom markers through resize, zoom and reacquisition (worker=%s)', async (worker) => {
    const source = sheet('A');
    source.colWidths = Object.freeze({ ...source.colWidths, 1: 20.125, 2: 20.125 });
    const { viewer, engine, frame } = mount([source, sheet('B')], { worker, cellScale: .85 });
    await engine.showSheet(0);
    let current = await frame();
    expectSize(current, TWO, 274, 102);
    expectSize(current, ONE, 274, 102);
    resizeColumnA(engine, 250);
    current = await frame();
    expectSize(current, ONE, 274, 102);
    expect(current.get(TWO)![2]).toBeGreaterThan(274);
    viewer.setScale(1.1);
    expectSize(await frame(), ONE, 274 / .85 * 1.1, 102 / .85 * 1.1);
    await engine.showSheet(1);
    await engine.showSheet(0);
    expectSize(await frame(), ONE, 274 / .85 * 1.1, 102 / .85 * 1.1);
    viewer.destroy();
  });

  it.each([false, true])('commits a late parser placeholder even when the provisional model had retained anchors (worker=%s)', async (worker) => {
    const source = sheet('A');
    const { viewer, engine, frame } = mount([source], { worker });
    let complete!: (value: Worksheet) => void;
    const completion = new Promise<Worksheet>((resolve) => { complete = resolve; });
    const workbook = engine.wb as Record<string, unknown>;
    workbook.acquireWorksheetPreviewLease = async () => ({
      worksheet: source, release: vi.fn(), partial: true, completion,
      waitForRows: vi.fn(async () => undefined), releaseFirstPaint: vi.fn(),
    });
    engine.scheduleRender = vi.fn();
    const onError = vi.fn();
    (engine.opts as { onError: typeof onError }).onError = onError;
    await engine.showSheet(0);
    complete({ ...sheet('A'), images: [], shapeGroups: [], parseError: 'invalid later cell' });
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.currentWorksheet?.parseError).toBe('invalid later cell');
    expect((await frame()).has(ONE)).toBe(false);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });

  it('drops an old acquisition before it can overwrite the new workbook displayed projection', async () => {
    const source = sheet('A');
    const { viewer, engine, open } = mount([source], { hostRowHeight: 40 });
    await engine.showSheet(0);
    let finishOld!: (value: { worksheet: Worksheet; release: ReturnType<typeof vi.fn> }) => void;
    const pending = new Promise<{ worksheet: Worksheet; release: ReturnType<typeof vi.fn> }>((resolve) => { finishOld = resolve; });
    let acquired!: () => void;
    const acquisitionStarted = new Promise<void>((resolve) => { acquired = resolve; });
    (engine.wb as Record<string, unknown>).acquireWorksheetLease = () => { acquired(); return pending; };
    const stale = engine.showSheet(0);
    // Replace the workbook only after the old lease was actually requested.
    await acquisitionStarted;
    open(70);
    await engine.showSheet(0);
    const displayed = engine.currentWorksheet!;
    const expectedRows = displayed.rows;
    const oldRows: Worksheet['rows'] = [{ index: 1, height: null, cells: [] }];
    const release = vi.fn();
    finishOld({ worksheet: { ...source, images: [], shapeGroups: [], rows: oldRows }, release });
    await stale;
    expect(engine.currentWorksheet).toBe(displayed);
    expect(displayed.rows).toBe(expectedRows);
    expect(displayed.rowHeights[1]).toBe(pxToRowHeight(70));
    expect(release).toHaveBeenCalledOnce();
    viewer.destroy();
  });

  // B3: after a reload the superseded partial preview's row wait resolves; it
  // must not prepare heights, write a fallback reason or commit, and releases
  // its lease once (production showSheet path; hook spy proves no resume).
  it('drops a superseded partial preview when its row wait resolves after reload', async () => {
    const source = sheet('A');
    const { viewer, engine, open } = mount([source], { hostRowHeight: 40 });
    await engine.showSheet(0);
    const old = engine.wb as Record<PropertyKey, unknown>;
    const hook = vi.fn();
    old[prepareXlsxViewerRowHeights] = hook;
    let rowsReady!: () => void;
    const rows = new Promise<void>((resolve) => { rowsReady = resolve; });
    let waiting!: () => void;
    const waitStarted = new Promise<void>((resolve) => { waiting = resolve; });
    const release = vi.fn();
    old.acquireWorksheetPreviewLease = async () => ({
      worksheet: source, release, partial: true, completion: new Promise<Worksheet>(() => {}),
      waitForRows: () => { waiting(); return rows; }, releaseFirstPaint: vi.fn(),
    });
    const stale = engine.showSheet(0);
    await waitStarted;
    open(70);
    await engine.showSheet(0);
    const displayed = engine.currentWorksheet!;
    rowsReady();
    await stale;
    expect(hook).not.toHaveBeenCalled();
    expect(engine.previewFallbackReason).toBeNull();
    expect(engine.currentWorksheet).toBe(displayed);
    expect(displayed.rowHeights[1]).toBe(pxToRowHeight(70));
    expect(release).toHaveBeenCalledOnce();
    viewer.destroy();
  });

  // Acceptance: a partial preview waits for rows through the anchor's marker
  // rows below the viewport, then terminal completion after a manual edit
  // reuses the initial reference (never recaptured from the edited model).
  it('waits for anchor rows below the viewport and keeps the initial size through completion', async () => {
    const source = sheet('A');
    // Zero-based marker rows 0..15: the pictures end below the 600 px viewport.
    source.images = Object.freeze(source.images!.map((image) => Object.freeze({ ...image, toRow: 15 }))) as unknown as Worksheet['images'];
    const { viewer, engine, frame, lastFrame } = mount([source], { hostRowHeight: 40 });
    let complete!: (value: Worksheet) => void;
    const completion = new Promise<Worksheet>((resolve) => { complete = resolve; });
    const waitForRows = vi.fn(async (_rows: number) => undefined);
    (engine.wb as Record<string, unknown>).acquireWorksheetPreviewLease = async () => ({
      worksheet: source, release: vi.fn(), partial: true, completion, waitForRows, releaseFirstPaint: vi.fn(),
    });
    engine.scheduleRender = vi.fn();
    const onError = vi.fn();
    (engine.opts as { onError: typeof onError }).onError = onError;
    await engine.showSheet(0);
    expect(Math.max(...waitForRows.mock.calls.map(([rows]) => rows))).toBeGreaterThanOrEqual(16);
    // Host row 1 is 40 px, rows 2..15 authored 60 px.
    expectSize(lastFrame(), ONE, 320, 880);
    resizeColumnA(engine, 250);
    complete(source);
    await Promise.resolve();
    await Promise.resolve();
    const current = await frame();
    expectSize(current, ONE, 320, 880);
    expectSize(current, TWO, 410, 870);
    expect(onError).not.toHaveBeenCalled();
    viewer.destroy();
  });
});

// CF dependencies of rows that determine a drawing's initial height can live
// outside the painted viewport. Only such sheets wait for complete materialization.
describe('prepared anchor conditional-format completion boundary (#1713)', () => {
  function dependentSheet(): Worksheet {
    const source = sheet('A');
    source.images = Object.freeze(source.images!.map(image => Object.freeze({ ...image, toRow: 50 }))) as unknown as Worksheet['images'];
    source.conditionalFormats = [{
      sqref: [{ top: 30, left: 1, bottom: 500, right: 1 }],
      rules: [{ type: 'cellIs', operator: 'greaterThan', formulas: ['A500'], dxfId: 0, priority: 1 }],
    }];
    return source;
  }

  it('waits for offscreen CF dependencies before capturing the prepared height', async () => {
    const source = dependentSheet();
    const { viewer, engine, frame } = mount([source], { hostRowHeight: 40 });
    let complete!: (value: Worksheet) => void;
    const completion = new Promise<Worksheet>(resolve => { complete = resolve; });
    let measured!: () => void;
    const measuredFirst = new Promise<void>(resolve => { measured = resolve; });
    const old = engine.wb as Record<PropertyKey, unknown>;
    old[prepareXlsxViewerRowHeights] = (view: Worksheet) => {
      const fullHeight = view.rows.some(row => row.index === 500) ? 80 : 40;
      view.rowHeights[1] = pxToRowHeight(view.colWidths[1] > pxToColWidth(160) ? 30 : fullHeight);
      GridGeometry.invalidate(view);
      measured();
    };
    old.acquireWorksheetPreviewLease = async () => ({
      worksheet: source, release: vi.fn(), partial: true, completion,
      waitForRows: vi.fn(async () => undefined), releaseFirstPaint: vi.fn(),
    });
    const shown = engine.showSheet(0);
    await measuredFirst;
    expect(engine.previewFallbackReason).toBe('drawing-dependency');
    complete({ ...source, rows: [{ index: 500, height: null, cells: [] }] });
    await shown;
    expect(engine.currentWorksheet!.rowHeights[1]).toBe(pxToRowHeight(80));
    const initial = await frame();
    expect(initial.get(ONE)).toEqual(initial.get(TWO));
    resizeColumnA(engine, 250);
    const resized = await frame();
    expect(resized.get(ONE)).toEqual(initial.get(ONE));
    expect(resized.get(TWO)![2]).toBeGreaterThan(initial.get(TWO)![2]);
    expect(resized.get(TWO)![3]).toBeLessThan(initial.get(TWO)![3]);
    viewer.destroy();
  });

  it('drops a stale CF completion before preparing or mutating the reloaded display', async () => {
    const source = dependentSheet();
    const { viewer, engine, open } = mount([source], { hostRowHeight: 40 });
    let complete!: (value: Worksheet) => void;
    const completion = new Promise<Worksheet>(resolve => { complete = resolve; });
    let measured!: () => void;
    const measuredFirst = new Promise<void>(resolve => { measured = resolve; });
    const old = engine.wb as Record<PropertyKey, unknown>;
    const hook = vi.fn((view: Worksheet) => {
      view.rowHeights[1] = pxToRowHeight(40);
      GridGeometry.invalidate(view);
      measured();
    });
    old[prepareXlsxViewerRowHeights] = hook;
    const release = vi.fn();
    old.acquireWorksheetPreviewLease = async () => ({
      worksheet: source, release, partial: true, completion,
      waitForRows: vi.fn(async () => undefined), releaseFirstPaint: vi.fn(),
    });
    const stale = engine.showSheet(0);
    await measuredFirst;
    expect(engine.previewFallbackReason).toBe('drawing-dependency');
    const preparations = hook.mock.calls.length;
    open(70);
    await engine.showSheet(0);
    const displayed = engine.currentWorksheet!;
    const displayedRows = displayed.rows;
    complete({ ...source, rows: [{ index: 500, height: null, cells: [] }] });
    await stale;
    expect(hook).toHaveBeenCalledTimes(preparations);
    expect(engine.currentWorksheet).toBe(displayed);
    expect(displayed.rows).toBe(displayedRows);
    expect(displayed.rowHeights[1]).toBe(pxToRowHeight(70));
    expect(engine.previewFallbackReason).toBeNull();
    expect(release).toHaveBeenCalledOnce();
    viewer.destroy();
  });
});

describe('superseded same-sheet acquisition ownership (#1713)', () => {
  it.each(['newer-navigation', 'destroy'] as const)('drops old work after %s without a model write or lease resurrection', async action => {
    const source = sheet('A');
    const { viewer, engine, frame } = mount([source], { hostRowHeight: 40 });
    await engine.showSheet(0);
    const workbook = engine.wb as { acquireWorksheetLease: (index: number) => Promise<{ worksheet: Worksheet; release: ReturnType<typeof vi.fn> }> };
    const acquire = workbook.acquireWorksheetLease;
    let complete!: (value: { worksheet: Worksheet; release: ReturnType<typeof vi.fn> }) => void;
    const pending = new Promise<{ worksheet: Worksheet; release: ReturnType<typeof vi.fn> }>(resolve => { complete = resolve; });
    let started!: () => void;
    const acquisitionStarted = new Promise<void>(resolve => { started = resolve; });
    let requested = false;
    workbook.acquireWorksheetLease = index => {
      if (requested) return acquire(index);
      requested = true;
      started();
      return pending;
    };
    const stale = engine.showSheet(0);
    await acquisitionStarted;
    if (action === 'destroy') viewer.destroy();
    else await engine.showSheet(0);
    const displayed = engine.currentWorksheet;
    const displayedRows = displayed?.rows;
    const release = vi.fn();
    complete({ worksheet: { ...source, images: [], rows: [{ index: 8, height: null, cells: [] }] }, release });
    await stale;
    expect(release).toHaveBeenCalledOnce();
    expect(engine.currentWorksheet).toBe(displayed);
    if (action === 'destroy') expect(displayed).toBeNull();
    else {
      expect(displayed!.rows).toBe(displayedRows);
      expectSize(await frame(), ONE, 320, 100);
      viewer.destroy();
    }
    expect(source.rows).toEqual([]);
    expect(source.rowHeights[1]).toBe(pxToRowHeight(60));
  });
});

describe('complete parser placeholder with an existing displayed projection (#1713)', () => {
  it.each([false, true])('replaces the old graph and preserves its reference only for valid reacquisition (worker=%s)', async worker => {
    const source = sheet('A');
    const { viewer, engine, frame } = mount([source], { worker, hostRowHeight: 40 });
    await engine.showSheet(0);
    resizeColumnA(engine, 250);
    const workbook = engine.wb as Record<string, unknown>;
    workbook.acquireWorksheetLease = async () => ({
      worksheet: { ...source, images: [], shapeGroups: [], parseError: 'invalid complete model' }, release: vi.fn(),
    });
    await engine.showSheet(0);
    expect(engine.currentWorksheet?.parseError).toBe('invalid complete model');
    expect((await frame()).has(ONE)).toBe(false);
    workbook.acquireWorksheetLease = async () => ({ worksheet: source, release: vi.fn() });
    await engine.showSheet(0);
    expect(engine.currentWorksheet?.parseError).toBeUndefined();
    expectSize(await frame(), ONE, 320, 100);
    viewer.destroy();
  });
});

it.each([false, true])('captures an unedited initial placeholder after valid same-sheet reacquisition (worker=%s)', async worker => {
  const valid = sheet('A');
  const placeholder = { ...valid, images: [], shapeGroups: [], parseError: 'degraded initial model' };
  const { viewer, engine, frame } = mount([valid], { worker });
  (engine.wb as Record<string, unknown>).acquireWorksheetLease = async () => ({ worksheet: placeholder, release: vi.fn() });
  try {
    await engine.showSheet(0);
    expect(engine.currentWorksheet?.parseError).toBe('degraded initial model');
    (engine.wb as Record<string, unknown>).acquireWorksheetLease = async () => ({ worksheet: valid, release: vi.fn() });
    await engine.showSheet(0);
    expect(engine.currentWorksheet?.parseError).toBeUndefined();
    expectSize(await frame(), ONE, 320, 120);
  } finally { viewer.destroy(); }
});

it('refuses to capture a replacement anchor baseline after manual placeholder edits', async () => {
  const valid = sheet('A');
  const placeholder = { ...valid, images: [], shapeGroups: [], parseError: 'degraded initial model' };
  const { viewer, engine } = mount([placeholder]);
  try {
    await engine.showSheet(0);
    resizeColumnA(engine, 250);
    (engine.wb as Record<string, unknown>).acquireWorksheetLease = async () => ({ worksheet: valid, release: vi.fn() });
    await expect(engine.showSheet(0)).rejects.toThrow('after view edits');
    expect(engine.currentWorksheet?.parseError).toBe('degraded initial model');
  } finally { viewer.destroy(); }
});

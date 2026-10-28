import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer, XlsxSheetViewer } from './viewer.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';
import type { Worksheet } from './types.js';
import { extractViewerRenderContext } from './worker-protocol.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A fake ImageBitmap: records close() and carries dimensions for the canvas
 *  resize path. */
function fakeBitmap(width = 800, height = 600): { width: number; height: number; close: ReturnType<typeof vi.fn> } {
  return { width, height, close: vi.fn() };
}

/** Minimal worksheet the render path can walk (empty grid, default sizes). */
function emptyWorksheet(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 64,
    defaultRowHeight: 20,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    charts: [],
    images: [],
    shapeGroups: [],
  } as unknown as Worksheet;
}

/**
 * Build a worker-mode viewer with its private state wired so `renderCurrentSheet`
 * reaches the worker branch: a nonzero canvasArea, a current worksheet, and a
 * fake workbook whose `renderViewportToBitmap` returns a caller-controlled
 * deferred promise. Returns handles to drive frames and inspect the canvas.
 */
function buildWorker(options: {
  scroll?: 'native' | 'logical';
  onViewportChange?: (offset: { x: number; y: number }) => void;
} = {}) {
  const doc = installDom();
  const frames = new Map<number, FrameRequestCallback>();
  let nextFrame = 0;
  if (options.scroll) Object.assign(doc.defaultView, {
    requestAnimationFrame: (callback: FrameRequestCallback) => {
      frames.set(++nextFrame, callback);
      return nextFrame;
    },
    cancelAnimationFrame: (id: number) => { frames.delete(id); },
  });
  const container = makeContainer();
  const v = options.scroll === 'logical'
    ? new XlsxSheetViewer(doc.createElement('canvas') as unknown as HTMLCanvasElement,
      { mode: 'worker', showScrollbars: false, onViewportChange: options.onViewportChange })
    : new XlsxViewer(container as unknown as HTMLElement,
      { mode: 'worker', onViewportChange: options.onViewportChange });

  // Each renderViewportToBitmap call parks its {resolve, bitmap} so the test can
  // resolve frames out of order.
  const inflight: Array<{ resolve: (b: unknown) => void; promise: Promise<unknown> }> = [];
  const renderViewportToBitmap = vi.fn(() => {
    let resolve!: (b: unknown) => void;
    const promise = new Promise<unknown>((r) => {
      resolve = r;
    });
    inflight.push({ resolve, promise });
    return promise;
  });

  const fakeWb = {
    renderViewportToBitmap,
    sheetNames: ['Sheet1'],
    sheetCount: 1,
    isHidden: () => false,
    destroy: vi.fn(),
  };

  // Inject private state (the vitest env is node; there is no real load()).
  const engine = options.scroll === 'logical'
    ? (v as unknown as { engine: unknown }).engine : v;
  const priv = engine as unknown as {
    wb: unknown;
    currentWorksheet: Worksheet;
    currentSheet: number;
    canvasArea: FakeEl;
    canvas: FakeEl;
    scrollHost: FakeEl;
    renderCurrentSheet: () => Promise<void>;
    updateSpacerSize(worksheet: Worksheet): void;
  };
  priv.wb = fakeWb;
  const currentWorksheet = emptyWorksheet();
  priv.currentWorksheet = currentWorksheet;
  priv.currentSheet = 0;
  priv.canvasArea.clientWidth = 800;
  priv.canvasArea.clientHeight = 600;

  return {
    v,
    inflight,
    renderViewportToBitmap,
    canvas: priv.canvas,
    currentWorksheet,
    render: () => priv.renderCurrentSheet(),
    scrollHost: priv.scrollHost,
    flushFrame: (time = 0) => {
      const callbacks = [...frames.values()];
      frames.clear();
      callbacks.forEach(callback => callback(time));
    },
    queuedFrames: () => frames.size,
    updateSpacerSize: () => priv.updateSpacerSize(currentWorksheet),
  };
}

/**
 * Worker-mode stale-frame dropping (improvement plan C4, commit 3): worker
 * bitmap renders overlap — a slow bitmap for an old scroll position can resolve
 * after a newer render was requested. The stale bitmap must be closed (GPU
 * memory freed) and NOT painted over the fresher frame. Single-canvas analogue
 * of the pptx scroll-viewer render epoch.
 */
describe('XlsxViewer worker-mode stale-frame drop (C4 commit 3)', () => {
  it.each(['native', 'logical'] as const)(
    'presents intermediate worker frames during continuous %s scrolling with aligned cell geometry',
    async (scroll) => {
      const { v, inflight, canvas, currentWorksheet, scrollHost, flushFrame, queuedFrames, updateSpacerSize } =
        buildWorker({ scroll });
      currentWorksheet.defaultColWidth = 8.43;
      currentWorksheet.defaultRowHeight = 15; // 20 CSS px at scale 1.
      currentWorksheet.rightToLeft = scroll === 'logical';
      currentWorksheet.freezeRows = 1;
      currentWorksheet.freezeCols = 1;
      scrollHost.clientWidth = 800;
      scrollHost.clientHeight = 600;
      scrollHost.scrollWidth = 10000;
      scrollHost.scrollHeight = 10000;
      updateSpacerSize();
      v.setSelection('C10:C12');
      const settle = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
      const scrollTo = (y: number) => {
        if (scroll === 'native') {
          scrollHost.scrollTop = y;
          scrollHost.scrollLeft = y / 2;
          scrollHost.dispatch('scroll');
        } else {
          vi.stubGlobal('WheelEvent', { DOM_DELTA_LINE: 1, DOM_DELTA_PAGE: 2 });
          scrollHost.dispatch('wheel', { deltaY: 20, deltaX: 10, deltaMode: 0,
            ctrlKey: false, metaKey: false, shiftKey: false, preventDefault() {} });
        }
      };
      flushFrame();
      inflight[0].resolve(fakeBitmap());
      await settle();
      const initial = v.getCellViewportRect('C10');
      expect(initial).not.toBeNull();
      for (let cycle = 0; cycle < 3; cycle++) {
        const start = cycle * 80 + 20;
        scrollTo(start);
        flushFrame();
        for (let y = start + 20; y <= start + 60; y += 20) scrollTo(y);
        // One active worker request; intervening viewports replace one pending callback.
        expect(inflight).toHaveLength(cycle + 2);
        const bitmap = fakeBitmap();
        inflight[cycle + 1].resolve(bitmap);
        await settle();
        expect(canvas._bitmapCtx?.lastBitmap).toBe(bitmap);
        expect(bitmap.close).not.toHaveBeenCalled();
        expect(v.getViewportOffset().y).toBe(start + 60);
        const rect = v.getCellViewportRect('C10');
        expect(rect?.y).toBe((initial?.y ?? 0) - start);
        expect(rect?.x).toBe((initial?.x ?? 0) + (scroll === 'logical' ? 1 : -1) * start / 2);
        expect(v.getCellViewportRect('A1')?.y).toBe(22); // frozen row stays fixed.
        // Selection geometry follows the displayed frame, rather than the newer scrollbar.
        const cutout = (canvas.parentElement as unknown as FakeEl)
          .querySelector('[data-xlsx-active-cell-cutout]');
        if (cycle === 0) expect(cutout?.getAttribute('y')).toBe(String(rect?.y));
        expect(queuedFrames()).toBe(1);
      }
      flushFrame();
      const final = fakeBitmap();
      inflight.at(-1)?.resolve(final);
      await settle();
      expect(canvas._bitmapCtx?.lastBitmap).toBe(final);
      expect(v.getCellViewportRect('C10')?.y).toBe((initial?.y ?? 0) - 240);
      // A header click must target the row/column in the displayed frame,
      // even while the scrollbar has advanced three more rows.
      scrollTo(260);
      flushFrame();
      for (const y of [280, 300, 320]) scrollTo(y);
      inflight.at(-1)?.resolve(fakeBitmap());
      await settle();
      const shown = v.getCellViewportRect('G30');
      if (!shown) throw new Error('Expected displayed cell geometry');
      const pointer = (clientX: number, clientY: number) => ({ button: 0, pointerId: 1,
        pointerType: 'mouse', clientX, clientY, shiftKey: false, ctrlKey: false,
        metaKey: false, preventDefault() {} });
      const extraAt = inflight.length;
      const row = pointer(scroll === 'logical' ? 795 : 5, shown.y + shown.height / 2);
      scrollHost.dispatch('pointerdown', row);
      scrollHost.dispatch('pointerup', row);
      expect(v.selectionState?.areas[0]).toEqual({ kind: 'rows', firstRow: 30, lastRow: 30 });
      const col = pointer(scroll === 'logical' ? shown.x + 8 : shown.x + shown.width - 8, 5);
      scrollHost.dispatch('pointerdown', col);
      scrollHost.dispatch('pointerup', col);
      expect(v.selectionState?.areas[0]).toEqual({ kind: 'columns', firstColumn: 7, lastColumn: 7 });
      v.destroy();
      for (const request of inflight.slice(extraAt)) request.resolve(fakeBitmap());
      await settle();
    },
  );

  it.each(['native', 'logical'] as const)(
    'hit-tests an ordinary %s drag against the displayed frame while a scroll frame is pending',
    async (scroll) => {
      const { v, inflight, currentWorksheet, scrollHost, flushFrame, updateSpacerSize, render } =
        buildWorker({ scroll });
      currentWorksheet.defaultColWidth = 8.43;
      currentWorksheet.defaultRowHeight = 15; // 20 CSS px at scale 1.
      scrollHost.clientWidth = 800;
      scrollHost.clientHeight = 600;
      scrollHost.scrollWidth = 10000;
      scrollHost.scrollHeight = 10000;
      updateSpacerSize();
      const settle = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
      const first = render();
      inflight[0].resolve(fakeBitmap());
      await first;
      await settle();
      const c5 = v.getCellViewportRect('C5');
      const d5 = v.getCellViewportRect('D5');
      if (!c5 || !d5) throw new Error('Expected displayed geometry');
      // Request a 100px vertical and one-column horizontal scroll; its frame stays pending.
      vi.stubGlobal('WheelEvent', { DOM_DELTA_LINE: 1, DOM_DELTA_PAGE: 2 });
      for (let step = 0; step < 5; step++) {
        if (scroll === 'native') {
          scrollHost.scrollTop = (step + 1) * 20;
          scrollHost.scrollLeft = (step + 1) * (d5.x - c5.x) / 5;
          scrollHost.dispatch('scroll');
        } else {
          scrollHost.dispatch('wheel', { deltaY: 20, deltaX: (d5.x - c5.x) / 5, deltaMode: 0,
            ctrlKey: false, metaKey: false, shiftKey: false, preventDefault() {} });
        }
      }
      flushFrame();
      expect(v.getViewportOffset().y).toBe(100);
      expect(v.getCellViewportRect('C5')?.y).toBe(c5.y); // still the displayed frame
      const pointer = (clientX: number, clientY: number) => ({ button: 0, pointerId: 1,
        pointerType: 'mouse', clientX, clientY, shiftKey: false, ctrlKey: false,
        metaKey: false, preventDefault() {} });
      const drag = (x: number, y: number) => {
        scrollHost.dispatch('pointerdown', pointer(x, y));
        scrollHost.dispatch('pointermove', pointer(x + 1, y));
        scrollHost.dispatch('pointerup', pointer(x + 1, y));
      };
      const midX = c5.x + c5.width / 2;
      const midY = c5.y + c5.height / 2; // clear of header resize borders
      drag(midX, midY);
      expect(v.selectionState?.areas[0]).toMatchObject(
        { kind: 'cells', top: 5, bottom: 5, left: 3, right: 3 });
      drag(5, midY);
      expect(v.selectionState?.areas[0]).toEqual({ kind: 'rows', firstRow: 5, lastRow: 5 });
      drag(midX, 5);
      expect(v.selectionState?.areas[0]).toEqual({ kind: 'columns', firstColumn: 3, lastColumn: 3 });
      for (const request of inflight) request.resolve(fakeBitmap());
      await settle();
      v.destroy();
    },
  );

  it.each(['native', 'logical'] as const)(
    'keeps extending a %s drag selection while autoscroll outruns slow worker frames',
    async (scroll) => {
      const { v, inflight, currentWorksheet, scrollHost, flushFrame, updateSpacerSize, render } =
        buildWorker({ scroll });
      currentWorksheet.defaultColWidth = 8.43;
      currentWorksheet.defaultRowHeight = 15; // 20 CSS px at scale 1.
      scrollHost.clientWidth = 800;
      scrollHost.clientHeight = 600;
      scrollHost.scrollWidth = 10000;
      scrollHost.scrollHeight = 10000;
      updateSpacerSize();
      const settle = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
      const first = render();
      inflight[0].resolve(fakeBitmap());
      await first;
      await settle();
      const pointer = (clientX: number, clientY: number) => ({ button: 0, pointerId: 1,
        pointerType: 'mouse', clientX, clientY, shiftKey: false, ctrlKey: false,
        metaKey: false, preventDefault() {} });
      const start = v.getCellViewportRect('C5');
      if (!start) throw new Error('Expected C5 geometry');
      scrollHost.dispatch('pointerdown', pointer(start.x + 4, start.y + 4));
      // Drag below the viewport; autoscroll ticks advance the logical viewport.
      scrollHost.dispatch('pointermove', pointer(start.x + 4, 660));
      let time = 0;
      for (let tick = 1; tick <= 25; tick++) {
        time += 16;
        flushFrame(time);
        // Each worker response takes about five input frames.
        if (tick % 5 === 0) {
          for (const request of inflight) request.resolve(fakeBitmap());
          await settle();
        }
      }
      const scrolled = v.getViewportOffset().y;
      expect(scrolled).toBeGreaterThan(200);
      // The clamped pointer sits at the last visible data row of the logical viewport.
      const expectedRow = Math.floor((599 - 22 + scrolled) / 20) + 1;
      const area = v.selectionState?.areas[0];
      expect(area).toMatchObject({ kind: 'cells', top: 5 });
      expect((area as { bottom: number }).bottom).toBe(expectedRow);
      scrollHost.dispatch('pointerup', pointer(start.x + 4, 660));
      for (const request of inflight) request.resolve(fakeBitmap());
      await settle();
      v.destroy();
    },
  );

  it.each(['native', 'logical'] as const)(
    'notifies %s viewport consumers when a lagging worker frame moves displayed cell geometry',
    async (scroll) => {
      let annotationY: number | null = null;
      let calls = 0;
      let viewer: { getCellViewportRect(cell: string): { y: number } | null } | null = null;
      const { v, inflight, currentWorksheet, scrollHost, flushFrame, updateSpacerSize, render } =
        buildWorker({ scroll, onViewportChange: () => {
          calls++;
          annotationY = viewer?.getCellViewportRect('C10')?.y ?? null;
        } });
      viewer = v;
      currentWorksheet.defaultColWidth = 8.43;
      currentWorksheet.defaultRowHeight = 15; // 20 CSS px at scale 1.
      scrollHost.clientWidth = 800;
      scrollHost.clientHeight = 600;
      scrollHost.scrollWidth = 10000;
      scrollHost.scrollHeight = 10000;
      updateSpacerSize();
      const settle = async () => { for (let i = 0; i < 12; i++) await Promise.resolve(); };
      const first = render();
      inflight[0].resolve(fakeBitmap());
      await first;
      await settle();
      expect(v.getCellViewportRect('C10')?.y).toBe(202);
      vi.stubGlobal('WheelEvent', { DOM_DELTA_LINE: 1, DOM_DELTA_PAGE: 2 });
      for (let step = 0; step < 5; step++) {
        if (scroll === 'native') {
          scrollHost.scrollTop = (step + 1) * 20;
          scrollHost.dispatch('scroll');
        } else {
          scrollHost.dispatch('wheel', { deltaY: 20, deltaX: 0, deltaMode: 0,
            ctrlKey: false, metaKey: false, shiftKey: false, preventDefault() {} });
        }
      }
      expect(v.getViewportOffset().y).toBe(100);
      const scrollCalls = calls;
      flushFrame();
      inflight[1].resolve(fakeBitmap());
      await settle();
      expect(v.getCellViewportRect('C10')?.y).toBe(102);
      expect(annotationY).toBe(102);
      // One coalesced notification for the landed frame, none for idle frames.
      expect(calls).toBe(scrollCalls + 1);
      flushFrame();
      await settle();
      expect(calls).toBe(scrollCalls + 1);
      v.destroy();
    },
  );

  it('sends the Window-measured MDW as internal layout authority for worker paint', async () => {
    const { currentWorksheet, inflight, renderViewportToBitmap, render } = buildWorker();
    currentWorksheet.defaultFontFamily = 'Realm Divergence';
    currentWorksheet.defaultFontSize = 11;
    vi.stubGlobal('OffscreenCanvas', class {
      getContext() {
        return { font: '', measureText: () => ({ width: 13.2 }) };
      }
    });

    const pending = render();
    const call = renderViewportToBitmap.mock.calls[0] as unknown[] | undefined;
    const options = call?.[2];
    expect(extractViewerRenderContext(options as never).layoutMetrics).toEqual({
      maximumDigitWidth: 13,
    });
    inflight[0].resolve(fakeBitmap());
    await pending;
  });

  it('drops the older in-flight bitmap and paints only the latest', async () => {
    const { inflight, canvas, render } = buildWorker();

    // Frame A and frame B both dispatched before either resolves (a scroll
    // burst): two renders in flight.
    const pA = render();
    const pB = render();
    expect(inflight.length).toBe(2);

    const bmpA = fakeBitmap(800, 600);
    const bmpB = fakeBitmap(800, 600);

    // Resolve the NEWER frame (B) first: it is current, so it paints.
    inflight[1].resolve(bmpB);
    await pB;
    expect(bmpB.close).not.toHaveBeenCalled();
    expect(canvas._bitmapCtx?.lastBitmap).toBe(bmpB);

    // Now the OLDER frame (A) resolves late: it is stale (B superseded it), so
    // it must be closed and never painted.
    inflight[0].resolve(bmpA);
    await pA;
    expect(bmpA.close).toHaveBeenCalledTimes(1);
    // The canvas still shows B — the stale A was not transferred.
    expect(canvas._bitmapCtx?.lastBitmap).toBe(bmpB);
  });

  it('discards a pre-zoom bitmap even when scrolling can finish intermediate frames', async () => {
    const { v, inflight, canvas, render } = buildWorker();
    const pending = render();
    v.setScale(1.5);
    const stale = fakeBitmap();
    inflight[0].resolve(stale);
    await pending;
    expect(stale.close).toHaveBeenCalledOnce();
    expect(canvas._bitmapCtx?.lastBitmap).toBeNull();
    const zoomed = fakeBitmap();
    inflight[1].resolve(zoomed);
    for (let i = 0; i < 8; i++) await Promise.resolve();
    expect(canvas._bitmapCtx?.lastBitmap).toBe(zoomed);
    v.destroy();
  });

  it('paints a single in-flight frame normally (no false stale-drop)', async () => {
    const { inflight, canvas, render } = buildWorker();
    const p = render();
    const bmp = fakeBitmap(800, 600);
    inflight[0].resolve(bmp);
    await p;
    expect(bmp.close).not.toHaveBeenCalled();
    expect(canvas._bitmapCtx?.lastBitmap).toBe(bmp);
  });

  it('drops a bitmap that resolves after destroy() (no paint on a torn-down canvas)', async () => {
    const { v, inflight, canvas, render } = buildWorker();
    const p = render();
    v.destroy();
    const bmp = fakeBitmap(800, 600);
    inflight[0].resolve(bmp);
    await p;
    // destroy() advanced the generation ⇒ this frame is stale ⇒ closed, unpainted.
    expect(bmp.close).toHaveBeenCalledTimes(1);
    expect(canvas._bitmapCtx?.lastBitmap).toBeNull();
  });

  it('does not re-assign canvas width/height when a same-size frame paints (C4 commit 2, worker path)', async () => {
    const { inflight, canvas, render } = buildWorker();

    // First frame sizes the canvas.
    const p1 = render();
    inflight[0].resolve(fakeBitmap(800, 600));
    await p1;
    expect(canvas.width).toBe(800);
    expect(canvas.height).toBe(600);

    // Track further width/height writes on the canvas via a spy on the setters.
    let widthWrites = 0;
    let heightWrites = 0;
    Object.defineProperty(canvas, 'width', {
      get: () => 800,
      set: () => {
        widthWrites++;
      },
      configurable: true,
    });
    Object.defineProperty(canvas, 'height', {
      get: () => 600,
      set: () => {
        heightWrites++;
      },
      configurable: true,
    });

    // A same-size frame must NOT re-assign width/height.
    const p2 = render();
    inflight[1].resolve(fakeBitmap(800, 600));
    await p2;
    expect(widthWrites).toBe(0);
    expect(heightWrites).toBe(0);
  });
});

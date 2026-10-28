import { describe, expect, it, vi } from 'vitest';
import type { XlsxWorkbook } from '../workbook.js';
import {
  SelectionController,
  createSheetViewModel,
  SheetAcquisition,
  SheetRenderDispatcher,
  ViewportState,
} from './sheet-viewer-runtime.js';
import type { Worksheet } from '../types.js';

function workbook() {
  const destroy = vi.fn();
  return { value: { destroy } as unknown as XlsxWorkbook, destroy };
}

describe('XLSX viewer composition roles', () => {
  it('SheetAcquisition installs only the latest generation and closes every loser', async () => {
    const acquisition = new SheetAcquisition();
    const first = workbook();
    const second = workbook();
    let resolveFirst: ((value: XlsxWorkbook) => void) | undefined;
    const firstLoad = new Promise<XlsxWorkbook>((resolve) => { resolveFirst = resolve; });

    const stale = acquisition.replace(() => firstLoad);
    await acquisition.replace(() => Promise.resolve(second.value));
    resolveFirst?.(first.value);

    await expect(stale).resolves.toBeNull();
    expect(acquisition.current).toBe(second.value);
    expect(first.destroy).toHaveBeenCalledOnce();
    acquisition.destroy();
    expect(second.destroy).toHaveBeenCalledOnce();
  });

  it('SheetAcquisition is terminally closed and disposes a candidate that resolves late', async () => {
    const acquisition = new SheetAcquisition();
    let resolveLate: ((value: XlsxWorkbook) => void) | undefined;
    const pending = new Promise<XlsxWorkbook>((resolve) => { resolveLate = resolve; });
    const replacement = acquisition.replace(() => pending);

    acquisition.destroy();
    const late = workbook();
    resolveLate?.(late.value);

    await expect(replacement).rejects.toThrow('SheetAcquisition is closed');
    expect(late.destroy).toHaveBeenCalledOnce();
    await expect(acquisition.replace(() => Promise.resolve(workbook().value))).rejects.toThrow(
      'SheetAcquisition is closed',
    );
    expect(() => acquisition.install(workbook().value)).toThrow('SheetAcquisition is closed');
  });

  it('borrows an injected workbook without destroying the caller-owned resource', () => {
    const acquisition = new SheetAcquisition();
    const shared = workbook();

    acquisition.install(shared.value, false);
    acquisition.destroy();

    expect(shared.destroy).not.toHaveBeenCalled();
  });

  it('copies only viewer-mutable worksheet state', () => {
    const cell = { row: 1, col: 1 };
    const source = {
      rows: [{ index: 1, collapsed: false, cells: [cell] }],
      rowHeights: { 1: 20 },
      colWidths: { 1: 8 },
      colCollapsed: { 1: false },
    } as unknown as Worksheet;

    const view = createSheetViewModel(source);
    const viewColCollapsed = view.colCollapsed;
    if (!viewColCollapsed) throw new Error('Expected colCollapsed projection');
    view.rows[0].collapsed = true;
    view.rowHeights[1] = 30;
    view.colWidths[1] = 12;
    viewColCollapsed[1] = true;

    expect(source.rows[0].collapsed).toBe(false);
    expect(source.rowHeights[1]).toBe(20);
    expect(source.colWidths[1]).toBe(8);
    expect(source.colCollapsed?.[1]).toBe(false);
    expect(view.rows[0].cells[0]).toBe(cell);
  });

  it('SheetAcquisition reports terminal close when an in-flight loader rejects late', async () => {
    const acquisition = new SheetAcquisition();
    let rejectLate: ((reason: Error) => void) | undefined;
    const pending = new Promise<XlsxWorkbook>((_resolve, reject) => { rejectLate = reject; });
    const replacement = acquisition.replace(() => pending);

    acquisition.destroy();
    rejectLate?.(new Error('late parser failure'));

    await expect(replacement).rejects.toThrow('SheetAcquisition is closed');
  });

  it('ViewportState clamps logical offsets without a native scroll element', () => {
    const viewport = new ViewportState(1);
    viewport.setExtent(1000, 800);
    viewport.setViewportSize(300, 200);
    viewport.setOffset(900, 700);
    expect({ x: viewport.x, y: viewport.y }).toEqual({ x: 700, y: 600 });
    viewport.setScale(1.5);
    expect(viewport.scale).toBe(1.5);
  });

  it('SelectionController owns independent immutable cell coordinates', () => {
    const selection = new SelectionController();
    const cell = { row: 2, col: 3 };
    selection.select(cell);
    cell.row = 9;
    expect(selection.anchor).toEqual({ row: 2, col: 3 });
    selection.reset();
    expect(selection.active).toBeNull();
  });

  it('SelectionController only ends a drag for its active pointer', () => {
    const selection = new SelectionController();
    selection.beginDrag(7);
    expect(selection.dragging).toBe(true);
    expect(selection.draggingPointerId).toBe(7);

    selection.endDrag(9);
    expect(selection.dragging).toBe(true);
    expect(selection.draggingPointerId).toBe(7);

    selection.endDrag(7);
    expect(selection.dragging).toBe(false);
    expect(selection.draggingPointerId).toBeNull();
  });

  it('SelectionController owns range extension and renderer header projection', () => {
    const selection = new SelectionController();
    selection.select({ row: 2, col: 1 }, 'rows');
    selection.extend({ row: 5, col: 1 });
    expect(selection.snapshot()).toEqual({
      areas: [{ kind: 'rows', firstRow: 2, lastRow: 5 }],
      activeAreaIndex: 0,
      activeCell: { row: 2, col: 1 },
      extensionAnchor: { row: 2, col: 1 },
    });
    expect(selection.headerHighlight()).toEqual({
      selectedRowRange: { start: 2, end: 5, strong: true },
      selectedColRange: { start: 1, end: Number.MAX_SAFE_INTEGER, strong: false },
    });
  });

  it('SelectionController appends a new active area and extends only that area', () => {
    const selection = new SelectionController();
    selection.select({ row: 1, col: 1 });

    expect(selection.add({ row: 3, col: 4 })).toBe(true);
    selection.extend({ row: 5, col: 6 });

    expect(selection.snapshot()).toEqual({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 },
        { kind: 'cells', top: 3, left: 4, bottom: 5, right: 6 },
      ],
      activeAreaIndex: 1,
      activeCell: { row: 3, col: 4 },
      extensionAnchor: { row: 3, col: 4 },
    });
  });

  it('SelectionController activates an existing area without accumulating duplicates', () => {
    const selection = new SelectionController();
    selection.select({ row: 1, col: 1 });
    expect(selection.add({ row: 3, col: 4 })).toBe(true);
    selection.extend({ row: 5, col: 6 });

    expect(selection.add({ row: 4, col: 5 })).toBe(false);
    expect(selection.snapshot()).toEqual({
      areas: [
        { kind: 'cells', top: 1, left: 1, bottom: 1, right: 1 },
        { kind: 'cells', top: 3, left: 4, bottom: 5, right: 6 },
      ],
      activeAreaIndex: 1,
      activeCell: { row: 4, col: 5 },
      extensionAnchor: { row: 4, col: 5 },
    });
  });

  it('SheetRenderDispatcher invalidates in-flight generations on destroy', () => {
    const dispatcher = new SheetRenderDispatcher();
    const generation = dispatcher.begin();
    expect(dispatcher.isCurrent(generation)).toBe(true);
    dispatcher.destroy();
    expect(dispatcher.isCurrent(generation)).toBe(false);
  });

  it('keeps at most one active render and one latest pending render', async () => {
    const dispatcher = new SheetRenderDispatcher();
    let finishFirst: () => void = () => undefined;
    const first = new Promise<void>((resolve) => { finishFirst = resolve; });
    const calls: string[] = [];

    dispatcher.schedule(() => {
      calls.push('first');
      return first;
    });
    dispatcher.schedule(() => { calls.push('stale'); });
    dispatcher.schedule(() => { calls.push('latest'); });
    expect(calls).toEqual(['first']);

    finishFirst();
    await first;
    await Promise.resolve();
    await Promise.resolve();
    expect(calls).toEqual(['first', 'latest']);
    dispatcher.destroy();
  });

  it('schedules and cancels frames in the target canvas realm', () => {
    let popupFrame: FrameRequestCallback | null = null;
    const popupScheduler = {
      requestAnimationFrame: vi.fn((callback: FrameRequestCallback) => {
        popupFrame = callback;
        return 41;
      }),
      cancelAnimationFrame: vi.fn(),
    };
    const render = vi.fn();
    const dispatcher = new SheetRenderDispatcher(undefined, false, popupScheduler);

    dispatcher.schedule(render);
    expect(popupScheduler.requestAnimationFrame).toHaveBeenCalledOnce();
    expect(render).not.toHaveBeenCalled();

    (popupFrame as FrameRequestCallback | null)?.(0);
    expect(render).toHaveBeenCalledOnce();

    dispatcher.destroy();

    const pending = new SheetRenderDispatcher(undefined, false, popupScheduler);
    pending.schedule(render);
    pending.destroy();
    expect(popupScheduler.cancelAnimationFrame).toHaveBeenCalledWith(41);
  });

  it('invalidates an active worker bitmap when a non-scroll update is queued', async () => {
    const transfer = vi.fn();
    const canvas = {
      width: 0,
      height: 0,
      style: { width: '', height: '' },
      getContext: vi.fn(() => ({ transferFromImageBitmap: transfer })),
    } as unknown as HTMLCanvasElement;
    const dispatcher = new SheetRenderDispatcher(canvas, true);
    let finishActive: () => void = () => undefined;
    const active = new Promise<void>((resolve) => { finishActive = resolve; });
    let activeGeneration = 0;
    dispatcher.schedule(() => {
      activeGeneration = dispatcher.begin();
      return active;
    });

    dispatcher.schedule(() => undefined);
    const close = vi.fn();
    expect(dispatcher.commitBitmap(
      activeGeneration,
      { close, width: 1, height: 1 } as unknown as ImageBitmap,
      1,
      1,
    )).toBe(false);
    expect(close).toHaveBeenCalledOnce();
    expect(transfer).not.toHaveBeenCalled();

    finishActive();
    await active;
    await Promise.resolve();
    await Promise.resolve();
    dispatcher.destroy();
  });

  it('SheetRenderDispatcher owns stale ImageBitmap disposal', () => {
    const close = vi.fn();
    const dispatcher = new SheetRenderDispatcher();
    const stale = dispatcher.begin();
    dispatcher.begin();
    expect(dispatcher.commitBitmap(
      stale,
      { close, width: 1, height: 1 } as unknown as ImageBitmap,
      1,
      1,
    )).toBe(false);
    expect(close).toHaveBeenCalledOnce();
  });

  it('SheetRenderDispatcher closes an owned bitmap when commit fails', () => {
    const failure = new Error('context lost');
    const transfer = vi.fn(() => { throw failure; });
    const canvas = {
      width: 0,
      height: 0,
      style: { width: '', height: '' },
      getContext: vi.fn(() => ({ transferFromImageBitmap: transfer })),
    } as unknown as HTMLCanvasElement;
    const close = vi.fn();
    const bitmap = { close, width: 2, height: 3 } as unknown as ImageBitmap;
    const dispatcher = new SheetRenderDispatcher(canvas, true);
    const generation = dispatcher.begin();

    expect(() => dispatcher.commitBitmap(generation, bitmap, 20, 30)).toThrow(failure);
    expect(close).toHaveBeenCalledOnce();
  });
});

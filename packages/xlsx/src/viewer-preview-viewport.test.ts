import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import type { Worksheet } from './types.js';
import { installDom, makeContainer, type FakeEl } from './viewer-destroy-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/** A changed viewport cannot use row coverage proved for the initial frame. */
describe('XLSX provisional viewport', () => {
  it.each(['scroll', 'resize'] as const)('waits after %s before the pull completes', async (change) => {
    installDom();
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement);
    const paint = vi.fn(async () => undefined);
    const engine = viewer as unknown as {
      wb: unknown;
      currentSheet: number;
      currentWorksheet: Worksheet | null;
      previewCompletion: Promise<Worksheet> | null;
      firstPreviewRender: boolean;
      previewPreparedViewport: { width: number; height: number; scale: number } | null;
      canvasArea: FakeEl;
      scrollHost: FakeEl;
      viewport: { ensureExtent(width: number, height: number): void };
      viewportTop: number;
      renderCurrentSheet(): Promise<void>;
    };
    const worksheet = {
      name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
      defaultColWidth: 64, defaultRowHeight: 20, mergeCells: [],
      freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
    } as unknown as Worksheet;
    engine.wb = { renderViewport: paint, destroy: vi.fn() };
    engine.currentSheet = 0;
    engine.currentWorksheet = worksheet;
    engine.canvasArea.clientWidth = 800;
    engine.canvasArea.clientHeight = 600;
    engine.previewCompletion = new Promise<Worksheet>(() => undefined);
    engine.firstPreviewRender = true;
    engine.previewPreparedViewport = { width: 800, height: 600, scale: 1 };
    if (change === 'scroll') {
      engine.scrollHost.clientHeight = 600;
      engine.scrollHost.scrollHeight = 10_000;
      engine.viewport.ensureExtent(800, 10_000);
      engine.viewportTop = 100;
    }
    else engine.canvasArea.clientWidth = 700;

    await engine.renderCurrentSheet();
    expect(paint).not.toHaveBeenCalled();
    viewer.destroy();
  });
});

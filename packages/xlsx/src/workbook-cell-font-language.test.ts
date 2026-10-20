import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
import { resolvedThemeCellFont } from './renderer.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';
import { createSizeOverriddenWorksheet } from './worker-protocol.js';
import type { ParsedWorkbook, Worksheet } from './types.js';

const font = { name: 'Calibri', scheme: 'minor' as const, bold: false, italic: false };
const metadata = {
  workbook: { sheets: [{ name: 'Sheet1' }] },
  sharedStrings: [], styles: { fonts: [font], cellXfs: [], dxfs: [] },
} as unknown as ParsedWorkbook;
const worksheet: Worksheet = {
  name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {},
  defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
  freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
  themeJapaneseMinorFont: 'Meiryo',
};
const bytes = (value: unknown) => new TextEncoder().encode(JSON.stringify(value)).buffer;

/** Only archive transport is faked: real workbook loading and lazy model
 * materialization must retain the language chosen before any sheet is read. */
class WorkbookWorker {
  static instances: WorkbookWorker[] = [];
  readonly messages: Record<string, unknown>[] = [];
  private readonly listeners = new Set<(event: MessageEvent) => void>();
  constructor(private readonly mode: 'main' | 'worker') {
    WorkbookWorker.instances.push(this);
  }
  postMessage(message: Record<string, unknown>): void {
    this.messages.push(message);
    let response: unknown;
    if (message.type === 'parse') {
      response = this.mode === 'main'
        ? { type: 'parsed', id: message.id, workbookJson: bytes(metadata) }
        : { type: 'parsed', id: message.id, workbook: structuredClone(metadata) };
    } else if (message.type === 'openSheetSession') {
      response = { type: 'sheetSessionOpened', id: message.id };
    } else if (message.kind === 'pull') {
      const payload = bytes({ kind: 'finished', worksheet });
      response = { ...message, kind: 'chunk', payload, byteLength: payload.byteLength, done: true };
    } else if (message.kind === 'ack' || message.kind === 'close') {
      response = { ...message, kind: 'accepted', command: message.kind };
    }
    if (response) queueMicrotask(() => {
      for (const listener of this.listeners) listener({ data: response } as MessageEvent);
    });
  }
  addEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.delete(listener);
  }
  terminate(): void {}
}

vi.mock('./worker.ts?worker&inline', () => ({
  default: class { constructor() { return new WorkbookWorker('main'); } },
}));
vi.mock('./render-worker-host', () => ({
  createRenderWorker: () => new WorkbookWorker('worker'),
}));

afterEach(() => {
  vi.unstubAllGlobals();
  WorkbookWorker.instances = [];
});

describe('workbook cell-font language snapshot', () => {
  it.each(['main', 'worker'] as const)('captures at load before lazy sheet acquisition in %s mode', async (mode) => {
    vi.stubGlobal('Worker', WorkbookWorker);
    vi.stubGlobal('OffscreenCanvas', class {});
    vi.stubGlobal('location', { href: 'https://example.test/' });
    vi.stubGlobal('navigator', { language: 'en-US' });
    const source = () => new Uint8Array([0x50, 0x4b, 3, 4]).buffer;
    const english = await XlsxWorkbook.load(source(), { mode });
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    // No sheet was loaded in English: taking the snapshot at first layout or
    // lazy worksheet acquisition would select the wrong family here.
    const sheet = await english.getWorksheet(0);
    const view = createSheetViewModel(sheet);
    const workerProjection = createSizeOverriddenWorksheet(view, { rows: { 1: 30 } });
    for (const model of [sheet, view, workerProjection]) {
      expect(resolvedThemeCellFont(font, model).family).toBe('Calibri');
    }
    const japanese = await XlsxWorkbook.load(source(), { mode });
    const japaneseSheet = await japanese.getWorksheet(0);
    expect(resolvedThemeCellFont(font, japaneseSheet).family).toBe('Meiryo');
    vi.stubGlobal('navigator', { language: 'en-US' });
    expect(resolvedThemeCellFont(font, japaneseSheet).family).toBe('Meiryo');
    expect(resolvedThemeCellFont(font, sheet).family).toBe('Calibri');
    if (mode === 'worker') {
      const requests = WorkbookWorker.instances.map((worker) => worker.messages.find((msg) => msg.type === 'parse'));
      expect(requests.map((req) => req?.japaneseCellFonts)).toEqual([false, true]);
    }
    english.destroy();
    japanese.destroy();
  });
});

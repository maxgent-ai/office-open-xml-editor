import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OoxmlResourceUsageSnapshot, WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument } from './document.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { installStubCanvas } from './testing/synthetic-document.js';
import { noteContinuationBoundaryModel } from './testing/note-continuation-model.js';

// The public load option's library default for normal DOCX input: omitted
// continues paragraph footnotes like explicit true; explicit false keeps
// whole notes. Main/sliced loads run the real normalization, layout services
// and pagination; only the WASM parse is replaced by the shared model, as in
// the other main-mode load tests.

const globals = globalThis as Record<string, unknown>;
const originals = { Worker: globals.Worker, location: globals.location };
const USAGE: OoxmlResourceUsageSnapshot = {
  archiveEntryCount: 1, declaredInflatedBytes: 0, largestInflatedEntryBytes: 0,
  distinctInflatedBytes: 0, operationInflatedBytes: 0,
};

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void {}
}

/** Records worker-mode parse requests and answers them with empty metadata.
 * Wire contract only: it does not paginate. */
class ParseRecordingWorker {
  static parses: Record<string, unknown>[] = [];
  private readonly listeners = new Set<(event: MessageEvent) => void>();
  postMessage(message: unknown): void {
    const request = message as Record<string, unknown>;
    if (request.type !== 'parse' && request.type !== 'resourceUsage') return;
    if (request.type === 'parse') ParseRecordingWorker.parses.push(request);
    const reply = request.type === 'parse'
      ? { type: 'parsedMeta', id: request.id, meta: {
          pageCount: 0, revisions: [], comments: [], footnotes: [], endnotes: [], pageSizes: [],
          bookmarkPages: [], commentAnchorRanges: [], revisionAnchorRanges: [],
        } }
      : { type: 'resourceUsage', id: request.id, usage: undefined };
    queueMicrotask(() => { for (const listener of [...this.listeners]) listener({ data: reply } as MessageEvent); });
  }
  addEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.add(listener);
  }
  removeEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.delete(listener);
  }
  terminate(): void {}
}

beforeAll(() => { installStubCanvas(); });

afterEach(() => {
  vi.restoreAllMocks();
  globals.Worker = originals.Worker;
  globals.location = originals.location;
  ParseRecordingWorker.parses = [];
});

function installMainModeParse(): void {
  globals.Worker = SilentWorker;
  globals.location = { href: 'http://localhost/' };
  vi.spyOn(DocxDocument.prototype as unknown as { _parse(): Promise<void> }, '_parse')
    .mockImplementation(async function (this: DocxDocument) {
      const doc = this as unknown as { _document: unknown; _source: unknown; _meta: unknown };
      const model = noteContinuationBoundaryModel();
      doc._document = model;
      doc._source = layoutSourceStore(model);
      doc._meta = null;
    });
  vi.spyOn(DocxDocument.prototype as unknown as { _resourceUsage(): Promise<OoxmlResourceUsageSnapshot> },
    '_resourceUsage').mockResolvedValue(USAGE);
}

describe('DocxDocument.load footnote continuation default', () => {
  it.each([
    { option: undefined, sliceLayout: false, continued: true },
    { option: undefined, sliceLayout: true, continued: true },
    { option: true, sliceLayout: false, continued: true },
    { option: false, sliceLayout: false, continued: false },
    { option: false, sliceLayout: true, continued: false },
  ])('lays out footnote continuation $option in a main load (sliced=$sliceLayout)', async ({
    option, sliceLayout, continued,
  }) => {
    installMainModeParse();
    const document = await DocxDocument.load(new ArrayBuffer(0), {
      sliceLayout, currentDate: 0,
      ...(option === undefined ? {} : { allowFootnoteContinuation: option }),
    });
    try {
      const pages = await Promise.all(Array.from({ length: document.pageCount }, async (_, page) =>
        (await document.collectPageRuns(page)).map(run => run.text).join('')));
      // noteContinuationBoundaryModel arithmetic: continued notes keep one
      // head line beside B01; whole notes move B01 and the note to page 2.
      expect(pages).toEqual(continued
        ? ['F0F1F2F3F41B01note-0;', 'note-1;note-2;note-3;']
        : ['F0F1F2F3F4', '1B01note-0;note-1;note-2;note-3;']);
    } finally { document.destroy(); }
  });

  it.each([
    { option: undefined, wire: true },
    { option: true, wire: true },
    { option: false, wire: false },
  ])('sends footnote continuation $option to the worker parse as $wire', async ({ option, wire }) => {
    globals.Worker = ParseRecordingWorker;
    globals.location = { href: 'http://localhost/' };
    const document = await DocxDocument.load(new ArrayBuffer(0), {
      mode: 'worker',
      ...(option === undefined ? {} : { allowFootnoteContinuation: option }),
    });
    try {
      expect(ParseRecordingWorker.parses).toHaveLength(1);
      expect(ParseRecordingWorker.parses[0]).toMatchObject({ type: 'parse', allowFootnoteContinuation: wire });
    } finally { document.destroy(); }
  });
});

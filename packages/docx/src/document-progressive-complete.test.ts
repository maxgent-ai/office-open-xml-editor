import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OoxmlResourceUsageSnapshot, WorkerLike } from '@silurus/ooxml-core';
import { DocxDocument, docxViewerLoadSignal, type DocxViewerLoadControl, type LoadOptions } from './document.js';
import { activeDocxLayoutViewOf } from './document-layout-view.js';
import { PaginationAbortError } from './layout/pagination-scheduler.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
import { installDeterministicPaginationHost } from './testing/deterministic-pagination-host.js';
import { DocxViewer } from './viewer.js';
import { installDom, makeEl } from './scroll-viewer-test-dom.js';

// ─────────────────────────────────────────────────────────────────────────────
// The `onLayoutComplete` terminal-callback contract for main-mode progressive
// loads: exactly one success notification per load, whether or not the
// document was long enough to publish partials first.
//
// `layoutDocumentProgressively` only publishes a preview when the body has
// more than MIN_PROGRESSIVE_ENTRIES entries, so a short document's drain
// completes with `publishedLayout === null`. The success notification used to
// be gated on that local, so consumers of a fast document never learned the
// layout was complete — the terminal callback contract silently depended on
// document speed.
//
// These tests drive the REAL `DocxDocument.load` progressive block: `_parse`
// is stubbed to install a synthetic model/source (the WASM parse is not what
// is under test), and the final usage probe is stubbed because the inert
// inline worker cannot answer it.
// ─────────────────────────────────────────────────────────────────────────────

class SilentWorker implements WorkerLike {
  postMessage(): void {}
  addEventListener(): void {}
  removeEventListener(): void {}
  terminate(): void {}
}

const globals = globalThis as Record<string, unknown>;
const originals = {
  Worker: globals.Worker,
  location: globals.location,
};

const USAGE: OoxmlResourceUsageSnapshot = {
  archiveEntryCount: 1,
  declaredInflatedBytes: 0,
  largestInflatedEntryBytes: 0,
  distinctInflatedBytes: 0,
  operationInflatedBytes: 0,
};

beforeAll(() => {
  installStubCanvas();
});

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  globals.Worker = originals.Worker;
  globals.location = originals.location;
});

/** Stub the WASM parse with a synthetic document of `paragraphs` body entries. */
function installMainModeParse(paragraphs: number, kind: 'plain' | 'tracked' = 'plain'): void {
  globals.Worker = SilentWorker;
  globals.location = { href: 'http://localhost/' };
  vi.spyOn(
    DocxDocument.prototype as unknown as {
      _parse(
        buffer: ArrayBuffer,
        resourcePolicy: unknown,
        useGoogleFonts?: boolean,
        timeoutMs?: number,
        onUsage?: unknown,
        renderers?: unknown,
        progressive?: unknown,
      ): Promise<void>;
    },
    '_parse',
  ).mockImplementation(async function (this: DocxDocument) {
    const doc = this as unknown as {
      _document: unknown;
      _source: unknown;
      _meta: unknown;
    };
    const model = syntheticDocxModel(kind, { paragraphs, wordsPerParagraph: 120 });
    doc._document = model;
    doc._source = layoutSourceStore(model);
    doc._meta = null;
  });
  vi.spyOn(
    DocxDocument.prototype as unknown as {
      _resourceUsage(timeoutMs: number): Promise<OoxmlResourceUsageSnapshot>;
    },
    '_resourceUsage',
  ).mockResolvedValue(USAGE);
}

describe('main-mode progressive load: onLayoutComplete contract', () => {
  it('notifies completion exactly once for a fast document that publishes nothing early', async () => {
    // Three body entries stay under the preview threshold, so the drain
    // completes before any partial could be shown: no onLayoutPartial, and
    // load() resolves on the finished layout itself. The completion callback
    // must still fire — load() resolving is not a substitute for it.
    installMainModeParse(3);
    const completions: unknown[] = [];
    const partials: unknown[] = [];

    const doc = await DocxDocument.load(new ArrayBuffer(0), {
      progressiveLayout: true,
      onLayoutComplete: (error) => completions.push(error),
      onLayoutPartial: (partial) => partials.push(partial),
    });
    await doc.waitUntilLayoutComplete();

    expect(partials).toHaveLength(0);
    expect(completions).toEqual([undefined]);
    expect(doc.layoutComplete).toBe(true);
    doc.destroy();
  });

  it('notifies completion exactly once after partials for a long document', async () => {
    // Long enough to cross several page-count checkpoints: load() resolves on
    // the opening publication and the drain keeps publishing. The completion
    // callback fires once, after the last partial, never twice.
    installMainModeParse(600);
    const completions: unknown[] = [];
    const partials: number[] = [];

    const doc = await DocxDocument.load(new ArrayBuffer(0), {
      progressiveLayout: true,
      onLayoutComplete: (error) => completions.push(error),
      onLayoutPartial: (partial) => partials.push(partial.availableUnits),
    });
    await doc.waitUntilLayoutComplete();

    expect(partials.length).toBeGreaterThan(0);
    expect(completions).toEqual([undefined]);
    expect(doc.layoutComplete).toBe(true);
    doc.destroy();
  }, 300_000);
});

describe('ordinary main-mode sliced load ownership', () => {
  it('does not restart pagination when progress re-applies the unchanged view', async () => {
    // Regression from an onLayoutProgress UI that synchronizes its setting on
    // every update: the load must advance past its first paginator slice.
    installDom();
    installMainModeParse(20);
    const yields = installDeterministicPaginationHost();
    vi.spyOn(DocxViewer.prototype as unknown as { _render(): Promise<void> }, '_render')
      .mockResolvedValue(undefined);
    let progress = 0;
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      onLayoutProgress: () => {
        progress++;
        void viewer.setShowTrackedChanges(false);
        if (progress === 200) viewer.destroy();
      },
    });
    vi.stubGlobal('document', undefined);
    try {
      await viewer.load(new ArrayBuffer(0));
      expect(progress).toBeGreaterThan(0);
      expect(yields()).toBeGreaterThan(0);
    } finally {
      viewer.destroy();
    }
  }, 30_000);

  it('settles on the final view after repeated changes during pagination', async () => {
    installDom();
    installMainModeParse(60, 'tracked');
    const yields = installDeterministicPaginationHost();
    let loadedDoc!: DocxDocument;
    vi.spyOn(DocxDocument.prototype as unknown as {
      _resourceUsage(timeoutMs: number): Promise<OoxmlResourceUsageSnapshot>;
    }, '_resourceUsage').mockImplementation(function (this: DocxDocument) {
      loadedDoc = this;
      return Promise.resolve(USAGE);
    });
    vi.spyOn(DocxViewer.prototype as unknown as { _render(): Promise<void> }, '_render')
      .mockResolvedValue(undefined);
    const views = [true, false, true, false, true];
    let notifications = 0;
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      onLayoutProgress: () => {
        if (notifications < views.length) void viewer.setShowTrackedChanges(views[notifications]!);
        notifications++;
      },
    });
    vi.stubGlobal('document', undefined);
    try {
      await viewer.load(new ArrayBuffer(0));
      expect(notifications).toBeGreaterThanOrEqual(views.length);
      expect(yields()).toBeGreaterThan(0);
      expect(activeDocxLayoutViewOf(loadedDoc).showTrackedChanges).toBe(true);
    } finally {
      viewer.destroy();
    }
  }, 30_000);

  it('reconciles the active view when toggles return to their initial value during the final probe', async () => {
    installDom();
    installMainModeParse(100, 'tracked');
    vi.spyOn(DocxViewer.prototype as unknown as { _render(): Promise<void> }, '_render')
      .mockResolvedValue(undefined);
    let finishProbe!: (usage: OoxmlResourceUsageSnapshot) => void;
    let enteredProbe!: () => void;
    let loadedDoc!: DocxDocument;
    const probing = new Promise<void>((resolve) => { enteredProbe = resolve; });
    vi.spyOn(DocxDocument.prototype as unknown as {
      _resourceUsage(timeoutMs: number): Promise<OoxmlResourceUsageSnapshot>;
    }, '_resourceUsage').mockImplementation(function (this: DocxDocument) {
      loadedDoc = this;
      enteredProbe();
      return new Promise((resolve) => { finishProbe = resolve; });
    });
    let changed = false;
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      onLayoutProgress: () => {
        if (changed) return;
        changed = true;
        void viewer.setShowTrackedChanges(true);
      },
    });
    vi.stubGlobal('document', undefined);
    const loading = viewer.load(new ArrayBuffer(0));
    await probing;
    expect(activeDocxLayoutViewOf(loadedDoc).showTrackedChanges).toBe(true);
    await viewer.setShowTrackedChanges(false);
    finishProbe(USAGE);
    await loading;
    expect(activeDocxLayoutViewOf(loadedDoc).showTrackedChanges).toBe(false);
    expect(viewer.pageCount).toBe(loadedDoc.pageCount);
    viewer.destroy();
  }, 300_000);

  it('releases a font registered after a viewer-owned load is aborted', async () => {
    installMainModeParse(3);
    const faces = new Set<FontFace>();
    let finish!: () => void;
    let entered!: () => void;
    const started = new Promise<void>((resolve) => { entered = resolve; });
    vi.stubGlobal('document', { fonts: faces });
    vi.stubGlobal('FontFace', class {
      status = 'unloaded';
      constructor(public family: string) {}
      load(): Promise<this> {
        entered();
        return new Promise((resolve) => {
          finish = () => { this.status = 'loaded'; resolve(this); };
        });
      }
    });
    const abort = new AbortController();
    const loading = DocxDocument.load(new ArrayBuffer(0), {
      [docxViewerLoadSignal]: {
        signal: abort.signal,
        requestedView: () => undefined,
        subscribeViewChange: () => () => undefined,
      },
    } as LoadOptions);
    await started;
    expect(faces.size).toBe(1);
    abort.abort();
    finish();
    await expect(loading).rejects.toBeInstanceOf(PaginationAbortError);
    expect(faces.size).toBe(0);
  });

  it('cancels the active layout and resolves the original load with the latest tracked view', async () => {
    installMainModeParse(60);
    const yields = installDeterministicPaginationHost();
    let requested: boolean | undefined;
    let listener: (() => void) | null = null;
    let changed = false;
    let progress = 0;
    const control: DocxViewerLoadControl = {
      signal: new AbortController().signal,
      requestedView: () => requested,
      subscribeViewChange: (next) => {
        listener = next;
        return () => { if (listener === next) listener = null; };
      },
    };
    const opts = {
      [docxViewerLoadSignal]: control,
      onLayoutProgress: () => {
        progress++;
        if (changed) return;
        changed = true;
        requested = true;
        listener?.();
        listener?.();
      },
    } as LoadOptions;
    const doc = await DocxDocument.load(new ArrayBuffer(0), opts);
    expect(changed).toBe(true);
    expect(progress).toBeGreaterThan(0);
    expect(yields()).toBeGreaterThan(0);
    expect(activeDocxLayoutViewOf(doc).showTrackedChanges).toBe(true);
    expect(doc.layoutComplete).toBe(true);
    doc.destroy();
  }, 300_000);

  it('rejects a destroyed viewer-owned layout instead of installing a stale result', async () => {
    installMainModeParse(60);
    const abort = new AbortController();
    const control: DocxViewerLoadControl = {
      signal: abort.signal,
      requestedView: () => undefined,
      subscribeViewChange: () => () => undefined,
    };
    await expect(DocxDocument.load(new ArrayBuffer(0), {
      [docxViewerLoadSignal]: control,
      onLayoutProgress: () => abort.abort(),
    } as LoadOptions)).rejects.toBeInstanceOf(PaginationAbortError);
  }, 300_000);
});

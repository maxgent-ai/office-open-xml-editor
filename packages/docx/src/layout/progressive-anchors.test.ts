import { existsSync, readFileSync } from 'node:fs';
import init, { DocxArchive } from '../wasm/docx_parser.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';
import { describe, expect, it, vi } from 'vitest';

// Counts body pagination passes: each pass opens exactly one kernel session.
const passes = vi.hoisted(() => ({ opened: 0, finalizedPages: 0, sessionCalls: 0 }));
vi.mock('./page-factory.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./page-factory.js')>();
  return {
    ...actual,
    finalizeLayoutPage: (...args: Parameters<typeof actual.finalizeLayoutPage>) => {
      passes.finalizedPages += 1;
      return actual.finalizeLayoutPage(...args);
    },
  };
});
vi.mock('./runtime-state.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./runtime-state.js')>();
  return {
    ...actual,
    bodyLayoutKernelOf: (services: Parameters<typeof actual.bodyLayoutKernelOf>[0]) => {
      const kernel = actual.bodyLayoutKernelOf(services);
      return kernel && {
        openBodyLayoutSession: (...args: Parameters<typeof kernel.openBodyLayoutSession>) => {
          passes.opened += 1;
          // Every session call is layout work; counting them between host
          // turns measures how long the host waited.
          const session = kernel.openBodyLayoutSession(...args);
          return Object.fromEntries(Object.entries(session).map(([key, value]) => [
            key,
            typeof value === 'function'
              ? (...callArgs: unknown[]) => { passes.sessionCalls += 1; return value.apply(session, callArgs); }
              : value,
          ])) as typeof session;
        },
      };
    },
  };
});

import { createLayoutServices } from '../layout-runtime.js';
import { layoutSourceStore } from '../layout-source-model-adapter.js';
import type { BodyElement, DocxDocumentModel } from '../types.js';
import { paginateBody } from './body-paginator.js';
import { layoutFingerprint } from './invariants.js';
import { normalizeLayoutOptions } from './options.js';
import { PaginationAbortError } from './pagination-scheduler.js';
import { layoutDocumentProgressively, type ProgressiveLayoutPreview } from './progressive.js';
import type { DocumentLayout, LayoutPage } from './types.js';

// Page-owned anchors (§20.4.3.5 positionV relativeFrom="margin") and
// page-positioned floating tables (§17.4.57 tblpPr vertAnchor="page") are
// resolved by exact-state convergence: the first pass prescans anchors in
// source order, later passes apply the destinations and table page floors the
// previous pass proved. A page published from a pass that a later pass
// supersedes would be revised. These cases pin that such documents still
// publish progressively and that every published page is already the page of
// the final layout.
//
// Pagination here is independent of glyph metrics, so the pass counts and
// publication points are identical on every host: each paragraph is a single
// short glyph on an exact 24 pt line (§17.3.1.33), measured by an injected
// linear-metric context, and every anchor and table has explicit geometry.
// The 468 pt text column (US Letter, 1 in margins) holds 27 lines per page.

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const context = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return context as unknown as CanvasRenderingContext2D;
}

function line(extraRuns: readonly unknown[] = []): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0,
    lineSpacing: { value: 24, rule: 'exact', explicit: true },
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'x', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'NotInMetrics',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }, ...extraRuns],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as BodyElement;
}

/** A full-width, margin-relative image that text must clear (wrapTopAndBottom,
 * §20.4.2.20). Its paragraph is a page-owned anchor owner. */
function anchoredImageLine(heightPt: number): BodyElement {
  return line([{
    type: 'image', imagePath: 'word/media/anchor.png', mimeType: 'image/png',
    widthPt: 468, heightPt, anchor: true, anchorXPt: 0, anchorYPt: 0,
    anchorXFromMargin: true, anchorYFromPara: false,
    wrapMode: 'topAndBottom', wrapSide: 'bothSides',
    anchorXRelativeFrom: 'margin', anchorYRelativeFrom: 'margin',
  }]);
}

/** A page-positioned floating table of empty exact-height rows. */
function pageFloatingTable(
  widthPt: number,
  tblpY: number,
  rows: number,
  rowHeightPt: number,
  topFromText: number,
): BodyElement {
  const noBorders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
  return {
    type: 'table',
    colWidths: [widthPt],
    rows: Array.from({ length: rows }, () => ({
      cells: [{
        content: [], colSpan: 1, vMerge: null, borders: noBorders,
        background: null, vAlign: 'top', widthPt: null,
      }],
      rowHeight: rowHeightPt, rowHeightRule: 'exact', isHeader: false,
    })),
    borders: noBorders,
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText, bottomFromText: 0,
      horzAnchor: 'margin', horzSpecified: true, vertAnchor: 'page', tblpX: 0, tblpY,
    },
  } as unknown as BodyElement;
}

/** An ordinary (in-flow) table of empty exact-height rows across the column. */
function ordinaryTable(rows: number, rowHeightPt: number): BodyElement {
  const { tblpPr: _tblpPr, ...table } = pageFloatingTable(468, 0, rows, rowHeightPt, 0) as unknown as
    Record<string, unknown>;
  return table as unknown as BodyElement;
}

/** `lines` single-line paragraphs; each insert lands at its body index, in order. */
function documentModel(
  lines: number,
  inserts: readonly (readonly [number, BodyElement])[],
): DocxDocumentModel {
  const body = Array.from({ length: lines }, () => line());
  for (const [index, element] of inserts) body.splice(index, 0, element);
  return {
    section: {
      pageWidth: 612, pageHeight: 792,
      marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
      headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: null,
    },
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function open(model: DocxDocumentModel) {
  const source = layoutSourceStore(model);
  return {
    input: source.bodyLayoutInput,
    services: createLayoutServices(source, { measureContext: measureContext() }),
    options: normalizeLayoutOptions(undefined, 1_700_000_000_000),
  };
}

function pageFingerprint(page: LayoutPage): string {
  return layoutFingerprint({ pages: [page], diagnostics: [] } as DocumentLayout);
}

function blockingLayout(model: DocxDocumentModel): Readonly<{ layout: DocumentLayout; passes: number }> {
  const before = passes.opened;
  const blocking = open(model);
  const layout = paginateBody(blocking.input, blocking.services, blocking.options);
  return { layout, passes: passes.opened - before };
}

async function progressiveRun(model: DocxDocumentModel) {
  const previews: ProgressiveLayoutPreview[] = [];
  const progressive = open(model);
  const final = await layoutDocumentProgressively(
    progressive.input,
    progressive.services,
    progressive.options,
    { onPreview: (preview) => { previews.push(preview); } },
  );
  return { previews, final };
}

/** The page holding the root of the body-level table at `bodyIndex`, and the
 * page where the flow reached it: the page holding the block before it. */
function tablePages(layout: DocumentLayout, bodyIndex: number) {
  const pagesWith = (predicate: (node: LayoutPage['layers']['body'][number]) => boolean) => (
    layout.pages.filter((page) => page.layers.body.some(predicate)).map((page) => page.pageIndex)
  );
  const placed = pagesWith((node) => node.kind === 'table' && node.source.path[0] === bodyIndex);
  const preceding = pagesWith((node) => node.source.path[0] === bodyIndex - 1);
  return { reachedPage: preceding.at(-1)!, placedPage: placed[0]! };
}

function expectPublishedPagesFinal(
  previews: readonly ProgressiveLayoutPreview[],
  final: DocumentLayout,
): void {
  const counts = previews.map((preview) => preview.layout.pages.length);
  expect(counts).toEqual([...counts].sort((left, right) => left - right));
  for (const preview of previews) {
    expect(preview.exact).toBe(false);
    expect(preview.layout.pages.length).toBeLessThanOrEqual(final.pages.length);
    preview.layout.pages.forEach((page, index) => {
      expect(pageFingerprint(page as LayoutPage))
        .toBe(pageFingerprint(final.pages[index] as LayoutPage));
    });
  }
}

// Line 60 sits on page 2 in source order. The first pass prescans its image
// on pages 0-2, the second applies it on the page it reached, and the third
// confirms the moved destination.
const anchoredImageModel = () => documentModel(300, [[60, anchoredImageLine(200)]]);

describe('progressive layout with page-owned anchors', () => {
  it('publishes only pages every later convergence pass reproduces', async () => {
    const model = anchoredImageModel();
    const blocking = blockingLayout(model);
    expect(blocking.passes).toBe(3);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // The anchored document now publishes before layout completes.
    expect(previews.length).toBeGreaterThan(1);
    expect(previews[0]!.layout.pages.length).toBeLessThan(final.pages.length);
    expectPublishedPagesFinal(previews, final);
  }, 300_000);

  it('keeps pages before a floored table publishable while its floor changes', async () => {
    // The full-width table's exclusion sends the lines before it on its page
    // past that page. The next pass therefore proves a minimum page for the
    // table (none, then one past its first page), and the converged pass reads
    // that floor where the flow reaches the table and advances past it.
    const tableIndex = 70;
    const model = documentModel(300, [[tableIndex, pageFloatingTable(468, 300, 4, 100, 20)]]);
    const blocking = blockingLayout(model);
    expect(blocking.passes).toBe(3);
    const { reachedPage, placedPage } = tablePages(blocking.layout, tableIndex);
    expect(placedPage).toBe(reachedPage + 1);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    expect(previews.length).toBeGreaterThan(0);
    expectPublishedPagesFinal(previews, final);
    // Every page before the reach page is published. A floor that acts at the
    // reach page bounds the pass there, so that page is never published early.
    expect(Math.max(...previews.map((preview) => preview.layout.pages.length)))
      .toBe(reachedPage);
  }, 300_000);

  it('lowers the bound to the reach page of a table placed at or past it', async () => {
    // Two page-positioned tables whose bands collide: the second is deferred
    // from the page where the flow reaches it to the next page (§17.4.56).
    // A checkpoint whose live page holds the deferred table must not publish
    // its reach page, because a later floor could still act there.
    const secondTableIndex = 62;
    const model = documentModel(300, [
      [60, pageFloatingTable(300, 72, 3, 120, 0)],
      [secondTableIndex, pageFloatingTable(300, 100, 3, 120, 0)],
    ]);
    const blocking = blockingLayout(model);
    const { reachedPage, placedPage } = tablePages(blocking.layout, secondTableIndex);
    expect(placedPage).toBe(reachedPage + 1);

    const { previews, final } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    expectPublishedPagesFinal(previews, final);
    const counts = previews.map((preview) => preview.layout.pages.length);
    // The checkpoint while the deferred table's page is live stops at the
    // reach page instead of publishing it (which would be `placedPage` pages),
    // and later checkpoints extend past both once that page has closed.
    expect(counts).toContain(reachedPage);
    expect(counts).not.toContain(placedPage);
    expect(counts.at(-1)!).toBeGreaterThan(placedPage);
  }, 300_000);

  it('publishes the opening pages of a document that opens with a page table during discovery', async () => {
    // A page-positioned table on the first page: the source-order first pass
    // cannot have registered it at the page start, so none of its pages is
    // final. The second pass, applying the plan of the pages the first has
    // closed, runs behind it and publishes before discovery finishes. A
    // three-page ordinary table makes one body entry of the second pass open
    // pages whose plan the first pass has not closed yet.
    const lines = 300;
    const model = documentModel(lines, [
      [2, pageFloatingTable(468, 300, 2, 100, 0)],
      [40, ordinaryTable(3 * 27, 24)],
    ]);
    const blocking = blockingLayout(model);
    expect(blocking.passes).toBe(2);
    expect(tablePages(blocking.layout, 2).placedPage).toBe(0);

    let steps = 0;
    let stepsAtFirstPreview: number | null = null;
    const previews: ProgressiveLayoutPreview[] = [];
    const run = open(model);
    const before = passes.opened;
    const final = await layoutDocumentProgressively(run.input, run.services, run.options, {
      scheduler: { onProgress: () => { steps += 1; } },
      onPreview: (preview) => {
        stepsAtFirstPreview ??= steps;
        previews.push(preview);
      },
    });
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // The overlapped second pass is the solver's second pass, not extra work.
    expect(passes.opened - before).toBe(blocking.passes);
    expectPublishedPagesFinal(previews, final);
    // The first pass alone suspends once per body entry; the opening preview
    // arrives well before it could have finished.
    expect(stepsAtFirstPreview).not.toBeNull();
    expect(stepsAtFirstPreview!).toBeLessThan(run.input.sequence.length / 4);
    expect(previews.at(-1)!.layout.pages.length).toBeGreaterThan(5);
  }, 300_000);

  it('keeps host turns frequent while the second pass waits on the first', async () => {
    // The ordinary table makes one second-pass entry wait for pages the first
    // pass has not closed; the many one-line paragraphs after it are what the
    // first pass must lay out meanwhile. Each must stay a suspension point.
    const model = documentModel(300, [
      [2, pageFloatingTable(468, 72, 3, 200, 0)],
      [6, pageFloatingTable(468, 72, 3, 200, 0)],
      [10, pageFloatingTable(468, 72, 3, 200, 0)],
      [40, ordinaryTable(4 * 27, 24)],
    ]);
    const blocking = blockingLayout(model);
    const run = open(model);
    let since = 0;
    let widest = 0;
    let last = passes.sessionCalls;
    const final = await layoutDocumentProgressively(run.input, run.services, run.options, {
      onPreview: () => {},
      scheduler: {
        now: () => Number.MAX_SAFE_INTEGER,
        sliceMs: 0,
        yieldToHost: () => {
          since = passes.sessionCalls - last;
          widest = Math.max(widest, since);
          last = passes.sessionCalls;
          return Promise.resolve();
        },
      },
    });
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blocking.layout));
    // One body entry here is at most about fifteen session calls. Waiting
    // synchronously for the first pass instead ran a whole page of entries
    // (over 180 calls) inside one host turn.
    expect(widest).toBeLessThan(40);
  }, 300_000);

  it('discovers first-pass destinations in work linear in the page count', async () => {
    const finalizations = async (pages: number) => {
      const model = documentModel(pages * 27, [[2, pageFloatingTable(468, 300, 2, 100, 0)]]);
      const run = open(model);
      const before = passes.finalizedPages;
      await layoutDocumentProgressively(run.input, run.services, run.options, { onPreview: () => {} });
      return passes.finalizedPages - before;
    };
    const small = await finalizations(11);
    const medium = await finalizations(21);
    const large = await finalizations(41);
    // Doubling the document may at most roughly double the finalized pages;
    // a full-prefix snapshot per closed page would quadruple them.
    expect(medium / small).toBeLessThan(2.5);
    expect(large / medium).toBeLessThan(2.5);
  }, 300_000);

  it('emits no stale page when cancelled during convergence', async () => {
    const model = anchoredImageModel();
    const { layout: final } = blockingLayout(model);
    const everyStep = (onYield: () => void) => ({
      now: () => Number.MAX_SAFE_INTEGER,
      sliceMs: 0,
      yieldToHost: () => { onYield(); return Promise.resolve(); },
    });

    let totalYields = 0;
    const complete = open(model);
    await layoutDocumentProgressively(complete.input, complete.services, complete.options, {
      onPreview: () => {},
      scheduler: everyStep(() => { totalYields += 1; }),
    });

    // Abort points inside the second and third passes, after publications
    // have started, as a destroyed viewer or a replaced load would.
    for (const fraction of [0.5, 0.7, 0.9]) {
      const abortAt = Math.floor(totalYields * fraction);
      const controller = new AbortController();
      const previews: ProgressiveLayoutPreview[] = [];
      let yields = 0;
      let publishedAfterAbort = false;
      const cancelled = open(model);
      await expect(layoutDocumentProgressively(
        cancelled.input,
        cancelled.services,
        cancelled.options,
        {
          onPreview: (preview) => {
            if (controller.signal.aborted) publishedAfterAbort = true;
            previews.push(preview);
          },
          scheduler: {
            ...everyStep(() => {
              yields += 1;
              if (yields === abortAt) controller.abort();
            }),
            signal: controller.signal,
          },
        },
      )).rejects.toBeInstanceOf(PaginationAbortError);
      expect(previews.length).toBeGreaterThan(0);
      expect(publishedAfterAbort).toBe(false);
      expectPublishedPagesFinal(previews, final);
    }
  }, 300_000);
});

describe('page-table destination cycle settlement', () => {
  const coupledTables = () => documentModel(33, [
    [13, pageFloatingTable(468, 72, 1, 360, 0)],
    [15, pageFloatingTable(468, 72, 1, 600, 0)],
    [33, pageFloatingTable(300, 72, 1, 600, 0)],
  ]);

  it('settles a near-page-tall table on the later page of its two-state cycle', () => {
    const model = documentModel(56, [
      [26, pageFloatingTable(524, 0, 1, 667, 0)],
      [55, pageFloatingTable(524, 0, 1, 667, 0)],
    ]);
    model.section.pageHeight = 842;
    const { layout } = blockingLayout(model);
    // Main alternates (0, 1) and (2, 3). Recovery pins those later floors;
    // subsequent ordinary collision admission may move a table farther.
    expect(tablePages(layout, 26).placedPage).toBe(2);
    expect(tablePages(layout, 55).placedPage).toBeGreaterThanOrEqual(3);
  });

  it('settles the coupled T1/T2/T3 cycle deterministically without overlapping body text', () => {
    const model = coupledTables();
    const { layout } = blockingLayout(model);
    expect(layoutFingerprint(blockingLayout(model).layout)).toBe(layoutFingerprint(layout));
    expect(layout.pages.flatMap((page) => page.layers.body)
      .filter((node) => node.kind === 'table')).toHaveLength(3);
    for (const page of layout.pages) {
      const tables = page.layers.body.filter((node) => node.kind === 'table');
      const lines = page.layers.body.flatMap((node) => node.kind === 'paragraph' ? node.lines : []);
      const bands = [...tables.map((table) => table.flowBounds), ...lines.map((line) => line.bounds)];
      for (const table of tables) {
        for (const b of bands) {
          const a = table.flowBounds;
          if (a === b) continue;
          expect(a.xPt >= b.xPt + b.widthPt || b.xPt >= a.xPt + a.widthPt
            || a.yPt >= b.yPt + b.heightPt || b.yPt >= a.yPt + a.heightPt).toBe(true);
        }
      }
    }
  });

  it('keeps cycle-settled progressive pages equal to the blocking result', async () => {
    const model = coupledTables();
    const { final, previews } = await progressiveRun(model);
    expect(layoutFingerprint(final)).toBe(layoutFingerprint(blockingLayout(model).layout));
    expectPublishedPagesFinal(previews, final);
  });
});

// Local Word controls exercise the parser-to-paginator path without publishing
// the non-redistributable fixtures. Placement fidelity is informational here:
// this regression guarantees termination, not Word's logical anchor rule.
describe('local page-table boundary controls', () => {
  for (const name of ['heights', 'source-boundaries', 'wrap-widths', 'zero-boundaries']) {
    const path = new URL(`../../public/private/docx/controls-1659/page-anchor-${name}.docx`, import.meta.url);
    it.skipIf(!existsSync(path))(`lays out the ${name} controls without error`, async () => {
      await init({ module_or_path: readFileSync(new URL('../wasm/docx_parser_bg.wasm', import.meta.url)) });
      const archive = new DocxArchive(readFileSync(path));
      let model: DocxDocumentModel;
      try {
        model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
      } finally {
        archive.free();
      }
      const { layout } = blockingLayout(model);
      const tables = layout.pages.flatMap((page) => page.layers.body
        .filter((node) => node.kind === 'table').map((node) => ({
          source: node.source.path, page: page.pageIndex + 1, bounds: node.flowBounds,
        })));
      expect(tables).toHaveLength(model.body.filter((node) => node.type === 'table').length);
      expect(layout.diagnostics.some((entry) => entry.code === 'NON_CONVERGENCE')).toBe(false);
    }, 30_000);
  }
});

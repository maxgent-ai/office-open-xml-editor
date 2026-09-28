import { describe, expect, it, vi } from 'vitest';

// Counts body pagination passes: each pass opens exactly one kernel session.
const passes = vi.hoisted(() => ({ opened: 0 }));
vi.mock('./runtime-state.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./runtime-state.js')>();
  return {
    ...actual,
    bodyLayoutKernelOf: (services: Parameters<typeof actual.bodyLayoutKernelOf>[0]) => {
      const kernel = actual.bodyLayoutKernelOf(services);
      return kernel && {
        openBodyLayoutSession: (...args: Parameters<typeof kernel.openBodyLayoutSession>) => {
          passes.opened += 1;
          return kernel.openBodyLayoutSession(...args);
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

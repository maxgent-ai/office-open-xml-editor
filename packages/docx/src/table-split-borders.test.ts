import { describe, it, expect, beforeAll } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableRow,
  DocxDocumentModel,
    SectionProps,
  BorderSpec,
} from './types';
import type { TableFragmentLayout } from './layout/table-pagination.js';

// ─────────────────────────────────────────────────────────────────────────────
// Split-edge border semantics for a MID-ROW page cut (fidelity round, agreed
// design). Measured Word ground truth across two document classes: a bordered
// form table DRAWS a full-width rule at the page-1 cut (even over a column
// whose cell has NO bottom border but a single top border), while a completely
// borderless table draws NOTHING at its cut. The unifying rule is ECMA-376
// §17.4.66: the cut is a shared horizontal edge between the leading piece's
// cell bottoms and a SYNTHETIC continuation sibling's cell tops (same source
// row specs, resolved as the next page's outer top) — none∨single → single,
// none∨none → nothing. The continuation piece still draws its own outer top on
// the next page. Row-boundary page cuts are untouched.
// ─────────────────────────────────────────────────────────────────────────────

interface StrokeSeg { x1: number; y1: number; x2: number; y2: number; width: number; color: string }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; strokes: StrokeSeg[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const strokes: StrokeSeg[] = [];
  let cur = { x: 0, y: 0 };
  let pending: { x1: number; y1: number; x2: number; y2: number } | null = null;
  let lineWidth = 1;
  let strokeStyle = '#000000';
  let fillStyle = '#000000';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get lineWidth() { return lineWidth; },
    set lineWidth(v: number) { lineWidth = v; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(v: string) { strokeStyle = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    textAlign: 'left' as CanvasTextAlign,
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() { pending = null; }, closePath() {},
    moveTo(x: number, y: number) { cur = { x, y }; },
    lineTo(x: number, y: number) { pending = { x1: cur.x, y1: cur.y, x2: x, y2: y }; cur = { x, y }; },
    stroke() {
      if (pending) strokes.push({ ...pending, width: lineWidth, color: strokeStyle });
    },
    fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, setTransform() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, strokes };
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
  };
});

const bs = (): BorderSpec => ({ style: 'single', width: 1, color: null } as BorderSpec);
const allEdges = () => ({ top: bs(), bottom: bs(), left: bs(), right: bs(), insideH: bs(), insideV: bs() });

function para(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 20, color: null, fontFamily: 'T',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 20, defaultFontFamily: 'T', widowControl: false,
  } as unknown as DocParagraph;
}

/** One-row bordered table whose single cell wraps to 4 lines (64 glyphs at 16/line, 80pt) — page body
 *  60pt tall, so the row splits mid-page at 3 lines. Grid 160pt == page width. */
function splitDoc(rows: DocTableRow[], pageHeight: number): DocxDocumentModel {
  const t: DocTable = {
    colWidths: [160], rows, borders: allEdges(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
  return {
    section: {
      pageWidth: 160, pageHeight,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...t } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { T: 'roman' },
  } as unknown as DocxDocumentModel;
}

const wrappingRow = (): DocTableRow => ({
  cells: [{
    content: [{ type: 'paragraph', ...para('あ'.repeat(64)) } as unknown as CellElement],
    colSpan: 1, vMerge: null, borders: allEdges(), background: null, vAlign: 'top', widthPt: null,
  }],
  rowHeight: null, rowHeightRule: 'auto', isHeader: false,
} as unknown as DocTableRow);

const shortRow = (text: string): DocTableRow => ({
  cells: [{
    content: [{ type: 'paragraph', ...para(text) } as unknown as CellElement],
    colSpan: 1, vMerge: null, borders: allEdges(), background: null, vAlign: 'top', widthPt: null,
  }],
  // This fixture isolates a page cut that lands exactly on a row boundary.
  // `exact` makes 20pt the complete §17.4.80 row box; an auto row also reserves
  // the resolved border footprint and therefore would not be a 20pt row.
  rowHeight: 20, rowHeightRule: 'exact', isHeader: false,
} as unknown as DocTableRow);

async function renderPage(layout: ReturnType<typeof layoutDocument>, pageIndex: number): Promise<StrokeSeg[]> {
  const { canvas, strokes } = makeRecordingCanvas();
  await paintLayoutPage(layout, pageIndex, canvas, { dpr: 1, scale: 1 });
  return strokes;
}

const horizontals = (strokes: StrokeSeg[]) =>
  strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.x2 - s.x1) > 10);

function retainedTableOnPage(layout: ReturnType<typeof layoutDocument>, pageIndex: number) {
  const table = layout.pages[pageIndex]?.layers.body.find((node) => node.kind === 'table');
  if (table?.kind !== 'table') throw new Error('expected retained table geometry');
  return table as TableFragmentLayout;
}

describe('mid-row page-cut border semantics (§17.4.66 conflict at the cut)', () => {
  it('draws the §17.4.66 winner at the page-1 cut and the continuation top on page 2', async () => {
    // A 4-line cell (80pt) in a 60pt body also needs the page-local top/cut
    // footprint. Two 20pt lines fit with the two half-rules; a third would make
    // the complete slice overflow.
    const model = splitDoc([wrappingRow()], 60);
    const layout = layoutDocument(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const firstFragment = retainedTableOnPage(layout, 0);
    expect(firstFragment.rows[0]?.cells[0]?.blocks[0]?.layout).toMatchObject({
      continuation: { lineStart: 0, continuesOnNext: true },
    });

    const p1 = await renderPage(layout, 0);
    const p2 = await renderPage(layout, 1);

    const p1H = horizontals(p1);
    // Top frame at y=0 present…
    expect(p1H.some((s) => Math.abs(s.y1 - 0) <= 1)).toBe(true);
    // …and the cut draws the conflict winner (single ∨ single) inside the
    // retained page-local fragment bottom (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT).
    const expectedCutY = firstFragment.flowBounds.yPt + firstFragment.flowBounds.heightPt - 0.5;
    const cut = p1H.filter((s) => Math.abs(s.y1 - expectedCutY) <= 0.1);
    expect(cut.length).toBe(1);
    expect(Math.min(cut[0].x1, cut[0].x2)).toBeLessThanOrEqual(1);
    expect(Math.max(cut[0].x1, cut[0].x2)).toBeGreaterThanOrEqual(159);

    // The continuation piece draws its own outer top on page 2 (y=0).
    const p2H = horizontals(p2);
    expect(p2H.some((s) => Math.abs(s.y1 - 0) <= 1)).toBe(true);
  });

  it('draws the cut rule over a NO-BOTTOM-border cell when its top border wins the conflict', async () => {
    // The decisive ground-truth datum: the form label cell has top/left/right
    // single but NO bottom — the cut still shows a full-width rule because the
    // synthetic continuation sibling's TOP (single) wins none ∨ single.
    const row = wrappingRow();
    (row.cells[0].borders as { bottom: BorderSpec | null }).bottom = null;
    const model = splitDoc([row], 60);
    const layout = layoutDocument(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const p1H = horizontals(await renderPage(layout, 0));
    const firstFragment = retainedTableOnPage(layout, 0);
    const expectedCutY = firstFragment.flowBounds.yPt + firstFragment.flowBounds.heightPt - 0.5;
    expect(p1H.filter((s) => Math.abs(s.y1 - expectedCutY) <= 0.1)).toHaveLength(1);
  });

  it('draws NOTHING at the cut of a borderless table', async () => {
    const row = wrappingRow();
    (row.cells[0] as { borders: unknown }).borders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
    const model = splitDoc([row], 60);
    (model.body[0] as unknown as DocTable).borders = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null } as DocTable['borders'];
    const layout = layoutDocument(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const p1H = horizontals(await renderPage(layout, 0));
    expect(p1H.filter((s) => Math.abs(s.y1 - 60) <= 1)).toHaveLength(0);
  });

  it('keeps the row-boundary page cut borders unchanged', async () => {
    // Three 20pt single-line rows in a 40pt page body: the cut falls BETWEEN
    // rows (no mid-row slice). Pin: a horizontal rule IS drawn at the page-1
    // fragment bottom, inside its box — the row's own boundary edge.
    const model = splitDoc([shortRow('a'), shortRow('b'), shortRow('c')], 40);
    const layout = layoutDocument(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    const firstFragment = retainedTableOnPage(layout, 0);
    expect(firstFragment.rows.every((row) => row.fragmentIndex === 0)).toBe(true);

    const p1 = await renderPage(layout, 0);
    const p1H = horizontals(p1);
    const fragmentBottomPt = firstFragment.flowBounds.yPt + firstFragment.flowBounds.heightPt;
    expect(p1H.filter((s) => Math.abs(s.y1 - (fragmentBottomPt - 0.5)) <= 0.1).length).toBeGreaterThan(0);
  });
});

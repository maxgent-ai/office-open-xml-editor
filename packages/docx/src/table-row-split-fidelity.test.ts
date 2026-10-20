import { beforeAll, describe, expect, it } from 'vitest';
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
} from './types';

interface TextCall {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly font: string;
}

interface StrokeSegment {
  readonly x1: number;
  readonly y1: number;
  readonly x2: number;
  readonly y2: number;
}

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  texts: TextCall[];
  strokes: StrokeSegment[];
  measureCount: () => number;
} {
  let font = '10px serif';
  let measures = 0;
  let current = { x: 0, y: 0 };
  let pending: StrokeSegment | null = null;
  let lineWidth = 1;
  let strokeStyle = '#000000';
  let fillStyle = '#000000';
  let translation = { x: 0, y: 0 };
  const translations: Array<{ x: number; y: number }> = [];
  const texts: TextCall[] = [];
  const strokes: StrokeSegment[] = [];
  const fontPx = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get lineWidth() { return lineWidth; },
    set lineWidth(value: number) { lineWidth = value; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(value: string) { strokeStyle = value; },
    get fillStyle() { return fillStyle; },
    set fillStyle(value: string) { fillStyle = value; },
    letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
    measureText(text: string) {
      measures++;
      const px = fontPx();
      return {
        width: [...text].length * px * 0.5,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() { translations.push({ ...translation }); },
    restore() { translation = translations.pop() ?? { x: 0, y: 0 }; },
    closePath() {}, fill() {}, fillRect() {}, strokeRect() {},
    clip() {}, rect() {}, scale() {}, setTransform() { translation = { x: 0, y: 0 }; },
    translate(x: number, y: number) {
      translation = { x: translation.x + x, y: translation.y + y };
    },
    rotate() {}, setLineDash() {},
    clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {}, drawImage() {},
    createLinearGradient() { return { addColorStop() {} }; },
    beginPath() { pending = null; },
    moveTo(x: number, y: number) { current = { x, y }; },
    lineTo(x: number, y: number) {
      pending = { x1: current.x, y1: current.y, x2: x, y2: y };
      current = { x, y };
    },
    stroke() {
      if (pending) strokes.push(pending);
    },
    fillText(text: string, x: number, y: number) {
      texts.push({ text, x: x + translation.x, y: y + translation.y, font });
    },
    strokeText(text: string, x: number, y: number) {
      texts.push({ text, x: x + translation.x, y: y + translation.y, font });
    },
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    texts,
    strokes,
    measureCount: () => measures,
  };
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
  };
});

function border() {
  return { style: 'single', width: 1, color: null } as const;
}

function borders() {
  return {
    top: border(), bottom: border(), left: border(), right: border(),
    insideH: border(), insideV: border(),
  };
}

function paragraph(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 20, color: null, fontFamily: 'T',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
      hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 20, defaultFontFamily: 'T', widowControl: false,
  } as unknown as DocParagraph;
}

function row(
  text: string,
  vAlign: 'top' | 'center' | 'bottom',
  rowHeight: number | null = null,
  rowHeightRule: 'auto' | 'exact' = 'auto',
): DocTableRow {
  return {
    cells: [{
      content: [{ type: 'paragraph', ...paragraph(text) } as unknown as CellElement],
      colSpan: 1, vMerge: null, borders: borders(), background: null, vAlign, widthPt: null,
    }],
    rowHeight,
    rowHeightRule,
    isHeader: false,
  } as unknown as DocTableRow;
}

function documentWithRow(
  tableRow: DocTableRow,
  pageHeight: number,
  pageWidth = 160,
): DocxDocumentModel {
  const table: DocTable = {
    colWidths: [pageWidth], rows: [tableRow], borders: borders(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
  return {
    section: {
      pageWidth, pageHeight,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...table } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { T: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderPage(
  layout: ReturnType<typeof layoutDocument>,
  pageIndex: number,
) {
  const recording = makeRecordingCanvas();
  await paintLayoutPage(layout, pageIndex, recording.canvas, { dpr: 1, scale: 1 });
  return recording;
}

function firstSliceOnPage(layout: ReturnType<typeof layoutDocument>, pageIndex: number) {
  const table = layout.pages[pageIndex]?.layers.body.find((node) => node.kind === 'table');
  if (table?.kind !== 'table') return undefined;
  const paragraph = table.rows[0]?.cells[0]?.blocks[0]?.layout;
  if (paragraph?.kind !== 'paragraph') return undefined;
  const continuation = paragraph.continuation;
  return continuation ? { start: continuation.lineStart, end: continuation.lineEnd } : undefined;
}

const splitText =
  '甲'.repeat(16) +
  '乙'.repeat(16) +
  '丙'.repeat(16) +
  '丁'.repeat(16);

describe('table row split fidelity — fragment-owned paint geometry', () => {
  it('retains width-balanced CJK label lines across table pagination', async () => {
    const labelRow = row('申　請　事　項　及　び　理　由', 'top');
    const model = documentWithRow(labelRow, 85, 12);
    model.settings = { balanceSingleByteDoubleByteWidth: true };
    model.section.docGridType = 'linesAndChars';
    model.section.docGridLinePitch = 20;
    model.section.docGridCharSpace = -4096;

    const layout = layoutDocument(model);
    const painted = await Promise.all(
      layout.pages.map((_, pageIndex) => renderPage(layout, pageIndex)),
    );
    const visibleByPage = painted.map((page) => page.texts
      .map((call) => call.text.replaceAll('　', ''))
      .join(''));
    expect({
      pages: layout.pages.length,
      slices: layout.pages.map((_, i) => firstSliceOnPage(layout, i)),
      visibleByPage,
    }).toEqual({
      pages: 2,
      slices: [{ start: 0, end: 4 }, { start: 4, end: 8 }],
      visibleByPage: ['申請事項', '及び理由'],
    });
    expect(visibleByPage.join('')).toBe('申請事項及び理由');
  });

  it('paints only the retained continuation [k, n) line window', async () => {
    const model = documentWithRow(row(splitText, 'top'), 60);
    const layout = layoutDocument(model);
    expect(layout.pages).toHaveLength(2);
    expect(firstSliceOnPage(layout, 0)).toEqual({ start: 0, end: 2 });
    expect(firstSliceOnPage(layout, 1)).toEqual({ start: 2, end: 4 });

    const page2 = await renderPage(layout, 1);
    const paintedText = page2.texts.map((call) => call.text).join('');
    expect(paintedText).toBe('丙'.repeat(16) + '丁'.repeat(16));
    expect(paintedText).not.toContain('甲');
    expect(paintedText).not.toContain('乙');
  });

  it('keeps every centered continuation baseline at or below its row-top rule', async () => {
    const model = documentWithRow(row(splitText, 'center'), 60);
    const layout = layoutDocument(model);
    expect(layout.pages).toHaveLength(2);
    expect(firstSliceOnPage(layout, 1)).toEqual({ start: 2, end: 4 });

    const page2 = await renderPage(layout, 1);
    const topRule = page2.strokes.find(
      (stroke) => Math.abs(stroke.y1 - stroke.y2) < 0.5 && Math.abs(stroke.x2 - stroke.x1) > 100,
    );
    expect(topRule).toBeDefined();
    expect(page2.texts.length).toBeGreaterThan(0);
    const minimumBaseline = Math.min(...page2.texts.map((call) => call.y));
    expect(minimumBaseline).toBeGreaterThanOrEqual(topRule?.y1 ?? 0);
  });

  it('keeps an unsplit centered cell deterministic and measurement-free', async () => {
    const model = documentWithRow(row('CENTER', 'center', 60, 'exact'), 80);
    const layout = layoutDocument(model);
    expect(layout.pages).toHaveLength(1);

    const first = await renderPage(layout, 0);
    const second = await renderPage(layout, 0);
    expect(second.texts).toEqual(first.texts);
    expect(first.texts.map(({ text, x, y }) => ({ text, x, y }))).toEqual([
      // Centred in the exact box below its 1pt top rule.
      { text: 'CENTER', x: 0, y: 36.5 },
    ]);
    expect(first.measureCount()).toBe(0);
    expect(second.measureCount()).toBe(0);
  });
});

import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  DocxDocumentModel,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocParagraph,
  SectionProps,
} from './types';

/**
 * ECMA-376 §17.4.50 `<w:tblInd>` — indentation added before the table's LEADING
 * edge (left in LTR, right in RTL/`bidiVisual`), shifting the table into the text
 * margin. A NEGATIVE value pulls the table OUTWARD past the leading margin toward
 * the page edge. Applies only when the resolved `jc` is left/leading.
 *
 * These tests pin the physical x-origin the renderer resolves for each direction
 * by capturing every `moveTo`/`lineTo` x-coordinate (cell borders are stroked via
 * retained border painter → `moveTo`/`lineTo`) and reading the table's left/right
 * extent from them. A vertical border is nudged ≤0.5 px by the crispness snap
 * (`crispOffset`), so the extent assertions allow a 0.75 px tolerance.
 */

/** Assert `actual` is within 0.75 px of `expected` (absorbs the ≤0.5 px crisp
 *  snap on a thin vertical border). */
function expectNear(actual: number, expected: number, msg: string): void {
  expect(Math.abs(actual - expected), `${msg} (got ${actual}, want ~${expected})`).toBeLessThanOrEqual(0.75);
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; xs: number[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const xs: number[] = [];
  let transform = { scaleX: 1, translateX: 0 };
  const transformStack: Array<typeof transform> = [];
  const record = (x: number) => xs.push(transform.translateX + x * transform.scaleX);
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { transformStack.push({ ...transform }); },
    restore() { transform = transformStack.pop() ?? { scaleX: 1, translateX: 0 }; },
    beginPath() {}, closePath() {},
    moveTo(x: number) { record(x); }, lineTo(x: number) { record(x); },
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {},
    scale(x: number) { transform.scaleX *= x; },
    translate(x: number) { transform.translateX += x * transform.scaleX; },
    setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, xs };
}

function solidBorder() { return { width: 0.5, color: '000000', style: 'single' }; }
function allSolid() {
  const b = solidBorder();
  return { top: b, bottom: b, left: b, right: b, insideH: b, insideV: b };
}
function bodyParagraph(text: string): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

/** A 1×1 solid-border table (single column `colW` pt wide) with the given
 *  `tblInd` (pt) and bidiVisual flag, on a 200×200 pt page with 20 pt margins
 *  (content [20, 180], width 160). */
function tableDoc(colW: number, tblInd: number | undefined, bidiVisual: boolean): DocxDocumentModel {
  const cell: DocTableCell = {
    content: [{ type: 'paragraph', ...bodyParagraph('x') }],
    colSpan: 1, vMerge: null, borders: allSolid(), background: null, vAlign: 'top',
    widthPt: colW,
  } as unknown as DocTableCell;
  const row: DocTableRow = {
    cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
  } as unknown as DocTableRow;
  const table: DocTable = {
    colWidths: [colW], rows: [row], borders: allSolid(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', tblInd, bidiVisual, widthPt: colW,
  } as unknown as DocTable;
  return {
    section: {
      pageWidth: 200, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...table }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('§17.4.50 tblInd — table indent from the leading margin', () => {
  it('retains the signed leading-edge origin for both visual directions', () => {
    const retained = (doc: DocxDocumentModel) => {
      const recording = makeRecordingCanvas();
      const services = createLayoutServices(doc, {
        measureContext: recording.canvas.getContext('2d') as CanvasRenderingContext2D,
      });
      return layoutDocument(doc, services, { currentDateMs: 0 }).pages[0]?.layers.body[0];
    };
    const ltr = retained(tableDoc(160, -10, false));
    const rtl = retained(tableDoc(160, -10, true));
    if (ltr?.kind !== 'table' || rtl?.kind !== 'table') {
      throw new Error('expected retained table geometry');
    }
    expect(ltr.flowBounds.xPt).toBeCloseTo(10, 6);
    expect(rtl.flowBounds.xPt).toBeCloseTo(30, 6);
  });

  it('LTR: negative tblInd pulls the LEFT origin outward past the left margin', async () => {
    // scale=1 (width 200 = pageWidth). content=[20,180]. colW=160 fits content.
    // No indent: left origin = contentX = 20. With tblInd=-10 → left origin = 10.
    const noInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, undefined, false), noInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.min(...noInd.xs), 20, 'LTR base left origin = contentX');

    const withInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, -10, false), withInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.min(...withInd.xs), 10, 'LTR left origin = contentX + tblInd = 20 + (-10)');
  });

  it('keeps exact-height row clipping open across the signed table extent', () => {
    const recording = makeRecordingCanvas();
    const doc = tableDoc(160, -10, false);
    const row = (doc.body[0] as DocTable).rows[0];
    if (!row) throw new Error('expected table row');
    row.rowHeight = 20;
    row.rowHeightRule = 'exact';
    const services = createLayoutServices(doc, {
      measureContext: recording.canvas.getContext('2d') as CanvasRenderingContext2D,
    });

    const retained = layoutDocument(doc, services, { currentDateMs: 0 }).pages[0]?.layers.body[0];

    if (retained?.kind !== 'table') throw new Error('expected retained table geometry');
    expect(retained.rows[0]?.cells[0]?.clipBounds).toEqual({
      xPt: 10,
      yPt: 20,
      widthPt: 170,
      heightPt: 20,
    });
  });

  it('preserves an authored width beyond the text band when the indented table fits the page', () => {
    // §17.18.87 allows AutoFit to override the preferred table width until the
    // table reaches the page width. The 180pt authored grid is wider than the
    // 160pt text band, but tblInd=-10 places it wholly inside the 200pt page:
    // [20 - 10, 20 - 10 + 180] = [10, 190]. It must not be pre-shrunk to the
    // text band before §17.4.50 placement is applied.
    const recording = makeRecordingCanvas();
    const doc = tableDoc(180, -10, false);
    const services = createLayoutServices(doc, {
      measureContext: recording.canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    const retained = layoutDocument(doc, services, { currentDateMs: 0 }).pages[0]?.layers.body[0];
    if (retained?.kind !== 'table') throw new Error('expected retained table geometry');

    expect(retained.flowBounds.xPt).toBeCloseTo(10, 6);
    expect(retained.flowBounds.widthPt).toBeCloseTo(180, 6);
    expect(retained.flowBounds.xPt + retained.flowBounds.widthPt).toBeCloseTo(190, 6);
  });

  it.each([
    ['LTR', false, { xPt: 10, widthPt: 190 }],
    ['RTL', true, { xPt: 0, widthPt: 190 }],
  ] as const)(
    '%s: caps AutoFit at the physical page edge after resolving a partial negative leading indent',
    (_direction, bidiVisual, expected) => {
      // The text band is [20, 180]. A -10pt leading indent moves the resolved
      // leading edge halfway into its 20pt page margin. AutoFit may therefore
      // use 190pt, but not the full 200pt page width: the latter would cross the
      // opposite physical page edge by 10pt in either visual direction.
      const recording = makeRecordingCanvas();
      const doc = tableDoc(220, -10, bidiVisual);
      const services = createLayoutServices(doc, {
        measureContext: recording.canvas.getContext('2d') as CanvasRenderingContext2D,
      });
      const retained = layoutDocument(doc, services, { currentDateMs: 0 }).pages[0]?.layers.body[0];
      if (retained?.kind !== 'table') throw new Error('expected retained table geometry');

      expect(retained.flowBounds.xPt).toBeCloseTo(expected.xPt, 6);
      expect(retained.flowBounds.widthPt).toBeCloseTo(expected.widthPt, 6);
      expect(retained.flowBounds.xPt).toBeGreaterThanOrEqual(0);
      expect(retained.flowBounds.xPt + retained.flowBounds.widthPt).toBeLessThanOrEqual(200);
    },
  );

  const autoWidthTable = (options: {
    sideMarginPt: number;
    indentPt: number;
    gridPt: number;
    compatibilityMode?: number;
    text?: string;
    preferredCell?: boolean;
    justification?: 'left' | 'center' | 'right';
  }) => {
    const source = tableDoc(options.gridPt, options.indentPt, false);
    const { widthPt: _preferred, ...sourceTable } = source.body[0] as DocTable;
    const sourceRow = sourceTable.rows[0]!;
    const { widthPt: _cellPreferred, ...sourceCell } = sourceRow.cells[0]!;
    const cell = options.text === undefined
      ? (options.preferredCell ? sourceRow.cells[0]! : sourceCell as DocTableCell)
      : { ...sourceCell, content: [{ type: 'paragraph', ...bodyParagraph(options.text) }] } as DocTableCell;
    const doc = {
      ...source,
      ...(options.compatibilityMode === undefined
        ? {}
        : { settings: { compatibilityMode: options.compatibilityMode } }),
      section: { ...source.section, marginLeft: options.sideMarginPt, marginRight: options.sideMarginPt },
      body: [{
        ...sourceTable,
        jc: options.justification ?? 'left',
        cellMarginLeft: 5.4,
        cellMarginRight: 5.4,
        rows: [{ ...sourceRow, cells: [cell] }],
      }],
    } as DocxDocumentModel;
    const recording = makeRecordingCanvas();
    const services = createLayoutServices(doc, {
      measureContext: recording.canvas.getContext('2d') as CanvasRenderingContext2D,
    });
    const retained = layoutDocument(doc, services, { currentDateMs: 0 }).pages[0]?.layers.body[0];
    if (retained?.kind !== 'table') throw new Error('expected retained table geometry');
    return retained.flowBounds;
  };

  it.each([
    // Word controls (auto tblW, 5.4pt outer cell margins, 468pt page, saved
    // grid past the band) scaled here to a 200pt page with a 210.8pt grid.
    // Mode 14 (and an omitted mode) hangs both outer margins outside the
    // band - tblInd text area; mode 15 keeps the table inside it.
    ['zero side margins, mode 14, overflow the page', undefined, 0, 0, { xPt: 0, widthPt: 210.8 }],
    ['a positive indent, mode 14', 14, 10, 20, { xPt: 30, widthPt: 170.8 }],
    ['a negative indent, mode 14, crosses the page edge', 14, 10, -20, { xPt: -10, widthPt: 210.8 }],
    ['zero side margins, mode 15', 15, 0, 0, { xPt: 0, widthPt: 200 }],
    ['a positive indent, mode 15', 15, 10, 20, { xPt: 30, widthPt: 160 }],
    ['a negative indent, mode 15, crosses the page edge', 15, 10, -20, { xPt: -10, widthPt: 200 }],
  ] as const)(
    'auto-width AutoFit: %s',
    (_case, compatibilityMode, sideMarginPt, indentPt, expected) => {
      const bounds = autoWidthTable({
        sideMarginPt, indentPt, gridPt: 210.8,
        ...(compatibilityMode === undefined ? {} : { compatibilityMode }),
      });
      // WORD_AUTOFIT_LEADING_INDENT_BAND: no physical page clamp.
      expect(bounds.xPt).toBeCloseTo(expected.xPt, 6);
      expect(bounds.widthPt).toBeCloseTo(expected.widthPt, 6);
    },
  );

  it('applies the mode ceiling independently of cell preferences', () => {
    // Cell preferences affect allocation, not the auto tblW occurrence limit.
    // Word PDF clipping bounds and matching-font line partitions also confirm
    // this for an all-dxa multirow table.
    const bounds = autoWidthTable({
      sideMarginPt: 10, indentPt: 5, gridPt: 190, compatibilityMode: 15,
      preferredCell: true,
    });
    expect(bounds.widthPt).toBeCloseTo(175, 6);
  });

  it('preserves the established ceiling for an unmeasured compatibility mode', () => {
    expect(autoWidthTable({
      sideMarginPt: 10, indentPt: 5, gridPt: 190, compatibilityMode: 12,
    }).widthPt).toBeCloseTo(190, 6);
  });

  it.each(['center', 'right'] as const)(
    '%s auto-width limit retains the indent in mode 14 and uses the full band in mode 15',
    (justification) => {
      const common = { sideMarginPt: 10, indentPt: 5, gridPt: 175, justification, text: 'x'.repeat(100) };
      expect(autoWidthTable({ ...common, compatibilityMode: 14 }).widthPt).toBeCloseTo(185.8, 6);
      expect(autoWidthTable({ ...common, compatibilityMode: 15 }).widthPt).toBeCloseTo(180, 6);
      // A wider grid separates the two mode-14 hypotheses, but no such Word
      // control exists. Preserve the established result outside their common
      // geometry instead of inferring a ceiling from these four controls.
      expect(autoWidthTable({ ...common, gridPt: 190, compatibilityMode: 14 }).widthPt).toBeCloseTo(190, 6);
      expect(autoWidthTable({ ...common, gridPt: 190, compatibilityMode: 15 }).widthPt).toBeCloseTo(190, 6);
    },
  );

  it.each([
    [14, 185.8],
    [15, 175],
  ] as const)(
    'mode %s: an unpreferred content column grows past a saved grid at band - tblInd',
    (compatibilityMode, expectedWidthPt) => {
      // The Word gridmatch controls save a grid of exactly band - tblInd. A
      // long first cell without tcW still fits to the mode's AutoFit ceiling.
      const bounds = autoWidthTable({
        sideMarginPt: 10, indentPt: 5, gridPt: 175, compatibilityMode,
        text: Array.from({ length: 40 }, () => 'ab').join(' '),
      });
      expect(bounds.xPt).toBeCloseTo(15, 6);
      expect(bounds.widthPt).toBeCloseTo(expectedWidthPt, 6);
    },
  );

  it('RTL (bidiVisual): negative tblInd pushes the RIGHT leading edge into the right margin', async () => {
    // No indent, bidiVisual, colW=160=content: table fills [20,180]; right edge 180.
    const noInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, undefined, true), noInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.max(...noInd.xs), 180, 'RTL base right edge = contentRight');

    // tblInd=-10, bidiVisual, colW=160. The table keeps its full 160 width and its
    // leading (RIGHT) edge sits 10 pt PAST the right margin: rightEdge =
    // contentX + contentW - tblInd = 20 + 160 - (-10) = 190. Left origin = 190-160 = 30.
    const withInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, -10, true), withInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.max(...withInd.xs), 190, 'RTL right edge = contentRight - tblInd = 190');
    expectNear(Math.min(...withInd.xs), 30, 'RTL left origin = rightEdge - tableW = 30');
  });

  it('applies positive tblInd after right alignment like Word ([MS-OI29500] 2.1.155)', async () => {
    // ECMA-376 §17.4.50 says jc='right' ignores tblInd, but Word explicitly
    // deviates: resolve right alignment first, then translate by the indent.
    const doc = tableDoc(100, 20, false);
    (doc.body[0] as unknown as DocTable).jc = 'right';
    const rec = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, rec.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.max(...rec.xs), 200, 'Word shifts the right-aligned table by tblInd');
  });
});

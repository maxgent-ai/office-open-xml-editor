import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { DocumentLayout } from './layout/types.js';
import type {
  BodyElement, DocTable, DocTableRow, DocTableCell,
  DocxTextRun, SectionProps, TblpPr, } from './types';

// ECMA-376 §17.4.57 — a FLOATING table (`<w:tblpPr>`) is positioned absolutely
// (out of flow) and is NOT confined to the text column: Word keeps its declared
// `<w:tblW>`/`<w:tblGrid>` width even when that exceeds the column band, letting
// the box extend into the page margins. We previously scaled such tables down to
// the text band via the block-table overflow cap. For a floating table the
// overflow cap is the page width (the physical constraint), not the column band;
// block tables keep the band cap unchanged.
//
// Deterministic stub canvas (glyph advance = charCount × fontPx). Copied from
// float-table-page-fit.test.ts.
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// pageWidth 600 / margins 75 ⇒ column band 450; page cap 600.
function section(): SectionProps {
  return {
    pageWidth: 600, pageHeight: 400,
    marginTop: 20, marginRight: 75, marginBottom: 20, marginLeft: 75,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false,
    background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
}

function cellPara(text: string): unknown {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  };
}

function tblp(): TblpPr {
  return {
    leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
    horzAnchor: 'margin', horzSpecified: true, tblpXSpec: 'center',
    vertAnchor: 'page', tblpX: 0, tblpY: 30,
  } as TblpPr;
}

/** A one-row table over `colWidthsPt` (pt). `layout`/`widthPt`/`tblpPr` control
 *  the branch under test. Short cell text keeps min-content below every grid col. */
function table(
  colWidthsPt: number[],
  opts: { layout?: 'fixed'; widthPt?: number; float?: boolean },
): BodyElement {
  const cells: DocTableCell[] = colWidthsPt.map((_, i) => ({
    content: [cellPara(`c${i}`)] as unknown as DocTableCell['content'],
    colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: null,
  } as DocTableCell));
  const rows: DocTableRow[] = [
    { cells, rowHeight: 20, rowHeightRule: 'exact', isHeader: false } as unknown as DocTableRow,
  ];
  const t: DocTable = {
    colWidths: colWidthsPt, rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    ...(opts.layout ? { layout: opts.layout } : {}),
    ...(opts.widthPt != null ? { widthPt: opts.widthPt } : {}),
    ...(opts.float ? { tblpPr: tblp() } : {}),
  } as unknown as DocTable;
  return { type: 'table', ...t } as unknown as BodyElement;
}

const resolvedWidth = (layout: DocumentLayout): number => {
  const table = layout.pages.flatMap((page) => page.layers.body).find((node) => node.kind === 'table');
  expect(table).toBeDefined();
  return table?.kind === 'table'
    ? table.columnWidthsPt.reduce((sum, widthPt) => sum + widthPt, 0)
    : 0;
};

describe('floating-table width cap (§17.4.57) — page width, not the column band', () => {
  it('keeps a FIXED-layout floating table at its full tblGrid width past the column band', () => {
    // Grid 520 > band 450, < page 600 ⇒ stays 520 (was scaled to 450).
    const layout = layoutBodyModel(
      [table([200, 120, 200], { layout: 'fixed', float: true })],
      section(), makeCtx(),
    );
    expect(resolvedWidth(layout)).toBeCloseTo(520, 1);
  });

  it('keeps an AUTOFIT floating table with a preferred tblW at its full grid width', () => {
    // Autofit + tblW=dxa (grid trusted): grid 520 > band 450 ⇒ stays 520.
    const layout = layoutBodyModel(
      [table([200, 120, 200], { widthPt: 520, float: true })],
      section(), makeCtx(),
    );
    expect(resolvedWidth(layout)).toBeCloseTo(520, 1);
  });

  it('clamps a floating table wider than the PAGE to the page width', () => {
    // Grid 700 > page 600 ⇒ scaled to 600 (the float cannot exceed the paper).
    const layout = layoutBodyModel(
      [table([300, 100, 300], { layout: 'fixed', float: true })],
      section(), makeCtx(),
    );
    expect(resolvedWidth(layout)).toBeCloseTo(600, 1);
  });

  it('keeps a NON-floating fixed table at its grid beyond the column band (#1881)', () => {
    // Windows Word 16.0 #1881 controls: in-flow fixed tracks are final.
    const layout = layoutBodyModel(
      [table([200, 120, 200], { layout: 'fixed' })],
      section(), makeCtx(),
    );
    expect(resolvedWidth(layout)).toBeCloseTo(520, 1);
  });

  it('leaves a floating table narrower than the band untouched (cap never bites)', () => {
    const layout = layoutBodyModel(
      [table([100, 100], { layout: 'fixed', float: true })],
      section(), makeCtx(),
    );
    expect(resolvedWidth(layout)).toBeCloseTo(200, 1);
  });
});

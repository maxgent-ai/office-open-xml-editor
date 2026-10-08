import { describe, expect, it, vi } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import type { DocumentLayout } from './layout/types.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
  TableBorders,
} from './types.js';

// An ordinary (in-flow) nested table laid out whole at its page origin
// because it holds §17.4.57 positioned children (table-pagination.ts
// layoutNestedWhole). Its rows' placement must cost work proportional to its
// rows, not to the square of them.
//
// Cost boundary: table row layout takes no layout service (layoutTable
// ignores `services`, and retained cell content is not re-measured), so no
// public diagnostic or service metric observes it. What these tests observe
// is the number of rows the real table layout engine is given, counted by a
// passthrough around its export: every call still runs and returns the real
// layout, and every caller in the pipeline is counted. That is a bound on
// full table layout calls only, not on the layout's whole work: the track
// arithmetic outside layoutTable (merge deficits, row tops summed over the
// whole table's final row heights) and the per-row boundary work around it are not
// observed, and nothing here measures CPU time. Their cost is reviewed
// statically, not proven by these counts.
// Synthetic linear metrics (one 10pt advance per character, 10pt lines) are a
// test measurement service only. Expected geometry comes from the authored
// operands and from the same document without the positioned children.

const work = vi.hoisted(() => ({ rows: 0 }));

vi.mock('./layout/table.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('./layout/table.js')>();
  return {
    ...actual,
    layoutTable: (...args: Parameters<typeof actual.layoutTable>) => {
      work.rows += args[0].rows.length;
      return actual.layoutTable(...args);
    },
  };
});

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * px(),
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

// A tall page so the whole outer row, and so the nested table, is laid out
// on one page: x∈[40,360], y∈[40,pageHeight−40] (2000 unless stated).
const section = (pageHeight = 2000) => ({
  pageWidth: 400, pageHeight,
  marginTop: 40, marginRight: 40, marginBottom: 40, marginLeft: 40,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
}) as SectionProps;

function paragraph(text: string): DocParagraph & { type: 'paragraph' } {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: text ? [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'NotInMetrics',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }] : [],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & { type: 'paragraph' };
}

const noBorders: TableBorders = {
  top: null, bottom: null, left: null, right: null, insideH: null, insideV: null,
};

function cell(content: CellElement[], vMerge: boolean | null = null): DocTableCell {
  return {
    content, colSpan: 1, vMerge, borders: { ...noBorders },
    background: null, vAlign: 'top', widthPt: null,
  };
}

function row(cells: DocTableCell[], exactPt: number | null = null): DocTableRow {
  return exactPt === null
    ? { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false }
    : { cells, rowHeight: exactPt, rowHeightRule: 'exact', isHeader: false };
}

function table(rows: DocTableRow[], extra: Partial<DocTable>): BodyElement {
  return {
    type: 'table', rows, borders: noBorders,
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 5, cellMarginRight: 5,
    jc: 'left', layout: 'fixed',
    ...extra,
  } as unknown as BodyElement;
}

/** A one-cell 50pt positioned table holding `text`, no margins or rules:
 * page x 300, and page y `yPt` or (null) its anchor paragraph's top. */
function positioned(text: string, yPt: number | null): CellElement {
  return table([row([cell([paragraph(text) as CellElement])])], {
    colWidths: [50], cellMarginLeft: 0, cellMarginRight: 0,
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'page', horzSpecified: true, tblpX: 300,
      vertAnchor: yPt === null ? 'text' : 'page', tblpY: yPt ?? 0,
    },
  } as Partial<DocTable>) as unknown as CellElement;
}

/**
 * The nested table: `blocks` repetitions of eight 10pt rows over a 60/60
 * grid. In each, column 1 of rows 2–4 is one vertical merge whose four lines
 * (40pt) outgrow its three rows (30pt), so its last row grows. With
 * `children`, row 0's first cell holds a page-positioned child at page
 * (300, 50) before n0, and the last repetition's row 5 — the first row below
 * that merge — holds a text-anchored child before its anchor.
 */
function nestedTable(blocks: number, children: boolean): BodyElement {
  const rows: DocTableRow[] = [];
  const last = 8 * (blocks - 1) + 5;
  for (let block = 0; block < blocks; block += 1) {
    for (let index = 0; index < 8; index += 1) {
      const r = 8 * block + index;
      const first: CellElement[] = [paragraph(`n${r}`) as CellElement];
      if (children && r === 0) first.unshift(positioned('pc', 50));
      if (children && r === last) first.unshift(positioned('tc', null));
      const second = index === 2
        ? cell(['a', 'b', 'c', 'd'].map((suffix) => paragraph(`m${block}${suffix}`) as CellElement), true)
        : index === 3 || index === 4
          ? cell([paragraph('') as CellElement], false)
          : cell([paragraph(`c${r}`) as CellElement]);
      rows.push(row([cell(first), second]));
    }
  }
  return table(rows, { colWidths: [60, 60] } as Partial<DocTable>);
}

/** A one-cell `widthPt` (50pt unless stated) page-positioned table holding
 * `text` at page (`xPt`, `yPt`), no margins or rules. */
function pageFixed(text: string, xPt: number, yPt: number, widthPt = 50): CellElement {
  return table([row([cell([paragraph(text) as CellElement])])], {
    colWidths: [widthPt], cellMarginLeft: 0, cellMarginRight: 0,
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'page', horzSpecified: true, tblpX: xPt,
      vertAnchor: 'page', tblpY: yPt,
    },
  } as Partial<DocTable>) as unknown as CellElement;
}

/** Page y of the nested table's top: below PRE's 10pt line at the 40pt margin. */
const NESTED_TOP = 50;

/**
 * The dense nested table: `rows` 10pt rows over a 60/60 grid. Column 1 is ONE
 * vertical merge from row 0 to the last row (so no row below row 0 is free of
 * a continuing merge), holding `rows + 2` 10pt lines: it outgrows its rows by
 * 20pt, which the last row takes. With `children`, every row's first cell
 * holds a page-positioned child before its own paragraph n<r>, beside the
 * grid at the row's own top (page y NESTED_TOP + 10r), at page x 240 or 300
 * on alternate rows: no two children meet, none reaches a cell or the outer
 * cell's text, so they change no row height and no line.
 */
function mergedTable(rows: number, children: boolean): BodyElement {
  return table(Array.from({ length: rows }, (_, r) => {
    const first: CellElement[] = [paragraph(`n${r}`) as CellElement];
    if (children) first.unshift(pageFixed(`k${r}`, r % 2 === 0 ? 240 : 300, NESTED_TOP + 10 * r));
    const second = r === 0
      ? cell(Array.from({ length: rows + 2 }, (__, line) => paragraph(`m${line}`) as CellElement), true)
      : cell([paragraph('') as CellElement], false);
    return row([cell(first), second]);
  }), { colWidths: [60, 60] } as Partial<DocTable>);
}

/** Row pitch of {@link exactMergedTable}. */
const EXACT_PT = 24;

/**
 * {@link mergedTable} with every row `exact` EXACT_PT: the same one merge
 * continuing through every row, whose `rows + 2` 10pt lines now fit its
 * rows, so no row grows. With `children`, row r's page-positioned child is at
 * its own row's top, page y NESTED_TOP + EXACT_PT·r (x 240 / 300 alternate):
 * none meets another, a cell or the outer cell's text.
 */
function exactMergedTable(rows: number, children: boolean): BodyElement {
  return table(Array.from({ length: rows }, (_, r) => {
    const first: CellElement[] = [paragraph(`n${r}`) as CellElement];
    if (children) first.unshift(pageFixed(`k${r}`, r % 2 === 0 ? 240 : 300, NESTED_TOP + EXACT_PT * r));
    const second = r === 0
      ? cell(Array.from({ length: rows + 2 }, (__, line) => paragraph(`m${line}`) as CellElement), true)
      : cell([paragraph('') as CellElement], false);
    return row([cell(first), second], EXACT_PT);
  }), { colWidths: [60, 60] } as Partial<DocTable>);
}

/**
 * {@link mergedTable} one level deeper: every row's first cell holds, instead
 * of a positioned child and n<r>, an ordinary in-flow one-cell 50pt table
 * (no margins or rules) holding them, so the positioned child is page-placed
 * content below the row's cells, not the dense table's own child. Column 1
 * is the same one merge continuing through every row. Each inner table is
 * one 10pt line high, so rows, merge growth and children are mergedTable's.
 */
function deepMergedTable(rows: number, children: boolean): BodyElement {
  return table(Array.from({ length: rows }, (_, r) => {
    const inner: CellElement[] = [paragraph(`n${r}`) as CellElement];
    if (children) inner.unshift(pageFixed(`k${r}`, r % 2 === 0 ? 240 : 300, NESTED_TOP + 10 * r));
    const holder = table([row([cell(inner)])], {
      colWidths: [50], cellMarginLeft: 0, cellMarginRight: 0,
    } as Partial<DocTable>) as unknown as CellElement;
    const second = r === 0
      ? cell(Array.from({ length: rows + 2 }, (__, line) => paragraph(`m${line}`) as CellElement), true)
      : cell([paragraph('') as CellElement], false);
    return row([cell([holder]), second]);
  }), { colWidths: [60, 60] } as Partial<DocTable>);
}

/** Lines of {@link mixedMergedTable}'s merge: more than any prefix of its
 * rows holds (row 0's 10pt line plus 24pt per exact row). */
const mixedMergeLines = (rows: number) => 3 * rows;

/**
 * {@link exactMergedTable} with row 0 `auto`: the same one merge continuing
 * through every row, now holding {@link mixedMergeLines} 10pt lines, more
 * than the rows prefixing any row — the whole table included — hold. Only
 * row 0 can take the deficit (every other row is exact), so it grows to the
 * merge's height less the exact rows'. With `children`, row r's first cell
 * holds a page-positioned child before n<r> at page y NESTED_TOP + EXACT_PT·r
 * (x 240 / 300 alternate): beside the grid, none meets another, a cell or
 * the outer cell's text.
 */
function mixedMergedTable(rows: number, children: boolean): BodyElement {
  return table(Array.from({ length: rows }, (_, r) => {
    const first: CellElement[] = [paragraph(`n${r}`) as CellElement];
    if (children) first.unshift(pageFixed(`k${r}`, r % 2 === 0 ? 240 : 300, NESTED_TOP + EXACT_PT * r));
    const second = r === 0
      ? cell(Array.from({ length: mixedMergeLines(rows) }, (__, line) => paragraph(`m${line}`) as CellElement), true)
      : cell([paragraph('') as CellElement], false);
    return row([cell(first), second], r === 0 ? null : EXACT_PT);
  }), { colWidths: [60, 60] } as Partial<DocTable>);
}

/**
 * Four auto rows over a 60/60 grid. Row r's first cell holds an ordinary
 * one-cell 50pt table (no margins or rules) holding n<r>; column 1 is one
 * vertical merge through all four rows holding six 10pt lines m0–m5 (60pt),
 * outgrowing the four rows. Only row 1 differs: its first cell is aligned
 * `vAlign`, and its inner table holds, before n1, a 50pt page-positioned
 * child k1 at page (40, 60) — the inner cell's own column and row 1's top —
 * so n1 must clear it.
 */
function alignedDeepTable(vAlign: DocTableCell['vAlign']): BodyElement {
  return table(Array.from({ length: 4 }, (_, r) => {
    const inner: CellElement[] = [paragraph(`n${r}`) as CellElement];
    if (r === 1) inner.unshift(pageFixed('k1', 40, 60));
    const holder = table([row([cell(inner)])], {
      colWidths: [50], cellMarginLeft: 0, cellMarginRight: 0,
    } as Partial<DocTable>) as unknown as CellElement;
    const second = r === 0
      ? cell(Array.from({ length: 6 }, (__, line) => paragraph(`m${line}`) as CellElement), true)
      : cell([paragraph('') as CellElement], false);
    const first = cell([holder]);
    return row([r === 1 ? { ...first, vAlign } : first, second]);
  }), { colWidths: [60, 60] } as Partial<DocTable>);
}

/**
 * Four auto rows over a 60/60 grid, every cell margin and rule 0. Column 1 is
 * one vertical merge through all four rows holding six 10pt lines m0–m5
 * (60pt). Row 0's first cell, aligned `vAlign`, holds its own (direct, no
 * holder table) 60pt page-positioned child k0 at page (40, 50) — over that
 * cell's whole column at row 0's top, left of column 1 — before n0, so n0
 * must clear it; rows 1–3 hold n1–n3.
 */
function directChildMergeTable(vAlign: DocTableCell['vAlign']): BodyElement {
  return table(Array.from({ length: 4 }, (_, r) => {
    const first = r === 0
      ? { ...cell([pageFixed('k0', 40, NESTED_TOP, 60), paragraph('n0') as CellElement]), vAlign }
      : cell([paragraph(`n${r}`) as CellElement]);
    const second = r === 0
      ? cell(Array.from({ length: 6 }, (__, line) => paragraph(`m${line}`) as CellElement), true)
      : cell([paragraph('') as CellElement], false);
    return row([first, second]);
  }), { colWidths: [60, 60], cellMarginLeft: 0, cellMarginRight: 0 } as Partial<DocTable>);
}

function layout(blocks: number, children: boolean): DocumentLayout {
  return layoutNested(nestedTable(blocks, children));
}

function layoutNested(nested: BodyElement, pageHeight?: number): DocumentLayout {
  return layoutBody([table([row([cell([
    paragraph('pre') as CellElement,
    nested as unknown as CellElement,
    paragraph('post') as CellElement,
  ])])], { colWidths: [320] } as Partial<DocTable>)], pageHeight);
}

/** `body` laid out directly in the section, with no outer cell. */
function layoutBody(body: readonly BodyElement[], pageHeight?: number): DocumentLayout {
  const model = {
    section: section(pageHeight),
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    footnotes: [], endnotes: [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  return layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
}

type Box = Readonly<{ text: string; x: number; y: number }>;

function runs(result: DocumentLayout): readonly Box[] {
  expect(result.pages).toHaveLength(1);
  return textRunGeometryForPage(result, 0).map((run) => ({
    text: run.placement.text,
    x: run.placement.bounds.xPt + run.pointToPage.e,
    y: run.placement.bounds.yPt + run.pointToPage.f,
  }));
}

function at(boxes: readonly Box[], text: string): Box {
  const matches = boxes.filter((box) => box.text === text);
  expect(matches, text).toHaveLength(1);
  return matches[0]!;
}

describe('a whole nested table holding positioned children', () => {
  const laidOut = (blocks: number) => {
    work.rows = 0;
    const result = layout(blocks, true);
    return { result, rows: work.rows };
  };

  it('places its rows, merges and children as the same table without children places its rows', () => {
    for (const blocks of [2, 16]) {
      const boxes = runs(layout(blocks, true));
      const control = runs(layout(blocks, false));
      // Every row top, the merges' growth and the outer cell's own
      // paragraphs: the positioned children take no flow and neither of
      // them reaches a cell (both lie right of the 120pt nested grid).
      expect(boxes.filter((box) => box.text !== 'pc' && box.text !== 'tc'), `${blocks}`).toEqual(control);
      // The page-positioned child is at its authored page position.
      expect(at(boxes, 'pc'), `${blocks}`).toMatchObject({ x: 300, y: 50 });
      // The text-anchored child (tblpY 0) is at its anchor's top: the row
      // below the last merge, whose top includes every merge's growth.
      const anchor = `n${8 * (blocks - 1) + 5}`;
      expect(at(boxes, 'tc'), `${blocks}`).toMatchObject({ x: 300, y: at(control, anchor).y });
    }
  });

  it('gives the table layout engine rows in proportion to its rows', () => {
    // 16 and 128 rows of one repeated pattern. Rows given to layoutTable that
    // grow with the row count, plus any fixed amount, grow at most eightfold;
    // laying out every row's prefix again grows with its square. Track
    // arithmetic outside layoutTable is not observed (see the cost boundary).
    const small = laidOut(2);
    const large = laidOut(16);
    expect(small.rows).toBeGreaterThan(0);
    expect(large.rows).toBeLessThanOrEqual(8 * small.rows);
  });
});

describe('a whole nested table with a child in every row under one continuing merge', () => {
  // The dense counterpart: the same ordinary nested whole-table path, but
  // every row hosts a positioned child and a vertical merge continues through
  // every row, so no prefix of the rows has the whole table's tracks: each
  // hosting row's offset is read from the whole table's final layout
  // (table-pagination.ts layoutNestedWholeAt). The bound is a resource
  // policy, not Word geometry.

  it('places every row, the merge and every child as the same table without children', () => {
    for (const rows of [16, 128]) {
      const boxes = runs(layoutNested(mergedTable(rows, true)));
      const control = runs(layoutNested(mergedTable(rows, false)));
      // Authored operands: every row 10pt from the nested top, the merge's
      // 20pt growth taken by the last row, POST below the grown table.
      for (let r = 0; r < rows; r += 1) {
        expect(at(control, `n${r}`), `${rows} n${r}`).toMatchObject({ y: NESTED_TOP + 10 * r });
      }
      expect(at(control, `m${rows + 1}`), `${rows}`).toMatchObject({ y: NESTED_TOP + 10 * (rows + 1) });
      expect(at(control, 'post'), `${rows}`).toMatchObject({ y: NESTED_TOP + 10 * (rows + 2) });
      // The children take no flow and meet no text: every other run, in
      // reading (source) order, is the control's.
      expect(boxes.filter((box) => !/^k\d+$/.test(box.text)), `${rows}`).toEqual(control);
      // Each child at its authored page position, placed once.
      for (let r = 0; r < rows; r += 1) {
        expect(at(boxes, `k${r}`), `${rows} k${r}`)
          .toMatchObject({ x: r % 2 === 0 ? 240 : 300, y: NESTED_TOP + 10 * r });
      }
    }
  });

  it('gives the table layout engine rows in proportion to its rows', () => {
    // Full table layout calls only (the metric above); track arithmetic is
    // not observed. 16 and 128 rows: rows given to layoutTable that grow
    // linearly, plus any fixed amount, grow at most eightfold; laying out the
    // prefix from the merge's start for every hosting row grows with the
    // square. The children-free control's own work is the baseline the
    // children may add to linearly.
    //
    // Not observed by this metric (no layout service, no row layout): how a
    // row finds its own children (table-pagination.ts floatingTableIndexOf,
    // indexed once per table by host cell, so per row it is proportional to
    // the row's cells and children) and each child's final-frame resolution
    // against a registry holding every earlier child (the page float
    // registry's collision rule, rows × children with a child per row). The
    // registry cost is named here, not bounded.
    const workOf = (rows: number, children: boolean) => {
      work.rows = 0;
      layoutNested(mergedTable(rows, children));
      return work.rows;
    };
    const small = workOf(16, true);
    const large = workOf(128, true);
    expect(small).toBeGreaterThan(workOf(16, false));
    expect(large).toBeLessThanOrEqual(8 * small);
  });
});

describe('a whole nested table with a child in every exact row under one continuing merge', () => {
  // The dense case with every row `exact`: an exact row can take no merge
  // deficit, so its top is not settled by the rule the auto rows above use,
  // and no row below the first is free of the continuing merge either. The
  // bound is the same resource policy, not Word geometry. 128 exact 24pt rows
  // need a page taller than the default: 3200pt, body y∈[40,3160].
  const PAGE_HEIGHT = 3200;

  it('places every row, the merge and every child as the same table without children', () => {
    for (const rows of [16, 128]) {
      const boxes = runs(layoutNested(exactMergedTable(rows, true), PAGE_HEIGHT));
      const control = runs(layoutNested(exactMergedTable(rows, false), PAGE_HEIGHT));
      // Authored operands: every row EXACT_PT from the nested top, the merge's
      // lines 10pt apart from its top (no growth), POST below the last row.
      for (let r = 0; r < rows; r += 1) {
        expect(at(control, `n${r}`), `${rows} n${r}`).toMatchObject({ y: NESTED_TOP + EXACT_PT * r });
      }
      expect(at(control, `m${rows + 1}`), `${rows}`).toMatchObject({ y: NESTED_TOP + 10 * (rows + 1) });
      expect(at(control, 'post'), `${rows}`).toMatchObject({ y: NESTED_TOP + EXACT_PT * rows });
      expect(boxes.filter((box) => !/^k\d+$/.test(box.text)), `${rows}`).toEqual(control);
      for (let r = 0; r < rows; r += 1) {
        expect(at(boxes, `k${r}`), `${rows} k${r}`)
          .toMatchObject({ x: r % 2 === 0 ? 240 : 300, y: NESTED_TOP + EXACT_PT * r });
      }
    }
  });

  it('gives the table layout engine rows in proportion to its rows', () => {
    // Full table layout calls only; track arithmetic is not observed.
    const workOf = (rows: number, children: boolean) => {
      work.rows = 0;
      layoutNested(exactMergedTable(rows, children), PAGE_HEIGHT);
      return work.rows;
    };
    const small = workOf(16, true);
    const large = workOf(128, true);
    expect(small).toBeGreaterThan(workOf(16, false));
    expect(large).toBeLessThanOrEqual(8 * small);
  });
});

describe('a whole nested table whose every row holds a table with a child, under one continuing merge', () => {
  // The dense case one level deeper: the dense table's rows hold no
  // positioned child of their own, but each first cell's in-flow table does,
  // so every row's inner table is placed whole at the page origin the dense
  // table's whole final layout gives it (table-pagination.ts
  // rowPagePlacement, read from layoutNestedWholeAt's layout, not from a
  // prefix of the rows). The bound is the same resource policy, not Word
  // geometry.

  it('places every row, the merge and every child as the same table without children', () => {
    for (const rows of [16, 128]) {
      const boxes = runs(layoutNested(deepMergedTable(rows, true)));
      const control = runs(layoutNested(deepMergedTable(rows, false)));
      // Authored operands, as the shallow dense case: every row 10pt from the
      // nested top, the merge's 20pt growth taken by the last row.
      for (let r = 0; r < rows; r += 1) {
        expect(at(control, `n${r}`), `${rows} n${r}`).toMatchObject({ y: NESTED_TOP + 10 * r });
      }
      expect(at(control, `m${rows + 1}`), `${rows}`).toMatchObject({ y: NESTED_TOP + 10 * (rows + 1) });
      expect(at(control, 'post'), `${rows}`).toMatchObject({ y: NESTED_TOP + 10 * (rows + 2) });
      expect(boxes.filter((box) => !/^k\d+$/.test(box.text)), `${rows}`).toEqual(control);
      for (let r = 0; r < rows; r += 1) {
        expect(at(boxes, `k${r}`), `${rows} k${r}`)
          .toMatchObject({ x: r % 2 === 0 ? 240 : 300, y: NESTED_TOP + 10 * r });
      }
    }
  });

  it('gives the table layout engine rows in proportion to its rows', () => {
    // Every full table layout call counts, the inner tables' included (one
    // row each, a fixed number of times per placed row); track arithmetic is
    // not observed.
    const workOf = (rows: number, children: boolean) => {
      work.rows = 0;
      layoutNested(deepMergedTable(rows, children));
      return work.rows;
    };
    const small = workOf(16, true);
    const large = workOf(128, true);
    expect(small).toBeGreaterThan(workOf(16, false));
    expect(large).toBeLessThanOrEqual(8 * small);
  });
});

describe('a whole nested table with a child in every row, an auto first row and exact rows under one continuing merge', () => {
  // The merge outgrows every prefix of the rows, the whole table included,
  // and only row 0 can take its deficit, so every row's top depends on the
  // merge's growth landing on row 0. The bound is the same resource policy,
  // not Word geometry. The cost this shape stresses is track arithmetic (the
  // deficit resolved over the rows and each row's top summed from the whole
  // table's final row heights), which runs outside layoutTable: no count here
  // observes it, so its bound is reviewed statically. 128 rows need a 4000pt
  // page, body y∈[40,3960].
  const PAGE_HEIGHT = 4000;

  it('places every row, the merge and every child as the same table without children', () => {
    for (const rows of [16, 128]) {
      const boxes = runs(layoutNested(mixedMergedTable(rows, true), PAGE_HEIGHT));
      const control = runs(layoutNested(mixedMergedTable(rows, false), PAGE_HEIGHT));
      // Authored operands: the table is the merge's height (its lines 10pt
      // apart from its top); row 0 takes all the growth, so it is that height
      // less the exact rows', and every later row is EXACT_PT below the last.
      const mergePt = 10 * mixedMergeLines(rows);
      const firstRowPt = mergePt - EXACT_PT * (rows - 1);
      expect(at(control, 'n0'), `${rows}`).toMatchObject({ y: NESTED_TOP });
      for (let r = 1; r < rows; r += 1) {
        expect(at(control, `n${r}`), `${rows} n${r}`)
          .toMatchObject({ y: NESTED_TOP + firstRowPt + EXACT_PT * (r - 1) });
      }
      for (let line = 0; line < mixedMergeLines(rows); line += 1) {
        expect(at(control, `m${line}`), `${rows} m${line}`).toMatchObject({ y: NESTED_TOP + 10 * line });
      }
      expect(at(control, 'post'), `${rows}`).toMatchObject({ y: NESTED_TOP + mergePt });
      // The children take no flow and meet no text.
      expect(boxes.filter((box) => !/^k\d+$/.test(box.text)), `${rows}`).toEqual(control);
      // Each child at its authored page position, placed once.
      for (let r = 0; r < rows; r += 1) {
        expect(at(boxes, `k${r}`), `${rows} k${r}`)
          .toMatchObject({ x: r % 2 === 0 ? 240 : 300, y: NESTED_TOP + EXACT_PT * r });
      }
    }
  });

  it('gives the table layout engine rows in proportion to its rows', () => {
    // Full table layout calls only. This does not observe the per-row track
    // arithmetic the shape stresses and does not show its cost linear; it
    // shows only that no row's prefix is laid out again in full.
    const workOf = (rows: number, children: boolean) => {
      work.rows = 0;
      layoutNested(mixedMergedTable(rows, children), PAGE_HEIGHT);
      return work.rows;
    };
    const small = workOf(16, true);
    const large = workOf(128, true);
    expect(small).toBeGreaterThan(workOf(16, false));
    expect(large).toBeLessThanOrEqual(8 * small);
  });
});

describe('a whole nested table whose aligned row holds a table with a child over its anchor', () => {
  // Row 1's inner table holds k1 over its own column at row 1's top, so n1
  // wraps below k1. Where the inner table sits depends on row 1's track and
  // its first cell's vAlign: placed against the rows above it alone, the
  // merge's 20pt deficit lands on row 1, a centered or bottom-aligned inner
  // table sits lower and n1 seems to clear k1; in the whole table the deficit
  // is the last row's. The inner table's page origin, and so n1's wrap, is
  // the one its final track and vAlign give. With n1 below k1 the inner table
  // is 20pt, as tall as row 1, so every alignment gives one geometry.

  it('wraps the anchor below the child at its top-aligned row geometry', () => {
    const boxes = runs(layoutNested(alignedDeepTable('top')));
    // Authored operands: PRE's 10pt line at the 40pt margin, the nested table
    // below it at NESTED_TOP; row 0 10pt; row 1 from 60, n1 below k1 (60–70)
    // so 20pt; rows 2 and 3 from 80 and 90, row 3 taking the merge's 10pt
    // deficit (60pt less 50pt) to end at 110, where POST is.
    expect(at(boxes, 'pre')).toMatchObject({ y: 40 });
    expect(at(boxes, 'k1')).toMatchObject({ x: 40, y: 60 });
    expect(at(boxes, 'n0')).toMatchObject({ y: NESTED_TOP });
    const n1 = at(boxes, 'n1');
    // k1 covers the inner cell's column, so n1 cannot sit beside it.
    expect(n1.x).toBeGreaterThanOrEqual(40);
    expect(n1.x).toBeLessThan(90);
    expect(n1.y).toBe(70);
    expect(at(boxes, 'n2')).toMatchObject({ y: 80 });
    expect(at(boxes, 'n3')).toMatchObject({ y: 90 });
    for (let line = 0; line < 6; line += 1) {
      expect(at(boxes, `m${line}`), `m${line}`).toMatchObject({ y: NESTED_TOP + 10 * line });
    }
    expect(at(boxes, 'post')).toMatchObject({ y: 110 });
  });

  it('wraps the anchor below the child when the row is centered or bottom-aligned', () => {
    const control = runs(layoutNested(alignedDeepTable('top')));
    for (const vAlign of ['center', 'bottom'] as const) {
      const boxes = runs(layoutNested(alignedDeepTable(vAlign)));
      // The child once, at its authored page position.
      expect(at(boxes, 'k1'), vAlign).toMatchObject({ x: 40, y: 60 });
      // n1 clears the child's bottom.
      expect(at(boxes, 'n1').y, vAlign).toBeGreaterThanOrEqual(70);
      // Every run, in reading (source) order, is the top-aligned row's.
      expect(boxes, vAlign).toEqual(control);
    }
  });
});

describe('a whole nested table whose aligned row holds its own child over its anchor, under a merge it starts', () => {
  // Row 0 hosts k0 directly and starts the merge. Its offset is 0 in every
  // layout, but its height is not: laid out alone, the merge restarting in it
  // gives row 0 all 60pt, so a centered or bottom-aligned cell sits n0 lower,
  // seemingly clear of k0; in the whole table the merge's deficit is the last
  // row's. The row's prepared content, and so n0's wrap, is the one its final
  // track and vAlign give. With n0 below k0 the cell's content is 20pt, as
  // tall as row 0, so every alignment gives one geometry.

  /** Page x of the nested grid's column 0: the outer table at the 40pt
   * margin, its 5pt cell margin, the nested table's 0pt one. */
  const NESTED_LEFT = 45;

  it('wraps the anchor below the child at its top-aligned row geometry', () => {
    const boxes = runs(layoutNested(directChildMergeTable('top')));
    // Authored operands: PRE's 10pt line at the 40pt margin, the nested table
    // below it at NESTED_TOP; k0 over column 0 from 50 to 60, so n0 at 60 and
    // row 0 20pt; rows 1–3 from 70, 80, 90, row 3 taking the merge's 10pt
    // deficit (60pt less 50pt) to end at 110, where POST is.
    expect(at(boxes, 'pre')).toMatchObject({ y: 40 });
    expect(at(boxes, 'k0')).toMatchObject({ x: 40, y: NESTED_TOP });
    expect(at(boxes, 'n0')).toMatchObject({ x: NESTED_LEFT, y: 60 });
    expect(at(boxes, 'n1')).toMatchObject({ x: NESTED_LEFT, y: 70 });
    expect(at(boxes, 'n2')).toMatchObject({ x: NESTED_LEFT, y: 80 });
    expect(at(boxes, 'n3')).toMatchObject({ x: NESTED_LEFT, y: 90 });
    for (let line = 0; line < 6; line += 1) {
      expect(at(boxes, `m${line}`), `m${line}`).toMatchObject({ y: NESTED_TOP + 10 * line });
    }
    expect(at(boxes, 'post')).toMatchObject({ y: 110 });
  });

  it('wraps the anchor below the child when the row is centered or bottom-aligned', () => {
    const control = runs(layoutNested(directChildMergeTable('top')));
    for (const vAlign of ['center', 'bottom'] as const) {
      const boxes = runs(layoutNested(directChildMergeTable(vAlign)));
      // The child once, at its authored page position.
      expect(at(boxes, 'k0'), vAlign).toMatchObject({ x: 40, y: NESTED_TOP });
      // n0 clears the child's bottom.
      expect(at(boxes, 'n0').y, vAlign).toBeGreaterThanOrEqual(60);
      // Every run, in reading (source) order, is the top-aligned row's.
      expect(boxes, vAlign).toEqual(control);
    }
  });
});

describe('a body table whose aligned row holds a table with a child over its anchor', () => {
  // alignedDeepTable directly in the body, not in an outer cell: the table
  // takes the paginated path (table-pagination.ts takeTableFragment) even
  // though all four rows fit the 2000pt page, and the fragment selected holds
  // every row. Row 1's inner table is placed whole at a page origin, and its
  // origin — so n1's wrap below k1 — is the one the selected fragment's final
  // track and vAlign give, not one from the rows above row 1 alone (where the
  // merge's deficit lands on row 1 and a centered or bottom-aligned inner
  // table sits lower). With n1 below k1 the inner table is 20pt, as tall as
  // row 1, so every alignment gives one geometry.
  const layoutRoot = (vAlign: DocTableCell['vAlign']) => layoutBody([
    paragraph('pre') as BodyElement,
    alignedDeepTable(vAlign),
    paragraph('post') as BodyElement,
  ]);

  /** Page x of the inner tables' cells: the table at the 40pt margin, its
   * 5pt cell margin, the inner table's 0pt one. */
  const ROOT_LEFT = 45;

  it('wraps the anchor below the child at its top-aligned row geometry', () => {
    const boxes = runs(layoutRoot('top'));
    // Authored operands: PRE's 10pt line at the 40pt margin, the table below
    // it at 50; row 0 10pt; row 1 from 60, n1 below k1 (60–70) so 20pt; rows
    // 2 and 3 from 80 and 90, row 3 taking the merge's 10pt deficit (60pt
    // less 50pt) to end at 110, where POST is.
    expect(at(boxes, 'pre')).toMatchObject({ y: 40 });
    expect(at(boxes, 'k1')).toMatchObject({ x: 40, y: 60 });
    expect(at(boxes, 'n0')).toMatchObject({ x: ROOT_LEFT, y: 50 });
    // k1 covers the inner cell's column, so n1 cannot sit beside it.
    expect(at(boxes, 'n1')).toMatchObject({ x: ROOT_LEFT, y: 70 });
    expect(at(boxes, 'n2')).toMatchObject({ x: ROOT_LEFT, y: 80 });
    expect(at(boxes, 'n3')).toMatchObject({ x: ROOT_LEFT, y: 90 });
    for (let line = 0; line < 6; line += 1) {
      expect(at(boxes, `m${line}`), `m${line}`).toMatchObject({ y: 50 + 10 * line });
    }
    expect(at(boxes, 'post')).toMatchObject({ y: 110 });
  });

  it('wraps the anchor below the child when the row is centered or bottom-aligned', () => {
    const control = runs(layoutRoot('top'));
    for (const vAlign of ['center', 'bottom'] as const) {
      const boxes = runs(layoutRoot(vAlign));
      // The child once, at its authored page position.
      expect(at(boxes, 'k1'), vAlign).toMatchObject({ x: 40, y: 60 });
      // n1 clears the child's bottom.
      expect(at(boxes, 'n1').y, vAlign).toBeGreaterThanOrEqual(70);
      // Every run, in reading (source) order, is the top-aligned row's.
      expect(boxes, vAlign).toEqual(control);
    }
  });
});

describe('a body table whose aligned row holds its own child over its anchor, under a merge it starts', () => {
  // directChildMergeTable directly in the body: the paginated path
  // (table-pagination.ts takeTableFragment), whose selected fragment holds
  // every row. Row 0 resolves k0 itself (finalFrameRow) before the fragment
  // is laid out; by itself row 0 would take the whole merge (60pt) and a
  // centered or bottom-aligned n0 would sit below k0, while the fragment's
  // final track gives the deficit to the last row. n0's wrap is the one the
  // materialized fragment's track and vAlign give. With n0 below k0 the
  // cell's content is 20pt, as tall as row 0, so every alignment gives one
  // geometry.
  const layoutRoot = (vAlign: DocTableCell['vAlign']) => layoutBody([
    paragraph('pre') as BodyElement,
    directChildMergeTable(vAlign),
    paragraph('post') as BodyElement,
  ]);

  it('wraps the anchor below the child at its top-aligned row geometry', () => {
    const boxes = runs(layoutRoot('top'));
    // Authored operands: PRE's 10pt line at the 40pt margin, the table below
    // it at 50, its column 0 at the 40pt margin (no cell margins); k0 over
    // column 0 from 50 to 60, so n0 at 60 and row 0 20pt; rows 1–3 from 70,
    // 80, 90, row 3 taking the merge's 10pt deficit to end at 110, where POST
    // is.
    expect(at(boxes, 'pre')).toMatchObject({ y: 40 });
    expect(at(boxes, 'k0')).toMatchObject({ x: 40, y: 50 });
    expect(at(boxes, 'n0')).toMatchObject({ x: 40, y: 60 });
    expect(at(boxes, 'n1')).toMatchObject({ x: 40, y: 70 });
    expect(at(boxes, 'n2')).toMatchObject({ x: 40, y: 80 });
    expect(at(boxes, 'n3')).toMatchObject({ x: 40, y: 90 });
    for (let line = 0; line < 6; line += 1) {
      expect(at(boxes, `m${line}`), `m${line}`).toMatchObject({ y: 50 + 10 * line });
    }
    expect(at(boxes, 'post')).toMatchObject({ y: 110 });
  });

  it('wraps the anchor below the child when the row is centered or bottom-aligned', () => {
    const control = runs(layoutRoot('top'));
    for (const vAlign of ['center', 'bottom'] as const) {
      const boxes = runs(layoutRoot(vAlign));
      // The child once, at its authored page position.
      expect(at(boxes, 'k0'), vAlign).toMatchObject({ x: 40, y: 50 });
      // n0 clears the child's bottom.
      expect(at(boxes, 'n0').y, vAlign).toBeGreaterThanOrEqual(60);
      // Every run, in reading (source) order, is the top-aligned row's.
      expect(boxes, vAlign).toEqual(control);
    }
  });
});

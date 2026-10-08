import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { tableFormatInput } from './parser-model.js';
import type { TableFragmentLayout } from './layout/table-pagination.js';
import type { LayoutPage, PaintNode } from './layout/types.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableRow,
  DocxTextRun,
  SectionProps,
  TblpPr,
  } from './types';

// Unit tests for the page-fit pagination of a page-overflowing FLOATING table
// (ECMA-376 §17.4.57 `<w:tblpPr>`).
//
// A page-overflowing vertAnchor="text" floating table follows Word's observed
// row-pagination behavior: it spills row-by-row onto continuation pages and the
// anchor paragraph flows beside the terminal continuation band from that page's
// body top. A page/margin-anchored floating table is still not split when its
// complete box fits the page's text region; computeFloatTableBox clamps that box
// into its absolute container instead.
//
// The stub canvas mirrors frame-keep-with-anchor.test.ts / pagination.test.ts:
// glyph advance = charCount × fontPx and the font box = 0.8/0.2 em, so a single
// line is exactly fontPx tall. Table row heights are pinned with
// rowHeightRule="exact" so each row's height is deterministic (independent of the
// stub's cell measurement), the table analogue of the frame's hRule="exact"/h.

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
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// pageWidth 200 / pageHeight 140, margins 20 ⇒ content band 160×100 (bodyTop 20).
function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

type DocRun = DocParagraph['runs'][number];

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(opts: { text?: string; fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

// Full TblpPr with the spec defaults; callers override only the axis under test.
// The default vertAnchor is 'page' (matching the Rust parser's fill for an
// absent w:vertAnchor), so the vAnchor gate is exercised by overriding it.
function tblp(over: Partial<TblpPr> = {}): TblpPr {
  return {
    leftFromText: 0,
    rightFromText: 0,
    topFromText: 0,
    bottomFromText: 0,
    horzAnchor: 'text',
    horzSpecified: true,
    vertAnchor: 'page',
    tblpX: 0,
    tblpY: 0,
    ...over,
  };
}

/** A single-cell row of the given exact pt height (`rowHeightRule="exact"` short-
 *  circuits content measurement, so the row height == `heightPt` regardless of the
 *  stub canvas — the table analogue of the frame's hRule="exact"/h). The optional
 *  `label` cell text lets a test read which rows landed on which page. */
function row(heightPt: number, label = ''): DocTableRow {
  return {
    cells: [
      {
        content: label ? [para({ text: label }) as unknown as CellElement] : [],
        colSpan: 1,
        vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null,
        vAlign: 'top',
        widthPt: null,
      },
    ],
    rowHeight: heightPt,
    rowHeightRule: 'exact',
    isHeader: false,
  };
}

/** A floating table (`w:tblpPr`) of `n` exact-height rows, each `rowHPt` tall and
 *  labelled `r1`, `r2`, … so a test can read the per-page row split. Its total
 *  height is `n × rowHPt`, deterministic under the stub. */
function floatTableRows(tp: TblpPr, n: number, rowHPt: number): BodyElement {
  const rows: DocTableRow[] = [];
  for (let i = 1; i <= n; i++) rows.push(row(rowHPt, `r${i}`));
  const t: DocTable = {
    colWidths: [80],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    tblpPr: tp,
  };
  return { type: 'table', ...t } as unknown as BodyElement;
}

function floatingTableWithHeader(
  tp: TblpPr,
  bodyRowCount: number,
  rowHeightPt: number,
): BodyElement {
  const rows = [
    { ...row(rowHeightPt, 'header'), isHeader: true } as DocTableRow,
    ...Array.from(
      { length: bodyRowCount },
      (_value, index) => row(rowHeightPt, `r${index + 1}`),
    ),
  ];
  const table: DocTable = {
    colWidths: [80],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    tblpPr: tp,
  };
  return { type: 'table', ...table } as unknown as BodyElement;
}

function borderedFloatTableRows(
  tp: TblpPr,
  n: number,
  rule: 'auto' | 'atLeast',
): BodyElement {
  const rows = Array.from({ length: n }, () => {
    const source = row(rule === 'auto' ? 0 : 20);
    return {
      ...source,
      cells: source.cells.map((sourceCell) => ({
        ...sourceCell,
        content: [],
        ...(rule === 'auto' ? { marginTop: 10, marginBottom: 10 } : {}),
      })),
      rowHeight: rule === 'auto' ? null : 20,
      rowHeightRule: rule,
    } as DocTableRow;
  });
  const outer = { style: 'single', width: 4, color: '#000000' } as const;
  const inside = { style: 'single', width: 1, color: '#000000' } as const;
  const table: DocTable = {
    colWidths: [80],
    rows,
    borders: {
      top: outer,
      bottom: outer,
      left: null,
      right: null,
      insideH: inside,
      insideV: null,
    },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    tblpPr: tp,
  };
  return { type: 'table', ...table } as unknown as BodyElement;
}

function mixedBoundaryFloatTable(tp: TblpPr): BodyElement {
  const rules = [
    { height: 10, rule: 'auto' },
    { height: 1, rule: 'exact' },
    { height: 10, rule: 'auto' },
    { height: 1, rule: 'exact' },
    { height: 10, rule: 'auto' },
  ] as const;
  const rows = rules.map(({ height, rule }) => {
    const source = row(rule === 'auto' ? 0 : height);
    return {
      ...source,
      cells: source.cells.map((sourceCell) => ({
        ...sourceCell,
        content: [],
        ...(rule === 'auto' ? { marginTop: 5, marginBottom: 5 } : {}),
      })),
      rowHeight: rule === 'auto' ? null : height,
      rowHeightRule: rule,
    } as DocTableRow;
  });
  const outer = { style: 'single', width: 12, color: '#000000' } as const;
  const table: DocTable = {
    colWidths: [80],
    rows,
    borders: {
      top: outer,
      bottom: outer,
      left: null,
      right: null,
      insideH: null,
      insideV: null,
    },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    tblpPr: tp,
  };
  return { type: 'table', ...table } as unknown as BodyElement;
}

/** A floating table (`w:tblpPr`) of total height `tableHPt`, laid out as one
 *  exact-height row so its measured extent is deterministic. */
function floatTable(tp: TblpPr, tableHPt: number): BodyElement {
  return floatTableRows(tp, 1, tableHPt);
}

/** An ordinary in-flow block table with one exact-height row. */
function blockTable(tableHPt: number): BodyElement {
  const table = floatTableRows(tblp(), 1, tableHPt) as unknown as DocTable & { type: 'table' };
  const { tblpPr: _floating, ...block } = table;
  return block as unknown as BodyElement;
}

function blockTableWithLines(lineCount: number): BodyElement {
  const runs: DocRun[] = [];
  for (let index = 0; index < lineCount; index += 1) {
    runs.push(textRun('x', 20));
    if (index + 1 < lineCount) runs.push({ type: 'break', breakType: 'line' });
  }
  const paragraph = {
    ...para(),
    runs,
    defaultFontSize: 20,
  } as unknown as CellElement;
  const table = blockTable(0) as unknown as DocTable & { type: 'table' };
  table.rows = [{
    ...table.rows[0]!,
    rowHeight: null,
    rowHeightRule: 'auto',
    cells: [{ ...table.rows[0]!.cells[0]!, content: [paragraph] }],
  }];
  return table as unknown as BodyElement;
}

const continuousBreak = (): BodyElement => ({
  type: 'sectionBreak', kind: 'continuous', columns: null,
} as BodyElement);

function layoutPages(
  body: BodyElement[],
  pageSection: SectionProps,
  measureContext: CanvasRenderingContext2D,
): readonly LayoutPage[] {
  return layoutBodyModel(body, pageSection, measureContext).pages;
}

/** True when a canonical table node originates from a floating body table. */
const isFloatTable = (node: PaintNode): node is Extract<PaintNode, { kind: 'table' }> =>
  node.kind === 'table' && !node.ordinaryFlow;

/** True when a page holds a floating table (any slice). */
const hasFloatTable = (page: LayoutPage): boolean => page.layers.body.some(isFloatTable);

/** The row labels (r1, r2, …) of the floating-table slice(s) on a page, in order. */
const textOf = (node: PaintNode): string => {
  if (node.kind !== 'paragraph') return '';
  return node.lines.flatMap((line) => line.placements)
    .filter((placement) => placement.kind === 'text')
    .map((placement) => placement.text)
    .join('');
};

const floatRowsOn = (page: LayoutPage): string[] => {
  const out: string[] = [];
  for (const node of page.layers.body) {
    if (!isFloatTable(node)) continue;
    for (const rowLayout of node.rows) {
      const paragraphLayout = rowLayout.cells[0]?.blocks.find((block) => block.layout.kind === 'paragraph')?.layout;
      out.push(paragraphLayout?.kind === 'paragraph' ? textOf(paragraphLayout) : '');
    }
  }
  return out;
};

/** True when a page holds the anchor paragraph (matched by its text). */
const hasAnchorText = (page: LayoutPage, text: string): boolean =>
  page.layers.body.some((node) => textOf(node) === text);

/** The newspaper column an element landed in. */
const colOf = (node: PaintNode | undefined): number | undefined => {
  const column = node && /:column:(\d+)$/u.exec(node.flowDomainId)?.[1];
  return column === undefined ? undefined : Number(column);
};

/** Find the (first) floating-table element on a page. */
const floatTableEl = (page: LayoutPage): PaintNode | undefined =>
  page.layers.body.find(isFloatTable);

const retainedFloatFragment = (node: PaintNode): TableFragmentLayout => {
  expect(node.kind).toBe('table');
  return node as TableFragmentLayout;
};

// These row-fit probes use a one-character following marker so it fits the
// clear side gap. Long-word gap admission is covered by float-gap-continuation;
// it must not introduce an unrelated extra text page into the table assertions.
describe('canonical layout — floating-table page-fit / row-split (§17.4.57, Word ground truth)', () => {
  // Content band 160×100, bodyTop 20. A vertAnchor="text" table's first slice sits
  // at its in-flow anchor (y=20N after N leading 20pt lines); it overflows once
  // 20N + tableH > 100 and is then SPLIT row-by-row, the remainder continuing at
  // the next page's body top.

  it('(a) splits a text-anchored floating table across pages, greedily filling page 1 from the anchor', () => {
    // 2 leading 20pt lines (y 20→40), then a 5-row table (each 20pt). Anchor at
    // y=40 ⇒ remaining band to the content bottom (100) is 60pt ⇒ 3 rows (r1–r3)
    // fit on page 1; r4–r5 continue at page 2's body top. The trailing anchor text
    // follows onto page 2 (kept with the final band).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // Page 1: the two leading lines + a slice holding exactly the rows that fit.
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2', 'r3']);
    // Page 2: the continuation slice with the remaining rows.
    expect(floatRowsOn(pages[1])).toEqual(['r4', 'r5']);
    // Every row appears exactly once across the split (no duplication, no loss).
    expect([...floatRowsOn(pages[0]), ...floatRowsOn(pages[1])]).toEqual([
      'r1', 'r2', 'r3', 'r4', 'r5',
    ]);
  });

  it.each(['auto', 'atLeast'] as const)(
    'reserves centered outer half-rules in %s floating-slice row tracks',
    (rule) => {
      const source = borderedFloatTableRows(
        tblp({ vertAnchor: 'text', tblpY: 0 }),
        3,
        rule,
      ) as unknown as DocTable;
      const body = [
        source as unknown as BodyElement,
        para({ text: 'anchor' }),
      ];

      expect(tableFormatInput(source).rows.map((format) => format.height?.rule ?? 'auto'))
        .toEqual([rule, rule, rule]);

      // Word charges half of each winning collapsed boundary to a non-exact row
      // track. A one-row page-local slice is therefore 20 + 2 + 2 = 24pt, so the
      // 44pt band admits one row per slice.
      const pages = layoutPages(body, section({ pageHeight: 84 }), makeCtx());
      const slices = pages
        .flatMap((page) => page.layers.body.filter(isFloatTable));
      const fragments = slices.map(retainedFloatFragment);

      expect(fragments).toHaveLength(3);
      expect(fragments.map((fragment) =>
        fragment.rows.map((rowLayout) => rowLayout.logicalRowIndex)))
        .toEqual([[0], [1], [2]]);
      expect(fragments.map((fragment) =>
        fragment.rows.map((rowLayout) => rowLayout.advancePt)))
        .toEqual([[24], [24], [24]]);
      expect(fragments.map((fragment) => fragment.advancePt)).toEqual([24, 24, 24]);
      expect(fragments.map((fragment) => fragment.flowBounds.heightPt)).toEqual([24, 24, 24]);
      expect(fragments.map((fragment) => fragment.inkBounds.heightPt)).toEqual([28, 28, 28]);
      for (const fragment of fragments) {
        expect(fragment.inkBounds.yPt).toBe(fragment.flowBounds.yPt - 2);
        expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
          expect.objectContaining({ widthPt: 4 }),
        ]);
        expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
          expect.objectContaining({ widthPt: 4 }),
        ]);
      }
    },
  );

  it('paginates mixed exact/auto floating rows with non-exact outer half-rules', () => {
    const source = mixedBoundaryFloatTable(
      tblp({ vertAnchor: 'text', tblpY: 0 }),
    ) as unknown as DocTable;
    const body = [
      source as unknown as BodyElement,
      para({ text: 'anchor' }),
    ];

    expect(tableFormatInput(source).rows.map((format) => format.height?.rule ?? 'auto'))
      .toEqual(['auto', 'exact', 'auto', 'exact', 'auto']);

    // The page-local top and bottom 12pt rules contribute 6pt half-rules to
    // non-exact edge rows. Exact rows remain authored complete boxes.
    const pages = layoutPages(body, section({ pageHeight: 72 }), makeCtx());
    const slices = pages.flatMap((page) => page.layers.body.filter(isFloatTable));
    const fragments = slices.map(retainedFloatFragment);
    expect(fragments).toHaveLength(2);
    expect(fragments.map((fragment) => fragment.rows.map((rowLayout) => rowLayout.logicalRowIndex)))
      .toEqual([[0, 1, 2, 3], [4]]);
    expect(fragments.map((fragment) => fragment.rows.map((rowLayout) => rowLayout.advancePt)))
      .toEqual([[16, 1, 10, 1], [22]]);
    expect(fragments.map((fragment) => fragment.advancePt)).toEqual([28, 22]);
    expect(fragments.map((fragment) => fragment.inkBounds.heightPt)).toEqual([40, 34]);
    for (const fragment of fragments) {
      expect(fragment.inkBounds.yPt).toBe(fragment.flowBounds.yPt - 6);
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 12 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 12 }),
      ]);
    }
  });

  it('repeats leading tblHeader rows with page-local ownership on floating continuations', () => {
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatingTableWithHeader(
        tblp({ vertAnchor: 'text', tblpY: 0 }),
        5,
        20,
      ),
      para({ text: 'anchor' }),
    ];

    const pages = layoutPages(body, section(), makeCtx());
    const fragments = pages.flatMap((page) =>
      page.layers.body.filter(isFloatTable).map(retainedFloatFragment));

    expect(pages).toHaveLength(2);
    expect(fragments.map((fragment) => fragment.rows.map((rowLayout) => [
      rowLayout.logicalRowIndex,
      rowLayout.ownership,
    ]))).toEqual([
      [[0, 'source'], [1, 'source'], [2, 'source']],
      [[0, 'repeated-header'], [3, 'source'], [4, 'source'], [5, 'source']],
    ]);
  });

  it('moves a following block table past a page-filling float across a continuous section', () => {
    const body = [
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      continuousBreak(),
      blockTable(90),
    ];
    const pages = layoutPages(body, section({ sectionStart: 'continuous' }), makeCtx());
    expect(pages.length).toBe(2);
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(pages[0].layers.body.filter((el) => el.kind === 'table' && !isFloatTable(el))).toHaveLength(0);
    expect(pages[1].layers.body.filter((el) => el.kind === 'table' && !isFloatTable(el))).toHaveLength(1);
  });

  it('does not overlap a following block table with a terminal float continuation slice', () => {
    const body = [
      // Five rows fill page 1; the sixth occupies y=0..20 on page 2.
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 6, 20),
      continuousBreak(),
      // It fits a clean 100pt page, but not page 2 below the 20pt float slice.
      blockTable(90),
    ];
    const pages = layoutPages(body, section({ sectionStart: 'continuous' }), makeCtx());
    expect(pages.length).toBe(3);
    expect(floatRowsOn(pages[1])).toEqual(['r6']);
    expect(pages[1].layers.body.filter((el) => el.kind === 'table' && !isFloatTable(el))).toHaveLength(0);
    expect(pages[2].layers.body.filter((el) => el.kind === 'table' && !isFloatTable(el))).toHaveLength(1);
  });

  it('rechecks the block-table position after each staggered float collision', () => {
    const body = [
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 1, 15),
      // The block table does not initially intersect this second band. Moving it
      // below the first float creates the second collision.
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 20 }), 1, 15),
      blockTable(10),
    ];
    const pages = layoutPages(body, section({ pageHeight: 80 }), makeCtx());

    expect(pages.length).toBe(2);
    expect(pages[0].layers.body.filter((el) => el.kind === 'table' && !isFloatTable(el))).toHaveLength(0);
    expect(pages[1].layers.body.filter((el) => el.kind === 'table' && !isFloatTable(el))).toHaveLength(1);
  });

  it('collides a split block-table row using only its remaining continuation extent', () => {
    const twoColumns = section({
      columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
    });
    const body = [
      floatTableRows(tblp({
        horzAnchor: 'page', vertAnchor: 'page', tblpX: 110, tblpY: 60,
      }), 1, 10),
      blockTableWithLines(6),
    ];

    const pages = layoutPages(body, twoColumns, makeCtx());
    const blockSlices = pages[0]?.layers.body.filter(
      (node): node is Extract<PaintNode, { kind: 'table' }> => (
        node.kind === 'table' && node.ordinaryFlow
      ),
    ) ?? [];

    expect(pages).toHaveLength(1);
    expect(blockSlices).toHaveLength(2);
    expect(blockSlices.map(colOf)).toEqual([0, 1]);
    expect(blockSlices[1]?.flowBounds.yPt).toBeCloseTo(20, 6);
  });

  it('(b) splits a tall floating table across pages until every row is placed', () => {
    // A single leading 20pt line (anchor at y=20) then an 8-row table (20pt each ⇒
    // 160pt total, > the 100pt content area, so it needs 2 pages). Page 1's band
    // runs from the anchor (y=20) to the bottom (100) = 80pt ⇒ 4 rows (r1–r4);
    // page 2 (fresh, full 100pt band) takes the remaining 4 (r5–r8).
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 8, 20),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2', 'r3', 'r4']);
    expect(floatRowsOn(pages[1])).toEqual(['r5', 'r6', 'r7', 'r8']);
  });

  it('(c) flows the trailing anchor paragraph from the TERMINAL page body top, not below the band', () => {
    // 2 leading lines (anchor at y=40), 5-row 20pt table split (r1–r3 | r4–r5).
    // The anchor paragraph must land on page 2 (the terminal continuation page) —
    // Word flows it beside the final band from that page's body top, so it never
    // stays on page 1 and never starts below the band.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(hasAnchorText(pages[0], 'anchor')).toBe(false);
    expect(hasAnchorText(pages[1], 'anchor')).toBe(true);
    // The anchor paragraph is the FIRST paragraph on page 2 after the continuation
    // slice (the slice is pushed before the anchor so the wrap band is registered
    // first) — i.e. it starts at the body top beside the band, not after it.
    const page2Types = pages[1].layers.body.map((el) => (isFloatTable(el) ? 'slice' : textOf(el)));
    expect(page2Types).toEqual(['slice', 'anchor']);
  });

  it('(a-multicol) splits a text-anchored floating table into the NEXT COLUMN in a multi-column section', () => {
    // 2 equal columns: colW = (160-20)/2 = 70; content height 100. Column 0 gets 2
    // leading lines (y 20→40); a 5-row 20pt table then overflows column 0 from the
    // anchor (40 + 60 = 100 exactly for 3 rows, r4–r5 spill). With a column still
    // available on the page, the remainder continues in COLUMN 1 — not a new page.
    const twoCol = section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', horzAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, twoCol, makeCtx());
    // Still a single page (column 1 absorbed the remaining rows + the anchor).
    expect(pages.length).toBe(1);
    // The first slice sits in column 0, the continuation slice in column 1.
    const slices = pages[0].layers.body.filter(isFloatTable);
    expect(slices.length).toBe(2);
    expect(colOf(slices[0])).toBe(0);
    expect(colOf(slices[1])).toBe(1);
    // The trailing anchor text follows the final band into column 1.
    const anchor = pages[0].layers.body.find((el) => textOf(el) === 'anchor');
    expect(anchor).toBeDefined();
    expect(colOf(anchor)).toBe(1);
  });

  it('(f) keeps a text-anchored floating table as ONE element when every row fits (no split)', () => {
    // Only 1 leading line (anchor at y=20). A 3-row 20pt table (60pt) fits within
    // [20,100] ⇒ no split. Everything stays on page 1 as a single float element
    // A small vertAnchor="text" float near the page top must not be divided or
    // relocated.
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 1 }), 3, 20),
      para({ text: 'x' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2', 'r3']);
    // Exactly ONE floating-table element (not sliced).
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
    expect(hasAnchorText(pages[0], 'x')).toBe(true);
  });

  it('moves the whole table to the next page when not even the first row fits the remaining band', () => {
    // 4 leading lines (anchor at y=80), then a 3-row 30pt table. The remaining band
    // (100−80 = 20pt) cannot hold even the first row (30pt), and a fuller band is
    // one page away ⇒ the WHOLE table moves to page 2 first (no page-1 slice), then
    // fills there (30+30 = 60 ≤ 100 ⇒ all 3 rows: 30·3 = 90 ≤ 100).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      para({ text: 'd' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 3, 30),
      para({ text: 'x' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // No slice on page 1 (the four leading lines only).
    expect(hasFloatTable(pages[0])).toBe(false);
    expect(pages[0].layers.body.map(textOf)).toEqual(['a', 'b', 'c', 'd']);
    // All rows land on page 2 (a fresh full band fits them).
    expect(floatRowsOn(pages[1])).toEqual(['r1', 'r2', 'r3']);
    expect(hasAnchorText(pages[1], 'x')).toBe(true);
  });

  it('terminates (no loop) and places a single over-tall row on the page it best fits', () => {
    // 3 leading lines (anchor at y=60), then a 1-row table 150pt tall — taller than
    // the whole 100pt content area, so it can never fit any band. The forward-
    // progress guarantee places it (overflowing) rather than looping. Because the
    // remaining page-1 band (40pt) can't hold it AND a fuller band is one page
    // away, it moves to page 2 first, then is placed there (over-tall rows are not
    // sub-divided — a floating table splits by ROW, and one row is atomic).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 1, 150),
      para({ text: 'x' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    // Terminates with a bounded page count; the over-tall single row is placed once.
    expect(pages.length).toBe(2);
    expect(floatRowsOn(pages[1])).toEqual(['r1']);
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
  });

  it('(d) does NOT split a page-anchored floating table (absolute y; clamped in place by geometry)', () => {
    // vertAnchor="page", tblpY=90: the table is pinned at page-y 90 with total
    // height 60 ⇒ it would reach y=150, past the 140 page edge. An absolute page
    // position is the SAME on any page, so it is NOT split or relocated — it stays
    // one element on page 1 (computeFloatTableBox clamps the box UP into the page;
    // see float-table-geometry.test.ts for the clamped y).
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTableRows(tblp({ vertAnchor: 'page', tblpY: 90 }), 2, 30),
      para({ text: 'x' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
    // Not sliced — both rows are on the single element.
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2']);
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
    expect(hasAnchorText(pages[0], 'x')).toBe(true);
  });

  it('(e) does NOT split a margin-anchored floating table (absolute y; clamped in place by geometry)', () => {
    // vertAnchor="margin", tblpY=70: the table is at margin-top(20)+70 = 90 with
    // total height 60 ⇒ bottom 150, past the bottom margin. Absolute ⇒ left as one
    // element (mirrors vertAnchor="page"); the box clamps up into the margin band.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      para({ text: 'c' }),
      floatTableRows(tblp({ vertAnchor: 'margin', tblpY: 70 }), 2, 30),
      para({ text: 'x' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(floatRowsOn(pages[0])).toEqual(['r1', 'r2']);
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(1);
  });

  it('(d2) SPLITS a page-anchored floating table that is TALLER than the body content area', () => {
    // vertAnchor="page", tblpY=10: 6 rows × 30pt = 180pt total. The body content
    // area is 100pt (pageH 140 − margins 20+20). A page-anchored table taller than
    // the text region cannot fit even when clamped to the top, so Word row-splits it.
    // Slice 1 sits at its absolute tblpY (page-y 10 ⇒ content-relative −10, i.e. it
    // starts in the top margin band down through the body); continuation slices flow
    // from the next page's body top (tblpY=0). This is distinct from (d): there the
    // table (60pt) FITS the 100pt text region and is instead clamped up in place.
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'page', tblpY: 10 }), 6, 30),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    // Taller than one page ⇒ paginated across ≥ 2 pages, every row placed once.
    expect(pages.length).toBeGreaterThanOrEqual(2);
    const allRows = pages.flatMap((p) => floatRowsOn(p));
    expect(allRows).toEqual(['r1', 'r2', 'r3', 'r4', 'r5', 'r6']);
    // More than one slice (the table was divided, not crammed onto one page).
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBeGreaterThanOrEqual(2);
    // The trailing anchor paragraph flows beside the FINAL band (last page).
    expect(hasAnchorText(pages[pages.length - 1], 'anchor')).toBe(true);
  });

  it('(e2) SPLITS a margin-anchored floating table taller than the body content area', () => {
    // vertAnchor="margin", tblpY=0: 6 rows × 30pt = 180pt > the 100pt text region ⇒
    // splits (mirrors (d2); the margin band top coincides with the body top here).
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'margin', tblpY: 0 }), 6, 30),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(pages.length).toBeGreaterThanOrEqual(2);
    const allRows = pages.flatMap((p) => floatRowsOn(p));
    expect(allRows).toEqual(['r1', 'r2', 'r3', 'r4', 'r5', 'r6']);
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBeGreaterThanOrEqual(2);
  });

  it('registers each slice as its own floating table so both pages report a float element', () => {
    // Sanity: a split leaves a floating-table element on BOTH pages (one slice
    // each), never a single element straddling the boundary.
    const body = [
      para({ text: 'a' }),
      para({ text: 'b' }),
      floatTableRows(tblp({ vertAnchor: 'text', tblpY: 0 }), 5, 20),
      para({ text: 'anchor' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    expect(hasFloatTable(pages[0])).toBe(true);
    expect(hasFloatTable(pages[1])).toBe(true);
    // Two slices total (one per page).
    const floatCount = pages.reduce((s, p) => s + p.layers.body.filter(isFloatTable).length, 0);
    expect(floatCount).toBe(2);
    // Each slice remains an independently retained table occurrence.
    for (const p of pages) {
      const el = floatTableEl(p);
      if (el) expect(el).toMatchObject({ kind: 'table', source: { path: [2] } });
    }
  });

  it('(g) DEFERS a page-anchored floating table when its raw band intersects an existing table float', () => {
    // Two page-anchored floating tables whose raw bands collide (§17.4.56 and
    // observed Word behavior).
    // Content band is 100pt (pageH 140 − margins 20+20), bodyTop 20.
    //
    // Table A: vertAnchor="page", tblpY=10, 4 rows × 30pt = 120pt > 100pt ⇒ it
    //   ROW-SPLITS (pageAnchoredOverflows). Slice 1 (A1–A3) fills page 1 from its
    //   absolute page-y 10; the continuation A4 flows onto page 2 at the body top
    //   (content-relative y 0 → 30).
    // Table B: vertAnchor="page", tblpY=25, 3 rows × 30pt = 90pt ≤ 100pt ⇒ B alone
    //   FITS the text region (single element, no split). Its raw absolute box is at
    //   page-y 25 (content-relative y 5), height 90 ⇒ it would occupy content-y
    //   5 → 95 on whatever page it lands on.
    //
    // If B is placed on page 2 (where A4's continuation band sits at content-y
    // 0 → 30), B's raw band (content-y 5 → 95) INTERSECTS A4's band — two floating
    // tables overlapping. Word defers the whole of B to the next page, where its
    // absolute tblpY=25 no longer collides with another table float.
    const body = [
      para({ text: 'a' }),
      floatTableRows(tblp({ vertAnchor: 'page', tblpY: 10 }), 4, 30), // Table A → r1..r4
      para({ text: 'mid' }),
      floatTableRows(tblp({ vertAnchor: 'page', tblpY: 25 }), 3, 30), // Table B → r1..r3
      para({ text: 'end' }),
    ];
    const pages = layoutPages(body, section(), makeCtx());
    // Page 2 (index 1) carries A's continuation (A4) but MUST NOT also carry any of
    // Table B — B is deferred so it never overlaps A4's band.
    const bandCollisionPage = pages[1];
    const tableFloatsOnPage2 = bandCollisionPage.layers.body.filter(isFloatTable);
    // Exactly one table float on page 2 (A's continuation slice), not two stacked.
    expect(tableFloatsOnPage2.length).toBe(1);
    // And it is A's continuation (a single row on that band), not B.
    expect(floatRowsOn(bandCollisionPage)).toEqual(['r4']);

    // Table B is deferred to a LATER page and lands there as ONE un-split element
    // (its 90pt fits the 100pt region once it is on a clean page).
    const bPageIdx = pages.findIndex(
      (p, i) => i > 1 && p.layers.body.some(isFloatTable),
    );
    expect(bPageIdx).toBeGreaterThan(1);
    const bPage = pages[bPageIdx];
    const bFloats = bPage.layers.body.filter(isFloatTable);
    expect(bFloats.length).toBe(1);
    expect(floatRowsOn(bPage)).toEqual(['r1', 'r2', 'r3']);

    // No page ever holds two DISTINCT table-float bands that vertically overlap:
    // assert every page has at most one floating-table element (each table's slice
    // owns its page's band alone here, since neither co-resident split occurs).
    for (const p of pages) {
      expect(p.layers.body.filter(isFloatTable).length).toBeLessThanOrEqual(1);
    }
  });
});

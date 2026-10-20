import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  BorderSpec,
  CellElement,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ECMA-376 §17.4.66 — end-to-end: render a table whose adjacent cells DISAGREE on
// their shared interior gridline and assert the drawn line uses the §17.4.66
// WINNER (not the later-painted cell). We record each border stroke as a segment
// with its width + colour, then read the shared vertical/horizontal line.

const TEST_FONT = 'Synthetic Untabled Serif';

interface StrokeSeg { x1: number; y1: number; x2: number; y2: number; width: number; color: string; }

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
    canvas: { width: 1000, height: 1000 },
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
        width: [...s].length * p,
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
    fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText() {}, strokeText() {},
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, strokes };
}

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
}

function paraOf(text: string): CellElement {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) }],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT, widowControl: false,
  } as unknown as CellElement;
}

type Edges = { top: BorderSpec | null; bottom: BorderSpec | null; left: BorderSpec | null; right: BorderSpec | null; insideH: BorderSpec | null; insideV: BorderSpec | null };
const NO_BORDERS: Edges = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };

function cell(text: string, borders: Partial<Edges> = {}): DocTableCell {
  return {
    content: [paraOf(text)], colSpan: 1, vMerge: null,
    borders: { ...NO_BORDERS, ...borders },
    background: null, vAlign: 'top', widthPt: 60,
  } as unknown as DocTableCell;
}

function tableOf(rows: DocTableRow[], tableBorders: Partial<Edges> = {}): DocTable {
  // Fixed 60 pt columns (one per logical column of the widest row). The page is
  // sized to the grid sum below so autofit never stretches the columns — the
  // shared interior vertical gridline then sits at a known x (60 for two cols).
  const cols = Math.max(...rows.map((r) => r.cells.reduce((sum, c) => sum + c.colSpan, 0)));
  return {
    colWidths: new Array(cols).fill(60), rows,
    borders: { ...NO_BORDERS, ...tableBorders },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

function rowOf(cells: DocTableCell[]): DocTableRow {
  // Exact rows keep the border-conflict assertions pinned to y=20; non-exact
  // rows intentionally reserve resolved border width and are covered by the
  // table-row-height geometry tests.
  return { cells, rowHeight: 20, rowHeightRule: 'exact', isHeader: false } as unknown as DocTableRow;
}

function docOf(t: DocTable): DocxDocumentModel {
  // Page width == grid sum so the content band exactly holds the fixed columns
  // (no autofit stretch). marginLeft 0 ⇒ grid origin at x=0.
  const gridSum = t.colWidths.reduce((s, w) => s + w, 0);
  return {
    section: {
      pageWidth: gridSum, pageHeight: 200,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...t } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(t: DocTable) {
  const { canvas, strokes } = makeRecordingCanvas();
  const gridSum = t.colWidths.reduce((s, w) => s + w, 0);
  await renderDocumentToCanvas(docOf(t), canvas, 0, { dpr: 1, width: gridSum });
  return strokes;
}

const bs = (over: Partial<BorderSpec>): BorderSpec => ({ style: 'single', width: 0.5, color: null, ...over });

// The two cells are 60 pt wide each at scale 1 ⇒ the shared vertical gridline is
// at x=60. Only vertical strokes near x=60 are the shared interior line.
function verticalAt(strokes: StrokeSeg[], x: number): StrokeSeg[] {
  return strokes.filter((s) => Math.abs(s.x1 - s.x2) < 0.5 && Math.abs(s.x1 - x) <= 1);
}
// A collapsed horizontal rule occupies the band starting at its row boundary
// (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT), so match its band top.
function horizontalAt(strokes: StrokeSeg[], y: number): StrokeSeg[] {
  return strokes.filter((s) => Math.abs(s.y1 - s.y2) < 0.5 && Math.abs(s.y1 - s.width / 2 - y) <= 1);
}

describe('§17.4.66 — adjacent cell border conflict, end-to-end render', () => {
  it('§17.4.15: gridBefore starts the first cell after the skipped grid columns', async () => {
    const middle = cell('middle', {
      top: bs({ width: 1 }),
      bottom: bs({ width: 1 }),
      left: bs({ width: 1 }),
      right: bs({ width: 1 }),
    });
    middle.widthPt = 40;
    const shiftedRow = {
      ...rowOf([middle]),
      gridBefore: 1,
      gridAfter: 1,
    } as unknown as DocTableRow;
    const t = tableOf([shiftedRow]);
    t.colWidths = [20, 40, 60];
    t.layout = 'fixed';
    t.widthPt = 120;

    const strokes = await render(t);
    const top = horizontalAt(strokes, 0);

    expect(top).toEqual(expect.arrayContaining([
      expect.objectContaining({ x1: 20, x2: 60 }),
    ]));
    expect(top.some((s) => Math.min(s.x1, s.x2) < 20)).toBe(false);
    expect(verticalAt(strokes, 20)).toHaveLength(1);
  });

  it('extends the inferred grid for a hand-built table without authored-grid metadata', async () => {
    const first = cell('first', {
      top: bs({ width: 1 }),
      left: bs({ width: 1 }),
    });
    const shiftedRow = { ...rowOf([first]), gridBefore: 3 } as unknown as DocTableRow;
    const t = tableOf([shiftedRow]);
    t.colWidths = [20, 40];
    t.layout = 'fixed';
    t.widthPt = 60;

    const strokes = await render(t);

    // The stable hand-built DocTable contract cannot state whether colWidths
    // came from an authored tblGrid. Only this compatibility path infers the
    // missing tracks; parsed documents retain authored-grid presence and apply
    // §17.4.15 before entering the solver.
    expect(horizontalAt(strokes, 0)).toEqual(expect.arrayContaining([
      expect.objectContaining({ x1: 30, x2: 60 }),
    ]));
    expect(verticalAt(strokes, 30)).toHaveLength(1);
  });

  it('draws a cell top border where the row above omits that grid slot', async () => {
    const upper = { ...rowOf([cell('upper')]), gridBefore: 1 } as unknown as DocTableRow;
    const lower = rowOf([cell('lower', { top: bs({ width: 1 }) }), cell('right')]);
    upper.cells[0]!.widthPt = 40;
    lower.cells[0]!.widthPt = 20;
    lower.cells[1]!.widthPt = 40;
    const t = tableOf([upper, lower]);
    t.colWidths = [20, 40];
    t.layout = 'fixed';
    t.widthPt = 60;

    const strokes = await render(t);

    expect(horizontalAt(strokes, 20)).toEqual(expect.arrayContaining([
      expect.objectContaining({ x1: 0, x2: 20 }),
    ]));
  });

  it('draws the physical-right boundary beside gridBefore in a bidiVisual row', async () => {
    const shifted = cell('shifted', { left: bs({ width: 1 }) });
    shifted.widthPt = 40;
    const shiftedRow = { ...rowOf([shifted]), gridBefore: 1 } as unknown as DocTableRow;
    const t = tableOf([shiftedRow]);
    t.colWidths = [20, 40, 60];
    t.layout = 'fixed';
    t.widthPt = 120;
    t.bidiVisual = true;

    const strokes = await render(t);

    expect(verticalAt(strokes, 100)).toHaveLength(1);
  });

  it('shared vertical gridline is drawn exactly ONCE (not once per cell)', async () => {
    // Both cells set their facing edge to the SAME single 1pt border. Previously
    // each cell drew it (two strokes at x=60); now it is drawn once.
    const strokes = await render(tableOf([
      rowOf([
        cell('a', { right: bs({ width: 1 }) }),
        cell('b', { left: bs({ width: 1 }) }),
      ]),
    ]));
    expect(verticalAt(strokes, 60)).toHaveLength(1);
  });

  it('uses Word border weights where dashed has a fixed weight of one', async () => {
    // [MS-OI29500] 2.1.169 fixes dashed/dotted weight at one regardless of
    // width. The 3pt single therefore outweighs the 1pt dashed border.
    const strokes = await render(tableOf([
      rowOf([
        cell('a', { right: bs({ style: 'single', width: 3, color: 'ff0000' }) }),
        cell('b', { left: bs({ style: 'dashed', width: 1, color: '0000ff' }) }),
      ]),
    ]));
    const shared = verticalAt(strokes, 60);
    expect(shared.length).toBeGreaterThanOrEqual(1);
    // The winner (single, red) must appear at x=60; the loser (dashed, blue) must not.
    const colors = shared.map((s) => s.color.toLowerCase());
    expect(colors.some((c) => c.includes('ff0000'))).toBe(true);
    expect(colors.some((c) => c.includes('0000ff'))).toBe(false);
    expect(shared.every((s) => Math.abs(s.width - 3) < 0.5)).toBe(true);
  });

  it('a cell border beats a table-level inside border (rule #1)', async () => {
    // Table insideV = thick 4pt; the left cell explicitly sets its right edge to a
    // single 1pt. A CELL border always beats a TABLE border (rule #1), so the
    // single wins — the right cell contributes no own edge (falls to table insideV
    // = thick), and cell-vs-table ⇒ cell.
    const strokes = await render(tableOf(
      [rowOf([
        cell('a', { right: bs({ style: 'single', width: 1, color: '00ff00' }) }),
        cell('b'),
      ])],
      { insideV: bs({ style: 'thick', width: 4, color: 'ff0000' }) },
    ));
    const shared = verticalAt(strokes, 60);
    const colors = shared.map((s) => s.color.toLowerCase());
    expect(colors.some((c) => c.includes('00ff00'))).toBe(true); // cell single wins
    expect(colors.some((c) => c.includes('ff0000'))).toBe(false); // table thick loses
  });

  it('nil loses to the opposing visible border', async () => {
    // ECMA-376 §17.4.66: when either conflicting cell border is nil or none,
    // the opposing border is displayed.
    const strokes = await render(tableOf([
      rowOf([
        cell('a', { right: bs({ style: 'nil' }) }),
        cell('b', { left: bs({ style: 'single', width: 1, color: '112233' }) }),
      ]),
    ]));
    const shared = verticalAt(strokes, 60);
    expect(shared).toEqual([
      expect.objectContaining({
        color: expect.stringContaining('112233'),
      }),
    ]);
  });

  it('does not paint a shared gridline when both sides are nil', async () => {
    const strokes = await render(tableOf([
      rowOf([
        cell('a', { right: bs({ style: 'nil' }) }),
        cell('b', { left: bs({ style: 'nil' }) }),
      ]),
    ]));
    expect(verticalAt(strokes, 60)).toHaveLength(0);
  });

  it('shared horizontal gridline uses the same Word border weights', async () => {
    // Thick uses width × Word border number, while dashed is fixed at one, so
    // the 4pt thick border wins at the 20pt row boundary.
    const strokes = await render(tableOf([
      rowOf([cell('a', { bottom: bs({ style: 'thick', width: 4, color: 'ff0000' }) })]),
      rowOf([cell('b', { top: bs({ style: 'dashed', width: 1, color: '00ff00' }) })]),
    ]));
    const shared = horizontalAt(strokes, 20);
    expect(shared.length).toBeGreaterThanOrEqual(1);
    const colors = shared.map((s) => s.color.toLowerCase());
    expect(colors.some((c) => c.includes('ff0000'))).toBe(true);
    expect(colors.some((c) => c.includes('00ff00'))).toBe(false);
  });

  it('uses the winning below-cell top edge extent when the above cell spans wider', async () => {
    const title = cell('title');
    title.colSpan = 2;
    const strokes = await render(tableOf([
      rowOf([title]),
      rowOf([
        cell('work', { top: bs({ width: 1, color: '123456' }) }),
        cell('gap'),
      ]),
    ]));
    const shared = horizontalAt(strokes, 20).filter((s) => s.color.toLowerCase().includes('123456'));
    expect(shared).toHaveLength(1);
    expect(Math.abs(shared[0].x2 - shared[0].x1)).toBeCloseTo(60, 5);
  });

  it('bidiVisual: shared vertical gridline still drawn once with the winner', async () => {
    // Under <w:bidiVisual> logical column 0 is at the PHYSICAL right. Two cells,
    // logical [a, b]: a is physically RIGHT (x=60..120), b physically LEFT
    // (x=0..60), so the shared line is at x=60. a.right = single 1pt (green),
    // b.left = single 1pt (green) — same spec ⇒ drawn once.
    const t = tableOf([
      rowOf([
        cell('a', { right: bs({ width: 1, color: '00aa00' }) }),
        cell('b', { left: bs({ width: 1, color: '00aa00' }) }),
      ]),
    ]);
    (t as unknown as { bidiVisual: boolean }).bidiVisual = true;
    const strokes = await render(t);
    // Exactly one shared interior vertical line at x=60 (not two, not zero).
    expect(verticalAt(strokes, 60)).toHaveLength(1);
    // Outer verticals at x=0 (b's physical left) and x=120 (a's physical right)
    // are each drawn once too.
    expect(verticalAt(strokes, 0).length + verticalAt(strokes, 120).length).toBe(0);
  });

  it('outer table edges are unaffected (non-regression: still drawn once)', async () => {
    // A single-cell table with an outer border on all sides. Each outer edge is
    // drawn once by the sole cell.
    const strokes = await render(tableOf([
      rowOf([cell('x', { top: bs({ width: 1 }), bottom: bs({ width: 1 }), left: bs({ width: 1 }), right: bs({ width: 1 }) })]),
    ]));
    // 4 edges (cell 60 wide, 20 tall): left x=0, right x=60, top y=0, bottom y=20.
    expect(verticalAt(strokes, 0)).toHaveLength(1);
    expect(verticalAt(strokes, 60)).toHaveLength(1);
    expect(horizontalAt(strokes, 0)).toHaveLength(1);
    expect(horizontalAt(strokes, 20)).toHaveLength(1);
  });
  it('§17.4.66 (#815): a colSpan cell resolves EACH below sub-segment against its own neighbour', async () => {
    // A title cell spanning both columns faces TWO below cells with DIFFERENT
    // top borders. The shared horizontal edge must be split at the column
    // boundary: left half → the left neighbour's border, right half → the right
    // neighbour's border. Before the fix only the span-origin (left) neighbour
    // was consulted and the right half's border was dropped.
    const title = cell('title');
    title.colSpan = 2;
    const strokes = await render(tableOf([
      rowOf([title]),
      rowOf([
        cell('l', { top: bs({ style: 'single', width: 1, color: '111111' }) }),
        cell('r', { top: bs({ style: 'thick', width: 4, color: '222222' }) }),
      ]),
    ]));
    const seg = horizontalAt(strokes, 20);
    const left = seg.filter((s) => Math.min(s.x1, s.x2) < 30);
    const right = seg.filter((s) => Math.max(s.x1, s.x2) > 90);
    expect(left.some((s) => s.color.toLowerCase().includes('111111'))).toBe(true);
    expect(right.some((s) => s.color.toLowerCase().includes('222222'))).toBe(true);
  });

  it('resolves each colSpan split segment independently when the spanning edge is nil', async () => {
    // §17.4.66 applies independently to each segment at a 3-wide spanning
    // owner → three independent cells boundary. The nil/nil left segment
    // remains open while the omitted opposing edges on the middle and right
    // segments admit the table insideH border.
    const spanning = cell('intro', { bottom: bs({ style: 'nil' }) });
    spanning.colSpan = 3;
    const table = tableOf(
      [
        rowOf([spanning]),
        rowOf([
          cell('merge-start', { top: bs({ style: 'nil' }) }),
          cell('label'),
          cell('value'),
        ]),
      ],
      { insideH: bs({ style: 'single', width: 1, color: '123456' }) },
    );
    table.layout = 'fixed';
    table.widthPt = 180;
    const strokes = await render(table);

    expect(horizontalAt(strokes, 20)).toEqual([
      expect.objectContaining({
        x1: 60,
        x2: 180,
        color: expect.stringContaining('123456'),
      }),
    ]);
  });

  it('retains each above-cell bottom segment when the row below spans the full grid', async () => {
    const bottom = bs({ style: 'single', width: 1, color: '123456' });
    const spanning = cell('spanning');
    spanning.colSpan = 3;
    const strokes = await render(tableOf(
      [
        rowOf([
          cell('left', { bottom }),
          cell('middle', { bottom }),
          cell('right', { bottom }),
        ]),
        rowOf([spanning]),
      ],
      { insideH: bs({ style: 'single', width: 1, color: '654321' }) },
    ));

    const shared = horizontalAt(strokes, 20);
    expect(shared).toEqual([
      expect.objectContaining({
        x1: 0,
        x2: 180,
        color: expect.stringContaining('123456'),
      }),
    ]);
  });

  it('§17.4.66 (#815): a vMerge cell resolves EACH right sub-segment against its own neighbour', async () => {
    // A tall cell (vMerge restart) spanning both rows faces TWO right neighbours
    // with DIFFERENT left borders. Its shared vertical edge must be split at the
    // row boundary: top half → the upper neighbour's border, bottom half → the
    // lower neighbour's border. Before the fix only the span-origin (upper)
    // neighbour was consulted and the bottom half's border was dropped.
    const tall = cell('tall');
    tall.vMerge = true;
    const cont = cell('');
    cont.vMerge = false;
    const strokes = await render(tableOf([
      rowOf([tall, cell('u', { left: bs({ style: 'single', width: 1, color: '111111' }) })]),
      rowOf([cont, cell('d', { left: bs({ style: 'thick', width: 4, color: '222222' }) })]),
    ]));
    const seg = verticalAt(strokes, 60);
    const top = seg.filter((s) => Math.min(s.y1, s.y2) < 10);
    const bot = seg.filter((s) => Math.max(s.y1, s.y2) > 30);
    expect(top.some((s) => s.color.toLowerCase().includes('111111'))).toBe(true);
    expect(bot.some((s) => s.color.toLowerCase().includes('222222'))).toBe(true);
  });

  it('resolves each vMerge split segment independently when the spanning edge is nil', async () => {
    // §17.4.66 applies independently to each segment of a three-row merged
    // owner whose nil right edge faces nil, omitted, and explicit opposing
    // edges. The results are open, table insideV, and the explicit edge.
    const restart = cell('merged', { right: bs({ style: 'nil' }) });
    restart.vMerge = true;
    const continuation1 = cell('', { right: bs({ style: 'nil' }) });
    continuation1.vMerge = false;
    const continuation2 = cell('', { right: bs({ style: 'nil' }) });
    continuation2.vMerge = false;
    const strokes = await render(tableOf(
      [
        rowOf([restart, cell('top', { left: bs({ style: 'nil' }) })]),
        rowOf([continuation1, cell('middle')]),
        rowOf([
          continuation2,
          cell('bottom', { left: bs({ style: 'single', width: 1, color: 'ff0000' }) }),
        ]),
      ],
      { insideV: bs({ style: 'single', width: 1, color: '0000ff' }) },
    ));

    const shared = verticalAt(strokes, 60);
    expect(shared.some((s) => Math.min(s.y1, s.y2) < 10)).toBe(false);
    expect(shared).toEqual(expect.arrayContaining([
      expect.objectContaining({
        y1: 20,
        y2: 40,
        color: expect.stringContaining('0000ff'),
      }),
      expect.objectContaining({
        y1: 40,
        y2: 60,
        color: expect.stringContaining('ff0000'),
      }),
    ]));
  });

  it('uses the final continuation cell border at the bottom of a vertical merge', async () => {
    // A vertical merge is represented by one restart cell followed by continuation
    // cells. The boundary at the bottom of the merged region is therefore the
    // continuation cell's bottom edge, not the restart cell's bottom edge. An
    // explicit nil on that final continuation must suppress the table insideH
    // fallback at the boundary below the merge.
    const restart = cell('merged');
    restart.vMerge = true;
    const continuation = cell('', { bottom: bs({ style: 'nil' }) });
    continuation.vMerge = false;
    const below = cell('below', { top: bs({ style: 'nil' }) });
    const strokes = await render(tableOf(
      [rowOf([restart]), rowOf([continuation]), rowOf([below])],
      { insideH: bs({ style: 'single', width: 1 }) },
    ));

    expect(horizontalAt(strokes, 40)).toHaveLength(0);
  });
});

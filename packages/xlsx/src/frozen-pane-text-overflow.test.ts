import { describe, expect, it } from 'vitest';
import { renderViewport } from './renderer.js';
import type { Cell, Styles, Worksheet, XlsxTextRunInfo } from './types.js';

// Excel paints each freeze pane as its own window onto the sheet: a value
// spills through empty neighbours in that pane's sheet coordinates whichever
// side of the freeze line it lies on (observed in Excel 16.114, columns A:C
// frozen, unscrolled and scrolled). See virtualizedTextOverflowOverscan.

const LONG = 'LEFT aaaa|bbbb|cccc|dddd|eeee|ffff|gggg|hhhh|iiii|jjjj|END';
const CENTERED = 'CENTER 1111|2222|3333|4444|END';
const RIGHT = 'RIGHT rrrr|ssss|tttt|uuuu|END';
const BORDER_COLOR = 'FF0000';

const FONT = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: 'Calibri',
};

const STYLES: Styles = {
  fonts: [FONT],
  fills: [],
  borders: [
    { left: null, right: null, top: null, bottom: null },
    { left: null, right: { style: 'thin', color: BORDER_COLOR }, top: null, bottom: null },
    { left: null, right: { style: 'hair', color: BORDER_COLOR }, top: null, bottom: null },
  ],
  cellXfs: [
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: 'left', wrapText: false },
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: 'right', wrapText: false },
    { fontId: 0, fillId: 0, borderId: 1, numFmtId: 0, alignH: null, wrapText: false },
    { fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: 'center', wrapText: false },
    { fontId: 0, fillId: 0, borderId: 2, numFmtId: 0, alignH: 'left', wrapText: false },
    { fontId: 0, fillId: 0, borderId: 1, numFmtId: 0, alignH: 'right', wrapText: false },
  ] as Styles['cellXfs'],
  numFmts: [],
  dxfs: [],
};

const LEFT_XF = 0;
const RIGHT_XF = 1;
const RIGHT_BORDER_XF = 2;
const CENTER_XF = 3;
const LEFT_HAIR_BORDER_XF = 4;
const RIGHT_ALIGNED_RIGHT_BORDER_XF = 5;

function cell(col: number, text: string, styleIndex = LEFT_XF): Cell {
  return { row: 1, col, styleIndex, value: { type: 'text', text } } as Cell;
}

function worksheet(cells: Cell[]): Worksheet {
  return {
    name: 'Frozen',
    rows: [{ index: 1, height: null, cells }],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 3,
    conditionalFormats: [],
    images: [],
    charts: [],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as Worksheet;
}

interface Rect { x: number; y: number; w: number; h: number }
interface TextPaint { text: string; x: number; clip: Rect }
interface LinePaint { x1: number; x2: number; color: string; clip: Rect }

function intersect(a: Rect, b: Rect): Rect {
  const x = Math.max(a.x, b.x);
  const y = Math.max(a.y, b.y);
  return {
    x,
    y,
    w: Math.max(0, Math.min(a.x + a.w, b.x + b.w) - x),
    h: Math.max(0, Math.min(a.y + a.h, b.y + b.h) - y),
  };
}

/** A canvas double that records painted text and vertical lines together with
 *  the clip rectangle in force, so each pane's visible output can be read. */
function recordingContext(width = 600): {
  ctx: CanvasRenderingContext2D;
  texts: TextPaint[];
  lines: LinePaint[];
} {
  const texts: TextPaint[] = [];
  const lines: LinePaint[] = [];
  const full: Rect = { x: -1e9, y: -1e9, w: 2e9, h: 2e9 };
  let clip = full;
  const stack: Rect[] = [];
  let path: Rect | null = null;
  let start: { x: number; y: number } | null = null;
  let segments: Array<{ x1: number; x2: number }> = [];
  const state: Record<string, unknown> = {
    canvas: { width, height: 60 },
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '11px Calibri',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    globalAlpha: 1,
    letterSpacing: '0px',
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    fillText: (text: string, x: number) => { texts.push({ text, x, clip }); },
    save: () => { stack.push(clip); },
    restore: () => { clip = stack.pop() ?? full; },
    beginPath: () => { path = null; segments = []; },
    rect: (x: number, y: number, w: number, h: number) => { path = { x, y, w, h }; },
    clip: () => { if (path) clip = intersect(clip, path); },
    moveTo: (x: number, y: number) => { start = { x, y }; },
    lineTo: (x: number) => {
      if (start) segments.push({ x1: start.x, x2: x });
      start = null;
    },
    stroke: () => {
      for (const s of segments) {
        lines.push({ ...s, color: String(state.strokeStyle), clip });
      }
    },
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  const noOp = () => {};
  const ctx = new Proxy(state, {
    get(target, property) {
      return property in target ? target[property as string] : noOp;
    },
    set(target, property, value) {
      target[property as string] = value;
      return true;
    },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, texts, lines };
}

/** Centre x of a column-header label (labels are drawn centred). */
function headerX(texts: TextPaint[], label: string): number {
  const header = texts.find((t) => t.clip.y === 0 && t.text === label);
  if (!header) throw new Error(`no ${label} header`);
  return header.x;
}

/** Frozen pane clips end at the freeze line; scrollable pane clips start there. */
function paneSplit(texts: TextPaint[], text: string): {
  frozen: TextPaint[];
  scroll: TextPaint[];
} {
  const painted = texts.filter((t) => t.text === text && t.clip.w > 0);
  // Column-header labels are clipped per pane; the scrollable pane's header
  // clip starts at the freeze line.
  const freezeLine = Math.max(...texts.filter((t) => t.clip.y === 0).map((t) => t.clip.x));
  return {
    frozen: painted.filter((t) => t.clip.x + t.clip.w <= freezeLine),
    scroll: painted.filter((t) => t.clip.x >= freezeLine),
  };
}

const VIEW_UNSCROLLED = { row: 1, col: 4, rows: 1, cols: 6 };
const VIEW_SCROLLED_TO_H = { row: 1, col: 8, rows: 1, cols: 6 };

describe('text overflow across a frozen-column boundary', () => {
  it('continues a frozen-column value into the unscrolled scrollable pane', () => {
    const { ctx, texts } = recordingContext();
    renderViewport(ctx, worksheet([cell(2, LONG)]), STYLES, VIEW_UNSCROLLED,
      { width: 600, height: 60, freezeCols: 3 });

    const { frozen, scroll } = paneSplit(texts, LONG);
    expect(frozen).toHaveLength(1);
    expect(scroll).toHaveLength(1);
    // Unscrolled, both panes place the value at the same sheet position, so
    // the two clipped halves join seamlessly at the freeze line.
    expect(scroll[0].x).toBe(frozen[0].x);
  });

  it('moves the spilled part with the scrollable pane when it scrolls', () => {
    const { ctx, texts } = recordingContext();
    renderViewport(ctx, worksheet([cell(2, LONG)]), STYLES, VIEW_SCROLLED_TO_H,
      { width: 600, height: 60, freezeCols: 3 });

    const { frozen, scroll } = paneSplit(texts, LONG);
    expect(frozen).toHaveLength(1);
    expect(scroll).toHaveLength(1);
    // Four uniform scrolled-off columns (D:G) separate the panes' origins.
    const columnPx = headerX(texts, 'B') - headerX(texts, 'A');
    expect(scroll[0].x).toBeCloseTo(frozen[0].x - 4 * columnPx, 6);
  });

  it('does not spill a frozen-column value through a non-empty cell', () => {
    // D1 blocks at the freeze line, both unscrolled and scrolled.
    for (const view of [VIEW_UNSCROLLED, VIEW_SCROLLED_TO_H]) {
      const { ctx, texts } = recordingContext();
      renderViewport(ctx, worksheet([cell(2, LONG), cell(4, 'x')]), STYLES, view,
        { width: 600, height: 60, freezeCols: 3 });
      expect(paneSplit(texts, LONG).scroll).toHaveLength(0);
    }

    // F1 lets the value reach D:E unscrolled, and stops it before it reaches
    // H once the pane has scrolled there.
    const unscrolled = recordingContext();
    renderViewport(unscrolled.ctx, worksheet([cell(2, LONG), cell(6, 'x')]), STYLES,
      VIEW_UNSCROLLED, { width: 600, height: 60, freezeCols: 3 });
    const [spill] = paneSplit(unscrolled.texts, LONG).scroll;
    expect(spill.clip.x + spill.clip.w).toBeCloseTo(headerX(unscrolled.texts, 'F')
      - (headerX(unscrolled.texts, 'F') - headerX(unscrolled.texts, 'E')) / 2, 6);

    const scrolled = recordingContext();
    renderViewport(scrolled.ctx, worksheet([cell(2, LONG), cell(6, 'x')]), STYLES,
      VIEW_SCROLLED_TO_H, { width: 600, height: 60, freezeCols: 3 });
    expect(paneSplit(scrolled.texts, LONG).scroll).toHaveLength(0);
  });

  it('spills a right-aligned scrollable value into the frozen pane', () => {
    const frozenX: number[] = [];
    for (const view of [VIEW_UNSCROLLED, VIEW_SCROLLED_TO_H]) {
      const { ctx, texts } = recordingContext();
      renderViewport(ctx, worksheet([cell(5, RIGHT, RIGHT_XF)]), STYLES, view,
        { width: 600, height: 60, freezeCols: 3 });

      // The frozen pane shows the value's leftward spill at its sheet
      // position, including after column E has scrolled out of view.
      const { frozen } = paneSplit(texts, RIGHT);
      expect(frozen).toHaveLength(1);
      frozenX.push(frozen[0].x);
    }
    expect(frozenX[1]).toBe(frozenX[0]);
  });

  it('spills past a value in a hidden column, as paint does', () => {
    // D1 is populated but its column is hidden (zero width); paint walks
    // only visible columns, so B1 continues into the scrollable pane at E.
    const { ctx, texts } = recordingContext();
    const ws = {
      ...worksheet([cell(2, LONG), cell(4, 'hidden')]),
      colWidths: { 4: 0 },
    } as Worksheet;
    renderViewport(ctx, ws, STYLES, { row: 1, col: 5, rows: 1, cols: 6 },
      { width: 600, height: 60, freezeCols: 3 });
    const { frozen, scroll } = paneSplit(texts, LONG);
    expect(frozen).toHaveLength(1);
    expect(scroll).toHaveLength(1);
    expect(scroll[0].x).toBe(frozen[0].x);
  });

  it('honours a column-inherited alignment when spilling into the frozen pane', () => {
    // E1 has no own style; column E is right-aligned (§18.3.1.13).
    const { ctx, texts } = recordingContext();
    const ws = {
      ...worksheet([{ ...cell(5, RIGHT), styleIndex: undefined } as Cell]),
      colStyleRanges: [{ min: 5, max: 5, styleIndex: RIGHT_XF }],
    } as Worksheet;
    renderViewport(ctx, ws, STYLES, VIEW_SCROLLED_TO_H, { width: 600, height: 60, freezeCols: 3 });
    expect(paneSplit(texts, RIGHT).frozen).toHaveLength(1);
  });

  it('splits a centred frozen-column value across the freeze line', () => {
    const unscrolled = recordingContext();
    renderViewport(unscrolled.ctx, worksheet([cell(3, CENTERED, CENTER_XF)]), STYLES,
      VIEW_UNSCROLLED, { width: 600, height: 60, freezeCols: 3 });
    const split = paneSplit(unscrolled.texts, CENTERED);
    expect(split.frozen).toHaveLength(1);
    expect(split.scroll).toHaveLength(1);
    expect(split.scroll[0].x).toBe(split.frozen[0].x);

    // Scrolled to H, its right half lies in the scrolled-off D:G.
    const scrolled = recordingContext();
    renderViewport(scrolled.ctx, worksheet([cell(3, CENTERED, CENTER_XF)]), STYLES,
      VIEW_SCROLLED_TO_H, { width: 600, height: 60, freezeCols: 3 });
    const scrolledSplit = paneSplit(scrolled.texts, CENTERED);
    expect(scrolledSplit.frozen).toHaveLength(1);
    expect(scrolledSplit.scroll).toHaveLength(0);
  });

  it('crosses the freeze line in frozen rows too', () => {
    // Row 1 is frozen as well: the frozen corner and the frozen-rows strip
    // are a pane pair of their own.
    const { ctx, texts } = recordingContext();
    renderViewport(ctx, worksheet([cell(2, LONG)]), STYLES, { row: 2, col: 4, rows: 1, cols: 6 },
      { width: 600, height: 60, freezeRows: 1, freezeCols: 3 });
    const { frozen, scroll } = paneSplit(texts, LONG);
    expect(frozen).toHaveLength(1);
    expect(scroll).toHaveLength(1);
    expect(scroll[0].x).toBe(frozen[0].x);
  });

  it('crosses the mirrored freeze line on a right-to-left sheet', () => {
    // §18.3.1.87 mirrors the grid: higher columns lie physically left, so a
    // right-aligned frozen B1 spills toward C1 and on into the scrollable pane.
    const { ctx, texts } = recordingContext();
    const ws = { ...worksheet([cell(2, LONG, RIGHT_XF)]), rightToLeft: true } as Worksheet;
    renderViewport(ctx, ws, STYLES, VIEW_UNSCROLLED, { width: 600, height: 60, freezeCols: 3 });

    const painted = texts.filter((t) => t.text === LONG && t.clip.w > 0);
    expect(painted).toHaveLength(2);
    expect(painted[0].x).toBe(painted[1].x);
  });

  it('reports a frozen value beyond the canvas with frozen-pane geometry', () => {
    // A:J are frozen but only part of them fits the canvas. E1 still belongs
    // to the frozen pane, whatever the canvas cuts off.
    const runs: XlsxTextRunInfo[] = [];
    const { ctx, texts } = recordingContext(240);
    renderViewport(ctx, worksheet([cell(5, RIGHT, RIGHT_XF)]), STYLES,
      { row: 1, col: 11, rows: 1, cols: 4 },
      { width: 240, height: 60, freezeCols: 10, onTextRun: (run) => runs.push(run) });

    expect(runs.map((run) => run.cellRef)).toEqual(['E1']);
    const columnPx = headerX(texts, 'B') - headerX(texts, 'A');
    const columnA = headerX(texts, 'A') - columnPx / 2;
    expect(runs[0].x).toBeCloseTo(columnA + 4 * columnPx, 6);
  });

  it('reports each value once, from the pane that owns its column', () => {
    const runs: XlsxTextRunInfo[] = [];
    const { ctx } = recordingContext();
    renderViewport(
      ctx,
      worksheet([cell(2, LONG), cell(5, RIGHT, RIGHT_XF)]),
      STYLES,
      VIEW_UNSCROLLED,
      { width: 600, height: 60, freezeCols: 3, onTextRun: (run) => runs.push(run) },
    );

    expect(runs.map((run) => run.cellRef).sort()).toEqual(['B1', 'E1']);
    const b = runs.find((run) => run.cellRef === 'B1')!;
    const e = runs.find((run) => run.cellRef === 'E1')!;
    // Unscrolled, owner-pane geometry is contiguous sheet geometry.
    expect(e.x - b.x).toBeCloseTo(3 * b.width, 6);
  });

  it('reports a value painted only in the other pane once, with owning-pane geometry', () => {
    // Scrolled to H, E1 is culled from its own (scrollable) pane but still
    // paints its leftward spill in the frozen pane.
    const runs: XlsxTextRunInfo[] = [];
    const { ctx, texts } = recordingContext();
    renderViewport(ctx, worksheet([cell(5, RIGHT, RIGHT_XF)]), STYLES, VIEW_SCROLLED_TO_H,
      { width: 600, height: 60, freezeCols: 3, onTextRun: (run) => runs.push(run) });

    expect(runs.map((run) => run.cellRef)).toEqual(['E1']);
    // In the scrollable pane, E lies three columns (E:G) left of the freeze line.
    const freezeLine = Math.max(...texts.filter((t) => t.clip.y === 0).map((t) => t.clip.x));
    expect(runs[0].x).toBeCloseTo(freezeLine - 3 * runs[0].width, 6);
  });

  it('does not re-stroke a border its painted overflow anchor already draws', () => {
    // No freeze: C1 overflows into the visible D1 and owns a right hairline.
    // C1 is painted as an off-screen anchor and strokes that edge itself; D1
    // must not stroke it a second time.
    const { ctx, lines } = recordingContext();
    const ws = { ...worksheet([cell(3, LONG, LEFT_HAIR_BORDER_XF)]), freezeCols: 0 } as Worksheet;
    renderViewport(ctx, ws, STYLES, VIEW_UNSCROLLED, { width: 600, height: 60 });

    const edges = lines.filter((line) => (
      line.x1 === line.x2 && /255,\s*0,\s*0|ff0000/i.test(line.color) && line.clip.w > 0 &&
      line.x1 >= line.clip.x - 1 && line.x1 <= line.clip.x + line.clip.w + 1
    ));
    expect(new Set(edges.map((line) => Math.round(line.x1))).size).toBe(edges.length);
  });

  it('does not move an overflow anchor border onto its neighbour in a right-to-left sheet', () => {
    // No freeze, mirrored grid: C1 (right border) spills into the visible D1.
    // Its edge borders B1, which is off-screen, so no red edge is visible.
    const { ctx, lines } = recordingContext();
    const ws = {
      ...worksheet([cell(3, LONG, RIGHT_ALIGNED_RIGHT_BORDER_XF)]),
      freezeCols: 0,
      rightToLeft: true,
    } as Worksheet;
    renderViewport(ctx, ws, STYLES, VIEW_UNSCROLLED, { width: 600, height: 60 });

    const visible = lines.filter((line) => (
      line.x1 === line.x2 && /255,\s*0,\s*0|ff0000/i.test(line.color) && line.clip.w > 0 &&
      line.x1 > line.clip.x && line.x1 < line.clip.x + line.clip.w
    ));
    expect(visible).toEqual([]);
  });

  it('keeps first-column edges of a scrolled right-to-left view without overflow', () => {
    // No text, no freeze: C1's right border lies on its scrolled-away side.
    // Only the viewport's own first column takes first-column edge handling.
    const probe = recordingContext();
    const ws = {
      ...worksheet([{ ...cell(3, ''), styleIndex: RIGHT_BORDER_XF, value: { type: 'empty' } } as Cell]),
      freezeCols: 0,
      rightToLeft: true,
    } as Worksheet;
    renderViewport(probe.ctx, ws, STYLES, VIEW_UNSCROLLED, { width: 600, height: 60 });
    const columnPx = headerX(probe.texts, 'E') - headerX(probe.texts, 'D');

    const { ctx, lines } = recordingContext();
    renderViewport(ctx, ws, STYLES, { row: 1, col: 3, rows: 1, cols: 6 },
      { width: 600, height: 60, scrollOffsetX: Math.abs(columnPx) });
    const visible = lines.filter((line) => (
      line.x1 === line.x2 && /255,\s*0,\s*0|ff0000/i.test(line.color) && line.clip.w > 0 &&
      line.x1 > line.clip.x && line.x1 < line.clip.x + line.clip.w
    ));
    expect(visible).toEqual([]);
  });

  it('keeps the frozen pane left gridline when its first column is hidden', () => {
    // Column A is hidden: B is the frozen pane's first column and still
    // draws the pane's left gridline.
    const { ctx, texts, lines } = recordingContext();
    const ws = { ...worksheet([]), colWidths: { 1: 0 } } as Worksheet;
    renderViewport(ctx, ws, STYLES, VIEW_UNSCROLLED, { width: 600, height: 60, freezeCols: 3 });
    // The frozen column-header clip starts where the frozen pane does.
    const paneLeft = Math.min(...texts.filter((t) => t.clip.y === 0).map((t) => t.clip.x));
    const leftEdge = lines.filter((line) => (
      line.x1 === line.x2 && line.color === '#d0d0d0' && line.clip.x === paneLeft &&
      Math.abs(line.x1 - paneLeft) <= 1
    ));
    expect(leftEdge.length).toBeGreaterThan(0);
  });

  it('keeps a scrolled-away merge border at the pane edge when another row overflows', () => {
    // F2:G2 is an empty merge whose G2 carries a right border; both have
    // scrolled out of the pane that starts at H. B1's overflow prepends
    // columns to the pane, which must not stop H2 from drawing that edge.
    const ws = {
      ...worksheet([cell(2, LONG)]),
      mergeCells: [{ top: 2, left: 6, bottom: 2, right: 7 }],
    } as Worksheet;
    ws.rows = [
      ...ws.rows,
      { index: 2, height: null, cells: [
        { row: 2, col: 7, styleIndex: RIGHT_BORDER_XF, value: { type: 'empty' } } as Cell,
      ] },
    ];
    const view = { row: 1, col: 8, rows: 2, cols: 6 };
    const { ctx, texts, lines } = recordingContext();
    renderViewport(ctx, ws, STYLES, view, { width: 600, height: 60, freezeCols: 3 });

    expect(paneSplit(texts, LONG).scroll).toHaveLength(1);
    const freezeLine = Math.max(...texts.filter((t) => t.clip.y === 0).map((t) => t.clip.x));
    const edge = lines.filter((line) => (
      line.x1 === line.x2 && /255,\s*0,\s*0|ff0000/i.test(line.color) &&
      line.clip.x === freezeLine && Math.abs(line.x1 - freezeLine) <= 1
    ));
    expect(edge.length).toBeGreaterThan(0);
  });

  it('keeps the left-pane neighbour border at the scrollable pane edge', () => {
    // C1 carries a right border. Pulling the frozen B1 value into the
    // scrollable pane must not stop D1, the pane's first visible column, from
    // drawing that shared edge at the freeze line.
    const { ctx, texts, lines } = recordingContext();
    renderViewport(
      ctx,
      worksheet([cell(2, LONG), { ...cell(3, ''), styleIndex: RIGHT_BORDER_XF, value: { type: 'empty' } } as Cell]),
      STYLES,
      VIEW_UNSCROLLED,
      { width: 600, height: 60, freezeCols: 3 },
    );

    const { scroll } = paneSplit(texts, LONG);
    expect(scroll).toHaveLength(1);
    const scrollClip = scroll[0].clip;
    const edge = lines.filter((line) => (
      line.x1 === line.x2 &&
      /255,\s*0,\s*0|ff0000/i.test(line.color) &&
      line.clip.x === scrollClip.x &&
      Math.abs(line.x1 - scrollClip.x) <= 1
    ));
    expect(edge.length).toBeGreaterThan(0);
  });
});

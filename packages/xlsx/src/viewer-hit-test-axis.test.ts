import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import {
  HEADER_W,
  HEADER_H,
  colWidthToPx,
  rowHeightToPx,
  getMdwForWorksheet,
} from './renderer.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * C3: getCellAt / getHeaderHit resolve the scrollable region via the O(log n)
 * {@link AxisMetrics.indexAt} instead of a linear scan from the first scrollable
 * index to the sheet limit (1,048,576 rows / 16,384 cols). These tests pin that
 * the fast path returns the SAME cell as the old linear scan across custom row
 * heights, custom column widths, and frozen panes — including the exact boundary
 * conditions the scan had (on a cell edge, at the frozen boundary, in the header
 * strip, and past the sheet's last cell).
 *
 * The oracle below is a verbatim re-implementation of the ORIGINAL linear scan
 * (`for (r = freezeRows + 1; r <= 1048576; r++) { acc += size; if (content <
 * acc) return r; } return null`) so the fast path is checked against the
 * behavior it replaced, with expectations fixed by that oracle. The viewer is
 * driven through its real `getCellAt`; `getBoundingClientRect()` returns
 * top/left = 0 in the fake DOM and the sheet is LTR, so a client (x, y) maps
 * directly to the logical layout coordinate the geometry math consumes.
 */

/** A worksheet with sparse custom sizes + a 2×2 frozen pane. Values chosen so
 *  the frozen extent and several scrollable cells sit at non-round pixels. */
function makeSheet(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    // Excel column-width units (chars); colWidthToPx converts with mdw.
    colWidths: { 1: 12, 2: 20, 4: 30, 7: 5 },
    // Row-height units (points); rowHeightToPx converts.
    rowHeights: { 1: 25, 2: 40, 5: 60, 9: 12 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 2,
    freezeCols: 2,
    conditionalFormats: [],
    images: [],
    charts: [],
  } as Worksheet;
}

/** Verbatim copy of the ORIGINAL linear-scan getCellAt (pre-C3), used as the
 *  oracle. `cs` (cellScale) = 1, LTR, rect origin (0,0). */
function oldGetCellAt(
  ws: Worksheet,
  clientX: number,
  clientY: number,
  scrollTop: number,
  scrollLeft: number,
): { row: number; col: number } | null {
  const lx = clientX;
  const ly = clientY;
  if (lx < HEADER_W || ly < HEADER_H) return null;
  const innerX = lx - HEADER_W;
  const innerY = ly - HEADER_H;
  const freezeRows = ws.freezeRows ?? 0;
  const freezeCols = ws.freezeCols ?? 0;
  const mdw = getMdwForWorksheet(ws);

  let frozenH = 0;
  const frozenRowH: number[] = [];
  for (let r = 1; r <= freezeRows; r++) {
    const h = rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    frozenRowH.push(h);
    frozenH += h;
  }
  let frozenW = 0;
  const frozenColW: number[] = [];
  for (let c = 1; c <= freezeCols; c++) {
    const w = colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw);
    frozenColW.push(w);
    frozenW += w;
  }

  let row: number;
  if (innerY < frozenH) {
    row = -1;
    let acc = 0;
    for (let r = 0; r < freezeRows; r++) {
      acc += frozenRowH[r];
      if (innerY < acc) { row = r + 1; break; }
    }
    if (row === -1) return null;
  } else {
    const contentY = innerY - frozenH + scrollTop;
    row = -1;
    let acc = 0;
    for (let r = freezeRows + 1; r <= 1048576; r++) {
      acc += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
      if (contentY < acc) { row = r; break; }
    }
    if (row === -1) return null;
  }

  let col: number;
  if (innerX < frozenW) {
    col = -1;
    let acc = 0;
    for (let c = 0; c < freezeCols; c++) {
      acc += frozenColW[c];
      if (innerX < acc) { col = c + 1; break; }
    }
    if (col === -1) return null;
  } else {
    const contentX = innerX - frozenW + scrollLeft;
    col = -1;
    let acc = 0;
    for (let c = freezeCols + 1; c <= 16384; c++) {
      acc += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw);
      if (contentX < acc) { col = c; break; }
    }
    if (col === -1) return null;
  }
  return { row, col };
}

/** Mount a viewer with the fixture worksheet injected and scroll set. */
function mountViewer(ws: Worksheet, scrollTop = 0, scrollLeft = 0): XlsxViewer {
  installDom();
  const v = new XlsxViewer(makeContainer() as unknown as HTMLElement);
  const priv = v as unknown as {
    currentWorksheet: Worksheet | null;
    scrollHost: { scrollTop: number; scrollLeft: number; scrollWidth: number; clientWidth: number };
  };
  priv.currentWorksheet = ws;
  priv.scrollHost.scrollTop = scrollTop;
  priv.scrollHost.scrollLeft = scrollLeft;
  return v;
}

describe('XlsxViewer.getCellAt — AxisMetrics fast path matches old linear scan', () => {
  it('keeps row and column header hits aligned with the rounded grid at non-integer zoom', () => {
    const ws = makeSheet();
    const v = mountViewer(ws, 37, 51);
    const internal = v as unknown as {
      viewport: { setScale(scale: number): void };
      selectionInput: {
        getHeaderHit(x: number, y: number):
          | { kind: 'corner' }
          | { kind: 'row'; row: number }
          | { kind: 'col'; col: number }
          | null;
      };
    };
    internal.viewport.setScale(1.25);
    const headerW = Math.round(HEADER_W * 1.25);
    const headerH = Math.round(HEADER_H * 1.25);

    for (const y of [headerH + 1, headerH + 42, headerH + 103]) {
      const cell = v.getCellAt(headerW + 1, y);
      const header = internal.selectionInput.getHeaderHit(1, y);
      expect(header?.kind).toBe('row');
      if (header?.kind === 'row') expect(header.row).toBe(cell?.row);
    }
    for (const x of [headerW + 1, headerW + 90, headerW + 211]) {
      const cell = v.getCellAt(x, headerH + 1);
      const header = internal.selectionInput.getHeaderHit(x, 1);
      expect(header?.kind).toBe('col');
      if (header?.kind === 'col') expect(header.col).toBe(cell?.col);
    }
    v.destroy();
  });

  it('matches the oracle over a grid sweep at several scroll offsets', () => {
    const ws = makeSheet();
    for (const [scrollTop, scrollLeft] of [
      [0, 0],
      [37, 51],
      [123, 200],
      [1000, 800],
    ] as const) {
      const v = mountViewer(ws, scrollTop, scrollLeft);
      // Sweep the header strip and the scrollable body at a fine granularity.
      for (let y = 0; y <= 260; y += 7) {
        for (let x = 0; x <= 320; x += 9) {
          const got = v.getCellAt(x, y);
          const want = oldGetCellAt(ws, x, y, scrollTop, scrollLeft);
          expect(got, `x=${x} y=${y} st=${scrollTop} sl=${scrollLeft}`).toEqual(want);
        }
      }
      v.destroy();
    }
  });

  it('resolves an exact-on-cell-edge scrollable point to the same cell as the oracle', () => {
    const ws = makeSheet();
    const mdw = getMdwForWorksheet(ws);
    const freezeRows = ws.freezeRows;
    const freezeCols = ws.freezeCols;

    // Frozen extents (unscaled px).
    let frozenH = 0;
    for (let r = 1; r <= freezeRows; r++) frozenH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
    let frozenW = 0;
    for (let c = 1; c <= freezeCols; c++) frozenW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw);

    // Height of the first scrollable row (row 3) and width of first scrollable col (col 3).
    const row3H = rowHeightToPx(ws.rowHeights[3] ?? ws.defaultRowHeight);
    const col3W = colWidthToPx(ws.colWidths[3] ?? ws.defaultColWidth, mdw);

    const v = mountViewer(ws, 0, 0);
    // The point exactly on the trailing edge of row 3 / col 3 (contentY == row3H)
    // belongs to row 4 / col 4 in the half-open `content < acc` convention — the
    // oracle decides, the fast path must agree.
    const yEdge = HEADER_H + frozenH + row3H; // exactly at row-3 bottom edge
    const xEdge = HEADER_W + frozenW + col3W; // exactly at col-3 right edge
    expect(v.getCellAt(xEdge, yEdge)).toEqual(oldGetCellAt(ws, xEdge, yEdge, 0, 0));
    // And just inside row 3 / col 3.
    expect(v.getCellAt(xEdge - 0.5, yEdge - 0.5)).toEqual(
      oldGetCellAt(ws, xEdge - 0.5, yEdge - 0.5, 0, 0),
    );
    v.destroy();
  });

  it('returns null in the header strip and null past the last cell, matching the oracle', () => {
    const ws = makeSheet();
    const v = mountViewer(ws, 0, 0);

    // Header strip (top-left corner region) → null for both.
    expect(v.getCellAt(1, 1)).toBeNull();
    expect(oldGetCellAt(ws, 1, 1, 0, 0)).toBeNull();

    // Scroll so far past the end that even the last row/col is above the point.
    // Row max = 1,048,576; scrolling the whole sheet height + slack lands past it.
    const bigScroll = 1_048_576 * rowHeightToPx(ws.defaultRowHeight) + 10_000;
    const bigScrollX = 16_384 * colWidthToPx(ws.defaultColWidth, getMdwForWorksheet(ws)) + 10_000;
    const vFar = mountViewer(makeSheet(), bigScroll, bigScrollX);
    const wsFar = (vFar as unknown as { currentWorksheet: Worksheet }).currentWorksheet;
    // Deep inside the scrollable body but scrolled past the last cell.
    const px = HEADER_W + 100;
    const py = HEADER_H + 100;
    expect(vFar.getCellAt(px, py)).toBeNull();
    expect(oldGetCellAt(wsFar, px, py, bigScroll, bigScrollX)).toBeNull();
    vFar.destroy();
    v.destroy();
  });
});

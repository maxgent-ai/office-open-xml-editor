import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import type { XlsxViewerOptions, CellAddress } from './viewer.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import { HEADER_W, HEADER_H } from './renderer.js';
import type { Hyperlink, Worksheet } from './types.js';
import * as core from '@silurus/ooxml-core';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX1 — clickable hyperlinks in the XLSX viewer. Unlike docx/pptx (which have a
 * span text-layer), xlsx activation is CELL hit-test based: a click resolves a
 * {row,col} via `getCellAt`, looks it up in the per-sheet hyperlink map, and
 * dispatches a `HyperlinkTarget`. These tests inject a worksheet with known
 * hyperlinks, drive the viewer through its real pointer listeners (rect origin
 * is (0,0) and the sheet is LTR in the fake DOM, so a client (x,y) maps directly
 * to the logical layout coordinate), and assert the callback / default handler.
 *
 * Hyperlink `row`/`col` are 1-based (parser `parse_cell_ref` / renderer
 * `hyperlinkMap` keying), matching `getCellAt`'s 1-based return.
 */

/** A worksheet with default cell sizes and the given hyperlinks. Default col
 *  width 8.43 / row height 15 give integer-ish px cells so a point a few px past
 *  the header lands in cell (row 1, col 1). */
function makeSheet(hyperlinks: Hyperlink[]): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
    hyperlinks,
  } as unknown as Worksheet;
}

interface Priv {
  hostWindow: { open?: typeof window.open };
  currentWorksheet: Worksheet | null;
  scrollHost: {
    scrollTop: number;
    scrollLeft: number;
    scrollWidth: number;
    scrollHeight: number;
    clientWidth: number;
    clientHeight: number;
    clientLeft: number;
    clientTop: number;
    dispatch(type: string, event?: unknown): void;
  };
  buildHyperlinkMap(ws: Worksheet): void;
  hyperlinks: {
    dispatch(cell: CellAddress): boolean;
    navigateInternal(location: string): Promise<void>;
  };
}

/** Mount a viewer, inject the fixture worksheet, and build its hyperlink map so
 *  hit-testing resolves against the injected sheet (bypasses async load). */
function mountViewer(ws: Worksheet, opts: XlsxViewerOptions = {}): { v: XlsxViewer; priv: Priv } {
  installDom();
  const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, opts);
  const priv = v as unknown as Priv;
  priv.currentWorksheet = ws;
  // Nonzero client size so the pointerdown scrollbar-gutter guard passes and the
  // press reaches selection/click handling.
  priv.scrollHost.clientWidth = 800;
  priv.scrollHost.clientHeight = 600;
  priv.buildHyperlinkMap(ws);
  return { v, priv };
}

/** Fire a full mouse click (pointerdown → pointerup) at client (x, y) through
 *  the viewer's real scrollHost listeners. `button:0`, `pointerType:'mouse'`. */
function clickAt(priv: Priv, x: number, y: number): void {
  const base = { button: 0, buttons: 1, pointerId: 1, pointerType: 'mouse', shiftKey: false, clientX: x, clientY: y, preventDefault() {} };
  priv.scrollHost.dispatch('pointerdown', base);
  priv.scrollHost.dispatch('pointerup', { ...base, buttons: 0 });
}

// A point a few px inside the first data cell (row 1, col 1): past the 50×22 header.
const CELL_A1_X = HEADER_W + 5;
const CELL_A1_Y = HEADER_H + 5;

describe('XlsxViewer IX1 hyperlink click', () => {
  it('fires onHyperlinkClick with {kind:external,url} for an external link', () => {
    const seen: HyperlinkTarget[] = [];
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { onHyperlinkClick: (t) => seen.push(t) });

    // Sanity: the click coordinate resolves to the hyperlinked cell.
    expect(v.getCellAt(CELL_A1_X, CELL_A1_Y)).toEqual({ row: 1, col: 1 });

    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(seen).toEqual([{ kind: 'external', url: 'https://example.com/' }]);
    v.destroy();
  });

  it('fires onHyperlinkClick with {kind:internal,ref} for a location link', () => {
    const seen: HyperlinkTarget[] = [];
    const ws = makeSheet([{ col: 1, row: 1, url: null, location: 'Sheet2!B7' }]);
    const { v, priv } = mountViewer(ws, { onHyperlinkClick: (t) => seen.push(t) });

    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(seen).toEqual([{ kind: 'internal', ref: 'Sheet2!B7' }]);
    v.destroy();
  });

  it('does not fire when the clicked cell carries no hyperlink', () => {
    const seen: HyperlinkTarget[] = [];
    const ws = makeSheet([{ col: 3, row: 3, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { onHyperlinkClick: (t) => seen.push(t) });

    // Click the top-left cell (no hyperlink there — the link is at C3).
    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(seen).toEqual([]);
    v.destroy();
  });

  it('passes a blocked scheme through to a custom callback verbatim (raw target)', () => {
    // A malicious `javascript:` url still reaches a custom callback — sanitising
    // is the DEFAULT handler's job, not the callback's. The callback owns policy.
    const seen: HyperlinkTarget[] = [];
    const ws = makeSheet([{ col: 1, row: 1, url: 'javascript:alert(1)', location: null }]);
    const { v, priv } = mountViewer(ws, { onHyperlinkClick: (t) => seen.push(t) });

    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(seen).toEqual([{ kind: 'external', url: 'javascript:alert(1)' }]);
    v.destroy();
  });
});

describe('XlsxViewer IX1 default hyperlink handler (no callback)', () => {
  it('opens a safe external url in a new tab via window.open', () => {
    const openSpy = vi.fn();
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { priv } = mountViewer(ws); // no onHyperlinkClick → default handler

    priv.hostWindow.open = openSpy;

    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(openSpy).toHaveBeenCalledTimes(1);
    expect(openSpy).toHaveBeenCalledWith('https://example.com/', '_blank', 'noopener,noreferrer');
  });

  it('does NOT navigate for a blocked javascript: scheme (default handler sanitises)', () => {
    const openSpy = vi.fn();
    const ws = makeSheet([{ col: 1, row: 1, url: 'javascript:alert(1)', location: null }]);
    const { priv } = mountViewer(ws); // default handler
    priv.hostWindow.open = openSpy;

    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(openSpy).not.toHaveBeenCalled();

    // Cross-check the core policy directly: the blocked scheme yields no
    // navigation (openExternalHyperlink returns false) with an injected window.
    const navigated = core.openExternalHyperlink('javascript:alert(1)', undefined, { open: openSpy });
    expect(navigated).toBe(false);
    expect(openSpy).not.toHaveBeenCalled();
  });

  it('navigates to the referenced sheet and cell for an internal Sheet!Cell location', async () => {
    const ws = makeSheet([{ col: 1, row: 1, url: null, location: 'Sheet2!A1' }]);
    const { v, priv } = mountViewer(ws);
    // Stub the workbook so sheetNames resolves and goToSheet is observable.
    const goSpy = vi.fn().mockResolvedValue(undefined);
    const scrollSpy = vi.fn().mockResolvedValue(undefined);
    Object.assign(v as unknown as { wb: unknown }, {
      wb: { sheetNames: ['Sheet1', 'Sheet2'], sheetCount: 2 },
    });
    (v as unknown as { goToSheet: unknown }).goToSheet = goSpy;
    (v as unknown as { scrollToCell: unknown }).scrollToCell = scrollSpy;

    await priv.hyperlinks.navigateInternal('Sheet2!A1');
    expect(goSpy).toHaveBeenCalledWith(1);
    expect(scrollSpy).toHaveBeenCalledWith('A1');
  });

  it('is a no-op for an internal location whose sheet is unknown', async () => {
    const ws = makeSheet([{ col: 1, row: 1, url: null, location: 'Ghost!A1' }]);
    const { v, priv } = mountViewer(ws);
    const goSpy = vi.fn().mockResolvedValue(undefined);
    Object.assign(v as unknown as { wb: unknown }, {
      wb: { sheetNames: ['Sheet1', 'Sheet2'], sheetCount: 2 },
    });
    (v as unknown as { goToSheet: unknown }).goToSheet = goSpy;

    await priv.hyperlinks.navigateInternal('Ghost!A1');
    expect(goSpy).not.toHaveBeenCalled();
  });

  it('resolves a bare defined name and navigates to its first cell', async () => {
    const ws = {
      ...makeSheet([{ col: 1, row: 1, url: null, location: 'MyName' }]),
      definedNames: [{ name: 'MyName', formula: "='Sheet 2'!$B$7:$B$9" }],
    };
    const { v, priv } = mountViewer(ws);
    const goSpy = vi.fn().mockResolvedValue(undefined);
    const scrollSpy = vi.fn().mockResolvedValue(undefined);
    Object.assign(v as unknown as { wb: unknown }, {
      wb: { sheetNames: ['Sheet1', 'Sheet 2'], sheetCount: 2 },
    });
    (v as unknown as { goToSheet: unknown }).goToSheet = goSpy;
    (v as unknown as { scrollToCell: unknown }).scrollToCell = scrollSpy;

    await priv.hyperlinks.navigateInternal('MyName');
    expect(goSpy).toHaveBeenCalledWith(1);
    expect(scrollSpy).toHaveBeenCalledWith('$B$7');
  });
});

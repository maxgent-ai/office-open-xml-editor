import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import type { XlsxViewerOptions, CellAddress } from './viewer.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import { HEADER_W, HEADER_H } from './renderer.js';
import type { Hyperlink, Worksheet } from './types.js';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX1 `enableHyperlinks` — a viewer-level policy switch (default `true`). When
 * `false`, XLSX hyperlink interactivity is disabled entirely: the cell hit-test
 * reports no hyperlink under any cell, so there is no pointer cursor over a link
 * (the pointermove cursor branch is literally `hyperlinks.at(hovered) ?
 * 'pointer' : ''`), no default open/navigate, and `onHyperlinkClick` is never
 * called. The single gate lives in `HyperlinkDispatcher.at`, which every consumer
 * (cursor affordance + click dispatch) already funnels through.
 */

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
  currentWorksheet: Worksheet | null;
  scrollHost: {
    clientWidth: number;
    clientHeight: number;
    dispatch(type: string, event?: unknown): void;
  };
  buildHyperlinkMap(ws: Worksheet): void;
  hyperlinks: { at(cell: CellAddress): Hyperlink | null };
}

function mountViewer(ws: Worksheet, opts: XlsxViewerOptions = {}): { v: XlsxViewer; priv: Priv } {
  installDom();
  const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, opts);
  const priv = v as unknown as Priv;
  priv.currentWorksheet = ws;
  priv.scrollHost.clientWidth = 800;
  priv.scrollHost.clientHeight = 600;
  priv.buildHyperlinkMap(ws);
  return { v, priv };
}

function clickAt(priv: Priv, x: number, y: number): void {
  const base = { button: 0, buttons: 1, pointerId: 1, pointerType: 'mouse', shiftKey: false, clientX: x, clientY: y, preventDefault() {} };
  priv.scrollHost.dispatch('pointerdown', base);
  priv.scrollHost.dispatch('pointerup', { ...base, buttons: 0 });
}

const CELL_A1_X = HEADER_W + 5;
const CELL_A1_Y = HEADER_H + 5;

describe('XlsxViewer — enableHyperlinks option', () => {
  it('does not hit-test a hyperlink cell when enableHyperlinks is false (no pointer cursor)', () => {
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { enableHyperlinks: false });
    // The hit-test reports nothing, so the pointermove cursor branch
    // (`hyperlinks.at(hovered) ? 'pointer' : ''`) can never show a pointer.
    expect(priv.hyperlinks.at({ row: 1, col: 1 })).toBeNull();
    v.destroy();
  });

  it('does not fire onHyperlinkClick on click when enableHyperlinks is false', () => {
    const seen: HyperlinkTarget[] = [];
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { enableHyperlinks: false, onHyperlinkClick: (t) => seen.push(t) });
    // Sanity: the click coordinate still resolves to the (link-bearing) cell.
    expect(v.getCellAt(CELL_A1_X, CELL_A1_Y)).toEqual({ row: 1, col: 1 });
    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(seen).toEqual([]);
    v.destroy();
  });

  it('does not run the default open handler on click when enableHyperlinks is false', () => {
    const openSpy = vi.fn();
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { enableHyperlinks: false });
    vi.stubGlobal('window', { devicePixelRatio: 1, open: openSpy });
    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(openSpy).not.toHaveBeenCalled();
    v.destroy();
  });

  it('hit-tests and dispatches normally when the option is omitted (default true)', () => {
    const seen: HyperlinkTarget[] = [];
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { onHyperlinkClick: (t) => seen.push(t) });
    expect(priv.hyperlinks.at({ row: 1, col: 1 })).not.toBeNull();
    clickAt(priv, CELL_A1_X, CELL_A1_Y);
    expect(seen).toEqual([{ kind: 'external', url: 'https://example.com/' }]);
    v.destroy();
  });

  it('hit-tests normally when enableHyperlinks is explicitly true', () => {
    const ws = makeSheet([{ col: 1, row: 1, url: 'https://example.com/', location: null }]);
    const { v, priv } = mountViewer(ws, { enableHyperlinks: true });
    expect(priv.hyperlinks.at({ row: 1, col: 1 })).not.toBeNull();
    v.destroy();
  });
});

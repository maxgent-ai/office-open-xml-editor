import { describe, expect, it } from 'vitest';
import { getGridGeometryForWorksheet, renderViewport } from './renderer.js';
import { createSheetViewModel } from './internal/sheet-viewer-runtime.js';
import {
  bindInitialAnchorSizes,
  captureInitialAnchorSizes,
} from './internal/initial-anchor-sizes.js';
import type { Styles, Worksheet } from './types.js';

// A viewer projection bound to captured initial anchor sizes must still paint
// its shape groups. The retained sizes are keyed by the worksheet's own anchor
// objects, so z-ordered per-leaf painting has to keep that identity: a copied
// anchor resolved to the non-renderable size and every group vanished while
// pictures (painted with the original anchor) stayed visible.
const EMU = 9525;
const EMPTY_STYLES = { fonts: [], fills: [], borders: [], cellXfs: [] } as unknown as Styles;
const RANGE = { row: 1, col: 1, rows: 20, cols: 20 };

function sheet(): Worksheet {
  const colWidths: Record<number, number> = {};
  const rowHeights: Record<number, number> = {};
  for (let index = 0; index <= 24; index++) {
    colWidths[index] = 20;
    rowHeights[index] = 45;
  }
  const leaf = (zOrder: number) => ({
    zOrder, x: 0, y: 0, w: 1, h: 1, rot: 0,
    strokeColor: '000000', strokeWidth: 1,
    geom: {
      type: 'custom',
      paths: [{ w: 10, h: 10, commands: [
        { op: 'moveTo', x: 0, y: 0 }, { op: 'lineTo', x: 10, y: 0 },
        { op: 'lineTo', x: 10, y: 10 }, { op: 'close' },
      ] }],
    },
  });
  return {
    name: 'Sheet1', rows: [], colWidths, rowHeights,
    defaultColWidth: 8.43, defaultRowHeight: 15, defaultRowHeightCustom: true, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], images: [],
    shapeGroups: [Object.freeze({
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 2, toColOff: 0, toRow: 2, toRowOff: 0,
      nativeExtCx: 100 * EMU, nativeExtCy: 50 * EMU,
      anchorTag: 'twoCellAnchor', editAs: 'oneCell',
      shapes: [leaf(1), leaf(2)],
    })],
  } as unknown as Worksheet;
}

function countLineTo(ws: Worksheet): number {
  let lines = 0;
  const state: Record<string, unknown> = { globalAlpha: 1 };
  const ctx = new Proxy(state, {
    get(_target, prop) {
      if (prop === 'canvas') return { width: 800, height: 600 };
      if (prop === 'measureText') return () => ({ width: 7 });
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      if (prop === 'lineTo') return () => { lines++; };
      if (typeof prop === 'string') return () => undefined;
    },
    set(target, prop, value) { target[String(prop)] = value; return true; },
  }) as unknown as CanvasRenderingContext2D;
  renderViewport(ctx, ws, EMPTY_STYLES, RANGE, { loadedImages: new Map() });
  return lines;
}

describe('shape groups on a projection bound to initial anchor sizes', () => {
  it('paints every leaf exactly like the unbound worksheet', () => {
    const unbound = countLineTo(createSheetViewModel(sheet()));
    expect(unbound).toBeGreaterThan(0);

    const view = createSheetViewModel(sheet());
    const reference = captureInitialAnchorSizes(view, getGridGeometryForWorksheet(view));
    expect(reference).toBeDefined();
    bindInitialAnchorSizes(view, reference!);
    expect(countLineTo(view)).toBe(unbound);
  });
});

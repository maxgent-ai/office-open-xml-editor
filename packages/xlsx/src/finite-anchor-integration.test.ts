import { describe, expect, it } from 'vitest';
import { normalizeXlsxWorksheetPolicy } from '@silurus/ooxml-core/worker';
import { getSheetRenderCache } from './renderer.js';
import { WorksheetViewProjectionCache } from './worker-protocol.js';
import { bindWorksheetPolicy, getWorksheetPolicy } from './worksheet-policy-context.js';
import type { InitialAnchorSizeReference } from './internal/initial-anchor-sizes.js';
import type { Worksheet } from './types.js';

function sourceSheet(): Worksheet {
  return {
    name: 'Finite anchored sheet', rows: [], colWidths: {}, rowHeights: {},
    defaultColWidth: 8.43, defaultRowHeight: 15, mergeCells: [],
    freezeRows: 0, freezeCols: 0, conditionalFormats: [], charts: [], shapeGroups: [],
    images: [{
      imagePath: 'xl/media/image.png', mimeType: 'image/png',
      anchorTag: 'twoCellAnchor', editAs: 'oneCell',
      nativeExtCx: 952500, nativeExtCy: 476250,
      fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
      toCol: 1, toColOff: 0, toRow: 1, toRowOff: 0,
    }],
  };
}

const reference: InitialAnchorSizeReference = {
  version: 1, imageCount: 1, shapeGroupCount: 0,
  entries: [{
    family: 'image', index: 0, cx: 952500, cy: 476250,
    fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
    toCol: 1, toColOff: 0, toRow: 1, toRowOff: 0,
  }],
};

const projection = { id: 71, revision: 1, initialAnchorSizes: reference };

describe('finite worksheet policy with retained anchor projections', () => {
  it('keeps a raised policy and renderer cap in an anchor-only worker projection', () => {
    const source = sourceSheet();
    const policy = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxCells: 300000 } });
    bindWorksheetPolicy(source, policy);
    const cache = new WorksheetViewProjectionCache();
    const first = cache.resolve(source, 0, projection, undefined);
    expect(first.worksheet).not.toBe(source);
    expect(getWorksheetPolicy(first.worksheet)).toBe(policy);
    expect(getSheetRenderCache(first.worksheet).coordinateIndexLimit).toBe(300000);
    const repeat = cache.resolve(source, 0, projection, undefined);
    expect(repeat.worksheet).toBe(first.worksheet);
    expect(repeat.created).toBe(false);
  });

  it('rebuilds the retained-anchor projection when the source policy identity changes', () => {
    const source = sourceSheet();
    const raised = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxCells: 300000 } });
    const lowered = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: { maxCells: 12 } });
    bindWorksheetPolicy(source, raised);
    const cache = new WorksheetViewProjectionCache();
    const first = cache.resolve(source, 0, projection, undefined);
    bindWorksheetPolicy(source, lowered);
    const second = cache.resolve(source, 0, projection, undefined);
    expect(second.worksheet).not.toBe(first.worksheet);
    expect(getWorksheetPolicy(second.worksheet)).toBe(lowered);
    expect(getWorksheetPolicy(first.worksheet)).toBe(raised);
    expect(cache.resolve(source, 0, projection, undefined).worksheet).toBe(second.worksheet);
  });
});

import { describe, expect, it } from 'vitest';
import viewer from '../viewer.ts?raw';
import geometry from './grid-geometry.ts?raw';
import runtime from './sheet-viewer-runtime.ts?raw';
import surface from './sheet-surface.ts?raw';

describe('XLSX viewer composition boundary', () => {
  it('both public facades reach one shared engine instead of copying sheet behavior', () => {
    expect(viewer).toContain("super(container, opts, { kind: 'composite' })");
    expect(viewer.match(/new XlsxViewerEngine\(/g)).toHaveLength(1);
    expect(viewer).not.toContain("parseSheetLocally");
  });

  it('the shared engine composes each state/surface role exactly once', () => {
    for (const role of [
      'SheetAcquisition',
      'ViewportState',
      'SelectionController',
      'SheetRenderDispatcher',
      'CanvasSurface',
      'SheetOverlayHost',
      'SheetViewEdits',
      'OutlineGutter',
      'SheetTabBar',
      'ZoomControl',
      'ValidationPanel',
      'CommentPopup',
      'HyperlinkDispatcher',
      'FindAdapter',
      'CopyController',
      'SelectionInput',
      'SelectionOverlay',
      'SelectionNotifier',
      'SelectionContextReader',
      'ChromeTheme',
    ]) {
      expect(viewer.match(new RegExp(`new ${role}\\(`, 'g')), role).toHaveLength(1);
    }
    expect(viewer).toContain('getGridGeometryForWorksheet');
  });

  it('keeps geometry pure and delegates static bitmap ownership to core', () => {
    expect(geometry).not.toContain('document.');
    expect(geometry).not.toContain('HTMLCanvasElement');
    expect(surface).not.toContain('XlsxWorkbook');
    expect(runtime).toContain('StaticCanvasRenderDispatcher');
    expect(runtime).not.toContain('transferFromImageBitmap');
    expect(runtime).not.toContain("getContext('bitmaprenderer')");
  });
});

import { loadPreviousPainters, type PreviousPainters } from '../../../tests/helpers/previous-painters';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { paintDrawingMLShape, resolveFill, trackPaintPath, withInheritedPatternScope,
  withPatternCoordinateSpace, type Stroke, type DrawingMLShapeGeometry } from '@silurus/ooxml-core';
import { withPatternPointScale } from '../../core/src/shape/paint';
import { renderSlide } from '../../pptx/src/renderer';
import { renderViewport } from '../../xlsx/src/renderer';
import type { Slide } from '@silurus/ooxml-pptx';
import type { Styles, Worksheet } from '@silurus/ooxml-xlsx';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const stops = [{ position: 0, color: '112233' }, { position: 1, color: 'DDEEFF' }];
type DecorationFill = NonNullable<Stroke['fill']>;
const fills: Record<string, DecorationFill> = {
  pattern: { fillType: 'pattern', preset: 'pct50', fg: 'FF0000', bg: '0000FF' },
  solid: { fillType: 'solid', color: '5478AB' },
  linear: { fillType: 'gradient', gradType: 'linear', angle: 25, stops },
  circle: { fillType: 'gradient', gradType: 'radial', path: 'circle', angle: 0,
    fillToRect: { l: .1, r: .5, t: .4, b: .1 }, stops },
  tiledLinear: { fillType: 'gradient', gradType: 'linear', angle: 25, stops, tileRect: { r: .5 } },
  tiledCircle: { fillType: 'gradient', gradType: 'radial', path: 'circle', angle: 0,
    fillToRect: { l: .1, r: .5, t: .4, b: .1 }, stops, tileRect: { r: .5 } },
};

// Compare every decoded pixel with the pinned production renderer running on
// this host. This protects paint frames and format wiring without treating one
// platform's native Skia output as a portable Office-fidelity reference.
// Connector tests include a filled triangle; custom paths have an open arrow.
// XLSX has no line decorations, so its integration covers the rightArrow body.
const routes = ['core-connector', 'core-custom', 'core-body', 'pptx-connector', 'pptx-custom', 'xlsx-body'];

async function renderCase(route: string, fill: DecorationFill, painters: Pick<PreviousPainters, 'paintDrawingMLShape' | 'renderSlide' | 'renderViewport'>): Promise<Uint8ClampedArray> {
  const c = new (skia as NonNullable<typeof skia>).Canvas(400, 400);
  const ctx = c.getContext('2d') as unknown as CanvasRenderingContext2D;
  const body = route.endsWith('body');
  const custom = route.endsWith('custom');
  const rect = body ? { x: 130, y: 140, w: 150, h: 100 } : { x: 200, y: 200, w: 50, h: 1 };
  const geometry: DrawingMLShapeGeometry = custom ? { kind: 'custom', subpaths: [[
    { cmd: 'moveTo', x: 0, y: 0 }, { cmd: 'lineTo', x: 1, y: 1 },
  ]] } : { kind: 'preset', name: body ? 'rightArrow' : 'line', adjustments: [] };
  const stroke = { color: '000000', width: 20, fill,
    ...(body ? {} : { tailEnd: { type: custom ? 'arrow' as const : 'triangle' as const, w: 'lg' as const, len: 'lg' as const } }),
  };
  if (route.startsWith('core')) {
    // Retained DOCX drawing commands enter this same shared production painter.
    painters.paintDrawingMLShape(ctx, { rect, geometry, fill: body ? fill : null, stroke,
      transform: { rotationDeg: 30, flipH: false, flipV: false } }, 1);
  } else if (route.startsWith('pptx')) {
    await painters.renderSlide(c as unknown as HTMLCanvasElement, {
      index: 0, slideNumber: 1, background: null, elements: [{ type: 'shape',
        x: rect.x * 9525, y: rect.y * 9525, width: rect.w * 9525, height: rect.h * 9525,
        rotation: 30, flipH: false, flipV: false, geometry: 'line', fill: null,
        stroke: { ...stroke, width: stroke.width * 9525 }, textBody: null,
        custGeom: geometry.kind === 'custom' ? geometry.subpaths : null, shadow: null,
      }],
    } as Slide, 400 * 9525, 400 * 9525, { width: 400, dpr: 1 });
  } else {
    painters.renderViewport(ctx, {
      name: 'Sheet1', isChartSheet: true, rows: [], colWidths: {}, rowHeights: {},
      freezeRows: 0, freezeCols: 0, defaultColWidth: 8.43, defaultRowHeight: 15,
      mergeCells: [], conditionalFormats: [], images: [], charts: [],
      defaultFontFamily: 'Calibri', defaultFontSize: 11,
      shapeGroups: [{ fromCol: 0, fromRow: 0, fromColOff: rect.x * 9525, fromRowOff: rect.y * 9525,
        toCol: 1, toRow: 1, toColOff: 0, toRowOff: 0, editAs: 'oneCell',
        nativeExtCx: rect.w * 9525, nativeExtCy: rect.h * 9525,
        shapes: [{ x: 0, y: 0, w: 1, h: 1, rot: 30, strokeWidth: 20,
          strokeColor: '000000', strokeFill: fill, fill,
          geom: { type: 'preset', name: 'rightArrow', adj: [] } }],
      }],
    } as Worksheet, { fonts: [], fills: [], borders: [], cellXfs: [], numFmts: [], dxfs: [] } as Styles,
    { row: 1, col: 1, rows: 1, cols: 1 });
  }
  return ctx.getImageData(0, 0, 400, 400).data;
}

describe.skipIf(!skia)('non-path-gradient paints retain main frames on rotated decorations and shapes', () => {
  let previous: PreviousPainters;
  beforeAll(async () => {
    vi.stubGlobal('OffscreenCanvas', (skia as NonNullable<typeof skia>).Canvas);
    previous = await loadPreviousPainters();
  }, 60000);
  afterAll(() => vi.unstubAllGlobals());
  it.each(routes)('%s retains pattern/solid/linear/circle and tiled native pixels', async route => {
    const current = { paintDrawingMLShape, renderSlide, renderViewport };
    for (const [name, fill] of Object.entries(fills)) {
      expect(await renderCase(route, fill, current), `${route}/${name}`)
        .toEqual(await renderCase(route, fill, previous));
    }
  }, 30000);

  it('preserves ordinary pattern scope through path tracking and effect-canvas inheritance', () => {
    const { Canvas } = skia as NonNullable<typeof skia>;
    const source = new Canvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D;
    const paint = (tracked: boolean) => {
      const target = new Canvas(100, 100).getContext('2d') as unknown as CanvasRenderingContext2D;
      target.translate(30, 20);
      target.rotate(Math.PI / 6);
      withPatternPointScale(source, 2, () => withPatternCoordinateSpace(source,
        { a: 1.2, b: .2, c: -.3, d: .8, e: 4, f: 7 }, () => {
          withInheritedPatternScope(tracked ? trackPaintPath(source) : source, target, () => {
            target.fillStyle = resolveFill(fills.pattern, target, 0, 0, 50, 50) as CanvasPattern;
            target.fillRect(0, 0, 50, 50);
          }, { x: 3, y: 5 });
        }));
      return target.getImageData(0, 0, 100, 100).data;
    };
    expect(paint(true)).toEqual(paint(false));
  });
});

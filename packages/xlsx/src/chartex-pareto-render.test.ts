import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { renderChartExChart } from '@silurus/ooxml-core/internal/chart-ex-renderer';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

const EMU_PER_PX = 9525;

const STYLES: Styles = {
  fonts: [{ bold: false, italic: false, underline: false, strike: false, size: 11, color: null, name: null }],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

function context(width: number, height: number): {
  ctx: CanvasRenderingContext2D;
  fills: Array<{ x: number; w: number; color: string }>;
  texts: Array<{ text: string; x: number }>;
} {
  const fills: Array<{ x: number; w: number; color: string }> = [];
  const texts: Array<{ text: string; x: number }> = [];
  const state: Record<string, unknown> = {
    canvas: { width, height },
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    font: '11px sans-serif',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    globalAlpha: 1,
    measureText: (text: string) => ({ width: [...text].length * 7 }),
    fillRect(x: number, _y: number, w: number) {
      fills.push({ x, w, color: String(state.fillStyle).toUpperCase() });
    },
    fillText(text: string, x: number) { texts.push({ text, x }); },
    createLinearGradient: () => ({ addColorStop: () => {} }),
  };
  const noop = () => {};
  const ctx = new Proxy(state, {
    get(target, property) { return property in target ? target[property as string] : noop; },
    set(target, property, value) { target[property as string] = value; return true; },
  });
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills, texts };
}

describe('XLSX ChartEx Pareto rendering', () => {
  it('keeps a sparse source-order Pareto category at its cached point index', () => {
    // Excel projection of a non-aggregated clusteredColumn owner whose value
    // cache omits idx=1: owner bars keep authored order (no descending sort).
    const chart: ChartModel = {
      chartType: 'pareto',
      title: null,
      categories: ['A', 'B', 'C'],
      series: [
        {
          name: 'Owner', color: null, values: [5, null, 3],
          dataPointOverrides: [0, 1, 2].map(idx => ({ idx, color: 'AA0000' })),
        },
        { name: 'Line', color: null, values: [], seriesType: 'line', useSecondaryAxis: true },
      ],
      chartexParetoOwnerIndex: 0,
      chartexParetoSortDescending: false,
      chartexParetoFlatEndpoint: false,
      showDataLabels: false,
      valMin: null,
      valMax: null,
      catAxisTitle: null,
      valAxisTitle: null,
      catAxisHidden: false,
      valAxisHidden: false,
      catAxisLineHidden: false,
      valAxisLineHidden: false,
      plotAreaBg: null,
      chartBg: null,
      showLegend: false,
      legendPos: null,
      catAxisCrossBetween: 'between',
      valAxisMajorTickMark: 'out',
      catAxisMajorTickMark: 'out',
      titleFontSizeHpt: null,
      titleFontColor: null,
      titleFontFace: null,
      catAxisFontSizeHpt: null,
      valAxisFontSizeHpt: null,
      dataLabelFontSizeHpt: null,
      subtotalIndices: [],
    };
    const worksheet = {
      name: 'Chart1',
      isChartSheet: true,
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
      charts: [{
        fromCol: 0, fromColOff: 0, fromRow: 0, fromRowOff: 0,
        toCol: 0, toColOff: 600 * EMU_PER_PX, toRow: 0, toRowOff: 360 * EMU_PER_PX,
        chart,
      }],
      defaultFontFamily: 'Calibri',
      defaultFontSize: 11,
    } as Worksheet;
    const recording = context(640, 400);

    renderViewport(
      recording.ctx, worksheet, STYLES, { row: 1, col: 1, rows: 40, cols: 40 },
      { chartEx: { render: renderChartExChart } },
    );

    const labels = recording.texts
      .filter(text => ['A', 'B', 'C'].includes(text.text))
      .sort((a, b) => a.x - b.x);
    expect(labels.map(label => label.text)).toEqual(['A', 'B', 'C']);
    const nearestLabel = (center: number) => labels.reduce((best, label) =>
      Math.abs(label.x - center) < Math.abs(best.x - center) ? label : best).text;
    const bars = recording.fills
      .filter(fill => fill.color === '#AA0000')
      .sort((a, b) => a.x - b.x)
      .map(fill => nearestLabel(fill.x + fill.w / 2));
    expect(bars).toEqual(['A', 'C']);
  });
});

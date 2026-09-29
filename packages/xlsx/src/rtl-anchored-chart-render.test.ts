import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { renderViewport } from './renderer.js';
import type { Styles, Worksheet } from './types.js';

type FillRectCall = { x: number; y: number; w: number; h: number; color: string };

function recordingContext(width = 800, height = 400): {
  ctx: CanvasRenderingContext2D;
  fills: FillRectCall[];
  scales: Array<[number, number]>;
  translations: Array<[number, number]>;
} {
  const fills: FillRectCall[] = [];
  const scales: Array<[number, number]> = [];
  const translations: Array<[number, number]> = [];
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
    fillRect(x: number, y: number, w: number, h: number) {
      fills.push({ x, y, w, h, color: String(state.fillStyle) });
    },
    scale(x: number, y: number) {
      scales.push([x, y]);
    },
    translate(x: number, y: number) {
      translations.push([x, y]);
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
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills, scales, translations };
}

const STYLES: Styles = {
  fonts: [
    {
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      size: 11,
      color: null,
      name: null,
    },
  ],
  fills: [],
  borders: [],
  cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 0 } as Styles['cellXfs'][number]],
  numFmts: [],
  dxfs: [],
};

const CHART: ChartModel = {
  chartType: 'clusteredBar',
  title: null,
  categories: [],
  series: [],
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
  chartBg: 'ABCDEF',
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

function worksheet(rightToLeft: boolean): Worksheet {
  return {
    name: 'Sheet1',
    rows: [],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 1,
    conditionalFormats: [],
    images: [],
    rightToLeft,
    charts: [
      {
        fromCol: 6,
        fromColOff: 0,
        fromRow: 1,
        fromRowOff: 0,
        toCol: 10,
        toColOff: 0,
        toRow: 10,
        toRowOff: 0,
        chart: CHART,
      },
    ],
    defaultFontFamily: 'Calibri',
    defaultFontSize: 11,
  } as Worksheet;
}

function chartBackground(rightToLeft: boolean) {
  const recording = recordingContext();
  renderViewport(
    recording.ctx,
    worksheet(rightToLeft),
    STYLES,
    { row: 1, col: 2, rows: 20, cols: 20 },
    { freezeCols: 1, scrollOffsetX: 17 },
  );
  const fills = recording.fills.filter((call) => call.color === '#ABCDEF');
  expect(fills).toHaveLength(1);
  return { background: fills[0], scales: recording.scales, translations: recording.translations };
}

describe('RTL anchored chart rendering', () => {
  it('mirrors the chart rectangle with scrolling and frozen columns', () => {
    const ltr = chartBackground(false);
    const rtl = chartBackground(true);

    expect(rtl.background.w).toBe(ltr.background.w);
    expect(rtl.background.h).toBe(ltr.background.h);
    // The chart keeps its viewport rectangle; pattern phase is scoped
    // separately so non-pattern chart geometry remains pixel-identical.
    expect(rtl.background.x).toBeCloseTo(800 - ltr.background.x - ltr.background.w);
  });

  it('does not flip the chart contents', () => {
    const rtl = chartBackground(true);
    expect(rtl.scales.some(([x, y]) => x < 0 || y < 0)).toBe(false);
  });
});

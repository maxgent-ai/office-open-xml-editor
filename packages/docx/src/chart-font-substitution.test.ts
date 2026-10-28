import { describe, expect, it } from 'vitest';
import type { ChartModel } from '@silurus/ooxml-core';
import { createCanonicalCanvasPaintResourceHandlers } from './paint/canonical-resource-handlers.js';

// DOCX charts paint the unscoped substitutions the document's font sources
// loaded (ChartModel.fontSubstitutions), as PPTX and XLSX charts do.

const CHART = {
  chartType: 'clusteredBar', title: 'Chart Title', categories: ['A', 'B'],
  series: [{ name: 'Series', color: null, values: [1, 2] }],
  showDataLabels: false, valMin: null, valMax: null, catAxisTitle: null, valAxisTitle: null,
  catAxisHidden: false, valAxisHidden: false, catAxisLineHidden: false, valAxisLineHidden: false,
  plotAreaBg: null, chartBg: null, showLegend: false, legendPos: null, catAxisCrossBetween: 'between',
  valAxisMajorTickMark: 'out', catAxisMajorTickMark: 'out', titleFontSizeHpt: null, titleFontColor: null,
  titleFontFace: 'Chart Face', catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
  dataLabelFontSizeHpt: null, subtotalIndices: [],
} as unknown as ChartModel;

function titleFont(substitutions?: Readonly<Record<string, string>>): string {
  const calls: Array<{ text: string; font: string }> = [];
  let font = '10px sans-serif';
  const noop = () => {};
  const ctx = new Proxy({
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => ({ width: text.length * 7, actualBoundingBoxAscent: 8, actualBoundingBoxDescent: 2, fontBoundingBoxAscent: 8, fontBoundingBoxDescent: 2 }),
    fillText: (text: string) => { calls.push({ text, font }); },
    createLinearGradient: () => ({ addColorStop: noop }),
    createRadialGradient: () => ({ addColorStop: noop }),
    createPattern: () => null,
    getTransform: () => ({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0, isIdentity: true }),
    getLineDash: () => [],
  } as Record<string | symbol, unknown>, {
    get(target, key) { return key in target ? target[key] : noop; },
    set(target, key, value) { target[key] = value; return true; },
  });
  const handlers = createCanonicalCanvasPaintResourceHandlers(undefined, undefined, undefined, undefined, substitutions);
  handlers.chart({ descriptor: { model: CHART } } as never, { xPt: 0, yPt: 0, widthPt: 400, heightPt: 300 } as never, ctx as never);
  return calls.find((call) => call.text === 'Chart Title')?.font ?? '';
}

describe('DOCX chart text and font sources', () => {
  it('paints the loaded unscoped substitution after the authored face', () => {
    expect(titleFont()).toContain('"Chart Face"');
    expect(titleFont()).not.toContain('Served Face');
    const substitutions = Object.freeze(Object.assign(Object.create(null) as Record<string, string>, { 'chart face': 'Served Face' }));
    expect(titleFont(substitutions)).toContain('"Chart Face", "Served Face"');
  });
});

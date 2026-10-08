import { describe, expect, it } from 'vitest';
import { renderChart, type ChartModel, type GradientFill } from '@silurus/ooxml-core';
import { loadSkiaForTests } from './test-imports';

const skia = await loadSkiaForTests();
const fill: GradientFill = {
  fillType: 'gradient', gradType: 'radial', path: 'rect', angle: 0,
  fillToRect: { l: .5, t: .5, r: .5, b: .5 },
  stops: [{ position: 0, color: 'FF0000' }, { position: .5, color: '00FF00' },
    { position: 1, color: '0000FF' }],
};

function chart(chartType: ChartModel['chartType']): ChartModel {
  return {
    chartType, title: null, categories: ['A', 'B'],
    series: [{ name: 'Series', color: 'FFFFFF', values: [4, 7],
      chartexStyle: { linePaints: [fill], linePaintAuthored: true }, lineWidthEmu: 8 * 12700 }],
    showDataLabels: false, valMin: 0, valMax: 10,
    catAxisTitle: null, valAxisTitle: null,
    catAxisHidden: false, valAxisHidden: false,
    catAxisLineHidden: false, valAxisLineHidden: false,
    catAxisCrossBetween: 'between',
    valAxisMajorTickMark: 'out', catAxisMajorTickMark: 'out',
    plotAreaBg: null, chartBg: null, showLegend: true, legendPos: 'b',
    titleFontSizeHpt: null, titleFontColor: null, titleFontFace: null,
    catAxisFontSizeHpt: null, valAxisFontSizeHpt: null,
    dataLabelFontSizeHpt: null, subtotalIndices: [],
  };
}

describe.skipIf(!skia)('chart path gradients preserve native paint', () => {
  it.each(['clusteredBar', 'clusteredBarH'] as const)('%s retains native series and legend outline pixels', chartType => {
    const Canvas = (skia as NonNullable<typeof skia>).Canvas;
    const actual = new Canvas(400, 300).getContext('2d') as unknown as CanvasRenderingContext2D;
    const expected = new Canvas(400, 300).getContext('2d') as unknown as CanvasRenderingContext2D;
    // Allocation-unavailable native fallback is the previous-renderer oracle.
    // Every Canvas drawing call still reaches the real backend; only auxiliary
    // allocation is unavailable. This independently exercises the complete
    // chart pipeline and avoids duplicating its layout or paint implementation.
    const nativeHost = new Proxy(expected, {
      get(target, property) {
        if (property === 'canvas') return { constructor: undefined };
        const value = Reflect.get(target, property);
        return typeof value === 'function' ? value.bind(target) : value;
      },
      set(target, property, value) { return Reflect.set(target, property, value); },
    });
    renderChart(actual, chart(chartType), { x: 0, y: 0, w: 400, h: 300 });
    renderChart(nativeHost, chart(chartType), { x: 0, y: 0, w: 400, h: 300 });
    const pixels = actual.getImageData(0, 0, 400, 300).data;
    expect(Buffer.from(pixels).equals(Buffer.from(expected.getImageData(0, 0, 400, 300).data))).toBe(true);
    // Ensure the controls exercise visible gradient outlines, not noFill paths.
    let colored = 0;
    for (let i = 0; i < pixels.length; i += 4) {
      if (Math.max(pixels[i], pixels[i + 1], pixels[i + 2])
          - Math.min(pixels[i], pixels[i + 1], pixels[i + 2]) > 30) colored++;
    }
    expect(colored).toBeGreaterThan(1000);
  });
});

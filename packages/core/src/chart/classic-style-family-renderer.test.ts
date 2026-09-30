import { describe, expect, it } from 'vitest';
import type { ChartModel, ChartSeries } from '../types/chart.js';
import { classicMarkerPaintWorkCount, renderChart } from './renderer.js';
import { classicDataMarkPaintWorkCount } from './classic-paint-work.js';
import { collectChartMarkerImageFills } from './image-fill.js';
import { withEffectiveChartStyleRoles } from './effective-style.js';
import { renderSimpleThreeDChart } from './three-d-renderer.js';
import type { ChartThreeDRenderer } from './three-d-contract.js';

const RECT = { x: 0, y: 0, w: 480, h: 300 };

function series(over: Partial<ChartSeries> = {}): ChartSeries {
  return { name: '', color: null, values: [1, 2, 3], ...over };
}

function model(over: Partial<ChartModel>): ChartModel {
  const base: ChartModel = {
    chartType: 'line',
    title: null,
    categories: ['A', 'B', 'C'],
    series: [series()],
    showDataLabels: false,
    valMin: null,
    valMax: null,
    catAxisTitle: null,
    valAxisTitle: null,
    showLegend: false,
    legendPos: null,
    catAxisHidden: true,
    valAxisHidden: true,
    catAxisLineHidden: true,
    valAxisLineHidden: true,
    valAxisMajorGridlines: false,
    plotAreaBg: null,
    chartBg: null,
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
  return { ...base, ...over };
}

function recordingContext(): {
  ctx: CanvasRenderingContext2D;
  fills: string[];
  rectFills: string[];
  strokes: string[];
  strokeWidths: number[];
  texts: string[];
} {
  const fills: string[] = [];
  const rectFills: string[] = [];
  const strokes: string[] = [];
  const strokeWidths: number[] = [];
  const texts: string[] = [];
  const state: Record<string, unknown> = {
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    lineCap: 'butt', lineJoin: 'miter', font: '10px sans-serif',
    textAlign: 'start', textBaseline: 'alphabetic', globalAlpha: 1,
    globalCompositeOperation: 'source-over',
  };
  const stack: Array<Record<string, unknown>> = [];
  const ctx = new Proxy(state, {
    get(target, prop: string) {
      if (prop in target) return target[prop];
      if (prop === 'save') return () => stack.push({ ...target });
      if (prop === 'restore') return () => Object.assign(target, stack.pop() ?? {});
      if (prop === 'fill') return () => fills.push(String(target.fillStyle));
      if (prop === 'fillRect') return () => rectFills.push(String(target.fillStyle));
      if (prop === 'stroke' || prop === 'strokeRect') {
        return () => {
          strokes.push(String(target.strokeStyle));
          strokeWidths.push(Number(target.lineWidth));
        };
      }
      if (prop === 'fillText') return (value: string) => texts.push(value);
      if (prop === 'measureText') return (text: string) => ({ width: text.length * 6 });
      if (prop === 'getLineDash') return () => [];
      if (prop === 'setLineDash') return () => {};
      if (prop === 'createLinearGradient' || prop === 'createRadialGradient') {
        return () => ({ addColorStop() {} });
      }
      return () => {};
    },
    set(target, prop: string, value) {
      target[prop] = value;
      return true;
    },
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, fills, rectFills, strokes, strokeWidths, texts };
}

describe('classic chart style family wiring', () => {
  describe('linked Chart Style data roles never paint classic series', () => {
    // Observed in Office (PowerPoint for the families below, plus Excel/Word
    // for clustered column and standard line): changing the linked dataPoint,
    // dataPoint3D, dataPointLine, dataPointMarker and hiLoLine paint leaves
    // classic series paint identical. upBar, downBar and dataPointWireframe
    // were not measured; the rule is extended to them without direct evidence.
    // Each case renders the same chart with and without a distinctive linked
    // role and requires identical paint, image preflight and paint-work.
    const DATA_ROLES = [
      'dataPoint', 'dataPoint3D', 'dataPointLine', 'dataPointMarker',
      'dataPointWireframe', 'upBar', 'downBar', 'hiLoLine',
    ] as const;
    const numericRole = {
      fillColors: ['156082'], fillPaintAuthored: true,
      lineColors: ['156082'], linePaintAuthored: true, lineWidthEmu: 28575,
    };
    const picture = {
      fillType: 'image' as const, imagePath: 'xl/media/linked.png',
      mimeType: 'image/png', stretch: true,
    };
    // Office-like linked role: an `effectRef idx=0` residue (effectNoStyle),
    // a picture recipe and no allow-no-paint modifiers.
    const linkedRole = {
      fillColors: ['00B050'], fillPaintAuthored: true, fillPaints: [picture],
      lineColors: ['00B050'], linePaintAuthored: true, lineWidthEmu: 76200,
      effectNoStyle: true, allowNoFillOverride: false, allowNoLineOverride: false,
    };
    const roleTable = (role: typeof numericRole) => Object.fromEntries(
      DATA_ROLES.map(name => [name, role]),
    );
    const GREEN = /#00B050|rgba\(0,\s*176,\s*80/i;
    const paint = (chart: ChartModel, threeD?: ChartThreeDRenderer) => {
      const rec = recordingContext();
      renderChart(rec.ctx, chart, RECT, 1, 0, threeD);
      return {
        fills: rec.fills, rectFills: rec.rectFills,
        strokes: rec.strokes, strokeWidths: rec.strokeWidths,
      };
    };
    function expectInert(chart: ChartModel, threeD?: ChartThreeDRenderer): void {
      const numeric = { ...chart, classicChartStyleRoles: roleTable(numericRole) };
      const baseline = paint(numeric, threeD);
      const linked = paint({ ...numeric, chartStyleRoles: roleTable(linkedRole) }, threeD);
      expect(JSON.stringify(baseline)).not.toMatch(GREEN);
      expect(linked).toEqual(baseline);
      expect(JSON.stringify(linked)).not.toMatch(GREEN);
      // Guard against a vacuous comparison: something was actually painted.
      expect(baseline.fills.length + baseline.rectFills.length + baseline.strokes.length)
        .toBeGreaterThan(0);
      // Image preflight and paint-work budgets follow the same role table.
      const withLinked = { ...numeric, chartStyleRoles: roleTable(linkedRole) };
      expect(collectChartMarkerImageFills(withLinked))
        .toEqual(collectChartMarkerImageFills(numeric));
      const bitmap = { width: 8, height: 8 } as unknown as CanvasImageSource;
      const work = (c: ChartModel) => {
        const prepared = withEffectiveChartStyleRoles(c);
        return [
          classicMarkerPaintWorkCount(prepared, () => bitmap, 1, RECT),
          classicDataMarkPaintWorkCount(prepared, () => bitmap, 1, RECT, threeD != null),
        ];
      };
      expect(work(withLinked)).toEqual(work(numeric));
    }
    const group = (
      kind: NonNullable<ChartModel['plotGroups']>[number]['kind'],
      seriesCount: number,
      over: Partial<NonNullable<ChartModel['plotGroups']>[number]> = {},
    ): NonNullable<ChartModel['plotGroups']>[number] => ({
      kind, seriesStart: 0, seriesCount,
      categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
      varyColors: false, ...over,
    });
    const two = [series({ name: 'A' }), series({ name: 'B', values: [3, 2, 1] })];

    it('bar and column: clustered column, stacked column and horizontal bar', () => {
      expectInert(model({
        chartType: 'clusteredBar', series: [series({ color: '156082' })],
        plotGroups: [group('bar', 1, { barDirection: 'col', grouping: 'clustered' })],
      }));
      expectInert(model({
        chartType: 'stackedBar', series: two,
        plotGroups: [group('bar', 2, { barDirection: 'col', grouping: 'stacked' })],
      }));
      expectInert(model({
        chartType: 'clusteredBarH', series: two,
        plotGroups: [group('bar', 2, { barDirection: 'bar', grouping: 'clustered' })],
      }));
    });

    it('line and area: standard line with markers, area and a column+line combination', () => {
      expectInert(model({
        series: [series({ color: '156082', showMarker: true, markerSymbol: 'circle' })],
        plotGroups: [group('line', 1, { grouping: 'standard' })],
      }));
      expectInert(model({
        chartType: 'area', series: two, plotGroups: [group('area', 2, { grouping: 'standard' })],
      }));
      expectInert(model({
        chartType: 'clusteredBar',
        series: [series({ seriesType: 'bar' }), series({ seriesType: 'line', values: [3, 2, 1] })],
        plotGroups: [
          group('bar', 1, { barDirection: 'col', grouping: 'clustered' }),
          group('line', 1, { seriesStart: 1, grouping: 'standard' }),
        ],
      }));
    });

    it('pie and doughnut: point-varying slices and rings', () => {
      for (const chartType of ['pie', 'doughnut'] as const) {
        expectInert(model({
          chartType, varyColors: true, series: [series({ values: [1, 2, 3] })],
          plotGroups: [group(chartType, 1, { varyColors: true })],
        }));
      }
    });

    it('scatter and bubble: markers and connecting lines', () => {
      expectInert(model({
        chartType: 'scatter', scatterStyle: 'lineMarker',
        series: [series({ categories: ['1', '2', '3'], showMarker: true, markerSymbol: 'circle' })],
        plotGroups: [group('scatter', 1, { scatterStyle: 'lineMarker' })],
      }));
      expectInert(model({
        chartType: 'bubble',
        series: [series({ categories: ['1', '2', '3'], bubbleSizes: [1, 2, 3] })],
        plotGroups: [group('bubble', 1)],
      }));
    });

    it('radar: marker/line and filled polygons', () => {
      expectInert(model({
        chartType: 'radar', series: [series({ showMarker: true, markerSymbol: 'circle' })],
        plotGroups: [group('radar', 1, { radarStyle: 'marker' })],
      }));
      expectInert(model({
        chartType: 'radar', radarStyle: 'filled', series: [series()],
        plotGroups: [group('radar', 1, { radarStyle: 'filled' })],
      }));
    });

    it('stock: series lines, high-low lines and up/down bars', () => {
      const ohlc = [
        series({ name: 'Open', values: [20, 45, 25] }),
        series({ name: 'High', values: [55, 57, 57] }),
        series({ name: 'Low', values: [11, 12, 13] }),
        series({ name: 'Close', values: [40, 30, 25] }),
      ];
      expectInert(model({
        chartType: 'stock', stockHiLowLines: true, stockUpDownBars: true,
        stockUpDownBarStyle: { gapWidthPercent: 100, up: {}, down: {} },
        series: ohlc, plotGroups: [group('stock', 4)],
      }));
    });

    it('surface: filled value bands and the wireframe mesh', () => {
      for (const surfaceWireframe of [false, true]) {
        expectInert(model({
          chartType: 'surface', surfaceWireframe, valMin: 0, valMax: 10, valAxisMajorUnit: 5,
          series: [series({ values: [1, 9, 2] }), series({ values: [9, 1, 8] })],
          plotGroups: [group('surface', 2)],
          threeD: { rotationX: 30, rotationY: 20, perspective: 30 },
        }));
      }
    });

    it('3-D: bar and area datum paint', () => {
      const threeD = { rotationX: 15, rotationY: 20, perspective: 30 };
      expectInert(model({
        chartType: 'clusteredBar', series: two, threeD,
        plotGroups: [group('bar3D', 2, { barDirection: 'col', grouping: 'clustered' })],
      }), { render: renderSimpleThreeDChart });
      expectInert(model({
        chartType: 'area', series: two, threeD,
        plotGroups: [group('area3D', 2, { grouping: 'standard' })],
      }), { render: renderSimpleThreeDChart });
    });

    it('keeps a direct noFill authoritative over numeric paint despite an Office-like linked role', () => {
      const direct = { fillHidden: true, fillPaintAuthored: true, lineHidden: true, linePaintAuthored: true };
      const bar = model({
        chartType: 'clusteredBar',
        series: [series({ chartexStyle: direct })],
        plotGroups: [group('bar', 1, { barDirection: 'col', grouping: 'clustered' })],
      });
      const threeD = { rotationX: 15, rotationY: 20, perspective: 30 };
      const bar3D = model({
        chartType: 'clusteredBar', threeD,
        series: [series({ chartexStyle: direct })],
        plotGroups: [group('bar3D', 1, { barDirection: 'col', grouping: 'clustered' })],
      });
      for (const [chart, renderer] of [
        [bar, undefined], [bar3D, { render: renderSimpleThreeDChart }],
      ] as const) {
        const numeric = { ...chart, classicChartStyleRoles: roleTable(numericRole) };
        const baseline = paint(numeric, renderer);
        const linked = paint({ ...numeric, chartStyleRoles: roleTable(linkedRole) }, renderer);
        expect(linked).toEqual(baseline);
        expect(JSON.stringify(linked)).not.toMatch(GREEN);
        expect(JSON.stringify(linked)).not.toMatch(/#156082|rgba\(21,\s*96,\s*130/i);
      }
    });

    it('classifies a hand-built classic model by chart type, not by optional layers', () => {
      // No numeric role table in either case. The first keeps plot groups and
      // also carries the linked role in the legacy `chartex*Style` aliases;
      // the second has neither plot groups nor aliases.
      const withGroups = model({
        chartType: 'clusteredBar', series: [series({ color: '156082' })],
        plotGroups: [group('bar', 1, { barDirection: 'col', grouping: 'clustered' })],
      });
      const bare = model({ chartType: 'clusteredBar', series: [series({ color: '156082' })] });
      const aliases = {
        chartexDataPointStyle: linkedRole,
        chartexDataPointLineStyle: linkedRole,
        chartexDataPointMarkerStyle: linkedRole,
      };
      for (const [chart, extra] of [[withGroups, aliases], [bare, {}]] as const) {
        const baseline = paint(chart);
        const withLinked = { ...chart, chartStyleRoles: roleTable(linkedRole), ...extra };
        expect(paint(withLinked)).toEqual(baseline);
        expect(JSON.stringify(paint(withLinked))).not.toMatch(GREEN);
        expect(collectChartMarkerImageFills(withLinked)).toEqual([]);
      }
    });

    it('keeps unmeasured linked roles, such as series lines, on the linked cascade', () => {
      const rec = recordingContext();
      renderChart(rec.ctx, model({
        chartType: 'ofPie',
        series: [series({ values: [40, 30, 20, 10] })],
        ofPie: {
          type: 'bar', splitType: 'pos', splitPos: 2,
          secondPieSizePercent: 75, gapWidthPercent: 100, seriesLines: true,
        },
        classicChartStyleRoles: roleTable(numericRole),
        chartStyleRoles: { seriesLine: { lineColors: ['778899'] } },
      }), RECT, 1);
      expect(rec.strokes).toContain('#778899');
    });
  });

  it('indexes a lone varyColors bar by point and honors direct noFill', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'clusteredBar',
      varyColors: true,
      series: [series({
        chartexFormatIdx: 8,
        dataPointOverrides: [{ idx: 1, fillHidden: true }],
      })],
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['888888', 'AA0000', '00AA00', '0000AA'],
          fillFormattingIndices: [8, 0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.rectFills).toContain('#AA0000');
    expect(rec.rectFills).toContain('#0000AA');
    expect(rec.rectFills).not.toContain('#00AA00');
  });

  it('uses a sparse source series index for a lone line', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      series: [series({ chartexFormatIdx: 8, showMarker: false })],
      classicChartStyleRoles: {
        dataPointLine: {
          lineColors: ['8899AA', '001100', '002200', '003300'],
          lineFormattingIndices: [8, 0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.strokes).toContain('#8899AA');
    expect(rec.strokes).not.toContain('#001100');
  });

  it('keeps a lone varyColors bar legend on the point-index palette', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'clusteredBar',
      varyColors: true,
      showLegend: true,
      legendPos: 'r',
      series: [series({ chartexFormatIdx: 8 })],
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['888888', 'AA0000', '00AA00', '0000AA'],
          fillFormattingIndices: [8, 0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.rectFills.filter(color => color === '#AA0000')).toHaveLength(2);
    expect(rec.rectFills.filter(color => color === '#00AA00')).toHaveLength(2);
    expect(rec.rectFills.filter(color => color === '#0000AA')).toHaveLength(2);
    expect(rec.rectFills).not.toContain('#888888');
  });

  it('replays a point-index style palette in every multi-series doughnut ring', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'doughnut',
      varyColors: true,
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Outer', values: [1, 1, 1], chartexFormatIdx: 8 }),
        series({ name: 'Inner', values: [1, 1, 1], chartexFormatIdx: 9 }),
      ],
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['AA0000', '00AA00', '0000AA'],
          fillFormattingIndices: [0, 1, 2],
        },
      },
    }), RECT, 1);

    expect(rec.fills.filter(color => color === '#AA0000')).toHaveLength(2);
    expect(rec.fills.filter(color => color === '#00AA00')).toHaveLength(2);
    expect(rec.fills.filter(color => color === '#0000AA')).toHaveLength(2);
  });

  it('uses a series-index style palette and series legend for varyColors=false doughnuts', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'doughnut',
      varyColors: false,
      showLegend: true,
      legendPos: 'r',
      categories: ['A', 'B', 'C'],
      series: [
        series({ name: 'Outer', values: [1, 1, 1], chartexFormatIdx: 8 }),
        series({ name: 'Inner', values: [1, 1, 1], chartexFormatIdx: 9 }),
      ],
      classicChartStyleRoles: {
        dataPoint: {
          fillColors: ['AA0000', '00AA00'],
          fillFormattingIndices: [8, 9],
        },
      },
    }), RECT, 1);

    expect(rec.fills.filter(color => color === '#AA0000')).toHaveLength(3);
    expect(rec.fills.filter(color => color === '#00AA00')).toHaveLength(3);
    expect(rec.texts).toContain('Outer');
    expect(rec.texts).toContain('Inner');
    expect(rec.texts).not.toContain('A');
  });

  it.each([
    ['line', {}],
    ['scatter', { scatterStyle: 'line' }],
    ['radar', {}],
    ['stock', { stockHiLowLines: false }],
  ] as const)('uses dataPointLine for the %s family', (chartType, extra) => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType,
      ...extra,
      series: chartType === 'stock'
        ? [series(), series({ values: [3, 4, 5] }), series({ values: [2, 3, 4] })]
        : [series({ showMarker: false })],
      classicChartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);

    expect(rec.strokes).toContain('#A1B2C3');
  });

  it('keeps direct line paint above the numeric style and lets direct noFill remove it', () => {
    const direct = recordingContext();
    renderChart(direct.ctx, model({
      series: [series({ lineColor: 'CC3300', showMarker: false })],
      classicChartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    expect(direct.strokes).toContain('#CC3300');
    expect(direct.strokes).not.toContain('#A1B2C3');

    const hidden = recordingContext();
    renderChart(hidden.ctx, model({
      series: [series({ chartexStyle: { lineHidden: true }, showMarker: false })],
      classicChartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    expect(hidden.strokes).not.toContain('#A1B2C3');
  });

  it('does not let numeric style paint revive classic series or point lines removed by direct spPr', () => {
    const markerOnly = recordingContext();
    renderChart(markerOnly.ctx, model({
      chartType: 'scatter',
      scatterStyle: 'lineMarker',
      series: [series({
        categories: ['1', '2', '3'],
        lineHidden: true,
        showMarker: true,
      })],
      classicChartStyleRoles: { dataPointLine: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    expect(markerOnly.strokes).not.toContain('#A1B2C3');

    const borderlessSlice = recordingContext();
    renderChart(borderlessSlice.ctx, model({
      chartType: 'pie',
      varyColors: true,
      series: [series({
        values: [1, 2, 3],
        dataPointOverrides: [{ idx: 1, lineHidden: true }],
      })],
      classicChartStyleRoles: { dataPoint: { lineColors: ['A1B2C3'] } },
    }), RECT, 1);
    // The style still outlines the other two slices, but the direct no-line
    // point must remove one of the three linked outlines.
    expect(borderlessSlice.strokes.filter(color => color === '#A1B2C3')).toHaveLength(2);
  });

  it.each(['area', 'bar combo'] as const)(
    'uses dataPoint fill and outline for %s area geometry',
    family => {
      const rec = recordingContext();
      renderChart(rec.ctx, model(family === 'area' ? {
        chartType: 'area',
        classicChartStyleRoles: {
          dataPoint: { fillColors: ['135724'], lineColors: ['246813'] },
        },
      } : {
        chartType: 'clusteredBar',
        series: [
          series({ seriesType: 'area' }),
          series({ seriesType: 'bar', values: [2, 3, 4] }),
        ],
        classicChartStyleRoles: {
          dataPoint: { fillColors: ['135724'], lineColors: ['246813'] },
        },
      }), RECT, 1);

      expect(rec.fills).toContain('#135724');
      expect(rec.strokes).toContain('#246813');
    },
  );

  it.each(['bar', 'area'] as const)(
    'keeps a varyColors line overlay point-indexed in a %s combo',
    host => {
      const rec = recordingContext();
      const hostSeries = series({
        seriesType: host,
        values: [2, 3, 4],
      });
      const line = series({
        seriesType: 'line',
        values: [1, 2, 3],
        showMarker: false,
      });
      renderChart(rec.ctx, model({
        chartType: host === 'bar' ? 'clusteredBar' : 'area',
        series: [hostSeries, line],
        plotGroups: [
          {
            kind: host, seriesStart: 0, seriesCount: 1,
            categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
          },
          {
            kind: 'line', seriesStart: 1, seriesCount: 1, varyColors: true,
            categoryAxis: 'primary', valueAxis: 'primary', seriesAxis: 'none',
          },
        ],
        varyingPointChartStyleRolesByGroup: [
          null,
          {
            dataPointLine: {
              lineColors: ['AA0000', '00AA00', '0000AA'],
              lineFormattingIndices: [0, 1, 2],
            },
          },
        ],
      }), RECT, 1);

      expect(rec.strokes).toContain('#00AA00');
      expect(rec.strokes).toContain('#0000AA');
    },
  );

  it('uses dataPoint fill for filled radar polygons', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'radar',
      radarStyle: 'filled',
      classicChartStyleRoles: { dataPoint: { fillColors: ['ABC123'] } },
    }), RECT, 1);
    expect(rec.fills).toContain('#ABC123');
  });

  it('uses dataPoint and seriesLine roles for ofPie marks and connectors', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'ofPie',
      series: [series({ values: [40, 30, 20, 10] })],
      ofPie: {
        type: 'bar', splitType: 'pos', splitPos: 2,
        secondPieSizePercent: 75, gapWidthPercent: 100, seriesLines: true,
      },
      classicChartStyleRoles: {
        dataPoint: { fillColors: ['110000', '002200', '000033', '444400'] },
      },
      // Unmeasured linked role: series lines still follow the linked style.
      chartStyleRoles: { seriesLine: { lineColors: ['778899'] } },
    }), RECT, 1);

    expect([...rec.fills, ...rec.rectFills]).toContain('#000033');
    expect(rec.strokes).toContain('#778899');
  });

  it('keeps a geometry-only point marker on its owning series style index', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'line',
      series: [series({
        chartexFormatIdx: 2,
        markerSymbol: 'circle',
        dataPointOverrides: [{ idx: 1, markerSize: 6 }],
      })],
      classicChartStyleRoles: {
        dataPointMarker: { fillColors: ['AA0000', '00AA00', '0000AA'] },
      },
    }), RECT, 1);

    expect(rec.fills.filter(color => color === '#0000AA')).toHaveLength(3);
    expect(rec.fills.filter(color => color === '#00AA00')).toHaveLength(0);
  });

  it('separates bubble point and series style palette indexes', () => {
    const rec = recordingContext();
    renderChart(rec.ctx, model({
      chartType: 'bubble',
      series: [series({
        chartexFormatIdx: 2,
        categories: ['1', '2'], values: [1, 2], bubbleSizes: [1, 1],
        chartexStyle: { fillColors: ['AA0000', '00AA00', '0000AA'] },
        dataPointOverrides: [{
          idx: 1,
          chartexStyle: { fillColors: ['CC0000', '00CC00', '0000CC'] },
        }],
      })],
    }), RECT, 1);

    expect(rec.fills).toEqual(['rgba(0,0,170,1)', 'rgba(0,204,0,1)']);
  });
});

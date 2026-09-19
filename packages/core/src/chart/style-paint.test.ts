import { describe, expect, it } from 'vitest';
import type { ChartModel } from '../types/chart.js';
import {
  chartStyleColor,
  chartStyleFillCascade,
  chartStyleFillDecision,
  chartStyleFillPaint,
  chartStyleFontColor,
  chartStyleLineCascade,
  chartThreeDSurfacePaint,
} from './style-paint.js';

describe('compact classic chart-style palettes', () => {
  it('selects sparse source formatting indexes without allocating through the largest index', () => {
    const style = {
      fillColors: ['222222', '888888'],
      fillPaints: [
        { fillType: 'solid' as const, color: 'AAAAAA' },
        { fillType: 'solid' as const, color: 'BBBBBB' },
      ],
      fillFormattingIndices: [2, 8],
    };
    expect(chartStyleColor(style, 'fill', 2)).toBe('222222');
    expect(chartStyleColor(style, 'fill', 8)).toBe('888888');
    expect(chartStyleFillPaint(style, 8)).toEqual({ fillType: 'solid', color: 'BBBBBB' });
  });

  it('falls through only the explicitly bounded Pattern2 overflow indexes', () => {
    const style = {
      fillPaintAuthored: true,
      fillColors: [null],
      fillSemanticFallbackIndices: [48],
    };
    expect(chartStyleFillDecision(style, 47)).toBeNull();
    expect(chartStyleFillDecision(style, 48)).toBeUndefined();
  });

  it('selects relative and fixed fontRef palette indexes', () => {
    expect(chartStyleFontColor({ fontColors: ['111111', '222222'] }, 1)).toBe('222222');
    expect(chartStyleFontColor({
      fontColors: ['111111', '222222'], fontColorIndex: 0,
    }, 1)).toBe('111111');
  });

  it('does not treat an omitted component in direct spPr as no paint', () => {
    const linked = {
      fillColors: ['112233'],
      lineColors: ['445566'],
      allowNoFillOverride: true,
      allowNoLineOverride: true,
    };
    expect(chartStyleFillCascade(linked, 0)).toEqual({ fillType: 'solid', color: '112233' });
    expect(chartStyleLineCascade(linked, 0)).toEqual({ fillType: 'solid', color: '445566' });
    expect(chartStyleFillCascade(linked, 0, { shapePropertiesPresent: true }))
      .toEqual({ fillType: 'solid', color: '112233' });
    expect(chartStyleLineCascade(linked, 0, { shapePropertiesPresent: true }))
      .toEqual({ fillType: 'solid', color: '445566' });
    expect(chartStyleFillCascade({ ...linked, allowNoFillOverride: false }, 0,
      { shapePropertiesPresent: true })).toEqual({ fillType: 'solid', color: '112233' });
  });
});

describe('3-D surface dash cascade', () => {
  const chart = (direct: object, linked: object): ChartModel => ({
    chartType: 'surface3D',
    series: [],
    threeD: { floor: { style: direct }, backWall: null, sideWall: null },
    chartStyleRoles: { floor: linked },
  } as unknown as ChartModel);

  it('keeps a direct preset dash atomic over a linked custom dash', () => {
    const model = chart(
      { lineDash: 'dash', lineDashAuthored: true },
      { lineCustomDash: [{ dash: 3, space: 2 }], lineDashAuthored: true },
    );
    expect(chartThreeDSurfacePaint(model, model.threeD?.floor, 'floor')).toMatchObject({
      lineDash: 'dash',
      lineCustomDash: undefined,
    });
  });

  it('keeps a direct custom dash atomic over a linked preset dash', () => {
    const custom = [{ dash: 3, space: 2 }];
    const model = chart(
      { lineCustomDash: custom, lineDashAuthored: true },
      { lineDash: 'dot', lineDashAuthored: true },
    );
    expect(chartThreeDSurfacePaint(model, model.threeD?.floor, 'floor')).toMatchObject({
      lineDash: undefined,
      lineCustomDash: custom,
    });
  });

  it('retains linked geometry beside a NoStyle line reference', () => {
    const model = chart({}, {
      lineHidden: true,
      lineNoStyle: true,
      lineWidthEmu: 25_400,
      lineDash: 'dash',
      lineCap: 'rnd',
      lineJoin: 'bevel',
    });
    expect(chartThreeDSurfacePaint(model, model.threeD?.floor, 'floor')).toMatchObject({
      line: undefined,
      lineWidthEmu: 25_400,
      lineDash: 'dash',
      lineCap: 'rnd',
      lineJoin: 'bevel',
    });
  });
});

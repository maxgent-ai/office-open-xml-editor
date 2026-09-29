// Classic chart chartex style helpers.
import type { ChartModel, ChartRect, ChartSeries, ChartStyleRole } from '../../types/chart';
import {
  chartExPointAuthorsLine,
  chartExPointFillDecision,
  chartStyleColor,
  chartStyleDirectFillDecision,
  chartStyleFillDecision,
  chartStyleLineDecision,
  type ChartExPointCarrier,
} from '../style-paint.js';
import type { Fill } from '../../types/common';
import { rawLinkedChartStyleRole } from '../effective-style.js';
import { resolveFill } from '../../shape/paint.js';
import { paintChartImageFill } from '../image-fill.js';
import { axisLineWidthPx } from '../axis-style.js';
import { CHARTEX_DEFAULT_PALETTE } from './palette.js';
import { dashPatternForLine } from './geometry.js';


// ═══════════════════════════════════════════════════════════════════════════
// Waterfall chart — subtotal bars filled, delta bars outlined.
// ═══════════════════════════════════════════════════════════════════════════

export type ChartExStyle = NonNullable<ChartModel['chartexDataPointStyle']>;


/** Resolve CT_DataLabels + indexed CT_DataLabel/dataLabelHidden without
 * renderer-specific precedence. Defaults describe the semantic label layer of
 * the chart type; authored visibility always overrides those defaults. */
export function chartExStyleColor(
  _chart: ChartModel,
  style: ChartExStyle | null | undefined,
  kind: 'fill' | 'line',
  index: number,
  _count: number,
): string | null {
  return chartStyleColor(style, kind, index);
}


export function chartExPaletteColor(
  chart: ChartModel,
  colors: ReadonlyArray<string | null | undefined>,
  colorIndex: number,
  _count: number,
): string | null {
  if (!colors.length) return null;
  const method = chart.chartexColorStyleMethod;
  const knownMethod = method === 'withinLinear'
    || method === 'acrossLinear'
    || method === 'withinLinearReversed'
    || method === 'acrossLinearReversed';
  // MS-ODRAWXML §2.8.4.1: unknown method strings have cycle semantics.
  if (!knownMethod) return colors[colorIndex % colors.length] ?? null;
  const within = method === 'withinLinear' || method === 'withinLinearReversed';
  // The specification defines which base color linear methods use, but does
  // not define the brightness range or color space. Preserve the authored
  // color here instead of inventing an Office compatibility curve. Once an
  // observed/approved rule exists, brightness belongs before styleClr/style
  // matrix transforms in the shared parser model, not as a post-paint tweak.
  return colors[within ? 0 : colorIndex % colors.length] ?? null;
}


export function chartExSemanticFill(chart: ChartModel, index: number, count: number): string {
  return (chart.chartexColorPalette
      ? chartExPaletteColor(chart, chart.chartexColorPalette, index, count)
      : null)
    ?? chart.chartexAccents?.[index % (chart.chartexAccents.length || 1)]
    ?? CHARTEX_DEFAULT_PALETTE[index % CHARTEX_DEFAULT_PALETTE.length];
}


export function chartExDataPointFill(
  chart: ChartModel,
  index: number,
  count: number,
  localStyle?: ChartExStyle | null,
): string {
  return chartExStyleColor(chart, localStyle, 'fill', index, count)
    ?? chartExStyleColor(chart, chart.chartexDataPointStyle, 'fill', index, count)
    ?? chartExSemanticFill(chart, index, count);
}


/** Line-paint counterpart of chartExStylePaintDecision. `undefined` means the
 * layer did not author a line paint, while `null` is an authored noFill or an
 * authored-but-unresolved paint that must suppress lower-precedence color. */
export function chartExStyleLinePaintDecision(
  chart: ChartModel,
  style: ChartExStyle | null | undefined,
  index: number,
  count: number,
): ChartModel['plotAreaLineFill'] | null | undefined {
  void chart;
  void count;
  return chartStyleLineDecision(style, index);
}


/** Resolve one ChartEx style paint layer. `undefined` means this layer supplied
 * no paint, while `null` records an explicit no-fill for consumers whose own
 * shape is governed by that layer. */
export function chartExStylePaintDecision(
  chart: ChartModel,
  style: ChartExStyle | null | undefined,
  index: number,
  count: number,
): Fill | null | undefined {
  void chart;
  void count;
  return chartStyleFillDecision(style, index);
}


export function chartExMarkerPaint(
  chart: ChartModel,
  index: number,
  count: number,
  localStyle: ChartExStyle | null | undefined,
  legacyColor: string | null | undefined,
  linkedStyle: ChartExStyle | null | undefined,
): Fill | null {
  const role = linkedStyle === chart.chartexDataPointMarkerStyle
    ? 'dataPointMarker'
    : linkedStyle === chart.chartexDataPointStyle
      ? 'dataPoint'
      : (Object.entries(chart.chartStyleRoles ?? {}).find(
          ([, style]) => style === linkedStyle,
        )?.[0] as ChartStyleRole | undefined);
  const rawLinkedStyle = role
    ? rawLinkedChartStyleRole(chart, role)
      ?? (chart.classicChartStyleRoles == null ? linkedStyle : undefined)
    : linkedStyle;
  const local = chartStyleDirectFillDecision(localStyle, rawLinkedStyle, index);
  if (local !== undefined) return local;
  if (legacyColor) return { fillType: 'solid', color: legacyColor };
  const linked = chartExStylePaintDecision(chart, linkedStyle, index, count);
  if (linked !== undefined) return linked;
  return { fillType: 'solid', color: chartExSemanticFill(chart, index, count) };
}


export function chartExDataPointPaint(
  chart: ChartModel,
  index: number,
  count: number,
  localStyle?: ChartExStyle | null,
  legacyColor?: string | null,
  linkedStyle: ChartExStyle | null | undefined = chart.chartexDataPointStyle,
): Fill | null {
  // ChartEx direct CT_Series / CT_DataPoint `spPr` owns each fill atom: a
  // positive paint and an explicit `a:noFill` both replace the linked
  // dataPoint Chart Style fill, which supplies only an omitted fill. Unlike
  // the classic cascade, `allowNoFillOverride` does not gate this (see
  // resolveChartExPointFill for the PowerPoint evidence).
  const local = chartStyleFillDecision(localStyle, index);
  if (local !== undefined) return local;
  if (localStyle && legacyColor) return { fillType: 'solid', color: legacyColor };
  if (legacyColor) return { fillType: 'solid', color: legacyColor };
  const linked = chartStyleFillDecision(linkedStyle, index);
  if (linked !== undefined) return linked;
  return { fillType: 'solid', color: chartExSemanticFill(chart, index, count) };
}


export function chartExFillStyle(
  ctx: CanvasRenderingContext2D,
  paint: Fill,
  x: number,
  y: number,
  w: number,
  h: number,
  fallbackColor: string,
  shapeRotationDeg = 0,
): string | CanvasGradient | CanvasPattern {
  // Keep solid ChartEx paints byte-compatible with the renderer's historical
  // `#RRGGBB` path. The shared resolver is needed only for structured fills;
  // routing solids through it would rewrite equivalent colors as rgba().
  if (paint.fillType === 'solid') {
    return paint.color.startsWith('#') ? paint.color : `#${paint.color}`;
  }
  return resolveFill(paint, ctx, x, y, w, h, shapeRotationDeg) ?? fallbackColor;
}


/** Paint an already-constructed classic mark path. Picture fills require the
 * current geometric path as an outer clip; the shared image painter then owns
 * crop/tile/stretch and its rectangular destination clip. Missing decoded
 * sources remain transparent rather than reviving the semantic fallback. */
export function paintClassicDataPointPath(
  ctx: CanvasRenderingContext2D,
  paint: Fill | null | undefined,
  bounds: ChartRect,
  fallbackColor: string,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (paint === null) return false;
  if (paint?.fillType === 'image') {
    ctx.save();
    ctx.clip();
    const painted = paintChartImageFill(
      ctx, paint, bounds.x, bounds.y, bounds.w, bounds.h, ptToPx, shapeRotationDeg,
    );
    ctx.restore();
    return painted;
  }
  ctx.fillStyle = paint
    ? chartExFillStyle(
        ctx, paint, bounds.x, bounds.y, bounds.w, bounds.h,
        fallbackColor, shapeRotationDeg,
      )
    : fallbackColor;
  ctx.fill();
  return true;
}


export function paintClassicDataPointRect(
  ctx: CanvasRenderingContext2D,
  paint: Fill | null | undefined,
  bounds: ChartRect,
  fallbackColor: string,
  ptToPx: number,
  shapeRotationDeg = 0,
): boolean {
  if (paint === null || !(bounds.w > 0) || !(bounds.h > 0)) return false;
  if (paint?.fillType === 'image') {
    ctx.beginPath();
    ctx.rect(bounds.x, bounds.y, bounds.w, bounds.h);
    return paintClassicDataPointPath(
      ctx, paint, bounds, fallbackColor, ptToPx, shapeRotationDeg,
    );
  }
  ctx.fillStyle = paint
    ? chartExFillStyle(
        ctx, paint, bounds.x, bounds.y, bounds.w, bounds.h,
        fallbackColor, shapeRotationDeg,
      )
    : fallbackColor;
  ctx.fillRect(bounds.x, bounds.y, bounds.w, bounds.h);
  return true;
}


export interface ResolvedChartExLineStyle {
  visible: boolean;
  color: string;
  paint: ChartModel['plotAreaLineFill'] | null | undefined;
  widthEmu: number | null;
  dash: string | null;
  customDash: ChartModel['plotAreaLineCustomDash'];
  cap: string | null;
  join: string | null;
  /** True when no direct or linked layer supplied the paint and the family's
   * semantic fallback outline is used. */
  semanticFallback?: boolean;
}


export type ChartExSeriesStyleCarrier = Pick<
  ChartSeries,
  'chartexStyle' | 'lineHidden' | 'lineColor' | 'lineWidthEmu'
>;


/** Resolve a ChartEx mark's effective outline once for both plot and legend.
 * Direct CT_Series formatting wins over the linked Chart Style role. `NoStyle`
 * is absence of decoration (and may expose a family semantic outline), while
 * an explicit noFill suppresses the outline. */
export function resolveChartExSeriesLineStyle(
  chart: ChartModel,
  linkedStyle: ChartExStyle | null | undefined,
  series: Partial<ChartExSeriesStyleCarrier> | null | undefined,
  index: number,
  count: number,
  fallbackColor: string,
  options: { linkedNoStyleFallback?: boolean } = {},
): ResolvedChartExLineStyle {
  void chart;
  void count;
  const local = series?.chartexStyle;
  // Direct ChartEx `spPr/a:ln` owns the outline paint, including an explicit
  // no-line, without the linked entry's `allowNoLineOverride` (see
  // resolveChartExPointFill for the PowerPoint evidence).
  const localPaint = chartStyleLineDecision(local, index);
  const legacyNoLine = series?.lineHidden === true ? null : undefined;
  const legacyPaintAuthored = series?.lineColor != null
    || legacyNoLine !== undefined;
  const linkedPaint = chartStyleLineDecision(linkedStyle, index);
  const selectedPaint = localPaint !== undefined
    ? localPaint
    : legacyPaintAuthored
      ? legacyNoLine !== undefined
        ? legacyNoLine
        : { fillType: 'solid' as const, color: series?.lineColor ?? fallbackColor }
      : linkedPaint !== undefined
        ? linkedPaint
        : undefined;
  // NoStyle applies only to paint. It deliberately exposes the semantic
  // family fallback while local geometry beside the NoStyle reference remains
  // part of the effective outline.
  const linkedNoStyleSuppressesSemanticPaint = selectedPaint === undefined
    && linkedStyle?.lineNoStyle === true
    && options.linkedNoStyleFallback !== true;
  return {
    semanticFallback: selectedPaint === undefined,
    visible: selectedPaint !== null && !linkedNoStyleSuppressesSemanticPaint,
    color: selectedPaint?.fillType === 'solid'
      ? selectedPaint.color
      : fallbackColor,
    paint: selectedPaint?.fillType === 'solid' ? undefined : selectedPaint,
    widthEmu: local?.lineWidthEmu
      ?? series?.lineWidthEmu ?? linkedStyle?.lineWidthEmu ?? null,
    dash: local?.lineCustomDash != null
      ? null : local?.lineDash ?? linkedStyle?.lineDash ?? null,
    customDash: local?.lineCustomDash ?? linkedStyle?.lineCustomDash ?? null,
    cap: local?.lineCap ?? linkedStyle?.lineCap ?? null,
    join: local?.lineJoin ?? linkedStyle?.lineJoin ?? null,
  };
}


/** Width PowerPoint uses for a ChartEx outline when neither the direct
 * `a:ln` nor the linked style entry supplies `w`: 0.75 pt in every measured
 * family (see resolveChartExPointFill). */
export const CHARTEX_DEFAULT_LINE_WIDTH_EMU = 9525;


/** Resolve one ChartEx data point's body fill.
 *
 * Precedence is per fill atom: the CT_DataPoint `spPr` fill, else the
 * CT_Series `spPr` fill, else the linked dataPoint Chart Style role, else the
 * family's semantic palette. An explicit direct `a:noFill` at either level
 * removes the fill.
 *
 * PowerPoint-observed (PowerPoint for Mac 16.113; waterfall, funnel,
 * histogram, pareto, box-and-whisker, treemap and sunburst; linked entries
 * `fillRef idx=1` with `lnRef` idx 0, 1 and 2, each with and without
 * `mods="allowNoFillOverride allowNoLineOverride"`):
 * - series or point `a:noFill` removes the fill, and series or point
 *   `a:ln/a:noFill` removes the outline, whether or not the modifiers are
 *   present;
 * - a point's positive fill or outline wins over the series, and an omitted
 *   fill or outline atom at either level inherits the next layer;
 * - an outline whose direct and linked `a:ln` both omit `w` is 0.75 pt.
 *   An `lnRef idx` >= 1 supplies the theme line's width and cap when `w` is
 *   omitted.
 * Box-and-whisker `dataPt` formatting had no visible effect on the box body,
 * so that family does not route bodies through point overrides. */
export function resolveChartExPointFill(
  chart: ChartModel,
  series: Pick<ChartSeries, 'chartexStyle' | 'color'> | null | undefined,
  point: ChartExPointCarrier | null | undefined,
  index: number,
  count: number,
  linkedStyle: ChartExStyle | null | undefined = chart.chartexDataPointStyle,
): Fill | null {
  const decision = chartExPointFillDecision(chart, series, point, index, linkedStyle);
  if (decision !== undefined) return decision;
  return { fillType: 'solid', color: chartExSemanticFill(chart, index, count) };
}


/** Resolve a ChartEx outline through ordered Chart Style role chains.
 * `paintRoles` are tried in order after the direct carrier until one supplies
 * line paint; `geometryRoles` supply `w`, dash, cap and join atoms the direct
 * `a:ln` omits (see the geometry-role selection below). A semantic-fallback
 * result takes role geometry only from its paint roles. */
export function resolveChartExLineChain(
  chart: ChartModel,
  carrier: Partial<ChartExSeriesStyleCarrier> | null | undefined,
  paintRoles: ReadonlyArray<ChartExStyle | null | undefined>,
  geometryRoles: ReadonlyArray<ChartExStyle | null | undefined>,
  index: number,
  count: number,
  fallbackColor: string,
  options: { linkedNoStyleFallback?: boolean } = {},
): ResolvedChartExLineStyle {
  let line = resolveChartExSeriesLineStyle(
    chart, paintRoles[0], carrier, index, count, fallbackColor, options,
  );
  for (const role of paintRoles.slice(1)) {
    if (!line.semanticFallback) break;
    const next = resolveChartExSeriesLineStyle(
      chart, role, null, index, count, fallbackColor, options,
    );
    if (!next.semanticFallback) line = next;
  }
  const direct = carrier?.chartexStyle;
  const directDash = direct?.lineDash != null || direct?.lineCustomDash != null;
  // The first role that carries a line supplies every omitted geometry atom
  // as a unit; atoms are not merged across roles. Excel output shows a
  // dataPoint `a:ln w` without `cap` keeping a flat cap even when the
  // dataPointLine role authors `cap="rnd"`. NoStyle (`lnRef idx=0`) affects
  // paint only, so geometry authored beside it still counts. A semantic
  // fallback outline takes geometry only from its paint roles: PowerPoint
  // draws no data-point outline from dataPointLine alone (round 2 R02).
  const geometryCandidates = line.semanticFallback ? paintRoles : geometryRoles;
  const geometryRole = geometryCandidates.find(role => role != null && (
    role.lineWidthEmu != null || role.lineCap != null || role.lineJoin != null
    || role.lineDash != null || role.lineCustomDash != null
    || (role.linePaintAuthored === true && role.lineNoStyle !== true)));
  const fromRoles = <T>(pick: (role: ChartExStyle) => T | null | undefined): T | null =>
    geometryRole ? pick(geometryRole) ?? null : null;
  const roleDash = geometryRole;
  return {
    ...line,
    widthEmu: direct?.lineWidthEmu ?? carrier?.lineWidthEmu
      ?? fromRoles(role => role.lineWidthEmu),
    dash: directDash
      ? direct?.lineCustomDash != null ? null : direct?.lineDash ?? null
      : roleDash?.lineCustomDash != null ? null : roleDash?.lineDash ?? null,
    customDash: directDash
      ? direct?.lineCustomDash ?? null
      : roleDash?.lineCustomDash ?? null,
    cap: direct?.lineCap ?? fromRoles(role => role.lineCap),
    join: direct?.lineJoin ?? fromRoles(role => role.lineJoin),
  };
}


/** Resolve one ChartEx data point's outline with the same point → series →
 * linked role precedence as resolveChartExPointFill. Paint belongs to the
 * layer that authors `a:ln`, else the linked dataPoint role. Geometry atoms a
 * point's `a:ln` omits fall back to the series outline, then the dataPoint
 * role (its `spPr` line or `lnRef` theme line), then the dataPointLine role's
 * geometry. PowerPoint-observed in every ChartEx family: dataPointLine never
 * paints a data-point outline, but supplies `w`/cap to a direct one whose
 * dataPoint role has no line (for example 2.25 pt round from the waterfall
 * default style); a dataPoint `lnRef idx` >= 1 or `spPr` line wins over it. */
export function resolveChartExPointLine(
  chart: ChartModel,
  series: Partial<ChartExSeriesStyleCarrier> | null | undefined,
  point: ChartExPointCarrier | null | undefined,
  index: number,
  count: number,
  fallbackColor: string,
  linkedStyle: ChartExStyle | null | undefined = chart.chartexDataPointStyle,
  options: { linkedNoStyleFallback?: boolean } = {},
): ResolvedChartExLineStyle {
  const geometryRoles = linkedStyle === chart.chartexDataPointStyle
    ? [linkedStyle, chart.chartexDataPointLineStyle]
    : [linkedStyle];
  if (!chartExPointAuthorsLine(point)) {
    return resolveChartExLineChain(
      chart, series, [linkedStyle], geometryRoles, index, count, fallbackColor, options,
    );
  }
  const pointStyle = point?.chartexStyle;
  const seriesStyle = series?.chartexStyle;
  const pointDashAuthored = pointStyle?.lineDash != null || pointStyle?.lineCustomDash != null;
  // Paint and geometry are separate atoms: a point `a:ln` that authors only
  // geometry keeps the series outline paint (the same per-atom inheritance the
  // controls show for fill versus line).
  // The parser records `lineHidden: false` for any authored `a:ln`, so only an
  // explicit no-line or actual paint atoms mean the point authors paint.
  const pointPaintAuthored = point?.lineColor != null || point?.lineHidden === true
    || pointStyle?.linePaintAuthored === true || pointStyle?.lineHidden === true
    || pointStyle?.linePaints?.some(paint => paint != null) === true
    || pointStyle?.lineColors?.some(color => color != null) === true;
  const paintSource = pointPaintAuthored
    ? { lineColor: point?.lineColor, lineHidden: point?.lineHidden, style: pointStyle }
    : { lineColor: series?.lineColor, lineHidden: series?.lineHidden, style: seriesStyle };
  const carrier: Partial<ChartExSeriesStyleCarrier> = {
    lineColor: paintSource.lineColor,
    lineHidden: paintSource.lineHidden,
    lineWidthEmu: point?.lineWidthEmu ?? series?.lineWidthEmu,
    chartexStyle: {
      ...pointStyle,
      linePaints: paintSource.style?.linePaints,
      lineColors: paintSource.style?.lineColors,
      lineColorIndex: paintSource.style?.lineColorIndex,
      linePaintAuthored: paintSource.style?.linePaintAuthored,
      lineHidden: paintSource.style?.lineHidden,
      lineNoStyle: paintSource.style?.lineNoStyle,
      lineWidthEmu: pointStyle?.lineWidthEmu ?? point?.lineWidthEmu
        ?? seriesStyle?.lineWidthEmu ?? series?.lineWidthEmu,
      lineDash: pointDashAuthored ? pointStyle?.lineDash : seriesStyle?.lineDash,
      lineCustomDash: pointDashAuthored
        ? pointStyle?.lineCustomDash : seriesStyle?.lineCustomDash,
      lineCap: pointStyle?.lineCap ?? seriesStyle?.lineCap,
      lineJoin: pointStyle?.lineJoin ?? seriesStyle?.lineJoin,
    },
  };
  return resolveChartExLineChain(
    chart, carrier, [linkedStyle], geometryRoles, index, count, fallbackColor, options,
  );
}


export function applyResolvedChartExLineStyle(
  ctx: CanvasRenderingContext2D,
  line: ResolvedChartExLineStyle,
  ptToPx: number,
  bounds?: ChartRect,
  shapeRotationDeg = 0,
): boolean {
  if (!line.visible) return false;
  // Structured (gradient/pattern) outline paint resolves against the stroked
  // shape's bounds; a paint Canvas cannot express leaves the outline unpainted
  // rather than reviving the solid fallback. Callers without bounds keep the
  // resolved solid colour.
  if (line.paint && line.paint.fillType !== 'solid' && bounds) {
    const stroke = resolveFill(
      line.paint, ctx, bounds.x, bounds.y, bounds.w, bounds.h, shapeRotationDeg,
    );
    if (!stroke) return false;
    ctx.strokeStyle = stroke;
  } else {
    ctx.strokeStyle = line.color.startsWith('#') ? line.color : `#${line.color}`;
  }
  // An authored outline without `w` is 0.75 pt. A family's semantic fallback
  // outline (e.g. the treemap tile separator standing in for Office's tile
  // gap) keeps its one-device-pixel rule.
  ctx.lineWidth = line.widthEmu != null
    ? axisLineWidthPx(line.widthEmu, ptToPx)
    : line.semanticFallback
      ? 1
      : axisLineWidthPx(CHARTEX_DEFAULT_LINE_WIDTH_EMU, ptToPx);
  ctx.setLineDash(dashPatternForLine(line.customDash, line.dash, ctx.lineWidth));
  ctx.lineCap = line.cap === 'rnd' ? 'round' : line.cap === 'sq' ? 'square' : 'butt';
  ctx.lineJoin = line.join === 'round' || line.join === 'bevel' ? line.join : 'miter';
  return true;
}


/** Apply CT_Series local shape properties before the linked Chart Style.
 *  [MS-ODRAWXML] 2.24.3.77 makes `<cx:series><cx:spPr>` the series' own
 *  OfficeArt formatting, so an authored line (including `noFill`) overrides
 *  the default data-point recipe instead of being merged underneath it. */
export function applyChartExSeriesLineStyle(
  ctx: CanvasRenderingContext2D,
  chart: ChartModel,
  style: ChartExStyle | null | undefined,
  series: Pick<ChartSeries, 'chartexStyle' | 'lineHidden' | 'lineColor' | 'lineWidthEmu'> | null | undefined,
  index: number,
  count: number,
  fallbackColor: string,
  ptToPx: number,
  options: { linkedNoStyleFallback?: boolean } = {},
): boolean {
  return applyResolvedChartExLineStyle(
    ctx,
    resolveChartExSeriesLineStyle(
      chart, style, series, index, count, fallbackColor, options,
    ),
    ptToPx,
  );
}


/** Build a synthetic legend series from the same resolved line contract used
 * by plot paint. `chartexStyle` carries dash/cap/join through the generic
 * legend pipeline without widening the public ChartSeries surface. */
export function chartExLegendSeries(
  chart: ChartModel,
  name: string,
  series: Partial<ChartExSeriesStyleCarrier> | null | undefined,
  linkedStyle: ChartExStyle | null | undefined,
  index: number,
  count: number,
  fillColor: string,
  semanticNoStyleFallback = false,
  inheritPlotOutline = true,
): ChartSeries {
  // Legend keys follow the same role chain as the plotted body so their
  // outline geometry (e.g. a dataPointLine 2.25 pt round rule) matches.
  const line = resolveChartExLineChain(
    chart,
    series,
    [linkedStyle],
    linkedStyle === chart.chartexDataPointStyle
      ? [linkedStyle, chart.chartexDataPointLineStyle]
      : [linkedStyle],
    index,
    count,
    fillColor,
    { linkedNoStyleFallback: semanticNoStyleFallback },
  );
  return {
    name,
    values: [],
    color: fillColor.replace(/^#/, ''),
    lineHidden: !inheritPlotOutline || !line.visible,
    lineColor: inheritPlotOutline && line.visible ? line.color.replace(/^#/, '') : null,
    // A visible authored outline without `w` is 0.75 pt on the body; give
    // the legend key the same default rather than the legend's 1 px rule.
    lineWidthEmu: inheritPlotOutline
      ? line.widthEmu ?? (line.visible && !line.semanticFallback
        ? CHARTEX_DEFAULT_LINE_WIDTH_EMU : null)
      : null,
    chartexStyle: {
      linePaints: inheritPlotOutline && line.paint !== undefined ? [line.paint] : null,
      linePaintAuthored: inheritPlotOutline && line.paint !== undefined ? true : null,
      lineDash: inheritPlotOutline ? line.dash : null,
      lineCustomDash: inheritPlotOutline ? line.customDash : null,
      lineCap: inheritPlotOutline ? line.cap : null,
      lineJoin: inheritPlotOutline ? line.join : null,
    },
  };
}

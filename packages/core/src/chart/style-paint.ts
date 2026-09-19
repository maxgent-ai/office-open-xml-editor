import type {
  ChartExElementStyle, ChartModel, ChartStockBarPaint,
} from '../types/chart.js';
import type { Fill } from '../types/common.js';
import { chartStyleDashChoice } from './effective-style.js';
import { compactStyleIndex, styleIndexSetHas } from './sparse-style-index.js';

function chartStylePaletteIndex(
  style: ChartExElementStyle,
  kind: 'font' | 'fill' | 'line' | 'effect',
  sourceIndex: number,
): number {
  const formattingIndices = kind === 'font'
    ? style.fontFormattingIndices
    : kind === 'fill' ? style.fillFormattingIndices
      : kind === 'line' ? style.lineFormattingIndices : style.effectFormattingIndices;
  const compactIndex = formattingIndices
    ? compactStyleIndex(formattingIndices, sourceIndex)
    : -1;
  return compactIndex >= 0 ? compactIndex : sourceIndex;
}

/** Select one already-expanded `fontRef/styleClr` palette entry. */
export function chartStyleFontColor(
  style: ChartExElementStyle | null | undefined,
  index: number,
): string | null {
  if (!style) return null;
  if (!style.fontColors?.length) return style.fontColor ?? null;
  const paletteIndex = style.fontColorIndex ?? chartStylePaletteIndex(style, 'font', index);
  return style.fontColors[paletteIndex % style.fontColors.length] ?? null;
}

/** Select one already-expanded linked Chart Style palette entry. */
export function chartStyleColor(
  style: ChartExElementStyle | null | undefined,
  kind: 'fill' | 'line',
  index: number,
): string | null {
  if (!style) return null;
  const colors = kind === 'fill' ? style.fillColors : style.lineColors;
  if (!colors?.length) return null;
  const fixedIndex = kind === 'fill' ? style?.fillColorIndex : style?.lineColorIndex;
  const paletteIndex = fixedIndex ?? chartStylePaletteIndex(style, kind, index);
  return colors[paletteIndex % colors.length] ?? null;
}

export function chartStyleFillPaint(
  style: ChartExElementStyle | null | undefined,
  index: number,
): Fill | null {
  if (!style) return null;
  const paints = style.fillPaints;
  if (!paints?.length) return null;
  const paletteIndex = style.fillColorIndex ?? chartStylePaletteIndex(style, 'fill', index);
  return paints[paletteIndex % paints.length] ?? null;
}

export function chartStyleLinePaint(
  style: ChartExElementStyle | null | undefined,
  index: number,
): ChartModel['plotAreaLineFill'] {
  if (!style) return null;
  const paints = style.linePaints;
  if (!paints?.length) return null;
  const paletteIndex = style.lineColorIndex ?? chartStylePaletteIndex(style, 'line', index);
  return paints[paletteIndex % paints.length] ?? null;
}

/** `undefined` means no authored fill at this layer; `null` means authored
 * no-fill or authored-but-unresolved and therefore suppresses fallback. */
export function chartStyleFillDecision(
  style: ChartExElementStyle | null | undefined,
  index: number,
): Fill | null | undefined {
  if (!style) return undefined;
  if (style.fillHidden) return style.fillNoStyle ? undefined : null;
  if (styleIndexSetHas(style.fillSemanticFallbackIndices, index)) return undefined;
  const paint = chartStyleFillPaint(style, index);
  if (paint) return paint;
  const color = chartStyleColor(style, 'fill', index);
  if (color) return { fillType: 'solid', color };
  return style.fillPaintAuthored === true ? null : undefined;
}

/** Line-paint counterpart of {@link chartStyleFillDecision}. */
export function chartStyleLineDecision(
  style: ChartExElementStyle | null | undefined,
  index: number,
): ChartModel['plotAreaLineFill'] | null | undefined {
  if (!style) return undefined;
  if (style.lineHidden) return style.lineNoStyle ? undefined : null;
  if (styleIndexSetHas(style.lineSemanticFallbackIndices, index)) return undefined;
  const paint = chartStyleLinePaint(style, index);
  if (paint) return paint;
  const color = chartStyleColor(style, 'line', index);
  if (color) return { fillType: 'solid', color };
  return style.linePaintAuthored === true ? null : undefined;
}

/** Resolve direct shape paint over a linked CT_StyleEntry. An omitted fill or
 * line in a present `spPr` still inherits: MS-ODRAWXML's `allowNo*Override`
 * permits an authored `noFill`/no-line choice to replace the style; it does
 * not turn an absent component into that choice. */
export function chartStyleFillCascade(
  linked: ChartExElementStyle | null | undefined,
  index: number,
  ...direct: Array<ChartExElementStyle | null | undefined>
): Fill | null | undefined {
  for (const style of direct) {
    const paint = chartStyleFillDecision(style, index);
    if (paint !== undefined) return paint;
  }
  return chartStyleFillDecision(linked, index);
}

export function chartStyleLineCascade(
  linked: ChartExElementStyle | null | undefined,
  index: number,
  ...direct: Array<ChartExElementStyle | null | undefined>
): ChartModel['plotAreaLineFill'] | null | undefined {
  for (const style of direct) {
    const paint = chartStyleLineDecision(style, index);
    if (paint !== undefined) return paint;
  }
  return chartStyleLineDecision(linked, index);
}

/** Resolve the fill atom of one classic up/down bar. Direct CT_UpDownBar
 * paint owns the component atomically; otherwise its local style overlays the
 * linked/numeric role. Keeping this decision outside the renderer makes image
 * prefetch and aggregate paint-work use the exact paint that reaches Canvas. */
export function chartStockBarFillDecision(
  chart: ChartModel,
  direct: ChartStockBarPaint,
  role: 'upBar' | 'downBar',
): Fill | null | undefined {
  if (direct.fillHidden === true) return null;
  if (direct.fill != null) return direct.fill;
  if (direct.fillColor != null) return { fillType: 'solid', color: direct.fillColor };
  const directStyleFill = chartStyleFillDecision(direct.style, 0);
  if (directStyleFill !== undefined) return directStyleFill;
  if (direct.fillPaintAuthored === true) return null;
  return chartStyleFillCascade(chart.chartStyleRoles?.[role], 0, direct.style);
}

export interface ChartThreeDSurfacePaint {
  fill: Fill | null | undefined;
  line: ChartModel['plotAreaLineFill'] | null | undefined;
  lineWidthEmu: number | null | undefined;
  lineDash: string | null | undefined;
  lineCustomDash: ChartModel['plotAreaLineCustomDash'];
  lineCap: string | null | undefined;
  lineJoin: string | null | undefined;
}

/** Resolve one authored CT_Surface against its dedicated linked Chart Style
 * role. Fill and outline paint are independent from line geometry; explicit
 * no-fill or unsupported authored paint suppresses lower paint without
 * discarding inherited width/dash/cap/join. */
export function chartThreeDSurfacePaint(
  chart: ChartModel,
  surface: NonNullable<ChartModel['threeD']>['floor'],
  role: 'floor' | 'wall',
): ChartThreeDSurfacePaint {
  const directStyle = surface?.style;
  const linkedStyle = chart.chartStyleRoles?.[role];
  let fill: Fill | null | undefined;
  if (surface?.fillHidden === true) fill = null;
  else if (surface?.fillColor) fill = { fillType: 'solid', color: surface.fillColor };
  else fill = chartStyleFillCascade(linkedStyle, 0, directStyle);

  let line: ChartModel['plotAreaLineFill'] | null | undefined;
  if (surface?.lineHidden === true) line = null;
  else if (surface?.lineColor) line = { fillType: 'solid', color: surface.lineColor };
  else line = chartStyleLineCascade(linkedStyle, 0, directStyle);
  // NoStyle is a paint sentinel. A linked entry may still contribute local
  // DrawingML geometry (width/dash/cap/join) beside its NoStyle lnRef.
  const linkedGeometry = linkedStyle;
  const directDash = surface?.lineDash != null
    ? { lineDash: surface.lineDash, lineDashAuthored: true }
    : undefined;
  const dash = chartStyleDashChoice(directDash, directStyle, linkedGeometry);
  return {
    fill,
    line,
    lineWidthEmu: surface?.lineWidthEmu
      ?? directStyle?.lineWidthEmu ?? linkedGeometry?.lineWidthEmu,
    lineDash: dash?.lineDash,
    lineCustomDash: dash?.lineCustomDash,
    lineCap: directStyle?.lineCap ?? linkedGeometry?.lineCap,
    lineJoin: directStyle?.lineJoin ?? linkedGeometry?.lineJoin,
  };
}

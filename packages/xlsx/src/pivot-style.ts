// Table and PivotTable style formatting (ECMA-376 §18.8.40-41, §18.5.1.5,
// §18.10.1.97).
//
// A Table style has the same shape: its elements apply in the Table style
// element order of §18.8.41 over the Table regions of §18.18.77 (whole table,
// column and row stripes over the data body, first and last column, header
// and total rows and their end cells). [MS-XLS] 2.5.176
// List12TableStyleClientInfo states the gating the XLSX tableStyleInfo
// booleans (§18.5.1.5) name: showFirstColumn / showLastColumn apply
// firstColumn / lastColumn, and the stripe flags apply their stripes; the
// header, total and corner-cell elements apply wherever those rows exist.
//
// A PivotTable style is a list of `tableStyleElement`s, each a differential
// format for one structured region of the PivotTable (§18.18.77, whose
// region diagrams define the areas below). Elements apply in the PivotTable
// style element order of §18.8.41; a later element wins per property. Each
// element is a differential format (§18.8.14-15, applied on top of what is
// already there), so a font toggle the element's dxf omits leaves the earlier
// value, and an explicit off (`<b val="0"/>`) turns an earlier on off. A
// region's border edges apply to its outline and its `horizontal` /
// `vertical` edges to its interior rules; an edge whose style is `none`
// (an explicitly cleared edge) overrides an earlier element's edge.
//
// Regions derive from the saved layout: the location offsets
// (§18.10.1.49), the row items (one per body row) and the column items
// (one per data column). Column subheadings need the column header layout,
// which the model does not carry; they are not drawn.
//
// Style option gating follows [MS-XLS] 2.4.273.107 SXAddl_SXCView_SXDTableStyleClient, which
// states the same Excel behavior the XLSX pivotTableStyleInfo booleans
// (§18.10.1.74) name: showRowHeaders applies firstColumn and the row
// subheadings; showColumnHeaders applies headerRow and the column
// subheadings; showLastColumn applies lastColumn; the stripe flags apply
// their stripes. firstHeaderCell, blank rows, subtotals and the grand total
// row are not gated.
import { DXF_FONT_TOGGLES, dxfFontToggle } from './dxf-font.js';
import {
  assertCoordinateRangeArea,
  setCoordinateIndexValue,
  type CoordinateIndexIdentity,
} from './renderer-coordinate-index.js';
import type {
  BorderEdge, CellFill, Dxf, PivotAxisItem, PivotTableMetadata, PivotTableStyleElement, TableInfo, Worksheet,
} from './types.js';
import { getWorksheetPolicy } from './worksheet-policy-context.js';

/** The merged Table / PivotTable style format of one cell. */
export interface StructuredCellFormat {
  fill?: CellFill;
  fontColor?: string;
  /** `fontColor` was set by a Table style layer (rich-text runs take only a
   *  Table style's color; see the renderer's `tableRunFontColor`). */
  fontColorFromTable?: boolean;
  /** The last layer naming a color named the automatic color; it applies over
   *  an inherited (not cell-owned) font color like any other style color. */
  fontColorAuto?: boolean;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  top?: BorderEdge;
  bottom?: BorderEdge;
  left?: BorderEdge;
  right?: BorderEdge;
}

/** §18.8.41 PivotTable Style Element Order (later wins). */
const ORDER = [
  'wholeTable',
  'pageFieldLabels',
  'pageFieldValues',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstRowStripe',
  'secondRowStripe',
  'firstColumn',
  'headerRow',
  'firstHeaderCell',
  'firstSubtotalColumn',
  'secondSubtotalColumn',
  'thirdSubtotalColumn',
  'blankRow',
  'firstSubtotalRow',
  'secondSubtotalRow',
  'thirdSubtotalRow',
  'firstColumnSubheading',
  'secondColumnSubheading',
  'thirdColumnSubheading',
  'firstRowSubheading',
  'secondRowSubheading',
  'thirdRowSubheading',
  'lastColumn',
  'totalRow',
] as const;

/** ST_ItemType values that are subtotals (§18.18.43). */
const SUBTOTALS = new Set([
  'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count',
  'stdDev', 'stdDevP', 'var', 'varP',
]);

interface Rect { top: number; bottom: number; left: number; right: number }

/**
 * The subtotal / subheading element level of an axis field at `depth`:
 * [MS-XLS] 2.4.321 TableStyleElement (tseType 0x10-0x12, 0x14-0x19) gives
 * the outermost field `first`, then alternates `second` (odd positions) and
 * `third` (even positions after the first) through the deeper fields.
 */
function level(depth: number): 'first' | 'second' | 'third' {
  if (depth === 0) return 'first';
  return depth % 2 === 1 ? 'second' : 'third';
}

/** Row bands cycling `first` (size a) and `second` (size b) rows. */
function bands(from: number, to: number, a: number, b: number): { first: Rect[]; second: Rect[] } {
  const out = { first: [] as Rect[], second: [] as Rect[] };
  let at = from;
  let odd = true;
  while (at <= to) {
    const size = Math.max(1, odd ? a : b);
    const end = Math.min(to, at + size - 1);
    (odd ? out.first : out.second).push({ top: at, bottom: end, left: 0, right: 0 });
    at = end + 1;
    odd = !odd;
  }
  return out;
}

/** The regions of every element kind for one PivotTable. */
function regions(p: PivotTableMetadata, sizes: Map<string, number>): Map<string, Rect[]> {
  const style = p.style!;
  const { top, bottom, left, right, firstDataRow, firstDataCol } = p.location;
  const bodyTop = top + firstDataRow;
  const dataLeft = left + firstDataCol;
  const out = new Map<string, Rect[]>();
  const add = (kind: string, rect: Rect) => {
    if (rect.top > rect.bottom || rect.left > rect.right) return;
    const list = out.get(kind);
    if (list) list.push(rect);
    else out.set(kind, [rect]);
  };
  add('wholeTable', { top, bottom, left, right });
  if (style.showColumnStripes) {
    const cols = bands(dataLeft, right, sizes.get('firstColumnStripe') ?? 1, sizes.get('secondColumnStripe') ?? 1);
    for (const band of cols.first) add('firstColumnStripe', { top: bodyTop, bottom, left: band.top, right: band.bottom });
    for (const band of cols.second) add('secondColumnStripe', { top: bodyTop, bottom, left: band.top, right: band.bottom });
  }
  if (style.showRowStripes) {
    const rows = bands(bodyTop, bottom, sizes.get('firstRowStripe') ?? 1, sizes.get('secondRowStripe') ?? 1);
    for (const band of rows.first) add('firstRowStripe', { ...band, left, right });
    for (const band of rows.second) add('secondRowStripe', { ...band, left, right });
  }
  if (style.showRowHeaders) add('firstColumn', { top, bottom, left, right: dataLeft - 1 });
  if (style.showColumnHeaders) add('headerRow', { top, bottom: bodyTop - 1, left, right });
  add('firstHeaderCell', { top, bottom: bodyTop - 2, left, right: dataLeft - 1 });
  const leafLevel = Math.max(0, p.rowFields.length - 1);
  (p.rowItems ?? []).forEach((item: PivotAxisItem, index) => {
    const row = bodyTop + index;
    if (row > bottom) return;
    const rect = { top: row, bottom: row, left, right };
    const lvl = level(item.depth);
    if (item.kind === 'blank') add('blankRow', rect);
    else if (item.kind === 'grand') add('totalRow', rect);
    else if (SUBTOTALS.has(item.kind)) add(`${lvl}SubtotalRow`, rect);
    else if (item.kind === 'data' && item.depth < leafLevel && style.showRowHeaders) {
      add(`${lvl}RowSubheading`, rect);
    }
  });
  (p.columnItems ?? []).forEach((item: PivotAxisItem, index) => {
    const col = dataLeft + index;
    if (col > right) return;
    const rect = { top, bottom, left: col, right: col };
    const lvl = level(item.depth);
    if (item.kind === 'grand') {
      if (style.showLastColumn) add('lastColumn', rect);
    } else if (SUBTOTALS.has(item.kind)) add(`${lvl}SubtotalColumn`, rect);
  });
  return out;
}

/** A Table / PivotTable style element's solid fill shows its bgColor: Excel's
 *  export of a table style whose firstRowStripe dxf carries a different
 *  fgColor paints the bgColor (sample-1), unlike a cell fill (§18.8.32). */
function styleElementFill(fill: CellFill): CellFill {
  return fill.patternType === 'solid' && fill.bgColor && fill.bgColor !== fill.fgColor
    ? { ...fill, fgColor: fill.bgColor }
    : fill;
}

function apply(
  target: StructuredCellFormat, dxf: Dxf, rect: Rect, row: number, col: number, fromTable: boolean,
): void {
  if (dxf.fill) target.fill = styleElementFill(dxf.fill);
  const font = dxf.font;
  if (font) {
    if (font.color) {
      target.fontColor = font.color;
      delete target.fontColorAuto;
      if (fromTable) target.fontColorFromTable = true;
      else delete target.fontColorFromTable;
    } else if (dxf.fontToggles?.automaticColor) {
      // An explicit automatic color replaces an earlier element's color.
      delete target.fontColor;
      target.fontColorAuto = true;
      if (fromTable) target.fontColorFromTable = true;
      else delete target.fontColorFromTable;
    }
    for (const key of DXF_FONT_TOGGLES) {
      const value = dxfFontToggle(dxf, key);
      if (value !== undefined) target[key] = value;
    }
  }
  const border = dxf.border;
  if (!border) return;
  const edge = (outer: BorderEdge | null | undefined, inner: BorderEdge | null | undefined, onOutline: boolean) =>
    onOutline ? outer : inner;
  const top = edge(border.top, border.horizontal, row === rect.top);
  const bottom = edge(border.bottom, border.horizontal, row === rect.bottom);
  const left = edge(border.left, border.vertical, col === rect.left);
  const right = edge(border.right, border.vertical, col === rect.right);
  if (top) target.top = top;
  if (bottom) target.bottom = bottom;
  if (left) target.left = left;
  if (right) target.right = right;
}

/** §18.8.41 Table Style Element Order (later wins). */
const TABLE_ORDER = [
  'wholeTable',
  'firstColumnStripe',
  'secondColumnStripe',
  'firstRowStripe',
  'secondRowStripe',
  'lastColumn',
  'firstColumn',
  'headerRow',
  'totalRow',
  'firstHeaderCell',
  'lastHeaderCell',
  'firstTotalCell',
  'lastTotalCell',
] as const;

/** Stripe bands over the visible lines `from..to`, cycling `first` (size a)
 *  and `second` (size b) visible lines; hidden lines take no band. Each band
 *  spans its first to last line (a band's outline edges are its region's). */
function visibleBands(
  from: number, to: number, a: number, b: number, hidden: (line: number) => boolean,
): { first: Rect[]; second: Rect[] } {
  const out = { first: [] as Rect[], second: [] as Rect[] };
  let odd = true;
  let left = Math.max(1, a);
  let band: Rect | null = null;
  for (let line = from; line <= to; line++) {
    if (hidden(line)) continue;
    if (!band) {
      band = { top: line, bottom: line, left: 0, right: 0 };
      (odd ? out.first : out.second).push(band);
    }
    band.bottom = line;
    left -= 1;
    if (left === 0) {
      odd = !odd;
      left = Math.max(1, odd ? a : b);
      band = null;
    }
  }
  return out;
}

/** The regions of every element kind for one Table (§18.18.77). Row stripes
 *  alternate over the visible body rows: Excel re-bands a table around hidden
 *  rows (a hidden first data row leaves the first stripe on the next visible
 *  row in the sample-5 Excel export). Column stripes keep their positions;
 *  no export has measured hidden columns. */
function tableRegions(
  t: TableInfo,
  sizes: Map<string, number>,
  hiddenRow: (row: number) => boolean,
): Map<string, Rect[]> {
  const { top, bottom, left, right } = t.range;
  const header = Math.max(0, t.headerRowCount ?? 1);
  const totals = Math.max(0, t.totalsRowCount ?? 0);
  const bodyTop = top + header;
  const bodyBottom = bottom - totals;
  const out = new Map<string, Rect[]>();
  const add = (kind: string, rect: Rect) => {
    if (rect.top > rect.bottom || rect.left > rect.right) return;
    const list = out.get(kind);
    if (list) list.push(rect);
    else out.set(kind, [rect]);
  };
  add('wholeTable', { top, bottom, left, right });
  if (t.showColumnStripes) {
    const cols = bands(left, right, sizes.get('firstColumnStripe') ?? 1, sizes.get('secondColumnStripe') ?? 1);
    for (const band of cols.first) add('firstColumnStripe', { top: bodyTop, bottom: bodyBottom, left: band.top, right: band.bottom });
    for (const band of cols.second) add('secondColumnStripe', { top: bodyTop, bottom: bodyBottom, left: band.top, right: band.bottom });
  }
  if (t.showRowStripes) {
    const rows = visibleBands(bodyTop, bodyBottom, sizes.get('firstRowStripe') ?? 1, sizes.get('secondRowStripe') ?? 1, hiddenRow);
    for (const band of rows.first) add('firstRowStripe', { ...band, left, right });
    for (const band of rows.second) add('secondRowStripe', { ...band, left, right });
  }
  if (t.showLastColumn) add('lastColumn', { top, bottom, left: right, right });
  if (t.showFirstColumn) add('firstColumn', { top, bottom, left, right: left });
  if (header > 0) {
    add('headerRow', { top, bottom: bodyTop - 1, left, right });
    add('firstHeaderCell', { top, bottom: bodyTop - 1, left, right: left });
    add('lastHeaderCell', { top, bottom: bodyTop - 1, left: right, right });
  }
  if (totals > 0) {
    add('totalRow', { top: bodyBottom + 1, bottom, left, right });
    add('firstTotalCell', { top: bodyBottom + 1, bottom, left, right: left });
    add('lastTotalCell', { top: bodyBottom + 1, bottom, left: right, right });
  }
  return out;
}

function layer(
  map: Map<string, StructuredCellFormat>,
  identity: CoordinateIndexIdentity,
  order: readonly string[],
  elementList: readonly PivotTableStyleElement[],
  areas: Map<string, Rect[]>,
  fromTable: boolean,
): void {
  const elements = new Map(elementList.map((element) => [element.kind, element]));
  for (const kind of order) {
    const element = elements.get(kind);
    const rects = areas.get(kind);
    if (!element || !rects) continue;
    for (const rect of rects) {
      for (let row = rect.top; row <= rect.bottom; row++) {
        for (let col = rect.left; col <= rect.right; col++) {
          const key = `${row}:${col}`;
          let format = map.get(key);
          if (!format) {
            format = {};
            setCoordinateIndexValue(map, key, format, identity);
          }
          apply(format, element.dxf, rect, row, col, fromTable);
        }
      }
    }
  }
}

/** A model without `styleElements` (a producer predating them) names its
 *  custom style's region elements by workbook dxf index; resolve those. A
 *  style that could not be applied as authored (`styleError`) stays off. */
function legacyTableElements(t: TableInfo, dxfs: readonly Dxf[] | undefined): PivotTableStyleElement[] {
  if (!dxfs || t.styleError) return [];
  const ids: Array<[string, number | undefined]> = [
    ['wholeTable', t.wholeTableDxf], ['headerRow', t.headerRowDxf], ['totalRow', t.totalRowDxf],
    ['firstColumn', t.firstColumnDxf], ['lastColumn', t.lastColumnDxf],
    ['firstRowStripe', t.band1HorizontalDxf], ['secondRowStripe', t.band2HorizontalDxf],
  ];
  return ids.flatMap(([kind, id]) => {
    const dxf = id === undefined ? undefined : dxfs[id];
    return dxf ? [{ kind, size: 1, dxf }] : [];
  });
}

/** Merged Table and PivotTable style formats by `row:col`. `dxfs` (the
 *  workbook's) resolves tables from producers predating `styleElements`. */
export function buildStructuredStyleMap(
  worksheet: Worksheet,
  dxfs?: readonly Dxf[],
): Map<string, StructuredCellFormat> {
  const map = new Map<string, StructuredCellFormat>();
  const identity: CoordinateIndexIdentity = {
    resource: 'worksheet-structured-style-index',
    operation: 'expand-structured-style-coordinates',
    limit: getWorksheetPolicy(worksheet).maxCoordinateIndexEntries,
  };
  const hiddenRows = new Set((worksheet.rows ?? []).filter((row) => row.hidden).map((row) => row.index));
  const hiddenRow = (row: number) => hiddenRows.has(row);
  for (const table of worksheet.tables ?? []) {
    const elements = table.styleElements ?? legacyTableElements(table, dxfs);
    if (elements.length === 0) continue;
    assertCoordinateRangeArea(table.range, identity);
    const sizes = new Map(elements.map((element) => [element.kind, element.size]));
    layer(map, identity, TABLE_ORDER, elements, tableRegions(table, sizes, hiddenRow), true);
  }
  for (const pivot of worksheet.pivotTables ?? []) {
    const style = pivot.style;
    if (!style || style.elements.length === 0) continue;
    const sizes = new Map(style.elements.map((element) => [element.kind, element.size]));
    assertCoordinateRangeArea(pivot.location, identity);
    layer(map, identity, ORDER, style.elements, regions(pivot, sizes), false);
  }
  return map;
}

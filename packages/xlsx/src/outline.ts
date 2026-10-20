/**
 * XL4 — row/column outline (grouping) geometry.
 *
 * Pure, DOM-free math that turns a worksheet's per-band outline metadata
 * (ECMA-376 §18.3.1.13 `<col>`, §18.3.1.73 `<row>`, §18.3.1.61
 * `<sheetPr><outlinePr>`) into a set of *group brackets* and *collapse toggles*
 * that the viewer draws in the outline gutter, plus the state changes a click on
 * a level button or a +/- toggle applies to the in-memory model.
 *
 * The geometry is deliberately separated from rendering so the level/summary
 * placement rules — which side of the group the summary band (and its +/-
 * button) sits on, per `summaryBelow` / `summaryRight` — are unit-testable
 * without a canvas.
 *
 * # Model
 *
 * Each band (row or column) carries an `outlineLevel` in 0..7. A *group* at
 * level `L` is a maximal run of consecutive bands whose level is `>= L`. The
 * bracket for that group is drawn in the gutter lane for level `L`, spanning the
 * detail run. Its collapse toggle sits on the adjacent *summary* band at level
 * `L-1`: the band immediately AFTER the run when the summary is below/right (the
 * default), or immediately BEFORE it when above/left.
 *
 * A summary band's `collapsed` flag (`<row collapsed>` / `<col collapsed>`) is
 * `true` when its one-level-deeper detail is hidden, so the toggle renders `+`
 * (expand) when collapsed and `-` (collapse) when expanded.
 */

/** Which axis a set of outline metadata describes. */
export type OutlineAxis = 'row' | 'col';

/** Width (row axis) / height (col axis) of one outline lane, in unscaled CSS px.
 *  Excel draws each nesting level in its own ~19px lane. */
export const OUTLINE_LANE_PX = 19;

/** Side of the level-number button bank / corner (unscaled CSS px). Holds the
 *  numbered "1 2 3" buttons that collapse the whole sheet to a level. */
export const OUTLINE_BUTTON_PX = 12;

/** Gap between adjacent numbered outline-level buttons. The level bank is
 *  independent from the wider bracket lanes: Excel packs these buttons into a
 *  compact strip instead of centering them in the 19px rail lanes. */
export const OUTLINE_BUTTON_GAP_PX = 2;

/** Leading inset of the numbered level-button bank. */
const OUTLINE_BUTTON_INSET_PX = 2;

/** Center coordinate of a numbered outline-level button within its bank. */
export function outlineLevelButtonCenterPx(level: number): number {
  return OUTLINE_BUTTON_INSET_PX
    + OUTLINE_BUTTON_PX / 2
    + (level - 1) * (OUTLINE_BUTTON_PX + OUTLINE_BUTTON_GAP_PX);
}

/** Extent (px) of the gutter for an axis with `maxLevel` nesting levels, or 0
 *  when the axis has no outlining. The lanes hold the group brackets; one extra
 *  lane holds the numbered level buttons (levels 1..maxLevel+1). */
export function gutterExtentPx(maxLevel: number): number {
  return maxLevel > 0 ? (maxLevel + 1) * OUTLINE_LANE_PX : 0;
}

/** Per-band outline metadata, indexed by 1-based band index. */
export interface BandOutline {
  /** 1-based band index (row number / column number). */
  index: number;
  /** Outline depth 0..7. */
  level: number;
  /** `collapsed` flag from the band (set on a summary band). */
  collapsed: boolean;
  /** Whether the band is currently hidden (collapsed-detail or user-hidden). */
  hidden: boolean;
}

/** A drawn group bracket: a bar in the gutter lane `level` spanning `[start,end]`
 *  detail bands, with an optional summary band that carries the +/- toggle. */
export interface OutlineGroup {
  /** Outline lane this bracket occupies (1-based; lane 1 is the outermost). */
  level: number;
  /** First detail band (inclusive, 1-based). */
  start: number;
  /** Last detail band (inclusive, 1-based). */
  end: number;
  /** The summary band index carrying the toggle, or `null` when the group has
   *  no summary band on the expected side (e.g. the run touches the sheet edge).
   *  When present the +/- button is drawn at this band. */
  summary: number | null;
  /** `true` when the summary band is collapsed ⇒ the toggle shows `+`. */
  collapsed: boolean;
}

/** The complete outline layout for one axis. */
export interface OutlineLayout {
  /** Maximum outline level present (0 ⇒ no outlining on this axis). */
  maxLevel: number;
  /** All group brackets, in ascending `(level, start)` order. */
  groups: OutlineGroup[];
}

/** One canvas segment of an outline bracket. Kept in the pure geometry layer
 * so the start hook cannot accidentally become a diagonal or migrate to the
 * summary end when row/column rendering changes. */
export interface OutlineBracketSegment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

/** Visible pane rectangle for one outline detail/summary run. Outline gutter
 * geometry shares the worksheet header + frozen-pane coordinate system, so a
 * scrolled run must be clipped at the frozen boundary instead of leaking into
 * the column/row header. Fractions are already resolved to CSS pixels by the
 * viewer before calling this helper. */
export function outlinePaneClipRect(
  axis: OutlineAxis,
  start: number,
  end: number,
  frozenBandCount: number,
  headerExtentPx: number,
  frozenExtentPx: number,
  canvasWidth: number,
  canvasHeight: number,
  rtl = false,
): { x: number; y: number; w: number; h: number } {
  if (axis === 'row') {
    const boundary = Math.min(canvasHeight, headerExtentPx + frozenExtentPx);
    let top = Math.min(canvasHeight, headerExtentPx);
    let bottom = canvasHeight;
    if (frozenBandCount > 0 && end <= frozenBandCount) bottom = boundary;
    else if (frozenBandCount > 0 && start > frozenBandCount) top = boundary;
    return { x: 0, y: top, w: canvasWidth, h: Math.max(0, bottom - top) };
  }

  if (!rtl) {
    const boundary = Math.min(canvasWidth, headerExtentPx + frozenExtentPx);
    let left = Math.min(canvasWidth, headerExtentPx);
    let right = canvasWidth;
    if (frozenBandCount > 0 && end <= frozenBandCount) right = boundary;
    else if (frozenBandCount > 0 && start > frozenBandCount) left = boundary;
    return { x: left, y: 0, w: Math.max(0, right - left), h: canvasHeight };
  }

  const headerLeft = Math.max(0, canvasWidth - headerExtentPx);
  const boundary = Math.max(0, headerLeft - frozenExtentPx);
  let left = 0;
  let right = headerLeft;
  if (frozenBandCount > 0 && end <= frozenBandCount) left = boundary;
  else if (frozenBandCount > 0 && start > frozenBandCount) right = boundary;
  return { x: left, y: 0, w: Math.max(0, right - left), h: canvasHeight };
}

/** Build the expanded-group bracket. Excel hooks the bracket toward the grid at
 * the FIRST detail band; the summary band at the opposite end carries the
 * minus button and therefore has no second hook. `detailStart` is the on-screen
 * coordinate of the group's first logical detail band (important for RTL
 * columns), while the long rail spans both detail endpoints. */
export function outlineBracketSegments(
  axis: OutlineAxis,
  laneCenter: number,
  detailStart: number,
  detailEnd: number,
  laneExtent: number,
): OutlineBracketSegment[] {
  const low = Math.min(detailStart, detailEnd);
  const high = Math.max(detailStart, detailEnd);
  if (axis === 'row') {
    return [
      { x1: laneCenter, y1: low, x2: laneCenter, y2: high },
      { x1: laneCenter, y1: detailStart, x2: laneCenter + laneExtent / 2, y2: detailStart },
    ];
  }
  return [
    { x1: low, y1: laneCenter, x2: high, y2: laneCenter },
    { x1: detailStart, y1: laneCenter, x2: detailStart, y2: laneCenter + laneExtent / 2 },
  ];
}

/**
 * Build the outline layout for one axis from its band metadata.
 *
 * `bands` need only contain bands with a non-zero level or a set `collapsed`
 * flag; every other band is treated as level 0 (ungrouped). `summaryBelow` is
 * the `<outlinePr>` flag for this axis (`summaryBelow` for rows, `summaryRight`
 * for columns) — `true` (default) puts the summary band after the detail run.
 */
export function buildOutlineLayout(bands: BandOutline[], summaryAfter: boolean): OutlineLayout {
  // Densify into a level lookup so gaps between recorded bands read as level 0.
  const levelOf = new Map<number, number>();
  const collapsedOf = new Map<number, boolean>();
  let maxIndex = 0;
  let maxLevel = 0;
  for (const b of bands) {
    if (b.level > 0) levelOf.set(b.index, b.level);
    if (b.collapsed) collapsedOf.set(b.index, true);
    if (b.index > maxIndex) maxIndex = b.index;
    if (b.level > maxLevel) maxLevel = b.level;
  }
  // A `collapsed` summary band can sit one past the deepest levelled band, so
  // scan one extra index for the summary lookup.
  const scanMax = maxIndex;

  const groups: OutlineGroup[] = [];
  if (maxLevel === 0) return { maxLevel: 0, groups };

  const lvl = (i: number) => levelOf.get(i) ?? 0;

  // For each lane L (1..maxLevel) find maximal runs of consecutive bands with
  // level >= L, then attach the summary band on the configured side.
  for (let L = 1; L <= maxLevel; L++) {
    let run: { start: number; end: number } | null = null;
    for (let i = 1; i <= scanMax + 1; i++) {
      const inRun = lvl(i) >= L;
      if (inRun) {
        if (run) run.end = i;
        else run = { start: i, end: i };
      } else if (run) {
        groups.push(makeGroup(L, run, summaryAfter, collapsedOf, lvl));
        run = null;
      }
    }
    if (run) groups.push(makeGroup(L, run, summaryAfter, collapsedOf, lvl));
  }
  return { maxLevel, groups };
}

function makeGroup(
  level: number,
  run: { start: number; end: number },
  summaryAfter: boolean,
  collapsedOf: Map<number, boolean>,
  lvl: (i: number) => number,
): OutlineGroup {
  // The summary band sits at level `level - 1`, immediately after the run
  // (summaryAfter) or immediately before it. It only qualifies as this group's
  // summary when it exists at the shallower level (i.e. is NOT part of the run).
  let summary: number | null = null;
  if (summaryAfter) {
    const cand = run.end + 1;
    if (cand >= 1 && lvl(cand) < level) summary = cand;
  } else {
    const cand = run.start - 1;
    if (cand >= 1 && lvl(cand) < level) summary = cand;
  }
  const collapsed = summary != null && (collapsedOf.get(summary) ?? false);
  return { level, start: run.start, end: run.end, summary, collapsed };
}

/**
 * Toggle a single group's collapse state, returning the set of detail band
 * indices whose hidden state must flip and the new `collapsed` value for the
 * summary band. This is the state a +/- button click produces; the viewer
 * applies it to the in-memory worksheet (never the file).
 *
 * Collapsing hides every band in `[start,end]` whose level is `>= group.level`.
 * Expanding reveals the direct children but keeps deeper nested groups that are
 * themselves collapsed hidden — Excel restores only one level per click. To keep
 * a viewer's behaviour predictable and reversible we reveal a child band unless
 * it is shadowed by a still-collapsed deeper summary inside this group.
 */
export function toggleGroupHidden(
  group: OutlineGroup,
  bands: BandOutline[],
): { hide: number[]; show: number[]; nowCollapsed: boolean } {
  const nowCollapsed = !group.collapsed;
  const byIndex = new Map<number, BandOutline>();
  for (const b of bands) byIndex.set(b.index, b);

  const hide: number[] = [];
  const show: number[] = [];
  if (nowCollapsed) {
    // Collapse: hide the whole detail run.
    for (let i = group.start; i <= group.end; i++) hide.push(i);
  } else {
    // Expand: reveal bands at exactly this group's level; leave bands belonging
    // to a still-collapsed deeper subgroup hidden.
    const collapsedSummaries = new Set<number>();
    for (const b of bands) {
      if (b.index >= group.start && b.index <= group.end && b.collapsed) {
        collapsedSummaries.add(b.index);
      }
    }
    for (let i = group.start; i <= group.end; i++) {
      if (isShadowedByCollapsedChild(i, group, byIndex, collapsedSummaries)) continue;
      show.push(i);
    }
  }
  return { hide, show, nowCollapsed };
}

/** Whether revealing band `i` should stay hidden because a deeper collapsed
 *  summary inside the group shadows it (i.e. `i` is detail of a nested group
 *  whose summary is still collapsed). */
function isShadowedByCollapsedChild(
  i: number,
  group: OutlineGroup,
  byIndex: Map<number, BandOutline>,
  collapsedSummaries: Set<number>,
): boolean {
  const band = byIndex.get(i);
  const level = band?.level ?? 0;
  // Bands at this group's own summary level are always revealed.
  if (level <= group.level) return false;
  // A deeper band is shadowed when a collapsed summary at a shallower-but-deeper
  // level than it sits adjacent within the run. We approximate with: any
  // collapsed summary strictly shallower than `level` covering `i`.
  for (const s of collapsedSummaries) {
    const sLevel = byIndex.get(s)?.level ?? 0;
    if (sLevel >= level) continue;
    if (sLevel < group.level) continue;
    // Summary below its detail: detail is at indices < s (and > previous summary).
    // Since exact run boundaries are not tracked here, treat any deeper band as
    // shadowed by the nearest collapsed summary at a shallower level.
    return true;
  }
  return false;
}

/**
 * Collect every band index that a "collapse to level N" (the numbered level
 * buttons at the corner of the gutter) hides or shows. Clicking level button `n`
 * shows all bands with level < `n` and hides all bands with level >= `n`.
 * Returns the target hidden state per affected band index.
 */
export function levelButtonHidden(
  bands: BandOutline[],
  targetLevel: number,
): { hide: number[]; show: number[] } {
  const hide: number[] = [];
  const show: number[] = [];
  for (const b of bands) {
    if (b.level >= targetLevel) hide.push(b.index);
    else show.push(b.index);
  }
  return { hide, show };
}

/** Minimal worksheet shape the axis extractors read (kept structural so tests
 *  can pass a plain object). */
export interface OutlineWorksheetLike {
  rows: { index: number; outlineLevel?: number; collapsed?: boolean; hidden?: boolean }[];
  colOutlineLevels?: Record<number, number>;
  colCollapsed?: Record<number, boolean>;
  colHidden?: Record<number, boolean>;
  outlinePr?: { summaryBelow: boolean; summaryRight: boolean };
}

/** Row-axis band metadata (only bands with a non-zero level or a set collapsed
 *  flag are surfaced). Keyed off the parser's per-row fields. */
export function rowBands(ws: OutlineWorksheetLike): BandOutline[] {
  const out: BandOutline[] = [];
  for (const r of ws.rows) {
    const level = r.outlineLevel ?? 0;
    const collapsed = r.collapsed ?? false;
    if (level === 0 && !collapsed) continue;
    out.push({ index: r.index, level, collapsed, hidden: r.hidden ?? false });
  }
  return out;
}

/** Column-axis band metadata, from the parallel `colOutlineLevels` /
 *  `colCollapsed` / `colHidden` maps. */
export function colBands(ws: OutlineWorksheetLike): BandOutline[] {
  const levels = ws.colOutlineLevels ?? {};
  const collapsed = ws.colCollapsed ?? {};
  const hidden = ws.colHidden ?? {};
  const indices = new Set<number>();
  for (const k of Object.keys(levels)) indices.add(Number(k));
  for (const k of Object.keys(collapsed)) indices.add(Number(k));
  const out: BandOutline[] = [];
  for (const i of [...indices].sort((a, b) => a - b)) {
    out.push({
      index: i,
      level: levels[i] ?? 0,
      collapsed: collapsed[i] ?? false,
      hidden: hidden[i] ?? false,
    });
  }
  return out;
}

/** `summaryBelow` for rows / `summaryRight` for columns, defaulting to `true`
 *  (ECMA-376 §18.3.1.61) when the sheet declares no `<outlinePr>`. */
export function summaryAfterFor(ws: OutlineWorksheetLike, axis: OutlineAxis): boolean {
  const pr = ws.outlinePr;
  if (!pr) return true;
  return axis === 'row' ? pr.summaryBelow : pr.summaryRight;
}

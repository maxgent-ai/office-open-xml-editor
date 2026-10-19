// Floating-table placement geometry (ECMA-376 §17.4.57 `<w:tblpPr>` /
// §17.4.56 `<w:tblOverlap>`).
//
// Pure placement math: given a `<w:tblpPr>` and the section geometry on
// `FloatRegistrationState`, resolve the table box in points, decide which side body text
// wraps on, and push the wrap-exclusion FloatRect onto `state.floats`. The
// anchor / alignment semantics line up 1:1 with a `<w:framePr>` text frame, so
// this module reuses frame-geometry's frameXContainer / frameYContainer /
// resolveAlignedPosH / resolveAlignedPosV and the shared pushFloatRect builder.
// Extracted from renderer.ts so the placement logic can be unit-reasoned in
// isolation (see float-table-geometry.test.ts / measure-column-geometry.test.ts).

import type { TblpPr } from './types.js';
import type { FloatRegistrationState } from './layout/acquisition-context.js';
import {
  FLOAT_OVERLAP_EPS,
  FLOAT_PAGE_RIGHT_SLACK,
  floatRectParticipant,
  floatingTableAvoidance,
  resolveFloatPlacement,
  type FloatPlacementParticipant,
} from './layout/floats.js';
import { resolvePointSpaceFloatingTableBoxPt } from './layout/floating-table-transaction.js';
export {
  resolveFloatingTableBoxPt,
  resolveFloatingTablePlacement,
} from './layout/floating-table-transaction.js';
import {
  frameXContainer,
  pushFloatRect,
} from './frame-geometry.js';

/** Resolved point-space placement of a floating table (`<w:tblpPr>`,
 *  ECMA-376 §17.4.57). `w`/`h` are the rendered table extent (sum of column
 *  widths × row heights). Exported for unit tests only — not package API. */
export interface FloatTableBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Resolve a floating table's top-left placement in points from its
 * `<w:tblpPr>` (ECMA-376 §17.4.57). `tableW`/`tableH` are the already-laid-out
 * table extent in points. The anchor / alignment semantics line up 1:1 with a
 * `<w:framePr>` text frame (horzAnchor↔hAnchor, vertAnchor↔vAnchor,
 * tblpXSpec↔xAlign, tblpYSpec↔yAlign, tblpX/tblpY↔x/y), so this mirrors
 * {@link computeFrameBox}'s placement math exactly — `frameXContainer` /
 * `frameYContainer` give the same per-anchor bands (text→column band,
 * margin→page content margin, page→physical edges), which is the #513
 * column-integrity guarantee.
 *
 * Exported for unit tests only (the float-table-geometry table) — not package API.
 */
export function computeFloatTableBox(
  tp: TblpPr,
  state: FloatRegistrationState,
  paraTop: number,
  tableW: number,
  tableH: number,
  /** When true, skip the renderer's implementation-defined
   *  vertAnchor=page/margin bottom clamp so the raw absolute box is returned.
   *  The paginator
   *  uses this to find where a page/margin-anchored table that overflows the text
   *  region must be row-split: the raw tblpY top drives slice 1's
   *  position, and clamping (which pins a too-tall box to the container top) would
   *  hide the overflow the split is meant to resolve. Placement (paint) keeps the
   *  clamp. Ignored for vertAnchor="text" (never clamped). */
  skipVClamp = false,
  /** Resolve against the current external registry before retained descendants
   *  are committed. The returned box may later be registered pre-resolved. */
  overlapResolution?: Readonly<{ allowOverlap: boolean }>,
): FloatTableBox {
  const textBand = frameXContainer('text', state);
  const box = resolvePointSpaceFloatingTableBoxPt({
    leftFromTextPt: tp.leftFromText,
    rightFromTextPt: tp.rightFromText,
    topFromTextPt: tp.topFromText,
    bottomFromTextPt: tp.bottomFromText,
    horzAnchor: tp.horzAnchor,
    horzSpecified: tp.horzSpecified,
    vertAnchor: tp.vertAnchor,
    xPt: tp.tblpX,
    yPt: tp.tblpY,
    ...(tp.tblpXSpec == null ? {} : { xAlign: tp.tblpXSpec }),
    ...(tp.tblpYSpec == null ? {} : { yAlign: tp.tblpYSpec }),
  }, {
    page: {
      xPt: 0,
      yPt: 0,
      widthPt: state.pageWidth,
      heightPt: state.pageH,
    },
    margin: {
      xPt: state.marginLeft,
      yPt: state.marginTop,
      widthPt: Math.max(0, state.pageWidth - state.marginLeft - state.marginRight),
      heightPt: Math.max(0, state.pageH - state.marginTop - state.marginBottom),
    },
    text: {
      xPt: textBand.left,
      yPt: paraTop,
      widthPt: Math.max(0, textBand.right - textBand.left),
      heightPt: tableH,
    },
  }, tableW, tableH, skipVClamp);
  if (!overlapResolution || box.w <= 0 || box.h <= 0) return box;
  const tableOverlap = overlapResolution.allowOverlap ? 'overlap' : 'never';
  const moving: FloatPlacementParticipant = {
    occurrenceId: 'display-moving-table',
    kind: 'table',
    tableOverlap,
    paragraphId: state.floatParaSeq,
    bounds: {
      xPt: box.x,
      yPt: box.y,
      widthPt: box.w,
      heightPt: box.h,
    },
    exclusionBounds: {
      xPt: box.x - tp.leftFromText,
      yPt: box.y - tp.topFromText,
      widthPt: box.w + tp.leftFromText + tp.rightFromText,
      heightPt: box.h + tp.topFromText + tp.bottomFromText,
    },
  };
  const resolved = resolveFloatPlacement({
    moving,
    blockers: state.floats.map(floatRectParticipant),
    avoidance: floatingTableAvoidance(),
    rightBoundaryPt: state.pageWidth,
    overlapEpsilonPt: FLOAT_OVERLAP_EPS,
    rightBoundarySlackPt: FLOAT_PAGE_RIGHT_SLACK,
  });
  return { ...box, x: resolved.bounds.xPt, y: resolved.bounds.yPt };
}

/**
 * Push the wrap-exclusion FloatRect for a resolved floating-table box onto
 * `state.floats` so following body text flows around the table (§17.4.57). The
 * exclusion is the table box padded by the *FromText dist values. Overlap
 * avoidance (§17.4.56) runs FIRST (mirroring registerShapeFloat: resolve x/y,
 * THEN build the exclusion from the resolved x/y) with allowOverlap =
 * `table.overlap !== 'never'` (default true ⇒ overlap permitted).
 *
 * `side` (which side text wraps on) is computed by the caller from the resolved
 * box vs the column band: the table sits to one side and text fills the other
 * (§17.4.57). The x-range is built from the box (which for horzAnchor="text" is
 * column-relative via frameXContainer), so resolveLineFloatWindow only
 * constrains the matching column (#513), consistent with registerFrameFloat.
 *
 * Exported for unit tests only (the float-table-geometry table) — not package API.
 */
export function registerTableFloat(
  box: FloatTableBox,
  tp: TblpPr,
  state: FloatRegistrationState,
  side: string,
  allowOverlap: boolean,
  /** The box was already resolved against the pre-descendant registry. */
  preResolved = false,
): void {
  if (box.w <= 0 || box.h <= 0) return;
  const dl = tp.leftFromText;
  const dr = tp.rightFromText;
  const dt = tp.topFromText;
  const db = tp.bottomFromText;
  const paraId = state.floatParaSeq++;

  // §17.4.56: resolve overlap against already-registered page floats before
  // fixing the exclusion rect. allowOverlap=false (tblOverlap="never") forces
  // avoidance of OTHER FLOATING TABLES only — §17.4.56 scopes "never" to other
  // floating tables, NOT DrawingML anchors (§20.4.2.3) or text frames. The
  // typed table participant makes resolveFloatPlacement limit normative
  // blockers to tables and retain this fact for tables placed later.
  // allowOverlap=true only avoids OTHER paragraphs' floats (the implementation-
  // defined scope named in compatibility.ts).
  pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    dl, dr, dt, db,
    kind: 'table',
    tableOverlap: allowOverlap ? 'overlap' : 'never',
    mode: 'square',
    side,
    imageKey: '', // non-image float: retained table paint owns it.
    paraId,
    avoidOverlap: !preResolved,
  });
}

/**
 * `tblpPr` supplies an exclusion rectangle but no left/right wrap-side choice.
 * Preserve both candidate flanks so `resolveLineFloatWindow` can apply the same
 * widest-free-gap geometry used for other square exclusions. Choosing a side
 * from the column midpoint discarded usable space and had no OOXML basis.
 *
 * Exported for unit tests only — not package API.
 */
export function floatTableWrapSide(_box: FloatTableBox, _state: FloatRegistrationState): string {
  return 'bothSides';
}

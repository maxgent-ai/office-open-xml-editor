import type { BorderSpec, CellBorders, TableBorders } from './types';
import { hex6ToRgb } from '@silurus/ooxml-core';
import {
  WORD_TABLE_BORDER_STYLE_PRECEDENCE,
  wordAuthoredBorderParticipates,
  wordTableBorderWeight,
} from './layout/table-compatibility.js';

/**
 * ECMA-376 §17.4.66 (tcBorders) — adjacent table cell border conflict resolution.
 *
 * When cell spacing is zero, two cells that share an interior gridline each
 * contribute a border for that edge. This module displays exactly ONE of them,
 * chosen by the following rules (applied in order), and is the pure kernel that
 * decides the winner. Canonical table border acquisition supplies the two
 * candidates for each shared edge and retains only the returned winner, so a
 * gridline is painted once with the correct spec (no more "last cell wins").
 *
 * Rules (ECMA-376 §17.4.66 plus the registered table-border weight
 * compatibility rule):
 *   0. `nil` and `none` lose to the opposing border.
 *   1. A CELL border always beats a TABLE(-level or table-style) border.
 *   2. `word-table-border-weight-precedence` supplies the border-number weight;
 *      dotted and dashed have weight 1 regardless of width.
 *   3. Equal weight ⇒ the style higher on the precedence list wins.
 *   4. Identical style ⇒ the darker colour wins, by three successive brightness
 *      formulas: R+B+2G, then B+2G, then G (smaller value wins each).
 *   5. Still identical ⇒ the border FIRST in reading order (`a`) is displayed.
 */

/** A conflict candidate: the resolved {@link BorderSpec} for one cell's edge plus
 *  whether it came from the cell's own formatting or from a table/table-style
 *  border (rule #1). `null` ⇒ this side contributes no candidate at all. */
export interface BorderCandidate {
  spec: BorderSpec;
  source: 'cell' | 'table';
}

/** Structural location of a cell in the table grid. The same edge cascade is
 * consumed by both border paint and row-footprint measurement. */
export interface CellEdgeFlags {
  topRow: boolean;
  bottomRow: boolean;
  leftCol: boolean;
  rightCol: boolean;
}

/** Cell/table cascade result before an adjacent-cell conflict is resolved. */
export interface ResolvedCellEdges {
  top: BorderCandidate | null;
  bottom: BorderCandidate | null;
  left: BorderCandidate | null;
  right: BorderCandidate | null;
}

/** ECMA-376 §17.4.38/§17.4.39/§17.4.66 — resolve one cell's own, inside,
 * and outer table border cascade. `mirror` maps logical left/right to physical
 * sides for `bidiVisual`; horizontal edges are unchanged. */
export function resolveCellEdges(
  cell: CellBorders,
  table: TableBorders,
  edges: CellEdgeFlags,
  mirror: boolean,
): ResolvedCellEdges {
  const horizontal = (
    own: BorderSpec | null,
    outer: boolean,
    tableOuter: BorderSpec | null,
  ): BorderCandidate | null => {
    if (own && wordAuthoredBorderParticipates(own.style)) {
      return { spec: own, source: 'cell' };
    }
    if (!outer && cell.insideH && wordAuthoredBorderParticipates(cell.insideH.style)) {
      return { spec: cell.insideH, source: 'cell' };
    }
    const inherited = outer ? tableOuter : table.insideH;
    return inherited && wordAuthoredBorderParticipates(inherited.style)
      ? { spec: inherited, source: 'table' }
      : null;
  };
  const vertical = (
    own: BorderSpec | null,
    outer: boolean,
    tableOuter: BorderSpec | null,
  ): BorderCandidate | null => {
    if (own && wordAuthoredBorderParticipates(own.style)) {
      return { spec: own, source: 'cell' };
    }
    if (!outer && cell.insideV && wordAuthoredBorderParticipates(cell.insideV.style)) {
      return { spec: cell.insideV, source: 'cell' };
    }
    const inherited = outer ? tableOuter : table.insideV;
    return inherited && wordAuthoredBorderParticipates(inherited.style)
      ? { spec: inherited, source: 'table' }
      : null;
  };

  const top = horizontal(cell.top, edges.topRow, table.top);
  const bottom = horizontal(cell.bottom, edges.bottomRow, table.bottom);
  const left = mirror
    ? vertical(cell.right, edges.rightCol, table.right)
    : vertical(cell.left, edges.leftCol, table.left);
  const right = mirror
    ? vertical(cell.left, edges.leftCol, table.left)
    : vertical(cell.right, edges.rightCol, table.right);
  return { top, bottom, left, right };
}

function borderWeight(spec: BorderSpec): number {
  return wordTableBorderWeight(spec.style, spec.width);
}
function precedenceIndex(style: string): number {
  const i = WORD_TABLE_BORDER_STYLE_PRECEDENCE.indexOf(
    style as (typeof WORD_TABLE_BORDER_STYLE_PRECEDENCE)[number],
  );
  return i === -1 ? WORD_TABLE_BORDER_STYLE_PRECEDENCE.length : i;
}

/** Parse a 6-hex colour to (r,g,b). `null`/auto ⇒ black (0,0,0): §17.4.66's
 *  darkness comparison uses the RENDERED colour, and auto resolves to black. A
 *  malformed value also falls back to black. */
function rgb(color: string | null): { r: number; g: number; b: number } {
  const parsed = color ? hex6ToRgb(color.replace(/^#/, '')) : null;
  if (!parsed) return { r: 0, g: 0, b: 0 };
  return { r: parsed[0], g: parsed[1], b: parsed[2] };
}

/** §17.4.66 rules #4a–#4c — compare two colours by brightness, DARKER wins.
 *  Returns <0 when `a` is darker (wins), >0 when `b` is darker, 0 when identical
 *  under all three formulas. */
function compareBrightness(a: string | null, b: string | null): number {
  const ca = rgb(a);
  const cb = rgb(b);
  const f1 = (c: { r: number; g: number; b: number }) => c.r + c.b + 2 * c.g;
  const f2 = (c: { r: number; g: number; b: number }) => c.b + 2 * c.g;
  const f3 = (c: { r: number; g: number; b: number }) => c.g;
  for (const f of [f1, f2, f3]) {
    const d = f(ca) - f(cb);
    if (d !== 0) return d;
  }
  return 0;
}

/**
 * Resolve the winning border for a shared cell edge per ECMA-376 §17.4.66. `a` is
 * the border FIRST in reading order (it wins a total tie, rule #5). Either side
 * may be `null` (that cell contributes no border to this edge). Returns the
 * winning candidate, or `null` when neither side paints (both absent or nil/none).
 */
export function resolveBorderConflict(
  a: BorderCandidate | null,
  b: BorderCandidate | null,
): BorderCandidate | null {
  // §17.4.66: nil and none are both non-painting candidates and lose to the
  // opposing border. They remain distinct during the per-cell cascade: none
  // admits style/table fallback, whereas authored nil blocks fallback on that
  // side before the two resolved candidates reach this kernel.
  const visible = (candidate: BorderCandidate | null): BorderCandidate | null => (
    candidate
      && candidate.spec.style !== 'nil'
      && candidate.spec.style !== 'none'
      ? candidate
      : null
  );
  const av = visible(a);
  const bv = visible(b);
  if (!av && !bv) return null;
  if (!av) return bv;
  if (!bv) return av;

  // Rule #1 — a cell border beats a table border.
  if (av.source === 'cell' && bv.source === 'table') return av;
  if (bv.source === 'cell' && av.source === 'table') return bv;

  // Rule #2 — heavier weight wins.
  const wa = borderWeight(av.spec);
  const wb = borderWeight(bv.spec);
  if (wa !== wb) return wa > wb ? av : bv;

  // Rule #3 — equal weight ⇒ higher on the precedence list (smaller index).
  const pa = precedenceIndex(av.spec.style);
  const pb = precedenceIndex(bv.spec.style);
  if (pa !== pb) return pa < pb ? av : bv;

  // Rule #4 — identical style ⇒ darker colour wins.
  const cmp = compareBrightness(av.spec.color, bv.spec.color);
  if (cmp !== 0) return cmp < 0 ? av : bv;

  // Rule #5 — fully identical ⇒ the first in reading order.
  return av;
}

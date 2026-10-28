import {
  addCoordinateIndexEntry,
  assertCoordinateRangeArea,
  type CoordinateIndexIdentity,
  type CoordinateRange,
} from './renderer-coordinate-index.js';

export interface MergeCoverageIndex {
  /** `row:col` of every authored merge's top-left cell. */
  anchors: Set<string>;
  /** `row:col` of every continuation cell (painted by some merge's anchor). */
  continuations: Set<string>;
}

/**
 * Build the merge topology the renderer paints from.
 *
 * ECMA-376 §18.3.1.55 lists merged ranges and stores a range's value and
 * format in its top-left cell; it says nothing about two `<mergeCell>`
 * entries naming the same or intersecting cells. Excel's handling of such a
 * collection has not been verified, so the library keeps one policy for
 * every consumer, shared with the Markdown projection
 * (`clear_merge_continuation_cells` in `parser/src/markdown.rs`):
 *
 * - a cell is a continuation cell exactly when some merge covers it without
 *   having it as its top-left — i.e. more merges cover it than start at it;
 * - identical ranges therefore mean the same as one range, and a range
 *   whose top-left lies inside another merge is not painted from that
 *   top-left by the main grid pass (its cells still count as covered).
 *
 * Work is independent of how often cells are re-covered. Each range is
 * validated arithmetically, then a row sweep keeps the covered columns in a
 * segment tree over the compressed column boundaries. Between two event
 * rows the covered set is constant, so its runs are enumerated once per
 * band — the enumeration visits O(log merges) tree nodes per covered
 * elementary column segment — and each covered cell of the band is visited
 * exactly once. Every visited cell is a new continuation entry
 * (charged against `continuationIdentity`) or one of the at most
 * `merges.length` top-left cells, and every enumerated run holds at least
 * one visited cell per band row, so the total is
 * O((merges + distinct covered cells) · log merges) and stops at the entry
 * budget. Re-walking every authored range instead made N repeats of one
 * large range cost N times its area while charging it only once.
 *
 * Refusal order: ranges are checked in document order — anchor entry, then
 * the range's own representability and area — and only then is the
 * combined coverage charged. A worksheet that breaks several limits
 * therefore reports a per-range refusal before a cumulative-coverage one.
 */
export function buildMergeCoverageIndex(
  merges: readonly CoordinateRange[],
  anchorIdentity: CoordinateIndexIdentity,
  continuationIdentity: CoordinateIndexIdentity,
): MergeCoverageIndex {
  const anchors = new Set<string>();
  const continuations = new Set<string>();
  const valid: CoordinateRange[] = [];
  for (const mc of merges) {
    addCoordinateIndexEntry(anchors, `${mc.top}:${mc.left}`, anchorIdentity);
    // Refuses non-integer coordinates and any single range whose
    // continuation cells alone exceed the budget, before any expansion.
    assertCoordinateRangeArea(mc, continuationIdentity, 1);
    if (mc.bottom >= mc.top && mc.right >= mc.left) valid.push(mc);
  }
  if (valid.length === 0) return { anchors, continuations };

  // Elementary column segments [bounds[i], bounds[i + 1] - 1].
  const bounds = [...new Set(valid.flatMap((mc) => [mc.left, mc.right + 1]))]
    .sort((a, b) => a - b);
  const boundIndex = new Map<number, number>();
  bounds.forEach((value, index) => boundIndex.set(value, index));
  const segments = bounds.length - 1;
  // count[n]: ranges assigned wholly to node n; covered[n]: some segment in
  // n's span is covered. A leaf's cover count is the sum along its path.
  const count = new Int32Array(segments * 4);
  const covered = new Uint8Array(segments * 4);
  const update = (node: number, lo: number, hi: number, l: number, r: number, delta: number): void => {
    if (r < lo || hi < l) return;
    if (l <= lo && hi <= r) {
      count[node] += delta;
    } else {
      const mid = (lo + hi) >> 1;
      update(node * 2, lo, mid, l, r, delta);
      update(node * 2 + 1, mid + 1, hi, l, r, delta);
    }
    covered[node] = count[node] > 0 || (lo < hi && (covered[node * 2] || covered[node * 2 + 1])) ? 1 : 0;
  };
  const coverCount = (segment: number): number => {
    let node = 1, lo = 0, hi = segments - 1, total = 0;
    for (;;) {
      total += count[node];
      if (lo === hi) return total;
      const mid = (lo + hi) >> 1;
      if (segment <= mid) { node *= 2; hi = mid; } else { node = node * 2 + 1; lo = mid + 1; }
    }
  };
  const coveredRuns = (node: number, lo: number, hi: number, out: Array<[number, number]>): void => {
    if (!covered[node]) return;
    if (count[node] > 0) {
      const last = out.at(-1);
      if (last && last[1] === bounds[lo] - 1) last[1] = bounds[hi + 1] - 1;
      else out.push([bounds[lo], bounds[hi + 1] - 1]);
      return;
    }
    const mid = (lo + hi) >> 1;
    coveredRuns(node * 2, lo, mid, out);
    coveredRuns(node * 2 + 1, mid + 1, hi, out);
  };

  const events: Array<{ row: number; l: number; r: number; delta: number }> = [];
  const starts = new Map<string, number>();
  const anchorsByRow = new Map<number, number[]>();
  for (const mc of valid) {
    const l = boundIndex.get(mc.left)!;
    const r = boundIndex.get(mc.right + 1)! - 1;
    events.push({ row: mc.top, l, r, delta: 1 }, { row: mc.bottom + 1, l, r, delta: -1 });
    const key = `${mc.top}:${mc.left}`;
    const prior = starts.get(key);
    starts.set(key, (prior ?? 0) + 1);
    if (prior === undefined) {
      const cols = anchorsByRow.get(mc.top);
      if (cols) cols.push(mc.left); else anchorsByRow.set(mc.top, [mc.left]);
    }
  }
  events.sort((a, b) => a.row - b.row);

  let next = 0;
  while (next < events.length) {
    const row = events[next].row;
    while (next < events.length && events[next].row === row) {
      const { l, r, delta } = events[next++];
      update(1, 0, segments - 1, l, r, delta);
    }
    // Every add has a later removal, so a covered band always ends at an event.
    if (!covered[1] || next >= events.length) continue;
    const bandEnd = events[next].row - 1;
    // Top-left cells covered only by the merges they start stay paintable.
    const ownAnchors = new Set<string>();
    for (const col of anchorsByRow.get(row) ?? []) {
      const key = `${row}:${col}`;
      if (coverCount(boundIndex.get(col)!) === starts.get(key)) ownAnchors.add(key);
    }
    const runs: Array<[number, number]> = [];
    coveredRuns(1, 0, segments - 1, runs);
    for (let r = row; r <= bandEnd; r++) {
      for (const [first, last] of runs) {
        for (let c = first; c <= last; c++) {
          const key = `${r}:${c}`;
          if (r === row && ownAnchors.has(key)) continue;
          addCoordinateIndexEntry(continuations, key, continuationIdentity);
        }
      }
    }
  }
  return { anchors, continuations };
}

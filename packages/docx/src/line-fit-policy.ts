/**
 * Greedy line-break selection shared by the body paragraph splitter and the
 * table-cell content splitter: returns the largest fitting prefix reached
 * before the first failed end in (start, limitEnd], using the caller's
 * cumulative fit measure for lines [start, end). Callers must
 * handle endpoints with a different reserve contract independently: the body
 * paragraph selector tests completion before its continuing-prefix walk.
 * Preserving the early stop keeps the prefix float-comparison order unchanged.
 *
 * ECMA-376 does not prescribe a line-selection algorithm for page/row filling;
 * greedy forward filling is this renderer's established layout convention
 * (§17.4.6 `cantSplit` and the header/exact-row rules gate WHETHER a row may
 * split and are enforced by the callers, not here).
 *
 * CONTRACTS:
 * - `fitAt` is invoked with STRICTLY INCREASING `end` values (start+1 ..
 *   limitEnd, no gaps until the walk stops), so a caller may accumulate its
 *   measure incrementally across calls — the body caller does, keeping the
 *   whole selection O(n).
 * - `fitAt` must return FINITE numbers. Line extents and collapsed spacing are
 *   finite by construction in both callers; non-finite values are outside the
 *   contract. (The `!(value <= available)` form stops the walk on NaN as a
 *   fail-safe — matching the body caller's historical `<=` loop condition; the
 *   cell caller's historical `>` break would have kept walking on NaN, an
 *   unreachable difference under this contract, resolved in favor of
 *   stopping.)
 */
export function selectLargestFittingEnd(
  start: number,
  limitEnd: number,
  available: number,
  fitAt: (end: number) => number,
): { end: number; fitValue: number } {
  let end = start;
  let fitValue = 0;
  for (let candidate = start + 1; candidate <= limitEnd; candidate++) {
    const candidateValue = fitAt(candidate);
    if (!(candidateValue <= available)) break;
    end = candidate;
    fitValue = candidateValue;
  }
  return { end, fitValue };
}

/**
 * ECMA-376 §17.3.1.44 (`widowControl`) widow/orphan adjustment following the
 * greedy selection. Pure decision mirroring the body pagination policy:
 * - drop the selection's last line when exactly ONE line would remain for the
 *   next page (a widow) and the selection keeps at least two lines;
 * - relocate the whole paragraph when only its FIRST line was selected (an
 *   orphan) and relocating can gain space (`canRelocate`).
 * Side effects (extent bookkeeping, the page advance and remeasure) stay with
 * the caller.
 */
export function adjustForWidowOrphan(input: {
  widowControl: boolean;
  start: number;
  end: number;
  totalLines: number;
  /** Whether moving the paragraph to the next column/page can gain space. The
   *  body caller passes its HISTORICAL condition `cursorY > 0` — the cursor
   *  sits below the PAGE content top. This is deliberately NOT "below the
   *  column top": in a continuous mid-page section (`colTop() > 0`) a
   *  paragraph sitting exactly at the column top still relocates, as it always
   *  did — preserving byte-identity with the pre-extraction behavior (which is
   *  why the VRT pixel baseline holds). Whether a column-relative condition is
   *  the better spec reading for mid-page regions is a separate question,
   *  tracked outside this refactor. */
  canRelocate: boolean;
}): { kind: 'keep' } | { kind: 'dropLastLine' } | { kind: 'relocate' } {
  if (!input.widowControl || input.end >= input.totalLines) return { kind: 'keep' };
  if (input.totalLines - input.end === 1 && input.end - input.start >= 2) {
    return { kind: 'dropLastLine' };
  }
  if (input.start === 0 && input.end - input.start === 1 && input.canRelocate) {
    return { kind: 'relocate' };
  }
  return { kind: 'keep' };
}

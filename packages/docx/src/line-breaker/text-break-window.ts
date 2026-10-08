/** A sorted acquired offset list with a constant-size slice. Queue tails retain
 * the same owned list: filtering/copying every suffix makes dense breaks quadratic.
 * Offsets are UTF-16; start/end index the list, origin rebases its source offsets. */
export interface TextBreakWindow {
  readonly offsets: readonly number[];
  readonly start: number;
  readonly end: number;
  readonly origin: number;
}

export function textBreakWindow(offsets: readonly number[]): TextBreakWindow | undefined {
  return offsets.length ? { offsets, start: 0, end: offsets.length, origin: 0 } : undefined;
}

export function textBreakOffsetAt(window: TextBreakWindow, index: number): number {
  return window.offsets[window.start + index]! - window.origin;
}

export function* textBreakOffsets(window: TextBreakWindow | undefined): Iterable<number> {
  if (!window) return;
  for (let i = 0; i < window.end - window.start; i++) yield textBreakOffsetAt(window, i);
}

export function sliceTextBreakWindow(
  window: TextBreakWindow | undefined, start: number, end: number,
): TextBreakWindow | undefined {
  if (!window) return undefined;
  const bound = (value: number, inclusive: boolean): number => {
    let lo = window.start, hi = window.end;
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (window.offsets[mid]! < value || inclusive && window.offsets[mid] === value) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  };
  const nextStart = bound(window.origin + start, true);
  const nextEnd = bound(window.origin + end, false);
  return nextStart < nextEnd
    ? { offsets: window.offsets, start: nextStart, end: nextEnd, origin: window.origin + start }
    : undefined;
}

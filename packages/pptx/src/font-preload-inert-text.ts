import { fontPreloadInertRanges } from './font-preload-inert-ranges.js';

/** Any concatenation of accepted assigned scalars is canonically unchanged:
 * no decomposition key, nonzero CCC or possible composition second operand.
 * UAX15 §12.1 stabilizes assigned strings across Unicode versions. Reject
 * malformed UTF16 and unknown/unassigned values; never call host normalize().
 * This optional exclusion policy changes no painter text or font selection. */
export function fontPreloadInertScalar(cp: number): boolean {
  if (!Number.isInteger(cp)) return false;
  const ranges = fontPreloadInertRanges();
  let lo = 0, hi = ranges.length / 2 - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >>> 1;
    if (cp < ranges[mid * 2]) hi = mid - 1;
    else if (cp > ranges[mid * 2 + 1]) lo = mid + 1;
    else return true;
  }
  return false;
}

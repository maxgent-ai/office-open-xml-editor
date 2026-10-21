import { JT_NAMES, JT_RANGE_STARTS, JT_RANGE_VALUE } from './joining-type.generated.js';

/** Unicode Joining_Type (UAX #44, DerivedJoiningType.txt). */
export type JoiningType = 'U' | 'C' | 'D' | 'L' | 'R' | 'T';

/** Return the Unicode Joining_Type for one code point. */
export function joiningType(cp: number): JoiningType {
  let lo = 0;
  let hi = JT_RANGE_STARTS.length - 1;
  let rangeIndex = -1;

  while (lo <= hi) {
    const mid = lo + ((hi - lo) >> 1);
    if (JT_RANGE_STARTS[mid] <= cp) {
      rangeIndex = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }

  if (rangeIndex < 0) return 'U';
  return JT_NAMES[JT_RANGE_VALUE[rangeIndex]] ?? 'U';
}

/** True when the code point can join to the following letter (D, L, C). */
export function joinsToFollowing(cp: number): boolean {
  const jt = joiningType(cp);
  return jt === 'D' || jt === 'L' || jt === 'C';
}

/** True when the code point can join to the preceding letter (D, R, C). */
export function joinsToPreceding(cp: number): boolean {
  const jt = joiningType(cp);
  return jt === 'D' || jt === 'R' || jt === 'C';
}

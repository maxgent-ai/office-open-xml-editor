/**
 * How many leading pages of a replacement layout equal the layout it replaces.
 *
 * Progressive layout publishes a growing prefix and then the authoritative
 * layout. Most publications only append pages, and a viewer that repaints
 * every mounted page for each of them throws away the opening page's bitmap
 * while it is still being rendered. A page record is the complete retained
 * paint input of that page (its layers, geometry and resolved fields), so two
 * structurally equal records paint the same pixels and a canvas painted from
 * one stays valid for the other. Records are plain, acyclic layout data.
 */
export function unchangedLeadingPageCount(
  previous: readonly unknown[] | null | undefined,
  next: readonly unknown[],
): number {
  if (!previous) return 0;
  const limit = Math.min(previous.length, next.length);
  let count = 0;
  while (count < limit && samePlainData(previous[count], next[count])) count += 1;
  return count;
}

function samePlainData(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (typeof left !== 'object' || typeof right !== 'object' || left === null || right === null) {
    return false;
  }
  if (Array.isArray(left)) {
    if (!Array.isArray(right) || left.length !== right.length) return false;
    for (let index = 0; index < left.length; index += 1) {
      if (!samePlainData(left[index], right[index])) return false;
    }
    return true;
  }
  if (Array.isArray(right)) return false;
  const leftKeys = Object.keys(left);
  if (leftKeys.length !== Object.keys(right).length) return false;
  for (const key of leftKeys) {
    if (!Object.prototype.hasOwnProperty.call(right, key)) return false;
    if (!samePlainData((left as Record<string, unknown>)[key], (right as Record<string, unknown>)[key])) {
      return false;
    }
  }
  return true;
}

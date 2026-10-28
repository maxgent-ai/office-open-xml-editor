import { joiningType, joinsToFollowing, joinsToPreceding } from '@silurus/ooxml-core/internal/joining-type';

/** Synthetic U+200D sides of a paint piece (see `TextPaintOp.joiners`). */
export type PaintJoiners = 'before' | 'after' | 'both';

/** Cursive joining across paint boundaries of one text (Unicode Joining_Type,
 * transparent marks skipped). Built in two linear passes; each query is O(1).
 * WORD_CONDENSED_SPACING_ADVANCE_FLOOR splits a cursive word into separately
 * painted pieces; each keeps its joined form across the new boundaries. */
export function cursiveJoinContext(text: string): (start: number, end: number) => PaintJoiners | undefined {
  const length = text.length;
  // Position of the first non-transparent code point at or after each index,
  // and of the last one before each index (-1 / length when none).
  const nextBase = new Int32Array(length + 1).fill(length);
  const previousBase = new Int32Array(length + 1).fill(-1);
  for (let index = length - 1; index >= 0; index -= 1) {
    const unit = text.charCodeAt(index);
    const trail = unit >= 0xDC00 && unit <= 0xDFFF && index > 0;
    if (trail) {
      nextBase[index] = nextBase[index + 1]!;
      continue;
    }
    nextBase[index] = joiningType(text.codePointAt(index)!) === 'T' ? nextBase[index + 1]! : index;
  }
  for (let index = 0; index < length;) {
    const cp = text.codePointAt(index)!;
    const width = cp > 0xFFFF ? 2 : 1;
    const base = joiningType(cp) === 'T' ? previousBase[index]! : index;
    for (let unit = 1; unit <= width && index + unit <= length; unit += 1) {
      previousBase[index + unit] = base;
    }
    index += width;
  }
  const cp = (position: number): number => text.codePointAt(position)!;
  return (start, end) => {
    const first = nextBase[start]!;
    const last = previousBase[end]!;
    const previous = previousBase[start]!;
    const next = nextBase[end]!;
    const inside = first < end && last >= start;
    const before = inside && previous >= 0 && joinsToFollowing(cp(previous)) && joinsToPreceding(cp(first));
    const after = inside && next < length && joinsToFollowing(cp(last)) && joinsToPreceding(cp(next));
    return before && after ? 'both' : before ? 'before' : after ? 'after' : undefined;
  };
}

/** `text` with the synthetic joiners named by `joiners`. */
export function withPaintJoiners(text: string, joiners: PaintJoiners | undefined): string {
  if (!joiners) return text;
  return `${joiners === 'after' ? '' : '‍'}${text}${joiners === 'before' ? '' : '‍'}`;
}

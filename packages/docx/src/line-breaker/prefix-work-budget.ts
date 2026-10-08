/** Library resource policy, independent of Office line-fitting behavior.
 * Exact prefix searches, joined admission and their grapheme scans (especially negative pitch) may revisit every candidate
 * on every queued tail. Bound their cumulative requested UTF-16 work per pass,
 * across semantic and emergency searches, before entering shaping. 16 Mi units
 * is a fixed work quota, not a width tolerance or a candidate-count truncation.
 * Grapheme scans are charged before allocation. Cache hits count too: constructing their text/context keys also scans text.
 * Exhaustion aborts this layout pass; no incomplete lines or alternate fit rule
 * are returned. The pass owns this counter, so retries and documents share no
 * mutable budget/cache and failure retains no input content. */
export function createPrefixWorkBudget(limit = 16 * 1024 * 1024): (utf16Units: number) => void {
  let used = 0;
  return utf16Units => {
    const next = used + utf16Units;
    if (!Number.isSafeInteger(utf16Units) || utf16Units < 0 || next > limit) {
      throw new RangeError('DOCX line-break prefix search exceeded its UTF-16 work quota');
    }
    used = next;
  };
}

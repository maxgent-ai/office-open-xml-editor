/**
 * The first `maxCodeUnits` UTF-16 code units of `value`, shortened by one when
 * the cut would split a surrogate pair, so a bounded excerpt never ends in a
 * lone high surrogate.
 */
export function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, maxCodeUnits);
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

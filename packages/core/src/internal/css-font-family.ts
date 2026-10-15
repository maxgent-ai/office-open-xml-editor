/** Serialize a named family as one CSS string (CSS Syntax 3 §4.3.5/4.3.7).
 * OOXML ST_TextTypeface is a string, so quotes and backslashes are data, not
 * font-list syntax. Generic family keywords are handled by callers separately.
 * Hex escapes for controls include a terminator so a following digit is data.
 */
export function quoteCssFontFamily(family: string): string {
  let escaped = '';
  for (const scalar of family) {
    const codePoint = scalar.codePointAt(0) ?? 0;
    if (scalar === '"' || scalar === '\\') escaped += `\\${scalar}`;
    else if (codePoint === 0) escaped += '\uFFFD';
    else if (codePoint <= 0x1f || codePoint === 0x7f) escaped += `\\${codePoint.toString(16)} `;
    else escaped += scalar;
  }
  return `"${escaped}"`;
}

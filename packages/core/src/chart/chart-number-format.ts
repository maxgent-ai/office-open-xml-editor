// Chart value number-formatting. Pure string/number logic (no canvas),
// extracted from the chart renderer so it can be unit-tested and reused.
// Implements the subset of ECMA-376 §18.8.30 number-format codes that chart
// axis ticks and data labels need: section syntax (positive;negative;zero),
// literal escapes, thousands separators, decimals, percent, and Excel serial
// dates.

import { excelSerialToUtcDate } from '../excel-date';
import {
  formatExcelDateTime,
  isDateFormatSection,
  splitFormatSections,
  textSectionIndex,
} from '../excel-number-format';
import { roundDecimalHalfUp } from '../text/round-decimal';

const localizedShortDateFormatters = new Map<string, Intl.DateTimeFormat>();

/** Excel built-in number format 14 is locale-sensitive rather than the
 * invariant `m/d/yy` pattern stored in a chart cache. */
export function formatLocalizedExcelShortDate(
  serial: number,
  date1904 = false,
  locale = typeof navigator === 'undefined' ? undefined : navigator.language,
): string {
  const cacheKey = locale ?? '';
  let formatter = localizedShortDateFormatters.get(cacheKey);
  if (!formatter) {
    formatter = new Intl.DateTimeFormat(locale, {
      year: 'numeric', month: 'numeric', day: 'numeric', timeZone: 'UTC',
    });
    localizedShortDateFormatters.set(cacheKey, formatter);
  }
  return formatter.format(excelSerialToUtcDate(serial, date1904));
}

/** Excel's default `formatCode="General"` for charts: raw numbers with no
 *  "k"/"M" abbreviation, trailing decimal zeros trimmed. */
export function formatChartVal(v: number): string {
  // Matches Excel's default `<c:valAx><c:numFmt formatCode="General">` which
  // shows raw numbers — no "k"/"M" abbreviation.
  if (Number.isInteger(v)) return String(v);
  // Cap at 6 decimals and trim trailing zeros (0.50 → "0.5"). Round half-UP on
  // the decimal value like Excel (matters only at a `.xxxxxx5` boundary; equal
  // to `toFixed(6)` otherwise), so this path never disagrees with the explicit
  // number-format sections below.
  return roundDecimalHalfUp(v, 6).replace(/\.?0+$/, '');
}

/**
 * Format a chart value with an Excel number-format code. Honors ECMA-376
 * §18.8.30 section syntax (positive;negative;zero;text), common literal
 * escapes (`"..."`, `\x`, `_x` → space), and numeric patterns built from
 * `#`, `0`, `.`, `,`. Unknown tokens are emitted verbatim so currency
 * symbols like `¥` or `$` keep working even when the workbook stored them
 * unquoted. Returns the default `formatChartVal` output when `code` is null
 * or an empty section tells the caller to hide the value.
 */
export function formatChartValWithCode(
  v: number,
  code: string | null | undefined,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). `true` resolves serial
   *  date codes against the 1904 epoch. Defaults to false (1900 system). */
  date1904 = false,
): string {
  // Absent `code`, or the reserved "General" keyword (ECMA-376 §18.8.30), both
  // mean the General number format. LibreOffice charts emit
  // `<c:numFmt formatCode="General">`; tokenizing it as a literal pattern would
  // render the word "General" instead of the value (issue #358).
  if (!code || code.trim().toLowerCase() === 'general') return formatChartVal(v);
  // Section selection per §18.8.30: positive;negative;zero, with the text
  // section left out (it never formats a number). When the negative section
  // is omitted, a negative number is formatted with the positive section and
  // a leading minus.
  const sections = splitFormatSections(code);
  const textIndex = textSectionIndex(sections);
  const numeric = sections.filter((_, i) => i !== textIndex);
  if (numeric.length === 0) return formatChartVal(v);
  let section: string;
  if (v > 0) section = numeric[0];
  else if (v < 0) section = numeric[1] ?? numeric[0];
  else section = numeric[2] ?? numeric[0];
  if (section === '') return '';
  const explicitNegative = v < 0 && numeric.length >= 2;
  // Date/time is a property of the selected section: the value is a serial
  // date (scatter / time-series axes), rendered by the same formatter as
  // worksheet cells.
  if (isDateFormatSection(section)) {
    return formatExcelDateTime(explicitNegative ? -v : v, section, date1904);
  }
  const needsLeadingMinus = v < 0 && !explicitNegative;
  return (needsLeadingMinus ? '-' : '') + applyChartNumberSection(Math.abs(v), section);
}

/**
 * Format a category-axis tick label. Categories arrive as the raw cached
 * strings the workbook stored (`<c:cat><c:strCache>` text, or the numeric
 * text of a `<c:numCache>` on a date/number axis). ECMA-376 §21.2.2.71
 * (`c:numFmt`) lets the category axis carry a number-format code: Excel
 * applies it to the tick labels just as it does on the value axis. When the
 * axis has such a code AND the raw category parses to a finite number, format
 * it (which routes serial dates through `formatExcelDateTime` automatically); a
 * missing code, `"General"`, or a non-numeric category string falls through
 * to the raw text unchanged — no new interpretation is invented.
 */
export function formatCategoryLabel(
  raw: string,
  code: string | null | undefined,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). Threaded to the date
   *  formatter for serial-date category labels. Defaults to false. */
  date1904 = false,
): string {
  if (!code) return raw;
  // Only numeric-looking categories are formatted. `Number('')` and
  // `Number('  ')` are 0 (falsely finite), so reject blank / whitespace text
  // explicitly; keep genuine string categories (e.g. "Q1", "North") verbatim.
  if (raw.trim() === '') return raw;
  const num = Number(raw);
  if (!Number.isFinite(num)) return raw;
  return formatChartValWithCode(num, code, date1904);
}

function applyChartNumberSection(abs: number, section: string): string {
  // Tokenize the section, separating numeric-pattern runs (`#`, `0`, `.`,
  // `,`, `?`) from literal runs so percent / decimal handling runs once.
  type Tok = { kind: 'lit' | 'num'; text: string };
  const toks: Tok[] = [];
  let i = 0;
  let pushedNum = false;
  let percent = false;
  while (i < section.length) {
    const c = section[i];
    if (c === '"') {
      i++;
      let s = '';
      while (i < section.length && section[i] !== '"') { s += section[i]; i++; }
      if (i < section.length) i++;
      toks.push({ kind: 'lit', text: s });
      continue;
    }
    if (c === '\\' && i + 1 < section.length) {
      toks.push({ kind: 'lit', text: section[i + 1] });
      i += 2;
      continue;
    }
    if (c === '_' && i + 1 < section.length) {
      // `_x` pads a width of x — render as a single space, matching Excel
      // alignment padding without caring about exact glyph metrics.
      toks.push({ kind: 'lit', text: ' ' });
      i += 2;
      continue;
    }
    if (c === '*' && i + 1 < section.length) {
      // `*x` fills the remaining column width with x; we can't know the
      // column width at this layer so we drop it.
      i += 2;
      continue;
    }
    if (c === '[') {
      i++;
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
      continue;
    }
    if (c === '%') { percent = true; toks.push({ kind: 'lit', text: '%' }); i++; continue; }
    if (c === '#' || c === '0' || c === '.' || c === ',' || c === '?') {
      let run = '';
      while (
        i < section.length &&
        (section[i] === '#' || section[i] === '0' || section[i] === '.' ||
         section[i] === ',' || section[i] === '?')
      ) { run += section[i]; i++; }
      toks.push({ kind: 'num', text: run });
      pushedNum = true;
      continue;
    }
    // Everything else (currency symbols like ¥, $, parens, spaces) is literal.
    toks.push({ kind: 'lit', text: c });
    i++;
  }
  if (!pushedNum) {
    // No numeric pattern at all — section is purely literal (e.g. `"N/A"`).
    return toks.map(t => t.text).join('');
  }
  const value = percent ? abs * 100 : abs;
  // Merge numeric tokens into one pattern — Excel treats `#,##0.00` as a
  // single pattern even when flanked by literals. We keep the literal tokens
  // where they are and replace the first num token with the formatted number,
  // dropping subsequent num tokens (they're all part of the same pattern).
  let pattern = '';
  for (const t of toks) if (t.kind === 'num') pattern += t.text;
  const formatted = formatNumericPattern(value, pattern);
  let seenNum = false;
  return toks.map(t => {
    if (t.kind === 'lit') return t.text;
    if (seenNum) return '';
    seenNum = true;
    return formatted;
  }).join('');
}

function formatNumericPattern(value: number, pattern: string): string {
  // Detect thousands separator (a `,` between digit placeholders) and the
  // number of decimal places (digit chars after `.`).
  const dotIdx = pattern.indexOf('.');
  const intPart = dotIdx >= 0 ? pattern.slice(0, dotIdx) : pattern;
  const fracPart = dotIdx >= 0 ? pattern.slice(dotIdx + 1) : '';
  const thousands = /,/.test(intPart);
  const fracDigits = (fracPart.match(/[#0?]/g) ?? []).length;
  // Minimum integer digits = count of `0` in integer part.
  const minIntDigits = (intPart.replace(/,/g, '').match(/0/g) ?? []).length;
  // Office rounds a chart label's decimals half-UP on the decimal value
  // (2.675 → "2.68"); `toFixed` would round the binary double down.
  const rounded = roundDecimalHalfUp(value, fracDigits);
  const [ints, fracs = ''] = rounded.split('.');
  const paddedInts = ints.padStart(minIntDigits, '0');
  const withSeparators = thousands ? paddedInts.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : paddedInts;
  if (fracDigits === 0) return withSeparators;
  return `${withSeparators}.${fracs.padEnd(fracDigits, '0')}`;
}

// Shared SpreadsheetML number-format grammar pieces (ECMA-376 §18.8.30) used
// by worksheet cells and chart labels alike: section splitting, the text
// section, date/time section detection and date/time rendering.

import { excelSerialToUtcDate } from './excel-date';

/** Index of the text section (§18.8.30), or -1: the fourth section when
 *  there are four, otherwise the last section when it holds an `@`
 *  placeholder outside quotes, escapes and pad / fill pairs. */
export function textSectionIndex(sections: string[]): number {
  if (sections.length >= 4) return 3;
  const last = sections[sections.length - 1];
  let i = 0;
  while (i < last.length) {
    const ch = last[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = last.indexOf('"', i + 1);
      i = end < 0 ? last.length : end + 1;
    } else if (ch === '[') {
      const end = last.indexOf(']', i);
      i = end < 0 ? last.length : end + 1;
    } else if (ch === '@') {
      return sections.length - 1;
    } else {
      i++;
    }
  }
  return -1;
}

/**
 * Split a whole format code into its `;`-separated sections. `;` never appears
 * inside quotes, escapes or `[...]` in a valid code, but we scan structurally
 * so a stray one inside those never splits the section (defensive; matches how
 * Excel lexes).
 */
export function splitFormatSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === '"') {
      cur += ch; i++;
      while (i < code.length && code[i] !== '"') cur += code[i++];
      if (i < code.length) cur += code[i++];
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      // An escape or a pad / fill pair: its operand is never a delimiter.
      cur += ch;
      if (i + 1 < code.length) cur += code[i + 1];
      i += 2;
    } else if (ch === '[') {
      cur += ch; i++;
      while (i < code.length && code[i] !== ']') cur += code[i++];
      if (i < code.length) cur += code[i++];
    } else if (ch === ';') {
      out.push(cur); cur = ''; i++;
    } else {
      cur += ch; i++;
    }
  }
  out.push(cur);
  return out;
}

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
/** Japanese short weekday names (aaa format code, e.g. "水"). */
const JP_WEEKDAY_SHORT = ['日', '月', '火', '水', '木', '金', '土'];
/** Japanese long weekday names (aaaa format code, e.g. "水曜日"). */
const JP_WEEKDAY_LONG = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

/** Japanese imperial eras, newest-first. First entry whose `start` is
 *  ≤ the target date wins (ECMA-376 §18.8.30 — g/gg/ggg and e/ee codes). */
const JP_ERAS: Array<{ start: Date; abbr: string; short: string; long: string }> = [
  { start: new Date(Date.UTC(2019, 4,  1)), abbr: 'R', short: '令', long: '令和' },
  { start: new Date(Date.UTC(1989, 0,  8)), abbr: 'H', short: '平', long: '平成' },
  { start: new Date(Date.UTC(1926, 11, 25)), abbr: 'S', short: '昭', long: '昭和' },
  { start: new Date(Date.UTC(1912, 6,  30)), abbr: 'T', short: '大', long: '大正' },
  { start: new Date(Date.UTC(1868, 0,  25)), abbr: 'M', short: '明', long: '明治' },
];

function resolveJpEra(date: Date): { abbr: string; short: string; long: string; year: number } {
  for (const era of JP_ERAS) {
    if (date.getTime() >= era.start.getTime()) {
      return {
        abbr: era.abbr,
        short: era.short,
        long: era.long,
        year: date.getUTCFullYear() - era.start.getUTCFullYear() + 1,
      };
    }
  }
  // Pre-Meiji: fall back to Gregorian year, keep Meiji names as a best effort.
  const last = JP_ERAS[JP_ERAS.length - 1];
  return { abbr: last.abbr, short: last.short, long: last.long, year: date.getUTCFullYear() };
}

/**
 * Format an Excel date serial with one date/time format section (callers
 * select the section; see splitFormatSections / isDateFormatSection).
 * Supports: y/yy/yyy/yyyy, m/mm/mmm/mmmm/mmmmm, d/dd/ddd/dddd,
 *           h/hh, m/mm (minutes when after h), s/ss, AM/PM, A/P,
 *           quoted literals, bracket escapes, _ padding, * fill.
 *
 * `date1904` selects the date system (`<workbookPr date1904>`, §18.2.28). The
 * serial → calendar-date conversion is delegated to the shared core
 * `excelSerialToUtcDate` (§18.17.4.1), which carries the 1900 Lotus
 * leap-year-bug compat and the 1904 epoch. It defaults to false so 1900-system
 * workbooks are unchanged (apart from the serial ≤ 59 leap-bug compat, which is
 * now correct in both systems).
 */
export function formatExcelDateTime(serial: number, section: string, date1904 = false): string {
  const date = excelSerialToUtcDate(serial, date1904);
  const yr = date.getUTCFullYear();
  const mo = date.getUTCMonth() + 1;   // 1-12
  const dy = date.getUTCDate();
  const wd = date.getUTCDay();          // 0=Sun
  // An elapsed-time section (`[h]:mm`, `[mm]:ss`) reads its clock fields as
  // remainders of the same absolute, millisecond-rounded duration as its
  // elapsed total, so a negative duration keeps consistent minutes.
  const elapsedSection = hasElapsedBracket(section);
  const absMs = Math.round(Math.abs(serial) * 86_400_000);
  const hr = elapsedSection ? Math.floor(absMs / 3_600_000) % 24 : date.getUTCHours();
  const mi = elapsedSection ? Math.floor(absMs / 60_000) % 60 : date.getUTCMinutes();
  const sc = elapsedSection ? Math.floor(absMs / 1_000) % 60 : date.getUTCSeconds();

  // Take the first section (positive / no-sign section)
  const hasAmPm = hasAmPmToken(section);
  const japaneseEra = sectionHasJapaneseLocale(section);
  let era: ReturnType<typeof resolveJpEra> | null = null;
  const getEra = (): ReturnType<typeof resolveJpEra> => era ?? (era = resolveJpEra(date));

  let result = '';
  let i = 0;
  let prevWasHour = false;

  while (i < section.length) {
    const ch = section[i];

    if (ch === '"') {
      // Quoted string literal
      i++;
      while (i < section.length && section[i] !== '"') result += section[i++];
      if (i < section.length) i++;
      prevWasHour = false;

    } else if (ch === '[') {
      // ECMA-376 §18.8.30: `[h]` / `[m]` / `[s]` are elapsed-time tokens that
      // suppress the h < 24 / m < 60 / s < 60 wrap-around and instead render
      // the full duration. Any other bracket content (locale IDs, colours,
      // conditions) is metadata and skipped.
      const end = section.indexOf(']', i);
      const inner = end > i ? section.slice(i + 1, end) : '';
      const elapsed = inner.match(/^([hms])\1*$/i);
      if (elapsed) {
        const kind = elapsed[1].toLowerCase();
        const sign = serial < 0 ? '-' : '';
        // Whole seconds of the millisecond-rounded duration, the same
        // rounding the clock fields above get from excelSerialToUtcDate.
        const absSec = Math.floor(absMs / 1000);
        let v: number;
        if      (kind === 'h') v = Math.floor(absSec / 3600);
        else if (kind === 'm') v = Math.floor(absSec / 60);
        else                   v = absSec;
        const padded = inner.length >= 2 ? String(v).padStart(inner.length, '0') : String(v);
        result += sign + padded;
        i = end + 1;
        prevWasHour = kind === 'h';
      } else {
        while (i < section.length && section[i] !== ']') i++;
        if (i < section.length) i++;
      }

    } else if (ch === '_') {
      i += 2; // _ followed by a padding character — skip both

    } else if (ch === '*') {
      i += 2; // * followed by fill character — skip both

    } else if (ch === '\\') {
      const escaped = section[i + 1];
      if (escaped !== undefined) result += escaped;
      i += 2;
      // An escaped time separator (`h\:mm`) keeps the hour context, like the
      // bare separators below.
      if (escaped !== ':' && escaped !== '/' && escaped !== '-' && escaped !== '.' && escaped !== ' ') {
        prevWasHour = false;
      }

    } else if (ch === 'y' || ch === 'Y') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'y') { n++; i++; }
      result += n <= 2 ? String(yr).slice(-2) : String(yr).padStart(4, '0');
      prevWasHour = false;

    } else if (ch === 'm' || ch === 'M') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'm') { n++; i++; }
      // Determine month vs minutes:
      //   minutes when immediately after h/hh, OR immediately before :s/:ss
      const rest = section.slice(i).replace(/\[[^\]]*\]/g, '');
      const isMinutes = prevWasHour || /^:s/i.test(rest);
      if (isMinutes) {
        result += n >= 2 ? String(mi).padStart(2, '0') : String(mi);
      } else {
        if      (n === 1) result += String(mo);
        else if (n === 2) result += String(mo).padStart(2, '0');
        else if (n === 3) result += MONTH_NAMES[mo - 1].slice(0, 3);
        else if (n === 4) result += MONTH_NAMES[mo - 1];
        else              result += MONTH_NAMES[mo - 1][0]; // mmmmm = first letter
      }
      prevWasHour = false;

    } else if (ch === 'd' || ch === 'D') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'd') { n++; i++; }
      if      (n === 1) result += String(dy);
      else if (n === 2) result += String(dy).padStart(2, '0');
      else if (n === 3) result += WEEKDAY_NAMES[wd].slice(0, 3);
      else              result += WEEKDAY_NAMES[wd];
      prevWasHour = false;

    } else if (ch === 'h' || ch === 'H') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'h') { n++; i++; }
      const h = hasAmPm ? (hr % 12 || 12) : hr;
      result += n >= 2 ? String(h).padStart(2, '0') : String(h);
      prevWasHour = true;

    } else if (ch === 's' || ch === 'S') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 's') { n++; i++; }
      result += n >= 2 ? String(sc).padStart(2, '0') : String(sc);
      prevWasHour = false;

    } else if (ch === 'g' || ch === 'G') {
      // Era name. Under a `[$-411]` (ja-JP) section, as Excel renders it:
      //   g → 'R' / 'H' / 'S' / 'T' / 'M', gg → '令' …, ggg → '令和' ….
      // Excel draws no era name in any other section (other LCIDs or none).
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'g') { n++; i++; }
      if (japaneseEra) {
        const e = getEra();
        if      (n === 1) result += e.abbr;
        else if (n === 2) result += e.short;
        else              result += e.long;
      }
      prevWasHour = false;

    } else if (ch === 'e' || ch === 'E') {
      // Era year: under `[$-411]` the Japanese era year (`ee` zero-padded);
      // Excel renders `e` and `ee` as the four-digit year in any other section.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'e') { n++; i++; }
      if (japaneseEra) {
        const y = getEra().year;
        result += n >= 2 ? String(y).padStart(2, '0') : String(y);
      } else {
        result += String(yr).padStart(4, '0');
      }
      prevWasHour = false;

    } else if (ch === 'r' || ch === 'R') {
      // Under `[$-411]`, `r` is `ee` and `rr` is `gggee` (§18.8.30); Excel
      // renders both as the four-digit year in any other section.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'r') { n++; i++; }
      if (japaneseEra) {
        const e = getEra();
        result += (n >= 2 ? e.long : '') + String(e.year).padStart(2, '0');
      } else {
        result += String(yr).padStart(4, '0');
      }
      prevWasHour = false;

    } else if (ch === 'A' || ch === 'a') {
      const upper = section.slice(i).toUpperCase();
      // Japanese weekday format codes (Excel ja locale). `aaaa` = "水曜日",
      // `aaa` = "水". Checked before AM/PM because those are shorter matches
      // and would otherwise swallow the leading 'a'.
      if (upper.startsWith('AAAA')) {
        result += JP_WEEKDAY_LONG[wd]; i += 4;
      } else if (upper.startsWith('AAA')) {
        result += JP_WEEKDAY_SHORT[wd]; i += 3;
      } else if (upper.startsWith('AM/PM')) {
        result += hr < 12 ? 'AM' : 'PM'; i += 5;
      } else if (upper.startsWith('A/P')) {
        result += hr < 12 ? 'A' : 'P'; i += 3;
      } else {
        result += ch; i++;
      }
      prevWasHour = false;

    } else {
      result += ch;
      i++;
      // Separators (:/-. space) don't reset the hour context for m/mm lookahead
      if (ch !== ':' && ch !== '/' && ch !== '-' && ch !== '.' && ch !== ' ') {
        prevWasHour = false;
      }
    }
  }

  return result;
}

/** Whether era codes in a section read as the Japanese era: the section
 *  carries a `[$…-LCID]` bracket (outside quotes) whose language ID is
 *  0x0411 ja-JP. Measured in Excel (ja-JP macOS): `[$-411]` renders the
 *  Japanese era, while `[$-404]`, `[$-409]` and codes without an LCID all
 *  render `g` as nothing and `e` / `ee` / `r` / `rr` as the four-digit year. */
function sectionHasJapaneseLocale(section: string): boolean {
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      if (end < 0) break;
      const lcid = /^\$[^-]*-([0-9a-f]+)$/i.exec(section.slice(i + 1, end));
      if (lcid && (parseInt(lcid[1], 16) & 0xffff) === 0x0411) return true;
      i = end + 1;
    } else {
      i++;
    }
  }
  return false;
}

/** Whether a section holds an AM/PM or A/P token outside quotes, escapes and
 *  pad / fill pairs (a quoted "AM/PM" is literal text). */
function hasAmPmToken(section: string): boolean {
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      i = end < 0 ? section.length : end + 1;
    } else if (/^(am\/pm|a\/p)/i.test(section.slice(i))) {
      return true;
    } else {
      i++;
    }
  }
  return false;
}

/** Whether a section holds an elapsed-time bracket `[h]` / `[mm]` / `[ss]`
 *  outside quotes, escapes and pad / fill pairs. */
function hasElapsedBracket(section: string): boolean {
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      if (end < 0) return false;
      if (/^([hms])\1*$/i.test(section.slice(i + 1, end))) return true;
      i = end + 1;
    } else {
      i++;
    }
  }
  return false;
}

/** Whether one format section (§18.8.30) is a date/time format. The body is
 *  scanned left to right so each token is read in context: `\x` escapes,
 *  `_x` padding and `*x` fill pairs (whose operand may itself be `"`), quoted
 *  literals and bracket content are skipped; what remains is date/time when
 *  it holds y / m / d / h / s, AM/PM or A/P, the Japanese weekday code
 *  `aaa+`, the Japanese era `g`, era year `e` and locale codes `r` / `rr`,
 *  or an elapsed-time bracket `[h]` / `[mm]` / `[ss]`. An `e` after a numeric
 *  placeholder (`0` `#` `?`) or before `+` / `-` / a placeholder is the
 *  scientific exponent, and `General` is the General keyword. */
export function isDateFormatSection(body: string): boolean {
  let i = 0;
  let afterPlaceholder = false;
  while (i < body.length) {
    const ch = body[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = body.indexOf('"', i + 1);
      i = end < 0 ? body.length : end + 1;
    } else if (ch === '[') {
      const end = body.indexOf(']', i);
      if (end < 0) return false;
      if (/^([hms])\1*$/i.test(body.slice(i + 1, end))) return true;
      i = end + 1;
    } else if (/^general/i.test(body.slice(i))) {
      i += 'general'.length;
    } else if (/[ymdhsgr]/i.test(ch)) {
      return true;
    } else if (ch === 'e' || ch === 'E') {
      if (!afterPlaceholder && !/[-+0#?]/.test(body[i + 1] ?? '')) return true;
      i++;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      afterPlaceholder = true;
      i++;
    } else if (/^(am\/pm|a\/p|a{3,})/i.test(body.slice(i))) {
      return true;
    } else {
      i++;
    }
  }
  return false;
}

// Shared Excel serial-date → JS `Date` conversion (ECMA-376 §18.17.4.1).
//
// A "serial date-time" is a signed number of days relative to a base date. Two
// base dates exist, selected by the `<workbookPr date1904>` attribute
// (§18.2.28) for cells and by `<c:date1904>` (§21.2.2.38) for charts:
//
//   • 1900 date system (default): base date is 1899-12-30 (serial 0), so
//     serial 1 → 1900-01-01. Excel additionally reproduces the Lotus 1-2-3
//     leap-year bug, treating the non-existent 1900-02-29 as serial 60. That
//     phantom day shifts every serial ≥ 60 forward by one relative to the
//     proleptic Gregorian calendar, so to recover the correct calendar date we
//     add one day back for serials < 60 (which sit before the phantom day and
//     are therefore off by one under a bug-free epoch). Serial 60 itself has no
//     real Gregorian date; we leave it unshifted (→ 1900-02-28), matching the
//     long-standing core behaviour rather than inventing a mapping.
//
//     The +1 compensation for serial < 60 is a deliberate choice to match
//     Excel's on-screen rendering, not the literal base-date arithmetic of
//     §18.17.4.1 (whose worked example puts serial 1.5 at 1899-12-31T12:00,
//     which would render serial 1 as 1899-12-31). We render serial 1 as
//     1900-01-01 because that is what Excel displays; the spec's base-date
//     formula and Excel's display diverge for serials below the phantom leap
//     day, and we follow Excel.
//
//   • 1904 date system: base date is 1904-01-01 (serial 0), so serial 1 →
//     1904-01-02. This system has no leap-year bug. The 1904 base is exactly
//     1462 days after the 1900 base date, which is why a Mac-authored (1904)
//     workbook shows dates 1462 days off when read with a 1900 epoch.
//
// Times are carried by the fractional part of the serial (§18.17.4.2): 0.5 =
// noon. All arithmetic is done in UTC so the caller reads `getUTC*` accessors
// without local-timezone drift.

const MS_PER_DAY = 86_400_000;

// Base dates as UTC epoch milliseconds.
//   1900: 1899-12-30 (serial 0)
//   1904: 1904-01-01 (serial 0)
const BASE_1900_MS = Date.UTC(1899, 11, 30);
const BASE_1904_MS = Date.UTC(1904, 0, 1);

/** Calendar/clock display admission for the Excel-compatible date systems.
 * Annex L.2.16.9.1 defines their last day; §18.17.4.2 allows time fractions
 * below 1. Library policy preserves serial zero (including existing 1900
 * rendering), while negative calendar/clock values use a width-independent
 * '#' marker. This is a display policy, not the wider signed temporal range
 * in §18.17.4.1, and does not restrict elapsed durations or date arithmetic.
 */
export function isExcelDisplayDateSerial(serial: number, date1904 = false): boolean {
  return Number.isFinite(serial) && serial >= 0 && serial < (date1904 ? 2_957_004 : 2_958_466);
}

function roundSerialDate(base: number, adjusted: number, serial: number, date1904: boolean): number {
  const ms = base + Math.round(adjusted * MS_PER_DAY);
  // A valid final-day fraction can round to midnight of year 10000. Saturate
  // only that rounding carry at the supported final millisecond; ordinary
  // dates still carry midnight, and out-of-range arithmetic is untouched.
  return isExcelDisplayDateSerial(serial, date1904)
    ? Math.min(ms, Date.UTC(10000, 0, 1) - 1)
    : ms;
}

/**
 * Convert an Excel date-time serial to a UTC `Date`.
 *
 * @param serial   Signed serial date-time (integer part = day, fraction = time).
 * @param date1904 `true` for the 1904 date system (Mac-authored workbooks /
 *                 `<c:date1904 val="1">`); `false`/omitted for the default 1900
 *                 date system.
 */
export function excelSerialToUtcDate(serial: number, date1904 = false): Date {
  // `new Date` truncates fractional milliseconds, and a stored time fraction
  // is rarely exact in binary (8:00 is 0.33333333333333331, a hair under
  // 28 800 000 ms), so round to the nearest millisecond first.
  if (date1904) {
    return new Date(roundSerialDate(BASE_1904_MS, serial, serial, date1904));
  }
  // 1900 system: apply the Lotus leap-year-bug compensation. Serials < 60 sit
  // before the phantom 1900-02-29 and are one day short under the bug-free
  // 1899-12-30 epoch, so add a day back. Serials ≥ 60 need no adjustment.
  const adjusted = serial < 60 ? serial + 1 : serial;
  return new Date(roundSerialDate(BASE_1900_MS, adjusted, serial, date1904));
}

/**
 * Convert a UTC `Date` back to an Excel date-time serial — the exact inverse of
 * {@link excelSerialToUtcDate} for every serial except the phantom 60.
 *
 * The fractional (time-of-day) part is preserved. For the 1900 date system the
 * Lotus leap-year-bug compensation is undone so the round-trip holds:
 *
 *   • Dates on/before 1900-02-28 (raw offset ≤ 60 days from the 1899-12-30
 *     base) map to serials 0…59 by subtracting the +1 day that
 *     `excelSerialToUtcDate` added — 1900-01-01 → 1, 1900-02-28 → 59.
 *   • Dates on/after 1900-03-01 (raw offset ≥ 61) map straight through —
 *     1900-03-01 → 61, and so on.
 *   • The phantom serial 60 (the non-existent 1900-02-29) is never produced:
 *     the two branches meet at raw offset 60 → serial 59 and 61 → serial 61,
 *     skipping 60 entirely.
 *
 * @param date     A UTC `Date` (read via its epoch milliseconds).
 * @param date1904 `true` for the 1904 date system; `false`/omitted for 1900.
 */
export function utcDateToExcelSerial(date: Date, date1904 = false): number {
  if (date1904) {
    return (date.getTime() - BASE_1904_MS) / MS_PER_DAY;
  }
  // 1900 system: `rawDays` is the offset from the 1899-12-30 base, i.e. the
  // "adjusted" value the forward function used. Raw offsets ≤ 60 correspond to
  // serials < 60 (which had +1 added), so undo it; offsets ≥ 61 pass through.
  // The boundary at 60 (→ serial 59) / 61 (→ serial 61) skips the phantom 60.
  const rawDays = (date.getTime() - BASE_1900_MS) / MS_PER_DAY;
  return rawDays <= 60 ? rawDays - 1 : rawDays;
}

import { describe, it, expect } from 'vitest';
import { excelSerialToUtcDate, utcDateToExcelSerial } from './excel-date';

// Helper: read back the calendar date a serial maps to, as YYYY-MM-DD in UTC.
function iso(serial: number, date1904: boolean): string {
  const d = excelSerialToUtcDate(serial, date1904);
  return `${d.getUTCFullYear().toString().padStart(4, '0')}-${(d.getUTCMonth() + 1)
    .toString()
    .padStart(2, '0')}-${d.getUTCDate().toString().padStart(2, '0')}`;
}

describe('excelSerialToUtcDate — 1900 date system (ECMA-376 §18.17.4.1)', () => {
  it('maps serial 1 to 1900-01-01 (base date 1899-12-30 is serial 0)', () => {
    expect(iso(1, false)).toBe('1900-01-01');
  });

  it('applies the Lotus 1900-02-29 leap-year bug: serial < 60 is shifted +1 day', () => {
    // Serial 59 is the last serial before the phantom 1900-02-29. Without the
    // +1 compat adjustment a naive epoch would render it 1900-02-27; Excel
    // renders 1900-02-28.
    expect(iso(59, false)).toBe('1900-02-28');
  });

  it('renders serial 61 as 1900-03-01 (the day after the phantom leap day)', () => {
    // Serial 60 is the phantom 1900-02-29 (no adjustment); serial 61 is the
    // first serial ≥ 60 and must land on 1900-03-01.
    expect(iso(61, false)).toBe('1900-03-01');
  });

  it('serial 60 keeps the pre-existing core behaviour (no +1 shift, → 1900-02-28)', () => {
    // 1900-02-29 does not exist in the proleptic Gregorian calendar. The
    // existing core `formatExcelDate` treated serials ≥ 60 without the +1
    // shift, so serial 60 lands on 1900-02-28. We deliberately preserve that.
    expect(iso(60, false)).toBe('1900-02-28');
  });

  it('maps a modern serial correctly (45292 → 2024-01-01)', () => {
    expect(iso(45292, false)).toBe('2024-01-01');
  });

  it('preserves the fractional time-of-day component', () => {
    // Serial 1.5 = midday on 1900-01-01.
    const d = excelSerialToUtcDate(1.5, false);
    expect(d.getUTCFullYear()).toBe(1900);
    expect(d.getUTCMonth() + 1).toBe(1);
    expect(d.getUTCDate()).toBe(1);
    expect(d.getUTCHours()).toBe(12);
  });

  it('rounds a stored time fraction to the nearest millisecond', () => {
    // 8:00 is stored as 0.33333333333333331 (28 799 999.999… ms); `new Date`
    // would truncate it to 7:59:59.999.
    const d = excelSerialToUtcDate(45292.33333333333, false);
    expect(d.getUTCHours()).toBe(8);
    expect(d.getUTCMinutes()).toBe(0);
    expect(excelSerialToUtcDate(45000.33333333333, true).getUTCHours()).toBe(8);
  });
});

describe('excelSerialToUtcDate — 1904 date system (ECMA-376 §18.17.4.1)', () => {
  it('maps serial 0 to the 1904 base date 1904-01-01', () => {
    expect(iso(0, true)).toBe('1904-01-01');
  });

  it('maps serial 1 to 1904-01-02 (no leap-year bug in the 1904 system)', () => {
    expect(iso(1, true)).toBe('1904-01-02');
  });

  it('is offset exactly 1462 days from the 1900 system for the same calendar date', () => {
    // 2024-01-01 is serial 45292 in the 1900 system and 43830 in the 1904
    // system. 45292 − 43830 = 1462 (4 years + 1 day = the classic Mac shift).
    expect(iso(43830, true)).toBe('2024-01-01');
    expect(45292 - 43830).toBe(1462);
  });

  it('preserves the fractional time-of-day component', () => {
    const d = excelSerialToUtcDate(0.5, true);
    expect(d.getUTCFullYear()).toBe(1904);
    expect(d.getUTCMonth() + 1).toBe(1);
    expect(d.getUTCDate()).toBe(1);
    expect(d.getUTCHours()).toBe(12);
  });
});

describe('utcDateToExcelSerial — inverse of excelSerialToUtcDate', () => {
  it('maps the 1900 leap-bug boundary dates to serials 1/59/61 (never 60)', () => {
    const serial = (y: number, m: number, d: number): number =>
      utcDateToExcelSerial(new Date(Date.UTC(y, m - 1, d)), false);
    expect(serial(1900, 1, 1)).toBe(1);   // base + 1 day, serial < 60 branch
    expect(serial(1900, 2, 28)).toBe(59); // last day before the phantom leap day
    expect(serial(1900, 3, 1)).toBe(61);  // first day after it — skips serial 60
  });

  it('round-trips every 1900-system serial except the phantom 60', () => {
    // serial → date → serial must be the identity for all valid serials. Serial
    // 60 has no real Gregorian date so it is deliberately excluded.
    for (let s = 0; s <= 200; s++) {
      if (s === 60) continue;
      const back = utcDateToExcelSerial(excelSerialToUtcDate(s, false), false);
      expect(back).toBe(s);
    }
    // A modern serial round-trips too.
    expect(utcDateToExcelSerial(excelSerialToUtcDate(45292, false), false)).toBe(45292);
  });

  it('the phantom serial 60 is never emitted by the inverse', () => {
    // Both the serial-59 date (1900-02-28) and the serial-61 date (1900-03-01)
    // sit on either side of the phantom day; neither reverse-maps to 60.
    expect(utcDateToExcelSerial(excelSerialToUtcDate(59, false), false)).toBe(59);
    expect(utcDateToExcelSerial(excelSerialToUtcDate(61, false), false)).toBe(61);
  });

  it('round-trips 1904-system serials (no leap-year bug)', () => {
    for (let s = 0; s <= 200; s++) {
      const back = utcDateToExcelSerial(excelSerialToUtcDate(s, true), true);
      expect(back).toBe(s);
    }
    expect(utcDateToExcelSerial(excelSerialToUtcDate(43830, true), true)).toBe(43830);
  });

  it('preserves the fractional time-of-day component', () => {
    // Serial 45292.5 = noon on 2024-01-01. The reverse must carry the .5 back.
    const back = utcDateToExcelSerial(excelSerialToUtcDate(45292.5, false), false);
    expect(back).toBe(45292.5);
  });
});

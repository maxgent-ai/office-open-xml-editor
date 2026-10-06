import { describe, expect, it } from 'vitest';
import {
  DEFAULT_XLSX_WORKSHEET_POLICY,
  normalizeXlsxWorksheetPolicy,
} from './xlsx-worksheet-policy.js';

const FIELDS = ['maxRows', 'maxCells', 'maxOwnedUtf8Bytes', 'maxJsonBytes'] as const;
const LARGEST_ACCEPTED = Number.MAX_SAFE_INTEGER - 1;

const EXPECTED_DEFAULT = {
  worksheet: {
    maxRows: 100_000,
    maxCells: 250_000,
    maxOwnedUtf8Bytes: 33_554_432,
    maxJsonBytes: 67_108_864,
  },
  cache: {
    maxRows: 200_000,
    maxCells: 500_000,
    maxOwnedUtf8Bytes: 67_108_864,
    maxJsonBytes: 134_217_728,
  },
  maxCoordinateIndexEntries: 250_000,
};

function normalize(limits: unknown) {
  return normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: limits as never });
}

function expectDeepFrozen(policy: unknown): void {
  const value = policy as { worksheet: object; cache: object };
  expect(Object.isFrozen(value)).toBe(true);
  expect(Object.isFrozen(value.worksheet)).toBe(true);
  expect(Object.isFrozen(value.cache)).toBe(true);
}

describe('normalizeXlsxWorksheetPolicy', () => {
  it('resolves omitted, undefined and empty limits to the frozen published defaults', () => {
    expect(DEFAULT_XLSX_WORKSHEET_POLICY).toEqual(EXPECTED_DEFAULT);
    expectDeepFrozen(DEFAULT_XLSX_WORKSHEET_POLICY);

    const allUndefined = Object.fromEntries(FIELDS.map((field) => [field, undefined]));
    for (const resolved of [
      normalizeXlsxWorksheetPolicy({}),
      normalize(undefined),
      normalize({}),
      normalize(allUndefined),
    ]) {
      expect(resolved).toEqual(EXPECTED_DEFAULT);
      expectDeepFrozen(resolved);
    }
  });

  it('applies explicit worksheet limits while keeping cache and coordinate-index floors', () => {
    // Partial override merges with the remaining defaults.
    expect(normalize({ maxRows: 7 })).toEqual({
      ...EXPECTED_DEFAULT,
      worksheet: { ...EXPECTED_DEFAULT.worksheet, maxRows: 7 },
    });

    // Lower explicit limits are honoured; cache/index keep their default floors.
    const lower = normalize({ maxRows: 1, maxCells: 2, maxOwnedUtf8Bytes: 3, maxJsonBytes: 4 });
    expect(lower).toEqual({
      worksheet: { maxRows: 1, maxCells: 2, maxOwnedUtf8Bytes: 3, maxJsonBytes: 4 },
      cache: EXPECTED_DEFAULT.cache,
      maxCoordinateIndexEntries: 250_000,
    });
    expectDeepFrozen(lower);

    // Higher limits raise cache and index to at least the worksheet limits;
    // there is no spreadsheet-grid maximum on rows/cells.
    const higher = normalize({
      maxRows: 2_000_000,
      maxCells: 20_000_000_000,
      maxOwnedUtf8Bytes: 100_000_000,
      maxJsonBytes: 300_000_000,
    });
    expect(higher).toEqual({
      worksheet: {
        maxRows: 2_000_000,
        maxCells: 20_000_000_000,
        maxOwnedUtf8Bytes: 100_000_000,
        maxJsonBytes: 300_000_000,
      },
      cache: {
        maxRows: 2_000_000,
        maxCells: 20_000_000_000,
        maxOwnedUtf8Bytes: 100_000_000,
        maxJsonBytes: 300_000_000,
      },
      maxCoordinateIndexEntries: 20_000_000_000,
    });

    // Only one field above its cache floor raises only that cache field.
    const mixed = normalize({ maxCells: 600_000 });
    expect(mixed.cache).toEqual({ ...EXPECTED_DEFAULT.cache, maxCells: 600_000 });
    expect(mixed.maxCoordinateIndexEntries).toBe(600_000);

    // MAX_SAFE_INTEGER - 1 is accepted for every counter (no arbitrary byte cap)
    // and its +1 overflow sentinel stays exactly representable.
    const largest = normalize(
      Object.fromEntries(FIELDS.map((field) => [field, LARGEST_ACCEPTED])),
    );
    for (const field of FIELDS) {
      expect(largest.worksheet[field]).toBe(LARGEST_ACCEPTED);
      expect(largest.cache[field]).toBe(LARGEST_ACCEPTED);
      expect(Number.isSafeInteger(largest.worksheet[field] + 1)).toBe(true);
      expect(largest.worksheet[field] + 1).toBe(Number.MAX_SAFE_INTEGER);
    }
    expect(largest.maxCoordinateIndexEntries).toBe(LARGEST_ACCEPTED);
    expect(largest.maxCoordinateIndexEntries + 1).toBe(Number.MAX_SAFE_INTEGER);
  });

  it('throws TypeError for non-object limits and non-number field values', () => {
    for (const shape of [null, [], [1], 42, 'limits', true, () => undefined]) {
      expect(() => normalize(shape), String(shape)).toThrow(TypeError);
    }
    for (const field of FIELDS) {
      for (const value of [null, '10', '', true, false, {}, [10]]) {
        expect(
          () => normalize({ [field]: value }),
          `${field}=${JSON.stringify(value)}`,
        ).toThrow(TypeError);
      }
    }
  });

  it('throws RangeError for numeric values that are not finite positive safe integers within range', () => {
    const invalid = [
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      0,
      -0,
      -1,
      0.5,
      1.5,
      Number.MAX_SAFE_INTEGER,
      Number.MAX_SAFE_INTEGER + 1,
      Number.MAX_VALUE,
    ];
    for (const field of FIELDS) {
      for (const value of invalid) {
        let thrown: unknown;
        try {
          normalize({ [field]: value });
        } catch (error) {
          thrown = error;
        }
        expect(thrown, `${field}=${String(value)}`).toBeInstanceOf(RangeError);
      }
    }
  });

  it('does not mutate inputs or the default policy and keeps sessions independent', () => {
    const defaultSnapshot = JSON.parse(JSON.stringify(DEFAULT_XLSX_WORKSHEET_POLICY));

    const frozenInput = Object.freeze({ maxRows: 10, maxCells: 20 });
    const first = normalize(frozenInput);
    expect(frozenInput).toEqual({ maxRows: 10, maxCells: 20 });

    const mutableInput: Record<string, number> = { maxRows: 300_000, maxJsonBytes: 5 };
    const second = normalize(mutableInput);
    mutableInput.maxRows = 1;
    mutableInput.maxCells = 1;
    expect(mutableInput).toEqual({ maxRows: 1, maxJsonBytes: 5, maxCells: 1 });

    expect(first).toEqual({
      worksheet: { ...EXPECTED_DEFAULT.worksheet, maxRows: 10, maxCells: 20 },
      cache: EXPECTED_DEFAULT.cache,
      maxCoordinateIndexEntries: 250_000,
    });
    expect(second).toEqual({
      worksheet: { ...EXPECTED_DEFAULT.worksheet, maxRows: 300_000, maxJsonBytes: 5 },
      cache: { ...EXPECTED_DEFAULT.cache, maxRows: 300_000 },
      maxCoordinateIndexEntries: 250_000,
    });
    expect(first.worksheet).not.toBe(second.worksheet);
    expect(first.cache).not.toBe(second.cache);

    // Resolved values cannot be altered by one session to influence another.
    expect(() => {
      (first.worksheet as { maxRows: number }).maxRows = 999;
    }).toThrow(TypeError);
    expect(() => {
      (first.cache as { maxCells: number }).maxCells = 999;
    }).toThrow(TypeError);
    expect(first.worksheet.maxRows).toBe(10);
    expect(second.worksheet.maxRows).toBe(300_000);

    expect(DEFAULT_XLSX_WORKSHEET_POLICY).toEqual(defaultSnapshot);
    expect(normalizeXlsxWorksheetPolicy({})).toEqual(EXPECTED_DEFAULT);
  });
});

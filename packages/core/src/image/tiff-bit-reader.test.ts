import { describe, expect, it } from 'vitest';
import { TiffBitReader } from './tiff.js';
import { TiffDecodeError } from './tiff-contract.js';

describe('TIFF strip bit reader', () => {
  it.each([2 ** 28, 2 ** 29, 2 ** 31, 2 ** 32])(
    'reads across byte position %i without signed or unsigned 32-bit wrap',
    (boundary) => {
      // Sparse byte storage exercises large cursor positions without allocating
      // a large strip. These bytes pack 9-bit ClearCode, literal 77, and EOI.
      const start = boundary - 1;
      const bytes: ArrayLike<number> = {
        length: start + 5,
        [start]: 0x80,
        [start + 1]: 0x13,
        [start + 2]: 0x60,
        [start + 3]: 0x20,
        [start + 4]: 0xff, // Outside the declared strip.
      };
      const reader = new TiffBitReader(bytes, start, 4, 'truncated strip');
      expect([reader.readBits(9), reader.readBits(9), reader.readBits(9)])
        .toEqual([256, 77, 257]);
      expect(Array.from({ length: 5 }, () => reader.readBit())).toEqual([0, 0, 0, 0, 0]);
      expect(() => reader.readBit()).toThrowError(new TiffDecodeError('truncated strip'));
    },
  );

  it('reads changing code widths and leaves the cursor intact on truncation', () => {
    // Independent bit string: 9 ones, 10 zeros, 11 ones, 12 alternating bits.
    const bytes = new Uint8Array([0xff, 0x80, 0x1f, 0xfe, 0xaa, 0x80]);
    const reader = new TiffBitReader(bytes, 0, bytes.length, 'truncated strip');
    expect([reader.readBits(9), reader.readBits(10), reader.readBits(11), reader.readBits(12)])
      .toEqual([511, 0, 2047, 0b101010101010]);
    expect(() => reader.readBits(9)).toThrowError('truncated strip');
    expect(reader.readBits(6)).toBe(0);
  });
});

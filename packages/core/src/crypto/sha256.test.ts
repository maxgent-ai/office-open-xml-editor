import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { sha256 } from './sha256';

const hex = (bytes: Uint8Array) => Buffer.from(bytes).toString('hex');

describe('sha256', () => {
  it('matches the platform digest across every padding boundary', () => {
    // Lengths 0..130 cover the 55/56-byte single-block limit and the
    // one-, two- and three-block paddings.
    for (let length = 0; length <= 130; length++) {
      const bytes = Uint8Array.from({ length }, (_, index) => (index * 131 + length) & 0xff);
      expect(hex(sha256(bytes)), `length ${length}`)
        .toBe(createHash('sha256').update(bytes).digest('hex'));
    }
  });
});

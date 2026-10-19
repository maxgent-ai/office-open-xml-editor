import { describe, expect, it } from 'vitest';
import { profileHeight, profileSlope } from './bevel-profiles';

describe('bevel profiles', () => {
  it('falls back to circle for unknown and prototype-like preset names', () => {
    for (const name of ['unknown', '__proto__', 'toString', 'constructor', 'hasOwnProperty']) {
      for (const t of [0, 0.25, 0.5, 1]) {
        expect(profileHeight(name, t)).toBe(profileHeight('circle', t));
        expect(profileSlope(name, t)).toBe(profileSlope('circle', t));
      }
    }
    expect(profileHeight('angle', 0.5)).not.toBe(profileHeight('circle', 0.5));
  });
});

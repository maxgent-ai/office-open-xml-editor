import { describe, expect, it } from 'vitest';
import { joiningType, joinsToFollowing, joinsToPreceding } from './joining-type.js';

describe('joiningType (UCD Joining_Type)', () => {
  it('classifies the canonical Arabic letters', () => {
    expect(joiningType(0x0640)).toBe('C'); // TATWEEL
    expect(joiningType(0x0628)).toBe('D'); // BEH
    expect(joiningType(0x0644)).toBe('D'); // LAM
    expect(joiningType(0x0627)).toBe('R'); // ALEF
    expect(joiningType(0x0631)).toBe('R'); // REH
    expect(joiningType(0x0648)).toBe('R'); // WAW
    expect(joiningType(0x064b)).toBe('T'); // FATHA
  });
  it('defaults non-joining scripts to U (Non_Joining)', () => {
    expect(joiningType(0x0041)).toBe('U'); // Latin A
    expect(joiningType(0x0020)).toBe('U'); // space
    expect(joiningType(0x3042)).toBe('U'); // あ (Hiragana)
  });
});

describe('joinsToFollowing / joinsToPreceding', () => {
  it('following-join is true for D / L / C only', () => {
    expect(joinsToFollowing(0x0628)).toBe(true); // BEH (D)
    expect(joinsToFollowing(0x0644)).toBe(true); // LAM (D)
    expect(joinsToFollowing(0x0640)).toBe(true); // TATWEEL (C)
    expect(joinsToFollowing(0x0627)).toBe(false); // ALEF (R)
    expect(joinsToFollowing(0x0631)).toBe(false); // REH (R)
    expect(joinsToFollowing(0x0041)).toBe(false); // Latin (U)
  });
  it('preceding-join is true for D / R / C only', () => {
    expect(joinsToPreceding(0x0628)).toBe(true); // BEH (D)
    expect(joinsToPreceding(0x0627)).toBe(true); // ALEF (R)
    expect(joinsToPreceding(0x0631)).toBe(true); // REH (R)
    expect(joinsToPreceding(0x0640)).toBe(true); // TATWEEL (C)
    expect(joinsToPreceding(0x0041)).toBe(false); // Latin (U)
  });
});

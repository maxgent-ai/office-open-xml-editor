import { describe, it, expect } from 'vitest';
import {
  hex6ToRgb,
  luminance601,
  duotoneImageData,
  type RgbaBuffer,
} from './duotone.js';

// ── Duotone effect (ECMA-376 §20.1.8.23) unit tests ──────────────────────────
// The pixel remap is a pure function over an ImageData-shaped buffer, so it is
// exercised here with hand-built RGBA arrays — no canvas needed. clr1 is the
// dark endpoint (luminance 0), clr2 the light endpoint (luminance 1).

describe('hex6ToRgb', () => {
  it('parses a 6-char hex to a byte triple', () => {
    expect(hex6ToRgb('DAB6BA')).toEqual([0xda, 0xb6, 0xba]);
    expect(hex6ToRgb('000000')).toEqual([0, 0, 0]);
    expect(hex6ToRgb('ffffff')).toEqual([255, 255, 255]);
  });
  it('rejects malformed input', () => {
    expect(hex6ToRgb('#DAB6BA')).toBeNull();
    expect(hex6ToRgb('DAB6B')).toBeNull();
    expect(hex6ToRgb('DAB6BAA')).toBeNull();
    expect(hex6ToRgb('GGGGGG')).toBeNull();
  });
});

describe('luminance601', () => {
  it('returns 0 for black and 1 for white', () => {
    expect(luminance601(0, 0, 0)).toBe(0);
    expect(luminance601(255, 255, 255)).toBeCloseTo(1, 6);
  });
  it('uses Rec. 601 weights', () => {
    // Pure green is the heaviest channel (0.587).
    expect(luminance601(0, 255, 0)).toBeCloseTo(0.587, 6);
    expect(luminance601(255, 0, 0)).toBeCloseTo(0.299, 6);
    expect(luminance601(0, 0, 255)).toBeCloseTo(0.114, 6);
  });
});

/** Build a 1×N RGBA buffer from `[r,g,b,a]` tuples. */
function buf(pixels: Array<[number, number, number, number]>): RgbaBuffer {
  const data = new Uint8ClampedArray(pixels.length * 4);
  pixels.forEach((p, i) => {
    data[i * 4] = p[0];
    data[i * 4 + 1] = p[1];
    data[i * 4 + 2] = p[2];
    data[i * 4 + 3] = p[3];
  });
  return { data, width: pixels.length, height: 1 };
}

describe('duotoneImageData', () => {
  const DARK = '000000'; // clr1
  const LIGHT = 'DAB6BA'; // clr2 (light pink), matching sample-9.xlsx

  it('maps a black pixel (t=0) to clr1', () => {
    const b = buf([[0, 0, 0, 255]]);
    duotoneImageData(b, DARK, LIGHT);
    expect([b.data[0], b.data[1], b.data[2]]).toEqual([0, 0, 0]);
    expect(b.data[3]).toBe(255); // alpha preserved
  });

  it('maps a white pixel (t=1) to clr2', () => {
    const b = buf([[255, 255, 255, 255]]);
    duotoneImageData(b, DARK, LIGHT);
    expect([b.data[0], b.data[1], b.data[2]]).toEqual([0xda, 0xb6, 0xba]);
    expect(b.data[3]).toBe(255);
  });

  it('maps a mid-grey (t=0.5) to the midpoint of the ramp', () => {
    // luminance601(128,128,128) = 128/255 ≈ 0.502.
    const b = buf([[128, 128, 128, 255]]);
    duotoneImageData(b, DARK, LIGHT);
    const t = 128 / 255;
    expect(b.data[0]).toBe(Math.round(0 + (0xda - 0) * t));
    expect(b.data[1]).toBe(Math.round(0 + (0xb6 - 0) * t));
    expect(b.data[2]).toBe(Math.round(0 + (0xba - 0) * t));
  });

  it('preserves per-pixel alpha while recoloring every pixel', () => {
    const b = buf([
      [255, 255, 255, 128], // semi-transparent white → clr2, alpha kept
      [0, 0, 0, 0], // fully transparent black → dark endpoint
    ]);
    duotoneImageData(b, 'FF0000', LIGHT);
    expect([b.data[0], b.data[1], b.data[2], b.data[3]]).toEqual([0xda, 0xb6, 0xba, 128]);
    expect([b.data[4], b.data[5], b.data[6], b.data[7]]).toEqual([255, 0, 0, 0]);
  });

  it('respects colour role order (swapping endpoints inverts the ramp)', () => {
    const white = buf([[255, 255, 255, 255]]);
    duotoneImageData(white, LIGHT, DARK); // clr1=pink, clr2=black
    // t=1 now maps to clr2=black.
    expect([white.data[0], white.data[1], white.data[2]]).toEqual([0, 0, 0]);
  });

  it('is a no-op when a colour is malformed', () => {
    const b = buf([[123, 45, 67, 255]]);
    duotoneImageData(b, 'nothex', LIGHT);
    expect([b.data[0], b.data[1], b.data[2]]).toEqual([123, 45, 67]);
  });

  it('near-white sample-9 corner (RGB 243–249) lands in the pink range', () => {
    const b = buf([[246, 246, 246, 255]]);
    duotoneImageData(b, DARK, LIGHT);
    // ~96% up the ramp toward DAB6BA → clearly pink, R>G, R>B, all high.
    expect(b.data[0]).toBeGreaterThan(200); // R near 0xDA
    expect(b.data[0]).toBeGreaterThan(b.data[1]); // pink: R>G
    expect(b.data[0]).toBeGreaterThan(b.data[2]); // pink: R>B
  });
});

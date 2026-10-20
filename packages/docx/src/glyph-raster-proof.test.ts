import { afterAll, describe, expect, it, vi } from 'vitest';
import { ownsGlyphs } from './glyph-raster-proof.js';

// The coverage probe face answers zero advance and zero ink for the glyphs it
// owns; everything else comes from a fallback face with real metrics.
const OWNED = new Set(['国', 'あ']);

class StubCanvas {
  constructor(public width: number, public height: number) {}
  getContext() {
    return {
      font: '',
      measureText: (text: string) => {
        const owned = OWNED.has(text);
        return {
          width: owned ? 0 : 25,
          actualBoundingBoxLeft: 0,
          actualBoundingBoxRight: owned ? 0 : 20,
          actualBoundingBoxAscent: owned ? 0 : 70,
          actualBoundingBoxDescent: 0,
        };
      },
    };
  }
}

vi.stubGlobal('OffscreenCanvas', StubCanvas);
afterAll(() => vi.unstubAllGlobals());

const request = (text: string) => ({
  text, coverageProbeFamily: 'probe', weight: 700, style: 'normal' as const,
});

describe('ownsGlyphs (LIBRARY_SYNTHESIZED_FACE_LINE_GEOMETRY)', () => {
  it('owns text whose every grapheme is the probe face', () => {
    expect(ownsGlyphs(request('国あ'))).toBe(true);
  });

  it('declines whitespace the face does not own', () => {
    expect(ownsGlyphs(request('国 あ'))).toBe(false);
  });

  it('declines empty text', () => {
    expect(ownsGlyphs(request(''))).toBe(false);
  });
});

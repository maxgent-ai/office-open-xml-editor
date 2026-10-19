import { afterEach, describe, expect, it, vi } from 'vitest';
import { classifyCjkFont, googleCjkFontAlias } from '../fonts/scripts';
import { buildPresetGeometryPath, hasPreset } from '../shape/preset-geometry/geometry';
import { hasTextWarp } from '../shape/text-warp';
import { resolveRegionMapFeature } from '../chart/region-map-renderer';
import { docxBorderDashArray, pptxPresetDashArray } from '../draw/dash';
import { GoogleFontPreloadLease } from '../fonts/preload';
import { GOOGLE_FONT_SUBSTITUTES } from '../fonts/google-fonts';

// Document strings that name Object.prototype members must miss every lookup
// table, exactly like any other unknown name, instead of resolving to the
// inherited function/object.
const PROTO_NAMES = ['constructor', '__proto__', 'toString', 'hasOwnProperty'];

describe('lookup tables keyed by document strings', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('treats prototype member names as unknown fonts, presets, warps, labels and styles', () => {
    const sink = new Proxy({}, { get: () => () => {} });
    for (const name of PROTO_NAMES) {
      expect(googleCjkFontAlias(name)).toBeNull();
      expect(classifyCjkFont(name)).toBeNull();
      expect(hasPreset(name)).toBe(false);
      expect(buildPresetGeometryPath(sink as never, name, 0, 0, 10, 10)).toBe(false);
      expect(hasTextWarp(name)).toBe(false);
      expect(resolveRegionMapFeature(name)).toBeUndefined();
      expect(docxBorderDashArray(name, 1)).toEqual([]);
      expect(pptxPresetDashArray(name, 1)).toEqual([]);
    }
  });

  it('never fetches a stylesheet for a prototype member font name', async () => {
    const fetch = vi.fn(async () => ({ ok: false, status: 404 }));
    vi.stubGlobal('fetch', fetch);
    vi.stubGlobal('FontFace', class {});
    const fonts = { add() {}, delete() {}, [Symbol.iterator]: function* () {} } as never;
    const lease = new GoogleFontPreloadLease(GOOGLE_FONT_SUBSTITUTES, fonts);
    await lease.ensure(PROTO_NAMES);
    expect(fetch).not.toHaveBeenCalled();
  });
});

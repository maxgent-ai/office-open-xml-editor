import { describe, expect, it, vi } from 'vitest';
import { decodeDib, decodePackedDib } from './dib.js';

vi.mock('./pixel-budget.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pixel-budget.js')>()),
  MAX_RASTER_DIMENSION: 16,
  MAX_RASTER_PIXELS: 32,
}));

/** A complete top-down 32bpp bitmap, or a header with a short pixel payload. */
function bitmap(width: number, height: number, bits = width * height * 4): DataView {
  const view = new DataView(new ArrayBuffer(40 + bits));
  view.setUint32(0, 40, true);
  view.setInt32(4, width, true);
  view.setInt32(8, -height, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 32, true);
  return view;
}

describe('DIB decoded-surface admission', () => {
  it('does not use bytes outside an explicit pixel span as proof of a surface', () => {
    const view = bitmap(17, 2);
    // The complete containing record has other data after four bitmap bytes.
    expect(decodeDib(view, 0, 40, 40, 4)).toBeNull();
  });

  it.each([
    ['explicit offsets (EMF)', (view: DataView) => decodeDib(view, 0, 40, 40, view.byteLength - 40)],
    ['packed bitmap (WMF)', (view: DataView) => decodePackedDib(view, 0, view.byteLength)],
  ] as const)('validates payloads before admitting exact boundaries with %s', (_name, decode) => {
    expect(decode(bitmap(16, 1))?.data).toHaveLength(64);
    expect(decode(bitmap(1, 16))?.data).toHaveLength(64);
    expect(decode(bitmap(8, 4))?.data).toHaveLength(128);
    for (const [width, height, metric, limit, observed] of [
      [17, 1, 'image-dimension', 16, 17],
      [1, 17, 'image-dimension', 16, 17],
      [11, 3, 'image-pixels', 32, 33],
    ] as const) {
      expect(() => decode(bitmap(width, height))).toThrowError(expect.objectContaining({
        code: 'ooxml-decoded-image-limit', metric, limit, observed,
      }));
    }
    // This multi-row payload was discarded by the old truncation check. Its
    // impossible dimensions must not turn that fallback into a budget error.
    expect(decode(bitmap(17, 2, 4))).toBeNull();
    expect(decode(bitmap(11, 3, 4))).toBeNull();
    // Even a single missing byte/row is not proof of an oversized surface.
    expect(decode(bitmap(17, 1, 17 * 4 - 1))).toBeNull();
    expect(decode(bitmap(11, 3, 11 * 3 * 4 - 1))).toBeNull();
    // Keep the historical partial-row tolerance for admitted surfaces.
    expect(decode(bitmap(16, 1, 16 * 4 - 1))?.data).toHaveLength(64);
  });

  it.each([8, 32])('requires declared color-table bytes before classifying an oversized %ibpp surface', (depth) => {
    const indexed = (width: number) => {
      const view = new DataView(new ArrayBuffer(48 + Math.ceil(width * depth / 32) * 4));
      view.setUint32(0, 40, true);
      view.setInt32(4, width, true);
      view.setInt32(8, 1, true);
      view.setUint16(12, 1, true);
      view.setUint16(14, depth, true);
      view.setUint32(32, 2, true);
      return view;
    };
    const boundary = indexed(16);
    expect(decodePackedDib(boundary, 0, boundary.byteLength)?.data).toHaveLength(64);
    const view = indexed(17);
    const expected = expect.objectContaining({ code: 'ooxml-decoded-image-limit', metric: 'image-dimension', limit: 16, observed: 17 });
    expect(() => decodePackedDib(view, 0, view.byteLength)).toThrowError(expected);
    expect(() => decodeDib(view, 0, 48, 48, view.byteLength - 48)).toThrowError(expected);
    // All pixels supplied, but one RGBQUAD missing from the declared BMI span.
    expect(decodeDib(view, 0, 44, 48, view.byteLength - 48)).toBeNull();
  });
});

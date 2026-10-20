import { afterEach, describe, expect, it, vi } from 'vitest';
import { EmfPlusPlayer, scanEmfPlus, type EmfPlusTarget } from './emf-plus.js';

// A 256-byte player ceiling makes the decoded-byte budget boundaries testable
// with bitmaps of a few pixels.
vi.mock('./pixel-budget.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('./pixel-budget.js')>()),
  HARD_MAX_DECODED_IMAGE_BYTES: 256,
  MAX_RASTER_DIMENSION: 16,
  MAX_RASTER_PIXELS: 32,
}));

afterEach(() => vi.unstubAllGlobals());

const u32 = (v: number) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
const f32 = (v: number) => [...new Uint8Array(new Float32Array([v]).buffer)];
/** One EMF+ record: Type, Flags, Size, DataSize, data (4-byte aligned). */
const plusRecord = (type: number, flags: number, data: number[] = []) => {
  const body = [...data];
  while (body.length % 4) body.push(0);
  return [type & 255, type >>> 8, flags & 255, flags >>> 8, ...u32(12 + body.length), ...u32(data.length), ...body];
};
/** An EMF record: iType, nSize, data. */
const emfRecord = (type: number, data: number[] = []) => [...u32(type), ...u32(8 + data.length), ...data];
/** EMR_COMMENT carrying EMF+ records. */
const comment = (...records: number[][]) => {
  const payload = records.flat();
  return emfRecord(70, [...u32(4 + payload.length), ...u32(0x2b464d45), ...payload]);
};
const file = (...records: number[][]) => new Uint8Array([...emfRecord(1, new Array(80).fill(0)), ...records.flat(), ...emfRecord(14)]);
const header = plusRecord(0x4001, 1, [...u32(0xdbc01002), ...u32(1), ...u32(96), ...u32(96)]);
/** EmfPlusImageAttributes (2.2.1.5), object 0. */
const attributes = plusRecord(0x4008, 0x0800, [...u32(0xdbc01002), ...u32(0), ...u32(0), ...u32(0xffffffff), ...u32(0), ...u32(0)]);
/** A 4×4 32bpp ARGB bitmap object: 64 decoded bytes. */
const bitmap4 = (id: number) => plusRecord(0x4008, 0x0500 | id, [
  ...u32(0xdbc01002), ...u32(1), ...u32(4), ...u32(4), ...u32(16), ...u32(0x0026200a), ...u32(0), ...new Array(64).fill(255),
]);
const bitmap = (id: number, width: number, height: number, missing = 0) => plusRecord(0x4008, 0x0500 | id, [
  ...u32(0xdbc01002), ...u32(1), ...u32(width), ...u32(height), ...u32(width * 4),
  ...u32(0x0026200a), ...u32(0), ...new Array(width * height * 4 - missing).fill(255),
]);
const drawImage = (image: number) =>
  plusRecord(0x401a, image, [...u32(0), ...u32(2), ...[0, 0, 4, 4].flatMap(f32), ...[0, 0, 4, 4].flatMap(f32)]);

describe('EMF+ decoded-byte budget', () => {
  it.each([24, 300])('rejects a 324-byte fragment overrunning TotalObjectSize %i before budget admission', (total) => {
    const oversized = plusRecord(0x4008, 0x8800, [...u32(total), ...u32(0xdbc01002), ...new Array(320).fill(0)]);
    const bytes = file(comment(header, bitmap4(1), oversized, drawImage(1)));
    const allocated: number[] = [];
    const Original = Uint8Array;
    vi.stubGlobal('Uint8Array', new Proxy(Original, {
      construct(target, args) {
        if (typeof args[0] === 'number') allocated.push(args[0]);
        return Reflect.construct(target, args);
      },
    }));
    const scan = scanEmfPlus(bytes);
    expect(scan.play).toBe(false);
    expect(scan.failures).toContain('EMF+ continued object (fragments exceed TotalObjectSize)');

    const target: EmfPlusTarget = {
      ctx: {} as CanvasRenderingContext2D,
      W: 4, H: 4, left: 0, top: 0, boundsW: 4, boundsH: 4, drew: false, unsupported: new Set(),
    };
    const at = 88; // the comment after the 88-byte EMR_HEADER
    const dv = new DataView(bytes.buffer);
    new EmfPlusPlayer(target).playComment(dv, at, at + dv.getUint32(at + 4, true));
    expect(target.unsupported).toContain('EMF+ continued object (fragments exceed TotalObjectSize)');
    vi.unstubAllGlobals();
    expect(allocated.every((n) => n <= 256)).toBe(true);
  });

  /** Exercise real playback as well as the allocation-free scanner. */
  function playback(records: number[][]) {
    const bytes = file(comment(header, ...records));
    const target: EmfPlusTarget = {
      ctx: null, W: 4, H: 4, left: 0, top: 0, boundsW: 4, boundsH: 4, drew: false, unsupported: new Set(),
    };
    const player = new EmfPlusPlayer(target);
    const dv = new DataView(bytes.buffer);
    player.playComment(dv, 88, 88 + dv.getUint32(92, true));
    player.finish();
    return [...target.unsupported];
  }

  function check(records: number[][]) {
    const scan = scanEmfPlus(file(comment(header, ...records)));
    expect(playback(records)).toEqual(scan.failures);
    return scan;
  }

  it.each([
    [16, 1, 17, 1, 'image-dimension', 16, 17],
    [1, 16, 1, 17, 'image-dimension', 16, 17],
    [8, 4, 11, 3, 'image-pixels', 32, 33],
  ] as const)('validates bitmap bytes before the %s×%s boundary', (w, h, overW, overH, metric, limit, observed) => {
    expect(check([bitmap(1, w, h)]).failures).toEqual([]);
    const oversized = [bitmap(1, overW, overH)];
    const expected = expect.objectContaining({ code: 'ooxml-decoded-image-limit', metric, limit, observed });
    expect(() => scanEmfPlus(file(comment(header, ...oversized)))).toThrowError(expected);
    expect(() => playback(oversized)).toThrowError(expected);
    // DataSize, rather than 4-byte record alignment padding, bounds pixels.
    expect(check([bitmap(1, overW, overH, 1)]).failures).toEqual(['EMF+ record 0x4008 (malformed)']);
  });

  it('keeps oversized unfinished, inconsistent and malformed continuations as content failures', () => {
    const fragment = (total: number, data: number[]) => plusRecord(0x4008, 0x8501, [...u32(total), ...data]);
    expect(check([fragment(300, [1, 2, 3, 4])]).failures).toContain('EMF+ continued object (unfinished)');
    expect(check([fragment(300, new Array(200).fill(0)), fragment(300, new Array(104).fill(0))]).failures)
      .toContain('EMF+ continued object (fragments exceed TotalObjectSize)');
    // Same ID/total with a changed kind cannot complete the preceding image
    // using an attributes record (or reinterpret its header as attributes).
    const valid = [...bitmap4(1).slice(12), ...new Array(208).fill(0)];
    expect(check([
      fragment(300, valid.slice(0, 10)),
      plusRecord(0x4008, 0x8801, [...u32(300), ...valid.slice(10)]),
    ]).failures)
      .toContain('EMF+ continued object (unfinished)');
    // Complete payload, but invalid dimensions; no assembly or RGBA is needed.
    const invalid = [...u32(0xdbc01002), ...u32(1), ...u32(0), ...u32(4), ...u32(16),
      ...u32(0x0026200a), ...u32(0), ...new Array(272).fill(0)];
    expect(check([fragment(300, invalid.slice(0, 10)), fragment(300, invalid.slice(10))]).failures)
      .toContain('EMF+ record 0x4008 (malformed)');
  });

  it('does not assemble continued attributes or unsupported object types', () => {
    const data = [...attributes.slice(12), ...new Array(276).fill(0)];
    const object = (kind: number) => plusRecord(0x4008, 0x8000 | (kind << 8), [...u32(data.length), ...data]);
    expect(check([object(8)]).failures).toEqual([]);
    expect(check([object(2)]).failures).toEqual(['EMF+ object type 2']);
  });

  it('admits the exact completed assembly plus decoded bitmap, and rejects one byte more', () => {
    const continued = (total: number) => {
      const data = [...bitmap4(1).slice(12), ...new Array(total - 92).fill(0)];
      return [data.slice(0, 10), data.slice(10)].map((part) => plusRecord(0x4008, 0x8501, [...u32(total), ...part]));
    };
    expect(check(continued(192)).failures).toEqual([]); // 192 assembly + 64 RGBA
    const expected = expect.objectContaining({
      code: 'ooxml-decoded-image-limit', metric: 'active-decoded-bytes', limit: 256, observed: 257,
    });
    expect(() => check(continued(193))).toThrowError(expected);
    expect(() => playback(continued(193))).toThrowError(expected);
    const unfinished = continued(193);
    // Drop one object byte, retaining the original TotalObjectSize. The
    // record's alignment padding cannot complete the continued payload.
    unfinished[1] = plusRecord(0x4008, 0x8501, unfinished[1].slice(12, 12 + 4 + 182));
    expect(check(unfinished).failures).toEqual(['EMF+ continued object (unfinished)']);
    // The old entry is released before its replacement is assembled/decoded.
    expect(check([bitmap4(1), ...continued(192)]).failures).toEqual([]);
  });

  it('admits the exact table/blit boundaries and rejects the next RGBA pixel', () => {
    // RGBA allocations advance in four-byte pixels; assembly above also
    // covers the byte-granular 256/257 boundary of the same admission gate.
    for (const records of [
      [bitmap4(1), bitmap4(2), bitmap4(3), bitmap4(4)], // retained = 256
      [attributes, bitmap4(1), bitmap4(2)], // retained + DrawImage scratch = 256
    ]) {
      const drawing = records[0] === attributes;
      const tail = drawing ? [drawImage(1)] : [];
      expect(check([...records, ...tail]).failures).toEqual([]);
      const oversized = [...records, bitmap(5, 1, 1), ...tail];
      const expected = expect.objectContaining({
        code: 'ooxml-decoded-image-limit', metric: 'active-decoded-bytes', limit: 256, observed: 260,
      });
      expect(() => scanEmfPlus(file(comment(header, ...oversized)))).toThrowError(expected);
      expect(() => playback(oversized)).toThrowError(expected);
      expect(check([...records, bitmap(5, 1, 1, 1), ...tail]).failures).toEqual(['EMF+ record 0x4008 (malformed)']);
      if (drawing) {
        const truncatedDraw = plusRecord(0x401a, 1, drawImage(1).slice(12, -1));
        // Enough retained data to cross only if a complete draw asks for its
        // intermediates. A truncated draw must remain a content failure.
        expect(check([...records, bitmap(5, 1, 1), truncatedDraw]).failures).toEqual(['EMF+ record 0x401a (malformed)']);
      }
    }
  });
});

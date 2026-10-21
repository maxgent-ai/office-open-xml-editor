import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  isWmf,
  isEmf,
  isMetafileMime,
  playWmf,
  renderWmfToBitmap,
  wmfRasterTarget,
} from './wmf.js';
import { decodeRasterOrMetafile } from './raster-or-metafile.js';
import { TiffDecodeError } from './tiff-contract.js';
import { isOptionalImageCodecUnavailableError } from './optional-image-fallback.js';

// ── WMF (Windows Metafile) player unit tests ────────────────────────────────
// The shared player replays WMF records into a recording Canvas context.
// Pixel-level clipping coverage is in node/src/core-paint-audit.test.ts.

// ── byte builders ───────────────────────────────────────────────────────────

/** Little-endian byte writer for crafting WMF records. */
class Writer {
  private bytes: number[] = [];
  u16(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff);
    return this;
  }
  i16(v: number) {
    return this.u16(v & 0xffff);
  }
  u32(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
    return this;
  }
  raw(...vals: number[]) {
    for (const v of vals) this.bytes.push(v & 0xff);
    return this;
  }
  build(): Uint8Array {
    return new Uint8Array(this.bytes);
  }
}

/** Standard (non-placeable) 18-byte WMF header. numObjects defaults small. */
function wmfHeader(numObjects = 8): Uint8Array {
  return new Writer()
    .u16(1) // mtType = 1 (in-memory? 1 or 2 both legal; we accept both)
    .u16(9) // mtHeaderSize (words)
    .u16(0x0300) // mtVersion
    .u32(0) // mtSize (words) — players ignore for our purposes
    .u16(numObjects) // mtNoObjects
    .u32(0) // mtMaxRecord (words)
    .u16(0) // mtNoParameters
    .build();
}

/** A WMF record: u32 sizeWords (incl. the 6-byte size+function header), u16
 *  function, then params. `paramWords` is the number of 16-bit param words. */
function record(fn: number, params: (w: Writer) => void): Uint8Array {
  const pw = new Writer();
  params(pw);
  const paramBytes = pw.build();
  if (paramBytes.length % 2 !== 0) throw new Error('param bytes must be even');
  const sizeWords = 3 + paramBytes.length / 2; // 3 words = u32 size + u16 fn
  const rec = new Writer().u32(sizeWords).u16(fn);
  const head = rec.build();
  const out = new Uint8Array(head.length + paramBytes.length);
  out.set(head, 0);
  out.set(paramBytes, head.length);
  return out;
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

// WMF record function codes
const FN = {
  EOF: 0x0000,
  SETPOLYFILLMODE: 0x0106,
  SETWINDOWORG: 0x020b,
  SETWINDOWEXT: 0x020c,
  SELECTOBJECT: 0x012d,
  DELETEOBJECT: 0x01f0,
  POLYGON: 0x0324,
  POLYLINE: 0x0325,
  POLYPOLYGON: 0x0538,
  RECTANGLE: 0x041b,
  CREATEPENINDIRECT: 0x02fa,
  CREATEBRUSHINDIRECT: 0x02fc,
  STRETCHDIBITS: 0x0f43,
} as const;

// ── recording mock ctx (records the draw calls + style mutations) ───────────

interface Call {
  op: string;
  args: Array<number | string>;
}
interface MockCtx {
  ctx: CanvasRenderingContext2D;
  calls: Call[];
  styles: { fill: string[]; stroke: string[]; lineWidth: number[]; fillRules: (string | undefined)[]; text: string[] };
}

function makeRecordingCtx(): MockCtx {
  const calls: Call[] = [];
  const styles = {
    fill: [] as string[],
    stroke: [] as string[],
    lineWidth: [] as number[],
    fillRules: [] as (string | undefined)[],
    text: [] as string[],
  };
  let _fill = '#000';
  let _stroke = '#000';
  let _lw = 1;
  const ctx = {
    get fillStyle() { return _fill; },
    set fillStyle(v: string) { _fill = v; },
    get strokeStyle() { return _stroke; },
    set strokeStyle(v: string) { _stroke = v; },
    get lineWidth() { return _lw; },
    set lineWidth(v: number) { _lw = v; },
    font: '10px sans-serif',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'top' as CanvasTextBaseline,
    lineJoin: 'miter' as CanvasLineJoin,
    lineCap: 'butt' as CanvasLineCap,
    save() { calls.push({ op: 'save', args: [] }); },
    restore() { calls.push({ op: 'restore', args: [] }); },
    beginPath() { calls.push({ op: 'beginPath', args: [] }); },
    closePath() { calls.push({ op: 'closePath', args: [] }); },
    moveTo(x: number, y: number) { calls.push({ op: 'moveTo', args: [x, y] }); },
    lineTo(x: number, y: number) { calls.push({ op: 'lineTo', args: [x, y] }); },
    rect(x: number, y: number, w: number, h: number) { calls.push({ op: 'rect', args: [x, y, w, h] }); },
    clip() { calls.push({ op: 'clip', args: [] }); },
    stroke() {
      calls.push({ op: 'stroke', args: [] });
      styles.stroke.push(_stroke);
      styles.lineWidth.push(_lw);
    },
    fill(rule?: string) {
      calls.push({ op: 'fill', args: [] });
      styles.fill.push(_fill);
      styles.fillRules.push(rule);
    },
    fillText(t: string, x: number, y: number) {
      calls.push({ op: 'fillText', args: [t, x, y] });
      styles.text.push(_fill);
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls, styles };
}

// ── isWmf / isEmf detection ─────────────────────────────────────────────────

describe('isWmf / isEmf detection', () => {
  it('detects a standard (non-placeable) WMF header', () => {
    const bytes = wmfHeader();
    expect(isWmf(bytes)).toBe(true);
    expect(isEmf(bytes)).toBe(false);
  });

  it('detects mtType=2 as WMF too', () => {
    const w = new Writer().u16(2).u16(9).u16(0x0300).u32(0).u16(8).u32(0).u16(0);
    expect(isWmf(w.build())).toBe(true);
  });

  it('detects a placeable WMF via the D7CDC69A magic', () => {
    // 22-byte placeable header: magic, handle, bbox(4×i16), inch, reserved, checksum
    const placeable = new Writer()
      .raw(0xd7, 0xcd, 0xc6, 0x9a) // magic
      .u16(0) // hWmf
      .i16(0).i16(0).i16(100).i16(100) // bbox
      .u16(96) // inch
      .u32(0) // reserved
      .u16(0); // checksum
    // followed by a standard header
    const full = concat(placeable.build(), wmfHeader());
    expect(isWmf(full)).toBe(true);
  });

  it('detects a true EMF (ENHMETAHEADER) and does NOT treat it as WMF', () => {
    // u32@0 = 1 (EMR_HEADER iType), u32@40 = 0x464D4520 (" EMF" signature).
    const buf = new Uint8Array(48);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 1, true);
    dv.setUint32(40, 0x464d4520, true);
    expect(isEmf(buf)).toBe(true);
    expect(isWmf(buf)).toBe(false);
  });

  it('rejects random bytes', () => {
    const rnd = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0xde, 0xad, 0xbe, 0xef]);
    expect(isWmf(rnd)).toBe(false);
    expect(isEmf(rnd)).toBe(false);
  });

  it('rejects a too-short buffer', () => {
    expect(isWmf(new Uint8Array([1, 0, 9]))).toBe(false);
    expect(isEmf(new Uint8Array([1, 0, 0, 0]))).toBe(false);
  });
});

// ── playWmf: minimal polyline replay with window mapping + pen select ────────

describe('playWmf — window mapping, pen, polyline', () => {
  it('replays a minimal WMF and maps logical→device coords with the current pen', () => {
    // Window org (0,0), ext (100,100); target bitmap 200×200 → scale ×2.
    // Pen: solid, color 0x00FF0000 = blue (COLORREF 0x00BBGGRR → R=0,G=0,B=0xFF).
    const file = concat(
      wmfHeader(),
      // SETWINDOWORG: y first, x second
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      // SETWINDOWEXT: y first, x second → yExt=100, xExt=100
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      // CREATEPENINDIRECT: style=0 (solid), widthX=1, widthY=0, color 0x00FF0000
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00ff0000)),
      // SELECTOBJECT idx 0
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // POLYLINE: 3 pts (10,20)-(30,40)-(50,60)
      record(FN.POLYLINE, (w) => w.i16(3).i16(10).i16(20).i16(30).i16(40).i16(50).i16(60)),
      record(FN.EOF, () => {}),
    );

    const m = makeRecordingCtx();
    const drew = playWmf(file, m.ctx, 200, 200);
    expect(drew).toBe(true);

    // Expect a moveTo to the first point then lineTo for the rest, mapped ×2.
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves.length).toBe(1);
    expect(moves[0].args).toEqual([20, 40]); // (10,20) × 2
    expect(lines.length).toBe(2);
    expect(lines[0].args).toEqual([60, 80]); // (30,40) × 2
    expect(lines[1].args).toEqual([100, 120]); // (50,60) × 2

    // The polyline strokes (no fill) with the selected blue pen.
    const strokes = m.calls.filter((c) => c.op === 'stroke');
    expect(strokes.length).toBe(1);
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#0000ff'); // R=0 G=0 B=255
    // Polyline never fills.
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false);
  });

  it('honors window origin and a negative ext (axis flip)', () => {
    // org (10,10); ext x=100, y=-100 (Y flips). Target 100×100 → |scale| ×1.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(10).i16(10)), // yOrg=10, xOrg=10
      record(FN.SETWINDOWEXT, (w) => w.i16(-100).i16(100)), // yExt=-100, xExt=100
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(10).i16(10).i16(60).i16(60)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    // logical (10,10) is the origin → device (0,0) even with the flip.
    expect(moves[0].args).toEqual([0, 0]);
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    // (60,60): dx=50 → x=50; dy=50, yExt=-100 → device y = 50 * (100 / -100) = -50.
    expect(lines[0].args[0]).toBe(50);
    expect(lines[0].args[1]).toBe(-50);
  });

  it('NULL-style pen (style 5) does not stroke', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(5).i16(1).i16(0).u32(0x00000000)), // PS_NULL
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(10).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(false);
  });
});

// ── playWmf: polygon / polypolygon fill + fill rule ─────────────────────────

describe('playWmf — polygon / polypolygon fill', () => {
  it('fills + strokes a POLYGON with the current brush + pen', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      // brush: SOLID (0), color green 0x0000FF00 (R=0,G=0xFF,B=0)
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x0000ff00).u16(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // pen: solid red 0x000000FF (R=0xFF)
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)),
      record(FN.SELECTOBJECT, (w) => w.u16(1)),
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playWmf(file, m.ctx, 10, 10)).toBe(true);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#00ff00'); // green brush
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000'); // red pen
  });

  it('NULL brush (style 1) does not fill', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(1).u32(0x00000000).u16(0)), // BS_NULL
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(1)),
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 10, 10);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });

  it('POLYPOLYGON honors SETPOLYFILLMODE (ALTERNATE → evenodd) for glyph holes', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(20).i16(20)),
      record(FN.SETPOLYFILLMODE, (w) => w.u16(1)), // ALTERNATE
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x00000000).u16(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // 2 sub-polys: outer 4-gon, inner 4-gon (a hole). u16 numPolys, u16 counts, then pts.
      record(FN.POLYPOLYGON, (w) =>
        w
          .u16(2)
          .u16(4)
          .u16(4)
          // outer
          .i16(0).i16(0).i16(20).i16(0).i16(20).i16(20).i16(0).i16(20)
          // inner hole
          .i16(5).i16(5).i16(15).i16(5).i16(15).i16(15).i16(5).i16(15),
      ),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playWmf(file, m.ctx, 20, 20)).toBe(true);
    // One fill spanning both sub-paths, with the evenodd rule.
    const fills = m.calls.filter((c) => c.op === 'fill');
    expect(fills.length).toBeGreaterThanOrEqual(1);
    expect(m.styles.fillRules.at(-1)).toBe('evenodd');
    // Both sub-polygons contributed moveTo's (8 vertices → 2 moveTo + 6 lineTo at least).
    expect(m.calls.filter((c) => c.op === 'moveTo').length).toBeGreaterThanOrEqual(2);
  });
});

// ── playWmf: authored boundary edges ──────────────────────────────────────

describe('playWmf — authored boundary edges', () => {
  it('a RECTANGLE coincident with the device bounds STILL strokes its outline by default', () => {
    // window org (0,0) ext (100,100), device 100×100 ⇒ logical==device. The rect
    // spans the full window; every authored edge is drawn.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // RECTANGLE params: bottom, right, top, left (full window).
      record(FN.RECTANGLE, (w) => w.i16(100).i16(100).i16(0).i16(0)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100);
    // Spec-clean: the boundary-coincident outline IS stroked as one closed path.
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
    expect(m.calls.filter((c) => c.op === 'moveTo').length).toBe(1);
    expect(m.calls.filter((c) => c.op === 'lineTo').length).toBe(3);
    expect(m.calls.some((c) => c.op === 'closePath')).toBe(true);
  });
});

describe('playWmf — clip and raster-state diagnostics', () => {
  it.each([false, true])('bounds exclude/offset Canvas work with drawing=%s', drawEach => {
    const draw = record(FN.RECTANGLE, w => w.i16(100).i16(100).i16(0).i16(0));
    const pair = [record(0x0415, w => w.i16(100).i16(60).i16(0).i16(50)),
      record(0x0220, w => w.i16(0).i16(1)), ...(drawEach ? [draw] : [])];
    const file = concat(wmfHeader(),
      record(FN.CREATEBRUSHINDIRECT, w => w.u16(0).u32(0xff).u16(0)),
      record(FN.SELECTOBJECT, w => w.u16(0)),
      ...Array.from({ length: 1000 }, () => pair).flat());
    const mock = makeRecordingCtx();
    playWmf(file, mock.ctx, 100, 100);
    expect(mock.calls.filter(c => c.op === 'clip').length).toBe(drawEach ? 1000 : 0);
    expect(mock.calls.filter(c => c.op === 'rect').length).toBeLessThanOrEqual(drawEach ? 4000 : 0);
  });

  it.each([{ params: [] }, { params: [9] }])('diagnoses an invalid or malformed mapping mode $params', ({ params }) => {
    const report = vi.fn();
    playWmf(concat(wmfHeader(), record(0x0103, w => params.forEach(p => w.u16(p)))),
      makeRecordingCtx().ctx, 100, 100, report);
    expect(report).toHaveBeenCalledWith([params.length ? 'META_SETMAPMODE invalid mode' : 'META_SETMAPMODE malformed']);
  });

  it('reports unsupported Boolean raster mixes instead of silently copying the pen', () => {
    const file = concat(wmfHeader(),
      record(FN.CREATEPENINDIRECT, w => w.u16(0).i16(0).i16(0).u32(0)),
      record(FN.SELECTOBJECT, w => w.u16(0)),
      record(0x0104, w => w.u16(7)),
      record(FN.RECTANGLE, w => w.i16(100).i16(100).i16(0).i16(0)),
      record(FN.EOF, () => {}));
    const report = vi.fn();
    const mock = makeRecordingCtx();
    expect(playWmf(file, mock.ctx, 100, 100, report)).toBe(false);
    expect(mock.calls.some(c => c.op === 'stroke')).toBe(false);
    expect(report).toHaveBeenCalledWith(['META_SETROP2 (unsupported binary mix)']);
  });

  it('reports malformed region scans through the existing incomplete-metafile channel', () => {
    const file = concat(wmfHeader(),
      record(0x06ff, w => w.u16(0).i16(6).u32(0).i16(34).i16(1).i16(2)
        .i16(0).i16(0).i16(100).i16(100)
        .u16(2).u16(0).u16(100).u16(0).u16(100).u16(4)),
      record(0x012c, w => w.u16(0)),
      record(FN.EOF, () => {}));
    const report = vi.fn();
    expect(playWmf(file, makeRecordingCtx().ctx, 100, 100, report)).toBe(false);
    expect(report).toHaveBeenCalledWith(expect.arrayContaining(['META_CREATEREGION malformed']));
  });
});

// ── playWmf: object table create/select/delete/reuse ────────────────────────

describe('playWmf — object table create / delete / slot reuse', () => {
  it('reuses the freed slot when a deleted object index is later recreated', () => {
    // Create pen#0 (blue), pen#1 (green); select#0; delete#0; create pen (red) →
    // must land in the freed slot 0; select#0 → current pen is red.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00ff0000)), // slot0 blue
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x0000ff00)), // slot1 green
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.DELETEOBJECT, (w) => w.u16(0)), // free slot0
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)), // → slot0 red
      record(FN.SELECTOBJECT, (w) => w.u16(0)), // select slot0 (now red)
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(10).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 10, 10);
    // The stroke uses the recreated red pen in the reused slot 0.
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000');
  });

  it('selecting a brush vs a pen routes by object kind', () => {
    // slot0 = brush(green), slot1 = pen(red). select#0 (brush) then #1 (pen).
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x0000ff00).u16(0)), // slot0 brush green
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)), // slot1 pen red
      record(FN.SELECTOBJECT, (w) => w.u16(0)), // brush
      record(FN.SELECTOBJECT, (w) => w.u16(1)), // pen
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 10, 10);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#00ff00');
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000');
  });
});

// ── playWmf: text-out labels ────────────────────────────────────────────────

// ── playWmf: graceful bail on malformed records ─────────────────────────────

describe('playWmf — robustness', () => {
  it('bails gracefully on a record claiming a bogus (too-small) size', () => {
    const bad = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(5).i16(5)),
      // a deliberately corrupt record: sizeWords = 1 (< 3) — must stop the loop.
      new Writer().u32(1).u16(FN.POLYLINE).build(),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    // Whatever was drawn before the corrupt record stands; no throw.
    expect(() => playWmf(bad, m.ctx, 10, 10)).not.toThrow();
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });

  it('returns false for non-WMF bytes', () => {
    const m = makeRecordingCtx();
    expect(playWmf(new Uint8Array([0xde, 0xad, 0xbe, 0xef, 0, 0, 0, 0]), m.ctx, 10, 10)).toBe(false);
  });
});

// ── playWmf: STRETCHDIBITS embedded raster DIB ──────────────────────────────

describe('playWmf — STRETCHDIBITS (embedded raster DIB)', () => {
  it('does not borrow a following record to admit a truncated packed DIB', () => {
    const bitmap = record(FN.STRETCHDIBITS, (w) => {
      w.u32(0x00cc0020);
      for (let i = 0; i < 9; i++) w.u16(1);
      w.u32(40).u32(40000).u32(2).u16(1).u16(32).u32(0).u32(0)
        .u32(0).u32(0).u32(0).u32(0).raw(0, 0, 0, 0);
    });
    const following = new Uint8Array(320006);
    const view = new DataView(following.buffer);
    view.setUint32(0, following.length / 2, true);
    view.setUint16(4, 0xffff, true); // unsupported record, skipped by size
    for (const tail of [new Uint8Array(0), following]) {
      const m = makeRecordingCtx();
      expect(() => playWmf(concat(
        wmfHeader(), bitmap, tail,
        record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0xff)),
        record(FN.SELECTOBJECT, (w) => w.u16(0)),
        record(FN.RECTANGLE, (w) => w.i16(10).i16(10).i16(0).i16(0)),
        record(FN.EOF, () => {}),
      ), m.ctx, 100, 100)).not.toThrow();
      expect(m.calls.some((call) => call.op === 'stroke')).toBe(true);
    }
  });

  it('propagates a packed DIB budget crossing out of WMF playback', () => {
    const bitmap = record(FN.STRETCHDIBITS, (w) => {
      w.u32(0x00cc0020);
      for (let i = 0; i < 9; i++) w.u16(1);
      w.u32(40).u32(40000).u32(1).u16(1).u16(24).u32(0).u32(0)
        .u32(0).u32(0).u32(0).u32(0);
      for (let i = 0; i < 120000; i++) w.raw(0);
    });
    expect(() => playWmf(concat(wmfHeader(), bitmap, record(FN.EOF, () => {})), makeRecordingCtx().ctx, 100, 100))
      .toThrowError(expect.objectContaining({
        code: 'ooxml-decoded-image-limit', metric: 'image-dimension', limit: 32767, observed: 40000,
      }));
  });

  // A META_STRETCHDIBITS record wraps a packed DIB. Its params (after the 6-byte
  // header) are: u32 RasterOp, i16 SrcHeight/SrcWidth/YSrc/XSrc, u16 UsageSrc,
  // i16 DestHeight/DestWidth/YDest/XDest, then the packed DIB. OffscreenCanvas is
  // absent in the node test env, so blitDibToCtx returns false (no draw) — we
  // assert the record parses without throwing and that later records still run.

  /** Append a packed top-down 2×2 24-bit BI_RGB DIB (40-byte header + 16 pixel
   *  bytes: 2 rows × 8-byte stride). Returns the same Writer for chaining. */
  function packed2x2Dib24(w: Writer): Writer {
    // BITMAPINFOHEADER (top-down: height = -2).
    w.u32(40).u32(2).u32((-2) >>> 0).u16(1).u16(24).u32(0).u32(0).u32(0).u32(0).u32(0).u32(0);
    // 2 rows, BGR pixels, each row padded to an 8-byte stride.
    w.raw(0, 0, 255, 0, 255, 0, 0, 0); // row0: red (B0 G0 R255), green (B0 G255 R0), pad
    w.raw(255, 0, 0, 30, 20, 10, 0, 0); // row1: blue, (10,20,30), pad
    return w;
  }

  it('parses a STRETCHDIBITS record without throwing and continues to later records', () => {
    const m = makeRecordingCtx();
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      // STRETCHDIBITS: draw the DIB into dest rect (XDest=0,YDest=0, 50×50).
      record(FN.STRETCHDIBITS, (w) => {
        w.u32(0x00cc0020); // RasterOperation SRCCOPY (ignored)
        w.i16(2).i16(2); // SrcHeight, SrcWidth
        w.i16(0).i16(0); // YSrc, XSrc
        w.u16(0); // UsageSrc (DIB_RGB_COLORS)
        w.i16(50).i16(50); // DestHeight, DestWidth
        w.i16(0).i16(0); // YDest, XDest
        packed2x2Dib24(w); // packed DIB
      }),
      // A polyline AFTER the blt must still execute (proves the loop advanced past
      // the STRETCHDIBITS record by its size, not by mis-parsing the DIB bytes).
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(10).i16(10).i16(20).i16(20)),
      record(FN.EOF, () => {}),
    );

    // No OffscreenCanvas in the node env ⇒ blitDibToCtx returns false (no draw);
    // the parse path must still run cleanly and the trailing polyline strokes.
    let drew = false;
    expect(() => {
      drew = playWmf(file, m.ctx, 100, 100);
    }).not.toThrow();
    // The trailing polyline drew, so playWmf reports true and the loop advanced
    // past the STRETCHDIBITS record correctly.
    expect(drew).toBe(true);
    const strokes = m.calls.filter((c) => c.op === 'stroke');
    expect(strokes.length).toBe(1);
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000');
    // The polyline endpoints mapped ×1 (window 100 → device 100).
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    expect(moves[0].args).toEqual([10, 10]);
  });
});

// ── wmfRasterTarget: supersampled, capped sizing ────────────────────────────

describe('wmfRasterTarget', () => {
  it('supersamples the intended pt size by 2×', () => {
    expect(wmfRasterTarget(100, 50)).toEqual({ w: 200, h: 100 });
  });

  it('falls back to a 300pt square (×2 → 600px) when size is unknown (0)', () => {
    expect(wmfRasterTarget(0, 0)).toEqual({ w: 600, h: 600 });
  });

  it('caps each dimension at 2000px', () => {
    expect(wmfRasterTarget(5000, 5000)).toEqual({ w: 2000, h: 2000 });
  });
});

// ── renderWmfToBitmap: OffscreenCanvas wrapper (browser/worker only) ─────────

describe('renderWmfToBitmap', () => {
  beforeEach(() => {
    // OffscreenCanvas + createImageBitmap don't exist in the node test env.
    const recorded: { ctx: unknown } = { ctx: null };
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) {
          this.width = w;
          this.height = h;
        }
        getContext() {
          const ctx = {
            fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
            lineJoin: 'miter', lineCap: 'butt',
            save() {}, restore() {}, beginPath() {}, closePath() {},
            moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
          };
          recorded.ctx = ctx;
          return ctx;
        }
      },
    );
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (src: { width: number; height: number }) => ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('rasterizes a minimal WMF to an ImageBitmap of the target size', async () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(50).i16(50)),
      record(FN.EOF, () => {}),
    );
    const bmp = await renderWmfToBitmap(file, 64, 48);
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(64);
    expect(bmp?.height).toBe(48);
  });

  it('returns null for non-WMF bytes', async () => {
    const bmp = await renderWmfToBitmap(new Uint8Array([1, 2, 3, 4]), 10, 10);
    expect(bmp).toBeNull();
  });

  it('returns null when nothing draws (empty metafile)', async () => {
    const file = concat(wmfHeader(), record(FN.EOF, () => {}));
    const bmp = await renderWmfToBitmap(file, 10, 10);
    expect(bmp).toBeNull();
  });
});

// ── decodeRasterOrMetafile: the shared raster/metafile decoder ───────────────

describe('decodeRasterOrMetafile', () => {
  beforeEach(() => {
    // OffscreenCanvas + createImageBitmap don't exist in the node test env.
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) {
          this.width = w;
          this.height = h;
        }
        getContext() {
          return {
            fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
            lineJoin: 'miter', lineCap: 'butt',
            save() {}, restore() {}, beginPath() {}, closePath() {},
            moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
          };
        }
      },
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('WMF bytes → rasterizes through the player (non-null bitmap), sized via wmfRasterTarget', async () => {
    // createImageBitmap here only ever sees the OffscreenCanvas (the WMF branch
    // never calls it on the blob), so report the canvas dims back as the bitmap.
    const cib = vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(50).i16(50)),
      record(FN.EOF, () => {}),
    );
    const blob = new Blob([file as BlobPart], { type: 'image/wmf' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 100, heightPt: 50 });
    expect(bmp).not.toBeNull();
    // wmfRasterTarget(100,50) = 200×100 (×2 supersample).
    expect(bmp?.width).toBe(200);
    expect(bmp?.height).toBe(100);
    // The WMF branch rasterizes the OffscreenCanvas, never the blob directly.
    expect(cib).toHaveBeenCalledTimes(1);
    const arg = cib.mock.calls[0][0] as { width: number; height: number };
    expect(arg.width).toBe(200);
    expect(arg.height).toBe(100);
  });

  it('EMF magic → null (skipped gracefully, no throw, no createImageBitmap on the blob)', async () => {
    const cib = vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const buf = new Uint8Array(48);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 1, true); // EMR_HEADER iType
    dv.setUint32(40, 0x464d4520, true); // " EMF" signature
    const blob = new Blob([buf as BlobPart], { type: 'image/emf' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 100, heightPt: 100 });
    expect(bmp).toBeNull();
    expect(cib).not.toHaveBeenCalled();
  });

  it('a PNG blob → the createImageBitmap path (returns the decoded bitmap)', async () => {
    const fake = { width: 7, height: 9, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (_blob: Blob) => fake);
    vi.stubGlobal('createImageBitmap', cib);
    // PNG magic 89 50 4E 47 0D 0A 1A 0A — not WMF, not EMF.
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 50, heightPt: 50 });
    expect(bmp).toBe(fake);
    expect(cib).toHaveBeenCalledTimes(1);
    expect(cib).toHaveBeenCalledWith(blob);
  });

  it('a PNG decode bomb rejects with a typed quota before createImageBitmap (RB1)', async () => {
    // The decode-bomb guard: a valid PNG header declaring a 60000×60000 image
    // (~14 GB RGBA once decoded) must be refused BEFORE createImageBitmap ever
    // allocates the surface. The fixture is a couple dozen bytes.
    const cib = vi.fn(async (_blob: Blob) => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    // 8-byte PNG signature + IHDR chunk declaring 60000×60000.
    const bomb = new Uint8Array(26);
    bomb.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    // IHDR length (13), "IHDR", then big-endian W=60000, H=60000.
    bomb.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const put = (o: number, v: number) => {
      bomb[o] = (v >>> 24) & 0xff;
      bomb[o + 1] = (v >>> 16) & 0xff;
      bomb[o + 2] = (v >>> 8) & 0xff;
      bomb[o + 3] = v & 0xff;
    };
    put(16, 60000);
    put(20, 60000);
    const blob = new Blob([bomb as BlobPart], { type: 'image/png' });
    await expect(decodeRasterOrMetafile(blob, { widthPt: 50, heightPt: 50 }))
      .rejects.toMatchObject({
        name: 'OoxmlDecodedImageLimitError',
        code: 'ooxml-decoded-image-limit',
        metric: 'image-pixels',
        observed: 60000 * 60000,
      });
    expect(cib).not.toHaveBeenCalled(); // never handed to the decoder
  });

  it('an in-budget PNG (1920×1080 header) still decodes (no false positive)', async () => {
    const fake = { width: 1920, height: 1080, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (_blob: Blob) => fake);
    vi.stubGlobal('createImageBitmap', cib);
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const put = (o: number, v: number) => {
      png[o] = (v >>> 24) & 0xff;
      png[o + 1] = (v >>> 16) & 0xff;
      png[o + 2] = (v >>> 8) & 0xff;
      png[o + 3] = v & 0xff;
    };
    put(16, 1920);
    put(20, 1080);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 50, heightPt: 50 });
    expect(bmp).toBe(fake);
    expect(cib).toHaveBeenCalledTimes(1);
  });

  it('downsamples an in-budget raster when both display target axes are smaller', async () => {
    const fake = { width: 960, height: 540, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => fake);
    vi.stubGlobal('createImageBitmap', cib);
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 1920);
    view.setUint32(20, 1080);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 960,
      targetHeightPx: 540,
    })).resolves.toBe(fake);
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 960,
      resizeHeight: 540,
      resizeQuality: 'high',
    });
  });

  it('downsamples a legitimate over-32MP raster to the requested display resolution', async () => {
    const sourceWidth = 12_000;
    const sourceHeight = 9_000; // 108 MP: legitimate poster / camera output.
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const put = (o: number, v: number) => {
      png[o] = (v >>> 24) & 0xff;
      png[o + 1] = (v >>> 16) & 0xff;
      png[o + 2] = (v >>> 8) & 0xff;
      png[o + 3] = v & 0xff;
    };
    put(16, sourceWidth);
    put(20, sourceHeight);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const resized = { width: 1_200, height: 900, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => resized);
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 1_200,
      targetHeightPx: 900,
    })).resolves.toBe(resized);
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 1_200,
      resizeHeight: 900,
      resizeQuality: 'high',
    });
  });

  it('uses the exact axis-wise DrawingML display grid', async () => {
    const sourceWidth = 12_090;
    const sourceHeight = 9_063;
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, sourceWidth);
    view.setUint32(20, sourceHeight);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => {
      const width = options?.resizeWidth ?? sourceWidth;
      return {
        width,
        height: options?.resizeHeight ?? sourceHeight,
        close() {},
      } as unknown as ImageBitmap;
    });
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 960,
      targetHeightPx: 720,
    })).resolves.toMatchObject({ width: 960, height: 720 });
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 960,
      resizeHeight: 720,
      resizeQuality: 'high',
    });
  });

  it('does not overshoot an exact integer browser target through floating-point roundoff', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 25);
    view.setUint32(20, 25);
    const bitmap = { width: 7, height: 7, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => bitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 7,
      targetHeightPx: 7,
      maxRetainedPixels: 49,
    })).resolves.toBe(bitmap);
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 7,
      resizeHeight: 7,
      resizeQuality: 'high',
    });
  });

  it('closes and rejects a decoder result that ignores a restricted retained-surface limit', async () => {
    const sourceWidth = 10_000;
    const sourceHeight = 1_000;
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, sourceWidth);
    view.setUint32(20, sourceHeight);
    const close = vi.fn();
    const ignoredResize = { width: sourceWidth, height: sourceHeight, close } as unknown as ImageBitmap;
    vi.stubGlobal('createImageBitmap', vi.fn(async () => ignoredResize));

    await expect(decodeRasterOrMetafile(
      new Blob([png as BlobPart], { type: 'image/png' }),
      {
        targetWidthPx: 1_000,
        targetHeightPx: 100,
        maxRetainedPixels: 1 << 23,
      },
    )).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 1 << 23,
      observed: sourceWidth * sourceHeight,
    });
    expect(close).toHaveBeenCalledOnce();
  });

  it('streams JPEG metadata until SOF instead of failing open after 64 KiB', async () => {
    const appLength = 65_535;
    const jpeg = new Uint8Array(2 + 2 + appLength + 2 + 17);
    let offset = 0;
    jpeg.set([0xff, 0xd8], offset); offset += 2; // SOI
    jpeg.set([0xff, 0xe1, 0xff, 0xff], offset); offset += 4; // APP1 incl. u16 length
    offset += appLength - 2; // payload; SOF now lies beyond byte 65536
    jpeg.set([0xff, 0xc0, 0x00, 0x11, 0x08, 0x23, 0x28, 0x2e, 0xe0], offset);
    const blob = new Blob([jpeg as BlobPart], { type: 'image/jpeg' });
    const resized = { width: 1_200, height: 900, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => resized);
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 1_200,
      targetHeightPx: 900,
    })).resolves.toBe(resized);
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 1_200,
      resizeHeight: 900,
      resizeQuality: 'high',
    });
  });

  it('keeps streamed JPEG EXIF orientation in the retained-surface decode plan', async () => {
    const exif = new Uint8Array(32);
    exif.set([0x45, 0x78, 0x69, 0x66, 0x00, 0x00], 0);
    exif.set([0x49, 0x49], 6);
    const exifView = new DataView(exif.buffer);
    exifView.setUint16(8, 42, true);
    exifView.setUint32(10, 8, true);
    exifView.setUint16(14, 1, true);
    exifView.setUint16(16, 0x0112, true);
    exifView.setUint16(18, 3, true);
    exifView.setUint32(20, 1, true);
    exifView.setUint16(24, 6, true); // coded 400x100 becomes browser-oriented 100x400

    const longAppPayload = 65_533;
    const jpeg = new Uint8Array(
      2 + (2 + 2 + exif.length) + (2 + 2 + longAppPayload) + 2 + 17,
    );
    let offset = 0;
    jpeg.set([0xff, 0xd8], offset); offset += 2;
    jpeg.set([0xff, 0xe1], offset); offset += 2;
    jpeg.set([0x00, exif.length + 2], offset); offset += 2;
    jpeg.set(exif, offset); offset += exif.length;
    jpeg.set([0xff, 0xe2, 0xff, 0xff], offset); offset += 4;
    offset += longAppPayload;
    jpeg.set([0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x64, 0x01, 0x90], offset);

    const blob = new Blob([jpeg as BlobPart], { type: 'image/jpeg' });
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => {
      const width = options?.resizeWidth ?? 100;
      const height = options?.resizeHeight ?? 400;
      return { width, height, close() {} } as unknown as ImageBitmap;
    });
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 25,
      targetHeightPx: 25,
      maxRetainedPixels: 2_500,
    })).resolves.toMatchObject({ width: 25, height: 25 });
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 25,
      resizeHeight: 25,
      resizeQuality: 'high',
    });
  });

  it('keeps the first streamed EXIF APP1 authoritative when its orientation is invalid', async () => {
    const exif = new Uint8Array(32);
    exif.set([0x45, 0x78, 0x69, 0x66, 0x00, 0x00, 0x49, 0x49], 0);
    const exifView = new DataView(exif.buffer);
    exifView.setUint16(8, 42, true);
    exifView.setUint32(10, 8, true);
    exifView.setUint16(14, 1, true);
    exifView.setUint16(16, 0x0112, true);
    exifView.setUint16(18, 3, true);
    exifView.setUint32(20, 1, true);
    exifView.setUint16(24, 0, true);
    const laterExif = exif.slice();
    new DataView(laterExif.buffer).setUint16(24, 6, true);
    const jpeg = new Uint8Array(2 + (4 + exif.length) + (4 + laterExif.length) + 9 + 2);
    let offset = 0;
    jpeg.set([0xff, 0xd8, 0xff, 0xe1, 0x00, exif.length + 2], offset); offset += 6;
    jpeg.set(exif, offset); offset += exif.length;
    jpeg.set([0xff, 0xe1, 0x00, laterExif.length + 2], offset); offset += 4;
    jpeg.set(laterExif, offset); offset += laterExif.length;
    jpeg.set([0xff, 0xc0, 0x00, 0x07, 0x08, 0x00, 0x64, 0x01, 0x90], offset); offset += 9;
    jpeg.set([0xff, 0xda], offset);
    const blob = new Blob([jpeg as BlobPart], { type: 'image/jpeg' });
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => {
      const width = options?.resizeWidth ?? 400;
      const height = options?.resizeHeight ?? 100;
      return { width, height, close() {} } as unknown as ImageBitmap;
    });
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 25,
      targetHeightPx: 25,
      maxRetainedPixels: 2_500,
    })).resolves.toMatchObject({ width: 25, height: 25 });
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 25,
      resizeHeight: 25,
      resizeQuality: 'high',
    });
  });

  it('finds browser-applied EXIF orientation after SOF and beyond the prefix', async () => {
    const exif = new Uint8Array(32);
    exif.set([0x45, 0x78, 0x69, 0x66, 0x00, 0x00], 0);
    exif.set([0x49, 0x49], 6);
    const exifView = new DataView(exif.buffer);
    exifView.setUint16(8, 42, true);
    exifView.setUint32(10, 8, true);
    exifView.setUint16(14, 1, true);
    exifView.setUint16(16, 0x0112, true);
    exifView.setUint16(18, 3, true);
    exifView.setUint32(20, 1, true);
    exifView.setUint16(24, 6, true);

    const longAppPayload = 65_533;
    const sofPayload = new Uint8Array(15);
    sofPayload.set([0x08, 0x00, 0x64, 0x01, 0x90]); // coded 400x100
    const jpeg = new Uint8Array(
      2 + (2 + 2 + sofPayload.length) + (2 + 2 + longAppPayload)
      + (2 + 2 + exif.length) + 2,
    );
    let offset = 0;
    jpeg.set([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11], offset); offset += 6;
    jpeg.set(sofPayload, offset); offset += sofPayload.length;
    jpeg.set([0xff, 0xe2, 0xff, 0xff], offset); offset += 4;
    offset += longAppPayload;
    jpeg.set([0xff, 0xe1, 0x00, exif.length + 2], offset); offset += 4;
    jpeg.set(exif, offset); offset += exif.length;
    jpeg.set([0xff, 0xda], offset); // stop metadata inspection at the first scan

    const blob = new Blob([jpeg as BlobPart], { type: 'image/jpeg' });
    const cib = vi.fn(async (_blob: Blob, options?: ImageBitmapOptions) => {
      const width = options?.resizeWidth ?? 100;
      const height = options?.resizeHeight ?? 400;
      return { width, height, close() {} } as unknown as ImageBitmap;
    });
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 25,
      targetHeightPx: 25,
      maxRetainedPixels: 2_500,
    })).resolves.toMatchObject({ width: 25, height: 25 });
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 25,
      resizeHeight: 25,
      resizeQuality: 'high',
    });
  });

  it('downsamples a panoramic source wider than the retained-canvas axis limit', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 40_000);
    view.setUint32(20, 2_000);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const resized = { width: 4_000, height: 200, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => resized);
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: 4_000,
      targetHeightPx: 200,
    })).resolves.toBe(resized);
    expect(cib).toHaveBeenCalledWith(blob, {
      resizeWidth: 4_000,
      resizeHeight: 200,
      resizeQuality: 'high',
    });
  });

  it('rejects before decode when the genuinely required target exceeds the pixel limit', async () => {
    const sourceWidth = 12_000;
    const sourceHeight = 9_000;
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, sourceWidth);
    view.setUint32(20, sourceHeight);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const cib = vi.fn();
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(blob, {
      targetWidthPx: sourceWidth,
      targetHeightPx: sourceHeight,
    })).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 1 << 25,
      observed: sourceWidth * sourceHeight,
    });
    expect(cib).not.toHaveBeenCalled();
  });

  it('still rejects an over-32MP raster when no bounded display target is known', async () => {
    const png = new Uint8Array(26);
    png.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    png.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(png.buffer);
    view.setUint32(16, 12_000);
    view.setUint32(20, 9_000);
    const cib = vi.fn();
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(new Blob([png as BlobPart], { type: 'image/png' })))
      .rejects.toMatchObject({
        code: 'ooxml-decoded-image-limit',
        metric: 'image-pixels',
        observed: 12_000 * 9_000,
      });
    expect(cib).not.toHaveBeenCalled();
  });

  it('keeps the encoded-source hard ceiling even when a tiny display target is requested', async () => {
    const bomb = new Uint8Array(26);
    bomb.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    bomb.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(bomb.buffer);
    view.setUint32(16, 30_000);
    view.setUint32(20, 30_000);
    const cib = vi.fn();
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(
      new Blob([bomb as BlobPart], { type: 'image/png' }),
      { targetWidthPx: 300, targetHeightPx: 300 },
    )).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      observed: 30_000 * 30_000,
    });
    expect(cib).not.toHaveBeenCalled();
  });

  it('admits a non-JPEG source axis above 65,535 when total pixels and output are bounded', async () => {
    const bomb = new Uint8Array(26);
    bomb.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    bomb.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(bomb.buffer);
    view.setUint32(16, 100_000);
    view.setUint32(20, 1_000);
    const bitmap = { width: 1_000, height: 10, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async () => bitmap);
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(
      new Blob([bomb as BlobPart], { type: 'image/png' }),
      { targetWidthPx: 1_000, targetHeightPx: 10 },
    )).resolves.toBe(bitmap);
    expect(cib).toHaveBeenCalledWith(expect.any(Blob), {
      resizeWidth: 1_000,
      resizeHeight: 10,
      resizeQuality: 'high',
    });
  });

  it('reports a retained-axis crossing separately when no downsample target is supplied', async () => {
    const panoramic = new Uint8Array(26);
    panoramic.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
    panoramic.set([0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52], 8);
    const view = new DataView(panoramic.buffer);
    view.setUint32(16, 40_000);
    view.setUint32(20, 1);
    const cib = vi.fn();
    vi.stubGlobal('createImageBitmap', cib);

    await expect(decodeRasterOrMetafile(
      new Blob([panoramic as BlobPart], { type: 'image/png' }),
    )).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-dimension',
      limit: 32_767,
      observed: 40_000,
    });
    expect(cib).not.toHaveBeenCalled();
  });

  it('distinguishes a missing optional TIFF codec from a configured codec failure', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 1, 2, 3, 4]);
    const blob = new Blob([bytes as BlobPart], { type: 'image/tiff' });
    const browserDecode = vi.fn();
    vi.stubGlobal('createImageBitmap', browserDecode);

    await expect(decodeRasterOrMetafile(blob)).rejects.toSatisfy(
      error => isOptionalImageCodecUnavailableError(error, 'tiff'),
    );
    expect(browserDecode).not.toHaveBeenCalled();

    const emptyRender = vi.fn(async () => null);
    await expect(decodeRasterOrMetafile(blob, { tiff: { render: emptyRender } }))
      .rejects.toBeInstanceOf(TiffDecodeError);
    expect(emptyRender).toHaveBeenCalledOnce();
    expect(browserDecode).not.toHaveBeenCalled();
  });

  it('wraps an arbitrary custom TIFF codec failure with a diagnostic TIFF error', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 1, 2, 3, 4]);
    const cause = new Error('custom codec internals');
    const render = vi.fn(async () => { throw cause; });

    const promise = decodeRasterOrMetafile(
      new Blob([bytes as BlobPart], { type: 'application/octet-stream' }),
      { tiff: { render } },
    );
    await expect(promise).rejects.toMatchObject({
      name: 'TiffDecodeError',
      code: 'ooxml-tiff-decode',
      cause,
    });
  });

  it('preserves a decoded-image quota error from another TIFF codec realm', async () => {
    const bytes = new Uint8Array([0x49, 0x49, 0x2a, 0x00, 1, 2, 3, 4]);
    const foreignQuota = {
      name: 'OoxmlDecodedImageLimitError',
      message: 'OOXML decoded image limit exceeded: image-pixels 50 > 49',
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 49,
      observed: 50,
    };
    const render = vi.fn(async () => { throw foreignQuota; });

    await expect(decodeRasterOrMetafile(
      new Blob([bytes as BlobPart], { type: 'application/octet-stream' }),
      { tiff: { render } },
    )).rejects.toBe(foreignQuota);
  });

  it('routes a MIME-identified unsupported TIFF container through the codec path', async () => {
    const bigTiffHeader = new Uint8Array([0x49, 0x49, 0x2b, 0x00, 8, 0, 0, 0]);
    const browserDecode = vi.fn();
    const render = vi.fn(async () => null);
    vi.stubGlobal('createImageBitmap', browserDecode);

    await expect(decodeRasterOrMetafile(
      new Blob([bigTiffHeader as BlobPart], { type: 'image/tiff' }),
      { tiff: { render } },
    )).rejects.toBeInstanceOf(TiffDecodeError);
    expect(render).toHaveBeenCalledOnce();
    expect(browserDecode).not.toHaveBeenCalled();
  });

  it('decodes a TIFF through the opt-in codec and forwards its retained target', async () => {
    const bytes = new Uint8Array(38);
    const view = new DataView(bytes.buffer);
    bytes.set([0x49, 0x49], 0);
    view.setUint16(2, 42, true);
    view.setUint32(4, 8, true);
    view.setUint16(8, 2, true);
    view.setUint16(10, 256, true);
    view.setUint16(12, 4, true);
    view.setUint32(14, 1, true);
    view.setUint32(18, 4_249, true);
    view.setUint16(22, 257, true);
    view.setUint16(24, 4, true);
    view.setUint32(26, 1, true);
    view.setUint32(30, 6_137, true);
    const blob = new Blob([bytes as BlobPart], { type: 'image/tiff' });
    const browserDecode = vi.fn();
    vi.stubGlobal('createImageBitmap', browserDecode);

    const bitmap = { width: 320, height: 463, close() {} } as unknown as ImageBitmap;
    const render = vi.fn(async (
      input: Uint8Array,
      options?: Readonly<{
        targetWidthPx?: number;
        targetHeightPx?: number;
        maxRetainedPixels?: number;
      }>,
    ) => {
      expect(Array.from(input)).toEqual(Array.from(bytes));
      return bitmap;
    });
    await expect(decodeRasterOrMetafile(blob, {
      tiff: { render },
      targetWidthPx: 320,
      targetHeightPx: 200,
      maxRetainedPixels: 1 << 23,
    })).resolves.toBe(bitmap);
    expect(render).toHaveBeenCalledTimes(1);
    expect(render).toHaveBeenCalledWith(bytes, {
      targetWidthPx: 320,
      targetHeightPx: 200,
      maxRetainedPixels: 1 << 23,
    });
    expect(browserDecode).not.toHaveBeenCalled();
  });

  it('rejects an oversized required TIFF target before invoking an injected codec', async () => {
    const bytes = new Uint8Array(38);
    const view = new DataView(bytes.buffer);
    bytes.set([0x49, 0x49], 0);
    view.setUint16(2, 42, true);
    view.setUint32(4, 8, true);
    view.setUint16(8, 2, true);
    view.setUint16(10, 256, true);
    view.setUint16(12, 4, true);
    view.setUint32(14, 1, true);
    view.setUint32(18, 12_000, true);
    view.setUint16(22, 257, true);
    view.setUint16(24, 4, true);
    view.setUint32(26, 1, true);
    view.setUint32(30, 9_000, true);
    const render = vi.fn();

    await expect(decodeRasterOrMetafile(
      new Blob([bytes as BlobPart], { type: 'image/tiff' }),
      { tiff: { render }, targetWidthPx: 12_000, targetHeightPx: 9_000 },
    )).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      observed: 12_000 * 9_000,
    });
    expect(render).not.toHaveBeenCalled();
  });

  it('closes and rejects an oversized decode when the header was unrecognized', async () => {
    const close = vi.fn();
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 8192, height: 8192, close }) as unknown as ImageBitmap),
    );
    const blob = new Blob([new Uint8Array([1, 2, 3])], { type: 'application/octet-stream' });

    await expect(decodeRasterOrMetafile(blob)).rejects.toMatchObject({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      observed: 8192 * 8192,
    });
    expect(close).toHaveBeenCalledTimes(1);
  });

  it('sniffs a header slice for raster: only the header is read, the Blob is handed to createImageBitmap whole', async () => {
    const fake = { width: 3, height: 4, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (_blob: Blob) => fake);
    vi.stubGlobal('createImageBitmap', cib);

    // A raster smaller than the sniff window. slice(0,N) must not throw (Blob.slice
    // clamps to the actual length) and the short `isWmf`/`isEmf` length guards must
    // both return false so the blob falls through to createImageBitmap. This 7-byte
    // "PNG-ish" buffer has no valid IHDR, so the dimension sniff returns null
    // (fail-open) and does not block it.
    const tiny = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]); // 7 bytes, PNG-ish
    const base = new Blob([tiny as BlobPart], { type: 'image/png' });

    // Spy on the two Blob reads: the raster path must slice (header sniff) but
    // must NOT read the whole blob via arrayBuffer() — createImageBitmap consumes
    // the Blob directly.
    let sliced = 0;
    let fullRead = 0;
    const wrapped = {
      type: base.type,
      size: base.size,
      slice(start?: number, end?: number) {
        sliced++;
        return base.slice(start, end);
      },
      async arrayBuffer() {
        fullRead++;
        return base.arrayBuffer();
      },
    } as unknown as Blob;

    const bmp = await decodeRasterOrMetafile(wrapped, { widthPt: 50, heightPt: 50 });
    expect(bmp).toBe(fake);
    expect(sliced).toBe(1); // header sniff happened
    expect(fullRead).toBe(0); // no whole-blob copy on the raster fast path
    expect(cib).toHaveBeenCalledTimes(1);
    expect(cib).toHaveBeenCalledWith(wrapped);
  });

  it('an empty WMF (no geometry) → null, not a throw', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap));
    const file = concat(wmfHeader(), record(FN.EOF, () => {}));
    const blob = new Blob([file as BlobPart], { type: 'image/wmf' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 100, heightPt: 100 });
    expect(bmp).toBeNull();
  });


});

describe('isMetafileMime', () => {
  it('is true only for the WMF/EMF metafile MIME types', () => {
    expect(isMetafileMime('image/wmf')).toBe(true);
    expect(isMetafileMime('image/emf')).toBe(true);
    expect(isMetafileMime('image/png')).toBe(false);
    expect(isMetafileMime('image/jpeg')).toBe(false);
    expect(isMetafileMime('image/svg+xml')).toBe(false);
    expect(isMetafileMime(undefined)).toBe(false);
  });
});

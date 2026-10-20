// ── DIB (Device-Independent Bitmap) decode + blit ────────────────────────────
//
// Both metafile players embed raster bitmaps as DIBs ([MS-WMF] 2.2.2.9
// DeviceIndependentBitmap / [MS-EMF] 2.2.2.9), so the decode+blit lives here and
// is shared by {@link ./emf.ts} (EMR_BITBLT / EMR_STRETCHDIBITS, which pass
// explicit bmi/bits offsets) and {@link ./wmf.ts} (META_STRETCHDIBITS /
// META_DIBBITBLT / META_DIBSTRETCHBLT, which store the DIB *packed* — header +
// palette + pixels contiguously — via {@link decodePackedDib}).

import { createAuxCanvas } from '../canvas/aux-canvas.js';
import { MAX_RASTER_DIMENSION, MAX_RASTER_PIXELS, OoxmlDecodedImageLimitError, isOoxmlDecodedImageLimitError } from './pixel-budget.js';

/**
 * Maximum width or height (px) accepted for a decoded DIB, and the megapixel
 * budget for the `width × height × 4` RGBA buffer. Shared with the standalone
 * raster header sniff (`./raster-dimensions.ts`) via `./pixel-budget` so a DIB
 * embedded in a metafile and a standalone blip are held to the SAME limits: a
 * DIB past the dimension cap could never be blitted (`blitDibToCtx` allocates an
 * aux canvas of exactly `dib.width × dib.height`), and one past the megapixel
 * cap (e.g. a crafted 65535×65535 header ≈ 17 GB RGBA) is refused before the
 * allocation. With both axes ≤ 32767 the product is exact in a double, so a
 * plain numeric comparison suffices — no BigInt.
 */
const MAX_DIB_DIMENSION = MAX_RASTER_DIMENSION;
const MAX_DIB_PIXELS = MAX_RASTER_PIXELS;

/** A decoded DIB as top-down RGBA (what `ImageData`/`putImageData` expects). */
export interface DecodedDib {
  width: number;
  height: number;
  data: Uint8ClampedArray; // RGBA, top-down
}

/**
 * Minimal DIB decoder: BITMAPINFOHEADER (40 bytes) + pixel data, supporting
 * BI_RGB(0) at 32bpp (BGRA), 24bpp (BGR, 4-byte row padding) and 8bpp palette
 * (RGBQUAD palette after the header). `bmiOff`/`bitsOff` are byte offsets into
 * `dv`. Returns `null` for unsupported headers/compression.
 */
export function decodeDib(
  dv: DataView,
  bmiOff: number,
  bmiLen: number,
  bitsOff: number,
  bitsLen: number,
): DecodedDib | null {
  // Offsets and declared spans must belong to the containing view (the EMF
  // player supplies a view ending at its record). Invalid ranges and encodings
  // are content fallback, before any decoded-surface resource admission.
  if (bmiOff < 0 || bitsOff < 0 || bitsLen < 0 || bmiLen < 40
    || bmiOff + bmiLen > dv.byteLength || bitsOff + bitsLen > dv.byteLength) return null;
  const biSize = dv.getUint32(bmiOff, true);
  if (biSize < 40 || biSize > bmiLen) return null; // complete BITMAPINFOHEADER (or larger)
  const biWidth = dv.getInt32(bmiOff + 4, true);
  const biHeightRaw = dv.getInt32(bmiOff + 8, true);
  const biBitCount = dv.getUint16(bmiOff + 14, true);
  const biCompression = dv.getUint32(bmiOff + 16, true);
  if (biCompression !== 0) return null; // BI_RGB only
  if (![1, 4, 8, 24, 32].includes(biBitCount)) return null;
  const topDown = biHeightRaw < 0;
  const width = Math.abs(biWidth);
  const height = Math.abs(biHeightRaw);
  if (width <= 0 || height <= 0) return null;
  // Preserve the existing short-data tolerance: at most one missing row can
  // produce a partial bitmap; more is discarded. Check that content fallback
  // before surface admission. Arithmetic (rather than 32-bit shifts) keeps
  // row sizes valid even for oversized headers that have not been admitted.
  const rowStride = Math.ceil(width * biBitCount / 32) * 4;
  // EMF records may carry unrelated data after the explicit pixel span too.
  // Only bitsLen proves bitmap bytes; the enclosing view is not pixel data.
  if (rowStride * height > bitsLen + rowStride) return null;
  // A tolerated partial row is not proof of a resource crossing. Preserve
  // partial decoding within budget, but oversized inputs need all pixel bytes
  // before a terminal resource error can replace malformed-content fallback.
  const oversized = width > MAX_DIB_DIMENSION || height > MAX_DIB_DIMENSION || width * height > MAX_DIB_PIXELS;
  if (oversized) {
    if (rowStride * height > bitsLen) return null;
    // MS-WMF 2.2.2.9: RGBQUAD entries are the indexed palette, or the
    // optional optimization table at >8bpp. Ordinary decoding tolerates a
    // short palette (black for missing indices) and ignores optimization
    // colors; absent declared bytes cannot prove a terminal policy crossing.
    const entries = dv.getUint32(bmiOff + 32, true) || (biBitCount <= 8 ? 1 << biBitCount : 0);
    if (entries * 4 > bmiLen - biSize) return null;
  }
  // Validate dimensions BEFORE the RGBA allocation. Reject non-positive dims,
  // any dimension past the max canvas size, and any total pixel count over the
  // megapixel budget — a crafted header (e.g. 65535×65535) otherwise reaches
  // `new Uint8ClampedArray(width*height*4)` and OOMs the tab (~17 GB). See
  // MAX_DIB_DIMENSION / MAX_DIB_PIXELS.
  // Library resource policy is terminal, including inside a metafile's
  // compatibility fallback. Unsupported encodings still return null above.
  if (width > MAX_DIB_DIMENSION || height > MAX_DIB_DIMENSION) {
    throw new OoxmlDecodedImageLimitError('image-dimension', MAX_DIB_DIMENSION, Math.max(width, height));
  }
  if (width * height > MAX_DIB_PIXELS) {
    throw new OoxmlDecodedImageLimitError('image-pixels', MAX_DIB_PIXELS, width * height);
  }

  const out = new Uint8ClampedArray(width * height * 4);

  // Palette for ≤8bpp follows the header.
  let palette: number[] | null = null;
  if (biBitCount <= 8) {
    let clrUsed = dv.getUint32(bmiOff + 32, true);
    if (clrUsed === 0) clrUsed = 1 << biBitCount;
    const palOff = bmiOff + biSize;
    palette = [];
    for (let i = 0; i < clrUsed; i++) {
      const o = palOff + i * 4;
      if (o + 4 > bmiOff + bmiLen) break;
      const b = dv.getUint8(o);
      const g = dv.getUint8(o + 1);
      const r = dv.getUint8(o + 2);
      palette.push((r << 16) | (g << 8) | b);
    }
  }

  const putPx = (dstRow: number, x: number, r: number, g: number, b: number, a: number) => {
    const di = (dstRow * width + x) * 4;
    out[di] = r;
    out[di + 1] = g;
    out[di + 2] = b;
    out[di + 3] = a;
  };

  let anyAlpha = false;
  for (let y = 0; y < height; y++) {
    const srcRow = topDown ? y : height - 1 - y; // bottom-up unless negative
    const dstRow = y;
    const rowOff = bitsOff + srcRow * rowStride;
    if (rowOff + rowStride > bitsOff + bitsLen) break;
    if (biBitCount === 32) {
      for (let x = 0; x < width; x++) {
        const o = rowOff + x * 4;
        const b = dv.getUint8(o);
        const g = dv.getUint8(o + 1);
        const r = dv.getUint8(o + 2);
        const a = dv.getUint8(o + 3);
        if (a !== 0) anyAlpha = true;
        putPx(dstRow, x, r, g, b, a);
      }
    } else if (biBitCount === 24) {
      for (let x = 0; x < width; x++) {
        const o = rowOff + x * 3;
        putPx(dstRow, x, dv.getUint8(o + 2), dv.getUint8(o + 1), dv.getUint8(o), 255);
      }
      anyAlpha = true;
    } else if (biBitCount === 8 && palette) {
      for (let x = 0; x < width; x++) {
        const idx = dv.getUint8(rowOff + x);
        const c = palette[idx] ?? 0;
        putPx(dstRow, x, (c >> 16) & 0xff, (c >> 8) & 0xff, c & 0xff, 255);
      }
      anyAlpha = true;
    } else if (biBitCount === 4 && palette) {
      // 4bpp: two palette indices per byte, high nibble first (MATLAB exports the
      // bar-chart fill as a 16-colour STRETCHDIBITS — sample-13 Fig.3 PR_VAR bars).
      for (let x = 0; x < width; x++) {
        const byte = dv.getUint8(rowOff + (x >> 1));
        const idx = (x & 1) === 0 ? (byte >> 4) & 0xf : byte & 0xf;
        const c = palette[idx] ?? 0;
        putPx(dstRow, x, (c >> 16) & 0xff, (c >> 8) & 0xff, c & 0xff, 255);
      }
      anyAlpha = true;
    } else if (biBitCount === 1 && palette) {
      // 1bpp: eight palette indices per byte, MSB first (monochrome pattern
      // brushes — gridlines/axes).
      for (let x = 0; x < width; x++) {
        const byte = dv.getUint8(rowOff + (x >> 3));
        const bit = (byte >> (7 - (x & 7))) & 1;
        const c = palette[bit] ?? 0;
        putPx(dstRow, x, (c >> 16) & 0xff, (c >> 8) & 0xff, c & 0xff, 255);
      }
      anyAlpha = true;
    } else {
      return null; // unsupported bit depth
    }
  }

  // 32bpp DIBs frequently store alpha=0 throughout (the BITBLT raster-op carries
  // opacity instead). Treat an all-zero alpha plane as fully opaque.
  if (biBitCount === 32 && !anyAlpha) {
    for (let i = 3; i < out.length; i += 4) out[i] = 255;
  }
  return { width, height, data: out };
}

/**
 * Decode a *packed* DIB — BITMAPINFOHEADER + optional palette + pixel bits, laid
 * out contiguously — which is how WMF stores it inline in a record (the EMF
 * records instead carry explicit bmi/bits offsets, so they call {@link decodeDib}
 * directly). Reads `biSize`/`biBitCount`/`biClrUsed` from the header, computes
 * the palette byte count, derives the pixel-bits offset, and delegates to
 * {@link decodeDib}. `dibOff`/`dibLen` bound the packed blob within `dv`.
 */
export function decodePackedDib(dv: DataView, dibOff: number, dibLen: number): DecodedDib | null {
  if (dibOff < 0 || dibLen < 40 || dibOff + dibLen > dv.byteLength) return null;
  // MS-WMF 2.3.1.6: the packed bitmap belongs to this record alone. Bound
  // every header/palette/pixel read before decodeDib tests truncation or admits
  // a surface; following records cannot supply this bitmap's missing bytes.
  dv = new DataView(dv.buffer, dv.byteOffset + dibOff, dibLen);
  dibOff = 0;
  const biSize = dv.getUint32(dibOff, true);
  if (biSize < 40) return null;
  const biBitCount = dv.getUint16(dibOff + 14, true);
  let palEntries = 0;
  if (biBitCount <= 8) {
    let clrUsed = dv.getUint32(dibOff + 32, true);
    if (clrUsed === 0) clrUsed = 1 << biBitCount;
    palEntries = clrUsed;
  } else {
    // A >8bpp DIB carries no indexed palette, but MAY still prepend an OPTIONAL
    // optimization color table when `biClrUsed > 0` (BITMAPINFOHEADER /
    // [MS-WMF] 2.2.2.9); the pixel bits then follow AFTER it. Skip those bytes so
    // the derived `bitsOff` stays aligned (0 = no table, the common case).
    palEntries = dv.getUint32(dibOff + 32, true);
  }
  const palBytes = palEntries * 4;
  const bmiLen = biSize + palBytes;
  const bitsOff = dibOff + bmiLen;
  const bitsLen = dibLen - bmiLen;
  if (bitsLen <= 0) return null;
  return decodeDib(dv, dibOff, bmiLen, bitsOff, bitsLen);
}

/**
 * Blit a decoded DIB into the device rect `[x0,y0]`–`[x1,y1]` on `ctx`, via a
 * temp OffscreenCanvas (`putImageData` the RGBA in, then `drawImage`-scale to the
 * dest). Returns `true` if it drew, `false` if OffscreenCanvas is unavailable (no
 * browser/worker canvas API) or any step throws. `x1<x0` / `y1<y0` are normalized
 * to absolute width/height (the top-left corner is the min of the two corners).
 */
export function blitDibToCtx(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  dib: DecodedDib,
  x0: number,
  y0: number,
  x1: number,
  y1: number,
): boolean {
  try {
    // Shared aux canvas (OffscreenCanvas, else a detached <canvas>); null in a
    // headless env ⇒ skip the blit and keep rendering, as before.
    const tmp = createAuxCanvas(dib.width, dib.height);
    if (!tmp) return false;
    const tctx = tmp.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
    if (!tctx) return false;
    // Allocate via createImageData (avoids the ImageData-constructor typed-array
    // overload mismatch) and copy the decoded RGBA in.
    const imgData = tctx.createImageData(dib.width, dib.height) as ImageData;
    imgData.data.set(dib.data);
    tctx.putImageData(imgData, 0, 0);
    const dx = Math.min(x0, x1);
    const dy = Math.min(y0, y1);
    const dw = Math.abs(x1 - x0);
    const dh = Math.abs(y1 - y0);
    ctx.drawImage(tmp, dx, dy, dw, dh);
    return true;
  } catch (error) {
    if (isOoxmlDecodedImageLimitError(error)) throw error;
    // Unsupported / missing image API → skip the blit, keep rendering.
    return false;
  }
}

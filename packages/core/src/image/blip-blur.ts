import { MAX_RASTER_PIXELS, MAX_RASTER_DIMENSION, MAX_IMAGE_EFFECT_PIXEL_WORK, OoxmlDecodedImageLimitError } from './pixel-budget.js';

import type { EffectBuffer } from './image-effect-frame.js';
export type { EffectBuffer } from './image-effect-frame.js';

export function blurPlan(width: number, height: number, radiusX: number, radiusY: number, grow: boolean) {
  // §20.1.8.15 specifies radius/all channels/grow, but no kernel. Library
  // rasterization policy: normalized Gaussian, sigma=radius/3, truncated at
  // ceil(radius) source-pixel samples. This finite support gives rad a precise bound and
  // keeps both separable passes and allocation preflight exact. No Office
  // pixel equivalence is claimed. Outside pixels are transparent black;
  // grow=false clips that same convolution to the original bounds.
  const rx = Math.ceil(radiusX), ry = Math.ceil(radiusY);
  const w = width + (grow ? 2 * rx : 0), h = height + (grow ? 2 * ry : 0);
  const taps = 2 * rx + 2 * ry + 2;
  const work = w * h * (taps + 2) + 2 * taps;
  const kernelBytes = (taps + 4) * Float64Array.BYTES_PER_ELEMENT;
  const pixels = w * h;
  if (![w, h, pixels, work].every(Number.isSafeInteger) || work > MAX_IMAGE_EFFECT_PIXEL_WORK) {
    throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK,
      Number.isSafeInteger(work) ? work : Number.MAX_SAFE_INTEGER);
  }
  if (w > MAX_RASTER_DIMENSION || h > MAX_RASTER_DIMENSION) {
    throw new OoxmlDecodedImageLimitError('image-dimension', MAX_RASTER_DIMENSION, Math.max(w, h));
  }
  // Local convolution owns its input, horizontal scratch and output. The
  // enclosing graph planner adds the original decoded source, retained masks,
  // fill resources and the final bitmap bake before admitting any allocation.
  const workspaceBytes = (width * height + 2 * pixels) * 4 + kernelBytes;
  if (workspaceBytes > MAX_RASTER_PIXELS * 4) {
    throw new OoxmlDecodedImageLimitError('active-decoded-bytes', MAX_RASTER_PIXELS * 4, workspaceBytes);
  }
  return { rx, ry, width: w, height: h, work, kernelBytes };
}
function weights(radius: number, r: number): Float64Array {
  const out = new Float64Array(2 * r + 1);
  if (radius === 0) { out[0] = 1; return out; }
  let sum = 0;
  for (let k = -r; k <= r; k++) {
    const v = Math.exp(-.5 * (k / (radius / 3)) ** 2);
    out[k + r] = v; sum += v;
  }
  for (let i = 0; i < out.length; i++) out[i] /= sum;
  return out;
}
export function blurPixels(input: EffectBuffer, radiusX: number, radiusY: number, grow: boolean): EffectBuffer {
  if (radiusX === 0 && radiusY === 0) return input;
  const plan = blurPlan(input.width, input.height, radiusX, radiusY, grow);
  const { rx, ry, width, height } = plan;
  const ox = grow ? rx : 0, oy = grow ? ry : 0;
  const wx = weights(radiusX, rx), wy = weights(radiusY, ry);
  // Byte scratch keeps the four-RGBA-surface envelope; each separable pass
  // rounds once to the same 8-bit grid as other ordered raster operations.
  // Filter premultiplied RGB so invisible colors cannot bleed into edges.
  const scratch = new Uint8ClampedArray(width * height * 4);
  const output = new Uint8ClampedArray(scratch.length);
  const sums = new Float64Array(4);
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    const sy = y - oy;
    if (sy < 0 || sy >= input.height) continue;
    sums.fill(0);
    for (let k = -rx; k <= rx; k++) {
      const sx = x - ox + k;
      if (sx < 0 || sx >= input.width) continue;
      const j = (sy * input.width + sx) * 4, weight = wx[k + rx];
      const a = input.data[j + 3] / 255;
      for (let c = 0; c < 3; c++) sums[c] += input.data[j + c] * a * weight;
      sums[3] += input.data[j + 3] * weight;
    }
    const i = (y * width + x) * 4;
    for (let c = 0; c < 4; c++) scratch[i + c] = Math.round(sums[c]);
  }
  for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
    sums.fill(0);
    for (let k = -ry; k <= ry; k++) {
      const sy = y + k;
      if (sy < 0 || sy >= height) continue;
      const j = (sy * width + x) * 4, weight = wy[k + ry];
      for (let c = 0; c < 4; c++) sums[c] += scratch[j + c] * weight;
    }
    const i = (y * width + x) * 4;
    const alpha = sums[3];
    for (let c = 0; c < 3; c++) output[i + c] = alpha > 0 ? Math.round(sums[c] * 255 / alpha) : 0;
    output[i + 3] = Math.round(alpha);
  }
  const f = input.frame ?? { x: 0, y: 0, width: input.width, height: input.height };
  return { data: output, width, height,
    frame: { ...f, x: f.x + ox, y: f.y + oy } };
}

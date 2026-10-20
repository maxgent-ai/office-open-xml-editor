import type { EffectBuffer } from './image-effect-frame.js';
import { MAX_RASTER_DIMENSION, MAX_RASTER_PIXELS, MAX_IMAGE_EFFECT_PIXEL_WORK, OoxmlDecodedImageLimitError } from './pixel-budget.js';

export type Affine = readonly [number, number, number, number, number, number];
export function affinePlan(input: Pick<EffectBuffer, 'width' | 'height' | 'frame'>, matrix: Affine) {
  const f = input.frame ?? { x: 0, y: 0, width: input.width, height: input.height };
  const [a, b, c, d, tx, ty] = matrix;
  const points = [[-f.x, -f.y], [input.width - f.x, -f.y], [-f.x, input.height - f.y], [input.width - f.x, input.height - f.y]]
    .map(([x, y]) => [a * x + c * y + tx, b * x + d * y + ty]);
  const left = Math.floor(Math.min(...points.map(p => p[0]))), top = Math.floor(Math.min(...points.map(p => p[1])));
  const width = Math.max(1, Math.ceil(Math.max(...points.map(p => p[0]))) - left);
  const height = Math.max(1, Math.ceil(Math.max(...points.map(p => p[1]))) - top);
  const pixels = width * height;
  if (![left, top, width, height, pixels].every(Number.isSafeInteger) || pixels > MAX_IMAGE_EFFECT_PIXEL_WORK) {
    throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK,
      Number.isSafeInteger(pixels) ? pixels : Number.MAX_SAFE_INTEGER);
  }
  if (width > MAX_RASTER_DIMENSION || height > MAX_RASTER_DIMENSION) {
    throw new OoxmlDecodedImageLimitError('image-dimension', MAX_RASTER_DIMENSION, Math.max(width, height));
  }
  if (pixels * 4 > MAX_RASTER_PIXELS) throw new OoxmlDecodedImageLimitError('image-pixels', MAX_RASTER_PIXELS, pixels * 4);
  return { width, height, frame: { x: left === 0 ? 0 : -left, y: top === 0 ? 0 : -top, width: f.width, height: f.height } };
}

/** §20.1.8.61: [sx tan(kx) tx; tan(ky) sy ty]. Coordinates are local
 * image pixels; callers convert EMU translations before entering the primitive.
 * Bilinear sampling of premultiplied RGBA is rasterization policy. */
export function affinePixels(input: EffectBuffer, matrix: Affine): EffectBuffer {
  const plan = affinePlan(input, matrix), data = new Uint8ClampedArray(plan.width * plan.height * 4);
  const [a, b, c, d, tx, ty] = matrix, det = a * d - b * c;
  if (det === 0) return { ...plan, data };
  const f = input.frame ?? { x: 0, y: 0 };
  for (let y = 0; y < plan.height; y++) for (let x = 0; x < plan.width; x++) {
    const px = x + .5 - plan.frame.x - tx, py = y + .5 - plan.frame.y - ty;
    const sx = (d * px - c * py) / det + f.x - .5;
    const sy = (-b * px + a * py) / det + f.y - .5;
    const x0 = Math.floor(sx), y0 = Math.floor(sy), fx = sx - x0, fy = sy - y0;
    let red = 0, green = 0, blue = 0, alpha = 0;
    for (let yy = 0; yy < 2; yy++) for (let xx = 0; xx < 2; xx++) {
      const ix = x0 + xx, iy = y0 + yy;
      if (ix < 0 || iy < 0 || ix >= input.width || iy >= input.height) continue;
      const i = (iy * input.width + ix) * 4;
      const w = (xx ? fx : 1 - fx) * (yy ? fy : 1 - fy);
      const aw = input.data[i + 3] * w;
      red += input.data[i] * aw; green += input.data[i + 1] * aw; blue += input.data[i + 2] * aw; alpha += aw;
    }
    const i = (y * plan.width + x) * 4;
    data[i] = alpha ? Math.round(red / alpha) : 0; data[i + 1] = alpha ? Math.round(green / alpha) : 0;
    data[i + 2] = alpha ? Math.round(blue / alpha) : 0; data[i + 3] = Math.round(alpha);
  }
  return { ...plan, data };
}

export function unionEffectGrids(a: Pick<EffectBuffer, 'width' | 'height' | 'frame'>,
  b: Pick<EffectBuffer, 'width' | 'height' | 'frame'>) {
  const af = a.frame ?? { x: 0, y: 0, width: a.width, height: a.height };
  const bf = b.frame ?? { x: 0, y: 0, width: b.width, height: b.height };
  const x = Math.max(af.x, bf.x), y = Math.max(af.y, bf.y);
  const width = x + Math.max(a.width - af.x, b.width - bf.x);
  const height = y + Math.max(a.height - af.y, b.height - bf.y);
  // The union is a new surface: enforce the declared per-axis dimension
  // quota at planning time, before any buffer or canvas is allocated.
  if (!(width <= MAX_RASTER_DIMENSION) || !(height <= MAX_RASTER_DIMENSION)) {
    throw new OoxmlDecodedImageLimitError('image-dimension', MAX_RASTER_DIMENSION,
      Number.isFinite(Math.max(width, height)) ? Math.max(width, height) : Number.MAX_SAFE_INTEGER);
  }
  return { width, height, frame: { x, y, width: af.width, height: af.height } };
}

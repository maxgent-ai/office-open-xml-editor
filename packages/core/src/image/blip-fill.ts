import type { Fill } from '../types/common.js';
import type { BlipEffectFill } from './blip-effects.js';
import type { EffectBuffer } from './blip-blur.js';
import { imageEffectFrame } from './image-effect-frame.js';
import { resolveFill } from '../shape/paint.js';
import { withStrictBrushSurfaces } from '../canvas/strict-brush-surface.js';
import { drawImageCropped, imageNaturalSize } from './crop.js';
import { defaultOffscreenFactory, type OffscreenFactory } from './duotone.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import { MAX_IMAGE_EFFECT_PIXEL_WORK, OoxmlDecodedImageLimitError } from './pixel-budget.js';

import type { ImageEffectWorkLedger } from './image-effect-work.js';

export type BlendMode = 'darken' | 'lighten' | 'mult' | 'over' | 'screen';
export interface BlipEffectRuntime {
  readonly offscreenFactory?: OffscreenFactory;
  readonly images?: ReadonlyMap<object, CanvasImageSource>;
  readonly baseImage?: CanvasImageSource;
  readonly workLedger?: ImageEffectWorkLedger;
}

/** §20.1.8.29 / §20.1.10.11. Separable blend colors, then Porter-Duff
 * source-over in premultiplied form; return straight RGBA. `over` means
 * overlaying the fill (source-over), not Photoshop's unrelated Overlay mode. */
export function blendPixel(base: ArrayLike<number>, fill: ArrayLike<number>, mode: BlendMode): number[] {
  const out = [0, 0, 0, 0];
  blendRgba(base, 0, fill, 0, out, 0, mode); return out;
}
/** One in-place-safe blend, with no pixel-loop allocations. Negative offsets
 * sample transparent pixels outside an effect's finite grid. */
export function blendRgba(base: ArrayLike<number>, bi: number, fill: ArrayLike<number>, fi: number,
  out: { [n: number]: number }, oi: number, mode: BlendMode): void {
  const ab = bi < 0 ? 0 : base[bi + 3] / 255, af = fi < 0 ? 0 : fill[fi + 3] / 255;
  const a = af + ab * (1 - af);
  for (let c = 0; c < 3; c++) {
    const b = bi < 0 ? 0 : base[bi + c] / 255, f = fi < 0 ? 0 : fill[fi + c] / 255;
    const blended = mode === 'darken' ? Math.min(b, f) : mode === 'lighten' ? Math.max(b, f)
      : mode === 'mult' ? b * f : mode === 'screen' ? b + f - b * f : f;
    out[oi + c] = a > 0 ? Math.round(255 * ((1 - af) * ab * b + (1 - ab) * af * f + ab * af * blended) / a) : 0;
  }
  out[oi + 3] = Math.round(a * 255);
}
// Resource policy: geometry must remain finite and retain integer tile parity.
// A finite work product alone cannot prove that an authored grid is iterable.
function rejectFillGeometry(): never {
  throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK, Number.MAX_SAFE_INTEGER);
}
function finiteGeometry(...values: number[]): void {
  if (!values.every(v => Number.isFinite(v) && Math.abs(v) <= Number.MAX_SAFE_INTEGER)) rejectFillGeometry();
}

/** Same finite grid used by planning and painting; no authored tile can escape
 * the graph's cumulative work ledger through a separate local draw loop. */
function effectTilePlan(fill: Extract<BlipEffectFill, { fillType: 'image' }>, image: CanvasImageSource, w: number, h: number,
  hostRotation = 0) {
  const tile = fill.tile as NonNullable<typeof fill.tile>, natural = imageNaturalSize(image);
  finiteGeometry(fill.dpi ?? 96);
  // §20.1.8.58. Missing placement facts use the shared image painter's
  // top-left, unscaled grid. The source's nominal 96 dpi is the existing
  // decoder policy when embedded DPI is unavailable.
  const frame = imageEffectFrame(image) ?? { x: 0, y: 0, width: natural.w, height: natural.h };
  // §20.1.8.58 repeats the authored image frame, windowed by srcRect
  // (§20.1.8.55; Annex L.4.8.4.3: the source rectangle is the content that
  // tiling duplicates), so the period is the cropped extent. Spatial-effect
  // coverage does not change that period (the crop painter owns its outset).
  const crop = fill.srcRect;
  const cropW = 1 - (crop?.l ?? 0) - (crop?.r ?? 0), cropH = 1 - (crop?.t ?? 0) - (crop?.b ?? 0);
  finiteGeometry(cropW, cropH);
  // A crop with no positive extent selects no content: nothing is tiled.
  if (!(cropW > 0 && cropH > 0)) return null;
  const periodW = frame.width * cropW, periodH = frame.height * cropH;
  const tw = periodW * 72 / (fill.dpi || 96) * (tile.sx ?? 1);
  const th = periodH * 72 / (fill.dpi || 96) * (tile.sy ?? 1);
  finiteGeometry(tw, th);
  if (!(tw > 0 && th > 0)) rejectFillGeometry();
  // An off-frame tile can still paint into the fill through its effect
  // outset. Reflected tiles swap the two edges, so bound both with the larger
  // support, measured in tile periods. These exact support bounds also govern
  // preflight draw work.
  // rotWithShape="0" counter-rotates the brush about the box centre (see
  // effectFillPixels); the clip then spans the inverse-rotated box, whose
  // bounding box exceeds w×h. Enumerate (and charge) the tiles covering it,
  // keeping the phase anchored to the authored box.
  let marginX = 0, marginY = 0;
  if (fill.rotWithShape === false && hostRotation !== 0) {
    finiteGeometry(hostRotation);
    const r = hostRotation * Math.PI / 180, c = Math.abs(Math.cos(r)), sn = Math.abs(Math.sin(r));
    marginX = Math.max(0, (w * c + h * sn - w) / 2);
    marginY = Math.max(0, (w * sn + h * c - h) / 2);
  }
  const padX = Math.ceil(Math.max(0, frame.x, natural.w - frame.x - frame.width) / periodW) + Math.ceil(marginX / tw);
  const padY = Math.ceil(Math.max(0, frame.y, natural.h - frame.y - frame.height) / periodH) + Math.ceil(marginY / th);
  const cols = Math.ceil(w / tw) + 2 + 2 * padX, rows = Math.ceil(h / th) + 2 + 2 * padY;
  const work = cols * rows;
  if (!Number.isSafeInteger(work) || work > MAX_IMAGE_EFFECT_PIXEL_WORK) {
    throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK,
      Number.isSafeInteger(work) ? work : Number.MAX_SAFE_INTEGER);
  }
  const alignment = tile.algn ?? 'tl';
  const ax = alignment.endsWith('r') || alignment === 'r' ? w - tw
    : ['t', 'ctr', 'b'].includes(alignment) ? (w - tw) / 2 : 0;
  const ay = alignment.startsWith('b') ? h - th
    : ['l', 'ctr', 'r'].includes(alignment) ? (h - th) / 2 : 0;
  const ox = ax + (tile.tx ?? 0) / 12700, oy = ay + (tile.ty ?? 0) / 12700;
  finiteGeometry(ox, oy);
  const x0 = Math.floor(-ox / tw) - padX, y0 = Math.floor(-oy / th) - padY;
  if (![x0, y0].every(Number.isSafeInteger)) rejectFillGeometry();
  // §20.1.8.58: translation changes phase, not the number of tiles.
  // Reduce coordinates before painting; retain the original index parity
  // for flip. Bounded counters always advance, even near IEEE-754 limits.
  const phase = (offset: number, size: number) => {
    const r = offset % size;
    return r > 0 ? r - size : r;
  };
  const px = phase(ox, tw) - padX * tw, py = phase(oy, th) - padY * th;
  finiteGeometry(px + (cols - 1) * tw, py + (rows - 1) * th);
  // Each output pixel can be covered by this many neighboring tile images.
  // Include one boundary tile per axis for pixel footprints/antialiasing;
  // counting draw calls alone undercharges grown sources with heavy overdraw.
  const overlap = (Math.ceil(natural.w / periodW) + 1) * (Math.ceil(natural.h / periodH) + 1);
  return { cols, rows, work, overlap, x0, y0, px, py, tw, th };
}

/** Additional fill generation/readback visits (the graph charges compositing
 * separately). Before decode, omit source-dependent tiles; the complete plan
 * is mandatory once nested sources are prepared, before any effect paints. */
export function effectFillWork(fill: BlipEffectFill, input: Omit<EffectBuffer, 'data'>,
  emuX: number, emuY: number, runtime?: BlipEffectRuntime, hostRotation = 0): number {
  const pixels = input.width * input.height;
  if (fill.fillType === 'none') return 0;
  if (fill.fillType === 'solid') return pixels;
  let work = pixels * 2; // rasterize and read back
  if (fill.fillType === 'image' && fill.tile && runtime) {
    const image = runtime.images?.get(fill);
    if (!image) throw new OoxmlError('unsupported-image-effect', 'Image fill source unavailable for work planning');
    const f = input.frame ?? { x: 0, y: 0, width: input.width, height: input.height };
    const plan = effectTilePlan(fill, image, f.width / (12700 / emuX), f.height / (12700 / emuY), hostRotation);
    if (plan) work += plan.work + pixels * plan.overlap;
  }
  return work;
}

export function effectFillPixels(
  fill: BlipEffectFill, input: EffectBuffer, emuX: number, emuY: number, runtime: BlipEffectRuntime, hostRotation = 0,
): EffectBuffer {
  if (fill.fillType === 'solid' || fill.fillType === 'none') {
    const data = new Uint8ClampedArray(input.width * input.height * 4);
    if (fill.fillType === 'solid') {
      const h = fill.color.replace(/^#/, '');
      const rgba = [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16));
      rgba.push(h.length === 8 ? parseInt(h.slice(6), 16) : 255);
      const frame = input.frame ?? { x: 0, y: 0, width: input.width, height: input.height };
      // A preceding grow changes the effect grid, not the authored fill box.
      // The fast solid path has the same bounds as resolveFill/fillRect below.
      for (let y = 0; y < input.height; y++) for (let x = 0; x < input.width; x++) {
        if (x + .5 < frame.x || x + .5 >= frame.x + frame.width || y + .5 < frame.y || y + .5 >= frame.y + frame.height) continue;
        data.set(rgba, (y * input.width + x) * 4);
      }
    }
    return { data, width: input.width, height: input.height };
  }
  const factory = runtime.offscreenFactory ?? defaultOffscreenFactory;
  const surface = factory(input.width, input.height);
  const ctx = surface?.getContext('2d') as CanvasRenderingContext2D | null | undefined;
  if (surface && !ctx) (surface as { width: number }).width = 0;
  if (!surface || !ctx) throw new OoxmlError('unsupported-image-effect', 'Image fill surface unavailable');
  const f = input.frame ?? { x: 0, y: 0, width: input.width, height: input.height };
  const sx = 12700 / emuX, sy = 12700 / emuY;
  const w = f.width / sx, h = f.height / sy;
  try {
    finiteGeometry(f.x, f.y, sx, sy, w, h);
    if (!(sx > 0 && sy > 0 && w >= 0 && h >= 0)) rejectFillGeometry();
    ctx.translate(f.x, f.y); ctx.scale(sx, sy);
    if (fill.fillType === 'image') {
      const image = runtime.images?.get(fill);
      if (!image) throw new OoxmlError('unsupported-image-effect', 'Image fill source unavailable');
      ctx.save(); ctx.beginPath(); ctx.rect(0, 0, w, h); ctx.clip();
      // CT_BlipFillProperties@rotWithShape has no schema default. An
      // explicit false counter-rotates the brush inside the unchanged clip.
      if (fill.rotWithShape === false && hostRotation !== 0) {
        ctx.translate(w / 2, h / 2); ctx.rotate(-hostRotation * Math.PI / 180); ctx.translate(-w / 2, -h / 2);
      }
      if (fill.alpha != null) ctx.globalAlpha *= fill.alpha;
      if (fill.tile) {
        const tile = fill.tile;
        const plan = effectTilePlan(fill, image, w, h, hostRotation);
        const { cols, rows, x0, y0, px, py, tw, th } = plan ?? { cols: 0, rows: 0, x0: 0, y0: 0, px: 0, py: 0, tw: 0, th: 0 };
        for (let row = 0; row < rows; row++) for (let col = 0; col < cols; col++) {
          const flipX = (tile.flip === 'x' || tile.flip === 'xy') && Math.abs((x0 % 2 + col) % 2) === 1;
          const flipY = (tile.flip === 'y' || tile.flip === 'xy') && Math.abs((y0 % 2 + row) % 2) === 1;
          ctx.save(); ctx.translate(px + col * tw, py + row * th);
          ctx.scale(flipX ? -1 : 1, flipY ? -1 : 1);
          drawImageCropped(ctx, image, fill.srcRect, flipX ? -tw : 0, flipY ? -th : 0, tw, th);
          ctx.restore();
        }
      } else {
        const r = fill.fillRect;
        const dx = (r?.l ?? 0) * w, dy = (r?.t ?? 0) * h;
        const dw = w * (1 - (r?.l ?? 0) - (r?.r ?? 0)), dh = h * (1 - (r?.t ?? 0) - (r?.b ?? 0));
        finiteGeometry(dx, dy, dw, dh, dx + dw, dy + dh);
        drawImageCropped(ctx, image, fill.srcRect, dx, dy, dw, dh);
      }
      ctx.restore();
    } else {
      // Gradients (including path shades), patterns and solid colors use the
      // existing shared fill painter, preserving its palette/geometry policy.
      const recipe: Fill = fill.fillType === 'gradient' ? { ...fill, stops: [...fill.stops] } : fill;
      // Strict scope: an unavailable tiled-gradient/path-shade surface is the
      // typed effect error, never the ordinary native-gradient degrade.
      const paint = withStrictBrushSurfaces(() => resolveFill(recipe, ctx, 0, 0, w, h, hostRotation, 1));
      // The ordinary pattern painter can fall back to a solid color when
      // its preset/surface is unavailable. Recursive effects must fail closed.
      if (!paint || (fill.fillType === 'pattern' && typeof paint === 'string')) throw new OoxmlError('unsupported-image-effect', 'Malformed or unavailable image-effect fill brush');
      ctx.fillStyle = paint; ctx.fillRect(0, 0, w, h);
    }
    return ctx.getImageData(0, 0, input.width, input.height);
  } finally {
    // Release the raster backing store before the next graph branch.
    (surface as { width: number; height: number }).width = 0;
  }
}

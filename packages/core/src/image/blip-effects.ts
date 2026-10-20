// Shared CT_Blip pixel effects (ECMA-376 §20.1.8.13) beyond the duotone:
// `<a:grayscl>` (§20.1.8.34), `<a:biLevel>` (§20.1.8.11) and `<a:clrChange>`
// (§20.1.8.16), applied in document order to the decoded raster. The parser
// emits complete chains; each duotone carries its own endpoints and alpha
// modulation is baked here, never reapplied through a global alpha field.
//
// Office behaviour that the standard leaves implicit, and the evidence used:
// - Luminance. ECMA-376 converts to gray "corresponding to their luminance"
//   and thresholds bi-level by luminance without defining the weights.
//   Input: pictures under grayscl followed by biLevel(50%), exported to PDF by
//   PowerPoint. Result: (2,167,223) turns white and (0,147,190), (0,126,229),
//   (9,74,178) black. Rec. 601 weights put the first below 50% (0.486);
//   Rec. 709 weights (0.532) put all four on the observed sides. Boundary
//   input: a JPEG picture under the same effects, compared with PowerPoint's
//   PDF at every pixel whose source block is a single colour that the
//   candidate rules classify differently (7,068 blocks of blues around the
//   threshold). Result: Rec. 709 luma of the stored (gamma-encoded) values,
//   with the grayscale truncated to a whole 8-bit level, matches all 7,068.
//   Rounding that level instead fails 30 (a luma of 127.59 must become 127,
//   black), unquantized Rec. 709 fails 30, Rec. 601 fails at least 1,534, the
//   channel average 4,553, and linear-light Rec. 709 is contradicted on nearly
//   every block. So grayscl writes floor(Rec. 709 luma) and biLevel compares
//   Rec. 709 luma with the threshold. (No block sat exactly on a threshold, so
//   >= follows the text: "values greater than or equal to the threshold are
//   set to white".)
// - clrChange. MS-OI29500 says Office leaves alpha alone unless useA is set.
//   Input: PowerPoint's own "set transparent colour" output (clrTo = the same
//   colour with alpha 0, useA absent) on two pictures, exported to PDF by
//   PowerPoint and compared through the PDF soft masks. Result: every block
//   that is pure clrFrom white becomes transparent (2,577 blocks, no
//   counterexample; 971 of them transparent where keeping the source alpha
//   would leave them opaque), and every non-matching block stays opaque.
//   Exact RGB matches therefore take clrTo's colour and alpha; with useA the
//   source alpha must match clrFrom's alpha as well.
// - lum. ECMA-376 §20.1.8.42 names brightness and contrast but gives no
//   formula. Evidence: PowerPoint's PDF export of a 256-step gray ramp under
//   a grid of bright × contrast values (±35%, ±70%, ±100%, and 0). Every
//   channel follows
//     out = clamp(k · (v + b/2 − ½) + ½ + b/2),
//   with v and out in 0–1 and b the brightness fraction. k = 1 + c for
//   contrast c ≤ 0 and 1 / (1 − c) for c > 0. So half of the brightness
//   shift is applied before the contrast scaling around mid-gray and half
//   after. This matches every measured level to within 0.7 of 255. The ramp
//   saved as a binary .ppt matches it too, with the stored contrast used
//   directly as k. At c = 1 the result is a threshold at v + b/2 = ½.

import { imageEffectError } from './image-effect-error.js';
import { ImageEffectWorkLedger } from './image-effect-work.js';
import { validateBlipEffect } from './blip-validation.js';
import { fillPaintWorkspaceBytes } from '../shape/paint.js';
import { affinePlan, affinePixels, unionEffectGrids, type Affine } from './blip-spatial.js';
import { expandBlipEffect, blipMaskOperations, effectDirection } from './blip-compound.js';
import type { SolidFill, NoFill, PatternFill, ImageFill, GradientFill, GradientStop, Shadow, Glow, Reflection } from '../types/common.js';
import { blendRgba, effectFillPixels, effectFillWork, type BlendMode, type BlipEffectRuntime } from './blip-fill.js';
import { blurPlan, blurPixels, type EffectBuffer } from './blip-blur.js';
import { rgbToHsl, hslToRgb, interpolateColorChannel } from '../shape/color-space.js';
import { OoxmlError } from '../errors/ooxml-error.js';
import type { Duotone } from './duotone';
import { duotoneImageData, hex6ToRgb } from './duotone';
import {
  MAX_IMAGE_EFFECT_PASSES,
  MAX_RASTER_PIXELS,
  MAX_IMAGE_EFFECT_PIXEL_WORK,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';

/** One CT_Blip pixel effect, as emitted by the parsers (camelCase JSON). */
export type BlipEffect =
  | { type: 'grayscale' }
  | { type: 'biLevel'; thresh: number }
  | {
      type: 'colorChange';
      from: string;
      fromAlpha: number;
      to: string;
      toAlpha: number;
      useAlpha: boolean;
    }
  | { type: 'luminance'; bright: number; contrast: number }
  // Optional endpoints accept legacy model-source markers; parsers always
  // emit this entry's own endpoints. Legacy markers use value.duotone.
  | { type: 'duotone'; clr1?: string; clr2?: string }
  | { type: 'alphaBiLevel'; thresh: number }
  | { type: 'alphaCeiling' | 'alphaFloor' }
  | { type: 'alphaInv'; color?: string }
  | { type: 'hsl'; hue: number; sat: number; lum: number }
  | { type: 'tint'; hue: number; amt: number }
  | { type: 'blur'; radius: number; grow: boolean }
  | { type: 'alphaMod'; container: BlipEffectContainer }
  | { type: 'container'; container: BlipEffectContainer }
  | { type: 'reference'; reference: string }
  | { type: 'definitions'; containers: readonly BlipEffectContainer[]; work: number }
  | { type: 'effectCountLimit'; observed: number }
  | { type: 'alphaOutset'; radius: number }
  | { type: 'blend'; container: BlipEffectContainer; blend: BlendMode }
  | { type: 'fill'; fill: BlipEffectFill }
  | { type: 'glow'; glow: Glow }
  | { type: 'shadow'; shadow: Shadow; inner: boolean }
  | { type: 'softEdge'; radius: number }
  | { type: 'reflection'; reflection: Reflection & { kx?: number; ky?: number; algn?: Shadow['algn']; fadeDir?: number; rotWithShape?: boolean } }
  | { type: 'transform'; sx: number; sy: number; kx: number; ky: number; tx: number; ty: number }
  | { type: 'relativeOffset'; tx: number; ty: number }
  | { type: 'presetShadow'; shadows: readonly Shadow[] }
  | { type: 'fillOverlay'; fill: BlipEffectFill; blend: BlendMode }
  | { type: 'alphaModFix'; amt: number }
  | { type: 'alphaRepl'; alpha: number }
  | { type: 'colorReplace'; color: string }
  | { type: 'unsupported'; effect: string };

export type BlipEffectImageFill = Omit<ImageFill, 'blipEffects'> & { readonly blipEffects?: readonly BlipEffect[] };
export type BlipEffectFill = SolidFill | NoFill | PatternFill
  | (Omit<GradientFill, 'stops'> & { readonly stops: readonly GradientStop[] }) | BlipEffectImageFill;

export interface BlipEffectContainer {
  readonly kind: 'sib' | 'tree';
  readonly name?: string;
  readonly effects: readonly BlipEffect[];
}

/** Ordered pixel effects plus the legacy single-duotone palette, retained
 *  for model-source callers that still emit a position-only marker. */
export interface BlipPixelEffects {
  readonly effects: readonly BlipEffect[];
  readonly duotone?: Duotone | null;
  /** Physical sampling of this image; defaults to 96 dpi when absent. */
  readonly emuPerPixelX?: number;
  readonly emuPerPixelY?: number;
  /** Host orientation in degrees, before its Canvas paint transform. These
   * inputs also own nested fills/containers and participate in cache identity. */
  readonly hostRotation?: number;
  readonly hostFlipH?: boolean;
  readonly hostFlipV?: boolean;
}

export function isBlipPixelEffects(value: unknown): value is BlipPixelEffects {
  return typeof value === 'object' && value !== null && Array.isArray((value as BlipPixelEffects).effects);
}

/** Nested fill resources, in authored order. The loader visits these
 * sequentially through the document's decoded owner, never Promise.all. */
export function blipEffectImageFills(effects: readonly BlipEffect[]): BlipEffectImageFill[] {
  const result: BlipEffectImageFill[] = [];
  for (const effect of blipEffectNodes(effects, false)) {
    if ((effect.type === 'fillOverlay' || effect.type === 'fill') && effect.fill.fillType === 'image') result.push(effect.fill);
  }
  return result;
}

/** Whether a chain can paint outside its authored frame: a growing blur
 * (§20.1.8.15), glow (§20.1.8.32), outer/preset shadow (§20.1.8.49/.52),
 * reflection (§20.1.8.50), relative offset (§20.1.8.51), transform
 * (§20.1.8.61) or a positive alpha outset (§20.1.8.4), at any container
 * depth. Pointwise chains never extend coverage, so hosts may keep culling
 * them by the picture rectangle. Nested image-fill recipes are clipped to
 * the picture and do not count. */
export function blipEffectsMayExtendCoverage(effects: readonly BlipEffect[]): boolean {
  const reach = blipEffectsCoverageReach(effects, 1, 1);
  return reach === null || reach.x > 0 || reach.y > 0;
}

/** Conservative reach of a chain's painted coverage beyond its authored
 * frame, per side, in the frame's units (`width`/`height`, EMU for hosts),
 * or null when it cannot be bounded statically (a scaling/skewing transform).
 * Every spatial node at any container depth is summed, so nested or chained
 * outsets are over-estimated, never under-estimated. Hosts use it to keep
 * viewport culling for effect pictures (pointwise chains reach 0). */
export function blipEffectsCoverageReach(
  effects: readonly BlipEffect[], width: number, height: number,
): { x: number; y: number } | null {
  let x = 0, y = 0;
  const finite = (v: number) => (Number.isFinite(v) ? Math.abs(v) : Infinity);
  for (const effect of blipEffectNodes(effects, false)) {
    switch (effect.type) {
      case 'blur': if (effect.grow) { x += finite(effect.radius); y += finite(effect.radius); } break;
      case 'alphaOutset': if (effect.radius > 0) { x += finite(effect.radius); y += finite(effect.radius); } break;
      case 'glow': x += finite(effect.glow.radius); y += finite(effect.glow.radius); break;
      case 'shadow':
        if (!effect.inner) {
          // A scaled or skewed shadow is an affine image of the frame about
          // its alignment point (§20.1.8.49); no cheap static bound, so the
          // host must not cull it.
          const sh = effect.shadow;
          if ((sh.sx ?? 1) !== 1 || (sh.sy ?? 1) !== 1 || (sh.kx ?? 0) !== 0 || (sh.ky ?? 0) !== 0) return null;
          const d = finite(sh.dist) + finite(sh.blur);
          x += d; y += d;
        }
        break;
      case 'presetShadow':
        for (const shadow of effect.shadows) {
          if ((shadow.sx ?? 1) !== 1 || (shadow.sy ?? 1) !== 1 || (shadow.kx ?? 0) !== 0 || (shadow.ky ?? 0) !== 0) return null;
          const d = finite(shadow.dist) + finite(shadow.blur);
          x += d; y += d;
        }
        break;
      case 'reflection': {
        // A reflection mirrors at most the whole frame beyond an edge.
        const r = effect.reflection;
        // Only the plain mirrored copy (|scale| <= 1, no skew) is bounded by
        // one frame extent beyond an edge.
        if (Math.abs(r.sx ?? 1) > 1 || Math.abs(r.sy ?? -1) > 1 || (r.kx ?? 0) !== 0 || (r.ky ?? 0) !== 0) return null;
        const d = finite(r.dist ?? 0) + finite(r.blur ?? 0);
        x += width + d; y += height + d;
        break;
      }
      // §20.1.8.51/.61 measure against the previous effect's grid, which
      // earlier spatial nodes may already have grown; no static bound.
      case 'relativeOffset':
        if (effect.tx !== 0 || effect.ty !== 0) return null;
        break;
      case 'transform':
        return null;
      // Named recipes may run more than once (§20.1.10.22 effectDag);
      // their multiplicity is not statically bounded here.
      case 'reference': case 'definitions':
        return null;
      default: break;
    }
  }
  // The evaluator grows its grid in whole source pixels, and a later alpha
  // operation (alphaInv, alphaRepl, …) can make that rounding margin opaque.
  // The source pixel size is unknown before decode but never exceeds the
  // frame, so each spatial node may add up to one frame extent per side.
  if (x > 0 || y > 0) {
    let spatialNodes = 0;
    for (const effect of blipEffectNodes(effects, false)) {
      if (['blur', 'alphaOutset', 'glow', 'shadow', 'presetShadow', 'reflection']
        .includes(effect.type)) spatialNodes++;
    }
    x += spatialNodes * width; y += spatialNodes * height;
  }
  // Report the same reach on both axes: a host may paint the frame rotated
  // (or as a rotated group leaf), so the bound must not depend on axis.
  const reach = Math.max(x, y);
  return Number.isFinite(reach) ? { x: reach, y: reach } : null;
}

/** Iterative ownership traversal. Image fills own independent decode recipes;
 * container/DAG recipes share the current evaluator. Both consume the same
 * finite projection envelope; active lists detect cycles without recursion. */
function* blipEffectNodes(effects: readonly BlipEffect[], includeImageEffects = true): Generator<BlipEffect> {
  const pending = [{ list: effects, index: 0 }];
  const active = new Set<readonly BlipEffect[]>();
  let count = 0;
  while (pending.length) {
    const frame = pending[pending.length - 1];
    if (frame.index === 0) {
      if (!Array.isArray(frame.list)) throw new OoxmlError('unsupported-image-effect', 'Invalid effect list');
      if (active.has(frame.list)) throw new OoxmlError('unsupported-image-effect', 'Cyclic effect or image fill');
      active.add(frame.list);
    }
    if (frame.index === frame.list.length) { active.delete(frame.list); pending.pop(); continue; }
    if (++count > MAX_IMAGE_EFFECT_PASSES) throw new OoxmlDecodedImageLimitError('image-effect-count', MAX_IMAGE_EFFECT_PASSES, count);
    const effect = frame.list[frame.index++];
    if (!effect || typeof effect !== 'object') throw new OoxmlError('unsupported-image-effect', 'Invalid effect node');
    const containers = effect.type === 'definitions' ? effect.containers
      : effect.type === 'alphaMod' || effect.type === 'container' || effect.type === 'blend' ? [effect.container] : [];
    if (!Array.isArray(containers) || containers.length > MAX_IMAGE_EFFECT_PASSES) throw new OoxmlError('unsupported-image-effect', 'Invalid effect definitions');
    for (let i = containers.length - 1; i >= 0; i--) {
      const c = containers[i];
      if (!c || !['tree', 'sib'].includes(c.kind)) throw new OoxmlError('unsupported-image-effect', 'Invalid effect container');
      pending.push({ list: c.effects, index: 0 });
    }
    if (effect.type === 'fillOverlay' || effect.type === 'fill') {
      if (!effect.fill) throw new OoxmlError('unsupported-image-effect', 'Invalid effect fill');
      if (includeImageEffects && effect.fill.fillType === 'image' && effect.fill.blipEffects) pending.push({ list: effect.fill.blipEffects, index: 0 });
    }
    yield effect;
  }
}

/** Stable cache-key suffix for a pixel transform. */
export function blipPixelEffectsKey(value: BlipPixelEffects): string {
  assertBlipPixelEffectsBudget(value, 0);
  // Delimiters such as '.' also occur in decimal numbers; effect names and
  // references can contain punctuation too. Encode the full ordered recipe
  // and both physical axes structurally, including named-definition work.
  try { return JSON.stringify([value.effects, value.duotone ?? null,
    value.emuPerPixelX ?? null, value.emuPerPixelY ?? null,
    value.hostRotation ?? 0, value.hostFlipH ?? false, value.hostFlipV ?? false], (_key, value: unknown) => {
    if (typeof value === 'number' && !Number.isFinite(value)) {
      throw new OoxmlError('unsupported-image-effect', 'Non-finite image effect parameter');
    }
    return value;
  }); } catch (error) {
    if (error instanceof OoxmlError) throw error;
    throw new OoxmlError('unsupported-image-effect', 'Image effect recipe is not serializable');
  }
}

/** Rec. 709 luma (0–1) of gamma-encoded sRGB bytes; see the evidence above. */
export function blipLuminance(r: number, g: number, b: number): number {
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
}

/** grayscl's 8-bit level: the Rec. 709 luma truncated to a whole level (see
 *  the evidence above). The epsilon keeps white at 255 despite the weights'
 *  binary rounding. */
export function blipGrayLevel(r: number, g: number, b: number): number {
  return Math.floor(blipLuminance(r, g, b) * 255 + 1e-6);
}

type EffectGrid = Pick<EffectBuffer, 'width' | 'height' | 'frame'>;
interface EffectContext {
  names: Map<string, BlipEffectContainer>;
  original?: EffectBuffer;
  charge(work: number): void;
  chargedDefinitions: Set<BlipEffect>;
}

interface EffectPlan { count: number; work: number; pixels: number; peak: number; width: number; height: number; frame?: EffectBuffer['frame'] }
function effectPlan(value: BlipPixelEffects, pixels: number, geometry?: EffectGrid, runtime?: BlipEffectRuntime): EffectPlan {
  if (value.duotone) validateBlipEffect({ type: 'duotone' }, value);
  const names = effectContainers(value.effects, value);
  const original = { width: geometry?.width ?? pixels, height: geometry?.height ?? 1,
    frame: geometry?.frame, pixels };
  const active = new Set<BlipEffectContainer>();
  let count = 0, work = 0;
  const chargedDefinitions = new Set<BlipEffect>();
  // A parser lookup that exhausted its budget must remain a typed quota,
  // including when it stopped before discovering the referenced definition.
  for (const effect of value.effects) if (effect.type === 'definitions') {
    if (!Number.isSafeInteger(effect.work) || effect.work < 0) throw new OoxmlError('unsupported-image-effect', 'Invalid effect-definition work record');
    work += effect.work; chargedDefinitions.add(effect);
  }
  if (work > MAX_IMAGE_EFFECT_PIXEL_WORK) throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK, work);
  const visit = (effects: readonly BlipEffect[], input: { width: number; height: number; pixels: number; frame?: EffectBuffer['frame'] }): EffectPlan => {
    let { width, height, pixels: current, frame } = input;
    let peak = current;
    for (const effect of effects) {
      if (effect.type === 'effectCountLimit') {
        if (Number.isSafeInteger(effect.observed) && effect.observed > MAX_IMAGE_EFFECT_PASSES) {
          throw new OoxmlDecodedImageLimitError('image-effect-count', MAX_IMAGE_EFFECT_PASSES, effect.observed);
        }
        throw new OoxmlError('unsupported-image-effect', 'Invalid effect-count record');
      }
      if (++count > MAX_IMAGE_EFFECT_PASSES) throw new OoxmlDecodedImageLimitError('image-effect-count', MAX_IMAGE_EFFECT_PASSES, count);
      if (effect.type === 'definitions') {
        if (!chargedDefinitions.has(effect)) { work += effect.work; chargedDefinitions.add(effect); }
        continue;
      }
      if (effect.type === 'unsupported') throw new OoxmlError('unsupported-image-effect', `Malformed CT_Blip effect: ${effect.effect}`);
      work += current;
      if (effect.type === 'luminance') work += 256; // channel lookup table construction
      const expanded = expandBlipEffect(effect, { width, height, frame }, value);
      if (expanded) {
        const p = visit(expanded, { width, height, frame, pixels: current });
        peak = Math.max(peak, p.peak); width = p.width; height = p.height; current = p.pixels; frame = p.frame;
      } else if (effect.type === 'transform' && current > 0) {
        const p = affinePlan({ width, height, frame }, blipAffine(effect, value));
        peak = Math.max(peak, current + p.width * p.height);
        width = p.width; height = p.height; frame = p.frame; current = width * height;
        work += current * 5; // four bilinear taps plus the optional ImageData copy
      } else if (blipMaskOperations(effect, { width, height, frame }, value) && current > 0) {
        const p = visit(blipMaskOperations(effect, { width, height, frame }, value) as readonly BlipEffect[],
          { width, height, frame, pixels: current });
        peak = Math.max(peak, current + p.peak);
        // Input clone, mask postprocessing and frame sampling. Outset can
        // grow the mask, so charge its output grid as well as the parent.
        work += current * 3 + p.pixels;
        if (effect.type !== 'shadow') { width = p.width; height = p.height; frame = p.frame; current = p.pixels; }
      } else if (effect.type === 'blur' && effect.radius !== 0 && current > 0) {
        const inputPixels = current;
        const p = blurPlan(width, height, effect.radius / (value.emuPerPixelX ?? 9525),
          effect.radius / (value.emuPerPixelY ?? 9525), effect.grow);
        const f = frame ?? { x: 0, y: 0, width, height };
        frame = { ...f, x: f.x + (effect.grow ? p.rx : 0), y: f.y + (effect.grow ? p.ry : 0) };
        width = p.width; height = p.height; current = width * height;
        work += p.work;
        peak = Math.max(peak, inputPixels + current * 2 + p.kernelBytes / 4);
      } else if (effect.type === 'alphaMod' || effect.type === 'container' || effect.type === 'blend' || effect.type === 'reference') {
        const container = effect.type === 'reference' ? names.get(effect.reference) : effect.container;
        if (!container) {
          if (effect.type !== 'reference' || !['fill', 'fillLine', 'line', 'children'].includes(effect.reference)) {
            throw new OoxmlError('unsupported-image-effect', 'Invalid effect container reference');
          }
          // §20.1.8.24: fill/fillLine refer to the original picture fill,
          // including its frame, not the preceding effect's transformed grid.
          // Keep the old grid live while copying the restored original.
          if (effect.reference === 'fill' || effect.reference === 'fillLine') {
            peak = Math.max(peak, current + original.pixels);
            work += original.pixels;
            ({ width, height, frame, pixels: current } = original);
          } else peak = Math.max(peak, current * 2);
          continue;
        }
        if (active.has(container)) throw new OoxmlError('unsupported-image-effect', 'Cyclic effect container');
        active.add(container);
        if (container.kind !== 'tree' && container.kind !== 'sib') throw new OoxmlError('unsupported-image-effect', 'Invalid effect container kind');
        // Every evaluated branch clones the parent before its own operations.
        work += current * (container.kind === 'tree' ? 1 : container.effects.length);
        const results = container.kind === 'tree'
          ? [visit(container.effects, { width, height, frame, pixels: current })]
          : container.effects.map(e => visit([e], { width, height, frame, pixels: current }));
        let mask = { width, height, frame };
        for (const p of results) {
          const oldPixels = mask.width * mask.height;
          const union = container.kind === 'tree' ? p : unionEffectGrids(mask, p);
          const unionPixels = union.width * union.height;
          const reallocates = union.width !== mask.width || union.height !== mask.height
            || (union.frame?.x ?? 0) !== (mask.frame?.x ?? 0) || (union.frame?.y ?? 0) !== (mask.frame?.y ?? 0);
          peak = Math.max(peak, current + p.peak + (container.kind === 'sib' ? oldPixels : 0));
          if (container.kind === 'sib') {
            work += unionPixels;
            peak = Math.max(peak, current + oldPixels + p.pixels + (reallocates ? unionPixels : 0));
          }
          mask = { width: union.width, height: union.height, frame: union.frame };
        }
        if (effect.type !== 'alphaMod') {
          if (effect.type === 'blend') {
            const union = unionEffectGrids({ width, height, frame }, mask);
            work += union.width * union.height;
            peak = Math.max(peak, current + mask.width * mask.height + union.width * union.height);
            mask = union;
          }
          width = mask.width; height = mask.height; frame = mask.frame;
          current = geometry ? width * height : current;
        }
        active.delete(container);
      } else if (effect.type === 'fillOverlay' || effect.type === 'fill') {
        peak = Math.max(peak, current * 3 + fillPaintWorkspaceBytes(effect.fill) / 4);
        work += fillWork(effect.fill, { width, height, frame }, value, runtime);
      }
    }
    return { count, work, width, height, frame, pixels: current, peak };
  };
  const result = visit(value.effects, original);
  if (value.duotone && !value.effects.some(e => e.type === 'duotone')) { result.count++; result.work += result.pixels; }
  // One decoded source remains owned while the readback pipeline runs, and
  // applyImageDataTransform keeps the original readback ImageData pinned
  // until the transformed buffer is written back, so both stay live through
  // every effect. Original-fill references hold a further snapshot only
  // during evaluation; it is released before the final bake.
  // A chain that allocates a replacement buffer (spatial effects, unions)
  // runs while the pinned readback is still live; an in-place pointwise
  // chain mutates that readback itself.
  const replacesReadback = result.peak > pixels || result.pixels !== pixels;
  if (usesOriginalImage(value.effects)) { result.peak += pixels; result.work += pixels; }
  result.peak = Math.max(result.peak, result.pixels * 3);
  result.peak += pixels * (replacesReadback ? 2 : 1);
  return result;
}
function effectContainers(effects: readonly BlipEffect[], value: BlipPixelEffects): Map<string, BlipEffectContainer> {
  for (const n of [value.emuPerPixelX ?? 9525, value.emuPerPixelY ?? 9525]) {
    if (!Number.isFinite(n) || n <= 0) throw new OoxmlError('unsupported-image-effect', 'Invalid effect sampling');
  }
  if (!Number.isFinite(value.hostRotation ?? 0) ||
    [value.hostFlipH, value.hostFlipV].some(v => v !== undefined && typeof v !== 'boolean')) {
    throw new OoxmlError('unsupported-image-effect', 'Invalid host orientation');
  }
  const result = new Map<string, BlipEffectContainer>();
  const container = (c: BlipEffectContainer) => {
    if (c.name) {
      const previous = result.get(c.name);
      if (previous && previous !== c) throw new OoxmlError('unsupported-image-effect', 'Duplicate effect container name');
      result.set(c.name, c);
    }
  };
  for (const effect of blipEffectNodes(effects)) {
    validateBlipEffect(effect, value);
    if (effect.type === 'unsupported') throw new OoxmlError('unsupported-image-effect', `Malformed CT_Blip effect: ${effect.effect}`);
    if (effect.type === 'effectCountLimit') {
      if (Number.isSafeInteger(effect.observed) && effect.observed > MAX_IMAGE_EFFECT_PASSES) {
        throw new OoxmlDecodedImageLimitError('image-effect-count', MAX_IMAGE_EFFECT_PASSES, effect.observed);
      }
      throw new OoxmlError('unsupported-image-effect', 'Invalid effect-count record');
    }
  }
  // Nested image fills own a separate evaluator/name table. Validate and
  // count their recipes above, but do not leak their definitions into this
  // image's namespace (a shared XML definition can be projected into both).
  for (const effect of blipEffectNodes(effects, false)) {
    if (effect.type === 'definitions') for (const c of effect.containers) container(c);
    else if (effect.type === 'alphaMod' || effect.type === 'container' || effect.type === 'blend') container(effect.container);
  }
  return result;
}

/** Complete sample/draw-iteration upper bound, including recursive branches.
 * Tiled fills require their decoded runtime sources: guessing a source grid
 * here could undercharge or incorrectly reject a valid document. */
export function blipPixelEffectWork(value: BlipPixelEffects, width: number, height: number, runtime: BlipEffectRuntime = {}): number {
  return effectPlan(value, width * height, { width, height }, runtime).work;
}

function usesOriginalImage(effects: readonly BlipEffect[], seen = new Set<BlipEffectContainer>()): boolean {
  for (const effect of effects) {
    if (effect.type === 'definitions') for (const container of effect.containers) {
      if (!seen.has(container)) { seen.add(container); if (usesOriginalImage(container.effects, seen)) return true; }
    }
    if (effect.type === 'reference' && ['fill', 'fillLine'].includes(effect.reference)) return true;
    if (effect.type === 'alphaMod' || effect.type === 'container' || effect.type === 'blend') {
      if (!seen.has(effect.container)) {
        seen.add(effect.container);
        if (usesOriginalImage(effect.container.effects, seen)) return true;
      }
    }
  }
  return false;
}

/**
 * Reject malformed chains, or a transform whose step count, workspace or work exceeds the
 * shared resource policy (`MAX_IMAGE_EFFECT_PASSES`,
 * `MAX_IMAGE_EFFECT_PIXEL_WORK` in pixel-budget.ts) before any pixel is
 * touched. This source-independent preflight is completed with decoded tile
 * sources by blipPixelEffectWork before evaluation. A crossing is an `OoxmlDecodedImageLimitError`, the same catchable
 * quota failure as an over-budget decode; effects are never dropped or
 * truncated to fit. `pixels` 0 checks the pass count alone (before decoding).
 */
export function assertBlipPixelEffectsBudget(value: BlipPixelEffects, pixels: number, geometry?: EffectGrid, additionalPixels = 0): void {
  const plan = effectPlan(value, pixels, geometry);
  if (plan.count > MAX_IMAGE_EFFECT_PASSES) throw new OoxmlDecodedImageLimitError('image-effect-count', MAX_IMAGE_EFFECT_PASSES, plan.count);
  if (!Number.isSafeInteger(plan.work) || plan.work > MAX_IMAGE_EFFECT_PIXEL_WORK) {
    throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK,
      Number.isSafeInteger(plan.work) ? plan.work : Number.MAX_SAFE_INTEGER);
  }
  if (geometry && plan.peak + additionalPixels > MAX_RASTER_PIXELS) {
    throw new OoxmlDecodedImageLimitError('active-decoded-bytes', MAX_RASTER_PIXELS * 4, Math.ceil((plan.peak + additionalPixels) * 4));
  }
}

/** Apply the effects in order. Point effects mutate the input; spatial effects
 * may return a new grid, which the caller must retain. Alpha effects mutate alpha at their
 *  authored position; clrChange replaces a matched pixel's color and alpha. Throws
 *  `OoxmlDecodedImageLimitError` before touching the buffer when the plan
 *  exceeds the effect budget (see {@link assertBlipPixelEffectsBudget}). Live
 *  execution also charges each operation before it runs; a planner mismatch
 *  aborts with a typed quota before executing unadmitted work. */
export function applyBlipPixelEffects(buf: EffectBuffer, value: BlipPixelEffects, runtime: BlipEffectRuntime = {}): EffectBuffer {
  const d = buf.data;
  let additional = 0;
  for (const image of new Set(runtime.images?.values())) {
    if (image === runtime.baseImage) continue;
    const size = image as { width: number; height: number };
    additional += size.width * size.height;
  }
  assertBlipPixelEffectsBudget(value, d.length / 4, buf, additional);
  const work = effectPlan(value, buf.width * buf.height, buf, runtime).work;
  if (!Number.isSafeInteger(work) || work > MAX_IMAGE_EFFECT_PIXEL_WORK) {
    throw new OoxmlDecodedImageLimitError('image-effect-work', MAX_IMAGE_EFFECT_PIXEL_WORK,
      Number.isSafeInteger(work) ? work : Number.MAX_SAFE_INTEGER);
  }
  const local = new ImageEffectWorkLedger(Math.min(work, MAX_IMAGE_EFFECT_PIXEL_WORK));
  const context: EffectContext = {
    names: effectContainers(value.effects, value), chargedDefinitions: new Set(),
    charge: amount => { local.charge(amount); runtime.workLedger?.charge(amount); },
  };
  try {
    if (usesOriginalImage(value.effects)) context.original = clonePixels(buf, context);
    return evaluateBlipPixels(buf, value, runtime, context);
  } catch (error) {
    throw imageEffectError(error);
  }
}
function evaluateBlipPixels(buf: EffectBuffer, value: BlipPixelEffects, runtime: BlipEffectRuntime,
  currentContext: EffectContext): EffectBuffer {
  let d = buf.data;
  let duotoneApplied = false;
  for (const effect of value.effects) {
    if (effect.type === 'definitions') {
      if (!currentContext.chargedDefinitions.has(effect)) {
        currentContext.charge(effect.work); currentContext.chargedDefinitions.add(effect);
      }
      continue;
    }
    // Charge from the live buffer before touching pixels. Control operations
    // conservatively reserve a parent-grid visit as in the early plan.
    currentContext.charge(buf.width * buf.height);
    const expanded = expandBlipEffect(effect, buf, value);
    if (expanded) {
      buf = evaluateBlipPixels(buf, { ...value, effects: expanded, duotone: undefined }, runtime, currentContext);
      d = buf.data; continue;
    }
    switch (effect.type) {
      case 'effectCountLimit': break; // rejected by resource preflight
      case 'transform': {
        const grid = affinePlan(buf, blipAffine(effect, value));
        currentContext.charge(grid.width * grid.height * 5);
        buf = affinePixels(buf, blipAffine(effect, value)); d = buf.data; break;
      }
      case 'alphaOutset': {
        const mask = evaluateBlipPixels(clonePixels(buf, currentContext), { ...value, duotone: undefined,
          effects: blipMaskOperations(effect, buf, value) as readonly BlipEffect[] }, runtime, currentContext);
        currentContext.charge(buf.width * buf.height * 2 + mask.width * mask.height);
        for (let y = 0; y < mask.height; y++) for (let x = 0; x < mask.width; x++) {
          const i = (y * mask.width + x) * 4;
          const rgb = sampleEffectPixel(buf, mask, x, y);
          for (let c = 0; c < 3; c++) mask.data[i + c] = rgb[c];
        }
        buf = mask; d = buf.data; break;
      }
      case 'reflection': {
        const r = effect.reflection;
        const faded = clonePixels(buf, currentContext);
        currentContext.charge(buf.width * buf.height * 2); // fade and mask postprocessing
        // The ramp spans the previous grid, including grown/transformed
        // pixels, in physical axes (§20.1.8.50). Pull page-fixed directions
        // back through the host just like the reflection's affine transform.
        const [vx, vy] = effectDirection(r.fadeDir ?? 90, r.rotWithShape === false, value);
        const wx = faded.width * (value.emuPerPixelX ?? 9525), hy = faded.height * (value.emuPerPixelY ?? 9525);
        const min = Math.min(0, vx * wx) + Math.min(0, vy * hy), extent = Math.abs(vx * wx) + Math.abs(vy * hy);
        for (let y = 0; y < faded.height; y++) for (let x = 0; x < faded.width; x++) {
          const position = ((x + .5) / faded.width * wx * vx + (y + .5) / faded.height * hy * vy - min) / extent;
          // §20.1.8.50 gives two positions on one ramp. Coincident
          // stops form a hard transition; inventing a unit-width interval
          // would retain opacity beyond the authored end position.
          const t = r.endPos === r.stPos ? (position >= r.endPos ? 1 : 0)
            : Math.max(0, Math.min(1, (position - r.stPos) / (r.endPos - r.stPos)));
          const i = (y * faded.width + x) * 4 + 3;
          faded.data[i] = Math.round(faded.data[i] * (r.stA + (r.endA - r.stA) * t));
        }
        buf = evaluateBlipPixels(faded, { ...value, duotone: undefined,
          effects: blipMaskOperations(effect, buf, value) as readonly BlipEffect[] }, runtime, currentContext);
        currentContext.charge(buf.width * buf.height); d = buf.data; break;
      }
      case 'shadow': { // Inner shadow is parent alpha * (1 - blurred shifted parent alpha).
        const shadow = effect.shadow;
        const mask = evaluateBlipPixels(clonePixels(buf, currentContext), { ...value, duotone: undefined,
          effects: blipMaskOperations(effect, buf, value) as readonly BlipEffect[] }, runtime, currentContext);
        currentContext.charge(buf.width * buf.height * 2 + mask.width * mask.height);
        const rgb = hex6ToRgb(shadow.color) ?? [0, 0, 0];
        for (let y = 0; y < buf.height; y++) for (let x = 0; x < buf.width; x++) {
          const i = (y * buf.width + x) * 4;
          const mi = sampleEffectIndex(mask, buf, x, y);
          d[i + 3] = Math.round(d[i + 3] * shadow.alpha * (1 - (mi < 0 ? 0 : mask.data[mi + 3]) / 255));
          for (let c = 0; c < 3; c++) d[i + c] = rgb[c];
        }
        break;
      }
      case 'fill':
        currentContext.charge(fillWork(effect.fill, buf, value, runtime));
      {
        // Native ImageData exposes data/width/height as prototype accessors,
        // so an object spread would drop them; copy the fields explicitly.
        const filled = effectFillPixels(effect.fill, buf, value.emuPerPixelX ?? 9525,
          value.emuPerPixelY ?? 9525, runtime, brushCounterRotation(value));
        buf = { data: filled.data, width: filled.width, height: filled.height, frame: buf.frame };
        d = buf.data; break;
      }
      case 'glow': case 'relativeOffset': case 'softEdge': case 'presetShadow': break; // expanded above

      case 'alphaMod': {
        const mask = evaluateContainer(buf, effect.container, value, runtime, currentContext);
        // §20.1.8.5: multiply parent coverage by the evaluated container's
        // alpha. The container consumes the current parent, not a white mask.
        for (let y = 0; y < buf.height; y++) for (let x = 0; x < buf.width; x++) {
          const mi = sampleEffectIndex(mask, buf, x, y);
          const a = mi < 0 ? 0 : mask.data[mi + 3];
          const i = (y * buf.width + x) * 4 + 3;
          d[i] = Math.round(d[i] * a / 255);
        }
        break;
      }
      case 'blend':
      case 'container':
        buf = effect.type === 'blend'
          ? composeEffectBuffers(buf, evaluateContainer(buf, effect.container, value, runtime, currentContext), effect.blend, currentContext)
          : evaluateContainer(buf, effect.container, value, runtime, currentContext);
        d = buf.data; break;
      case 'reference': {
        const target = currentContext.names.get(effect.reference);
        buf = target ? evaluateContainer(buf, target, value, runtime, currentContext)
          // CT_Blip's input is the embedded image: it has fill coverage and
          // no line or logical children. Picture outlines/child shapes belong
          // to the host's separate outer shape-effect tree.
          : ['fill', 'fillLine'].includes(effect.reference) ? clonePixels(currentContext.original ?? buf, currentContext)
            : { width: buf.width, height: buf.height, frame: buf.frame, data: new Uint8ClampedArray(d.length) };
        d = buf.data; break;
      }
      case 'fillOverlay': {
        currentContext.charge(fillWork(effect.fill, buf, value, runtime));
        const overlay = effectFillPixels(effect.fill, buf, value.emuPerPixelX ?? 9525,
          value.emuPerPixelY ?? 9525, runtime, brushCounterRotation(value));
        for (let i = 0; i < d.length; i += 4) {
          blendRgba(d, i, overlay.data, i, d, i, effect.blend);
        }
        break;
      }
      case 'blur':
        if (effect.radius !== 0) currentContext.charge(blurPlan(buf.width, buf.height,
          effect.radius / (value.emuPerPixelX ?? 9525), effect.radius / (value.emuPerPixelY ?? 9525), effect.grow).work);
        buf = blurPixels(buf, effect.radius / (value.emuPerPixelX ?? 9525),
          effect.radius / (value.emuPerPixelY ?? 9525), effect.grow);
        d = buf.data;
        break;
      case 'grayscale':
        for (let i = 0; i < d.length; i += 4) {
          const gray = blipGrayLevel(d[i], d[i + 1], d[i + 2]);
          d[i] = gray;
          d[i + 1] = gray;
          d[i + 2] = gray;
        }
        break;
      case 'biLevel':
        for (let i = 0; i < d.length; i += 4) {
          // "Values greater than or equal to the threshold are set to white."
          const level = blipLuminance(d[i], d[i + 1], d[i + 2]) >= effect.thresh ? 255 : 0;
          d[i] = level;
          d[i + 1] = level;
          d[i + 2] = level;
        }
        break;
      case 'colorChange': {
        const from = hex6ToRgb(effect.from);
        const to = hex6ToRgb(effect.to);
        if (!from || !to) break;
        const fromAlpha = Math.round(effect.fromAlpha * 255);
        const toAlpha = Math.round(effect.toAlpha * 255);
        for (let i = 0; i < d.length; i += 4) {
          if (d[i] !== from[0] || d[i + 1] !== from[1] || d[i + 2] !== from[2]) continue;
          if (effect.useAlpha && d[i + 3] !== fromAlpha) continue;
          d[i] = to[0];
          d[i + 1] = to[1];
          d[i + 2] = to[2];
          d[i + 3] = toAlpha;
        }
        break;
      }
      case 'luminance': {
        currentContext.charge(256);
        const table = luminanceTable(effect.bright, effect.contrast);
        for (let i = 0; i < d.length; i += 4) {
          d[i] = table[d[i]];
          d[i + 1] = table[d[i + 1]];
          d[i + 2] = table[d[i + 2]];
        }
        break;
      }
      case 'hsl':
      case 'tint': {
        // §20.1.8.39: H'=H+hue; S'=S(1+sat); L'=L(1+lum).
        // This follows ECMA's relative adjustment, using the same HSL
        // conversion as shared DrawingML lumMod/satMod colors. MS-OI29500
        // §2.1.1299 describes Office modulation differently; no pixel-match
        // claim is made here. Clamp S/L after adjustment, wrap hue.
        // §20.1.8.60: shift towards/away from the hue by signed amt.
        // RGB'=clamp(RGB + amt*(HSL(hue,1,.5)-RGB)). This is signed
        // color interpolation (the same channel operation as tint/shade),
        // not a hue rotation or Office-calibrated transfer function.
        const target = effect.type === 'tint' ? hslToRgb(effect.hue, 1, .5) : undefined;
        const rgb: [number, number, number] = [0, 0, 0], hsl: [number, number, number] = [0, 0, 0];
        for (let i = 0; i < d.length; i += 4) {
          if (effect.type === 'hsl') {
            rgb[0] = d[i]; rgb[1] = d[i + 1]; rgb[2] = d[i + 2];
            const [h, s, l] = rgbToHsl(rgb, hsl);
            hslToRgb(h + effect.hue, Math.max(0, Math.min(1, s * (1 + effect.sat))),
              Math.max(0, Math.min(1, l * (1 + effect.lum))), rgb);
            for (let c = 0; c < 3; c++) d[i + c] = Math.round(rgb[c]);
          } else {
            for (let c = 0; c < 3; c++) d[i + c] = interpolateColorChannel(d[i + c], (target as number[])[c], effect.amt);
          }
        }
        break;
      }
      // §20.1.8.1–4, 6, 8: alpha operations are pointwise and ordered.
      // Rounding to the nearest byte matches clrChange's alpha matching grid.
      case 'alphaBiLevel':
      case 'alphaCeiling':
      case 'alphaFloor':
      case 'alphaInv':
      case 'alphaModFix':
      case 'alphaRepl': {
        const inverseColor = effect.type === 'alphaInv' && effect.color ? hex6ToRgb(effect.color.slice(0, 6)) : null;
        for (let i = 3; i < d.length; i += 4) {
          const alpha = d[i];
          switch (effect.type) {
            case 'alphaBiLevel': d[i] = alpha / 255 >= effect.thresh ? 255 : 0; break;
            case 'alphaCeiling': d[i] = alpha > 0 ? 255 : 0; break;
            case 'alphaFloor': d[i] = alpha < 255 ? 0 : 255; break;
            case 'alphaInv': {
              // §20.1.8.4 + MS-OE376 §2.1.1367: the optional constant
              // color replaces RGB; alpha is inverted independently.
              const rgb = inverseColor;
              if (rgb) { d[i - 3] = rgb[0]; d[i - 2] = rgb[1]; d[i - 1] = rgb[2]; }
              d[i] = 255 - alpha; break;
            }
            case 'alphaModFix': d[i] = Math.min(255, Math.max(0, Math.round(alpha * effect.amt))); break;
            case 'alphaRepl': d[i] = Math.round(255 * effect.alpha); break;
          }
        }
        break;
      }
      case 'colorReplace': { // §20.1.8.18: RGB only; alpha is unaffected.
        const color = hex6ToRgb(effect.color);
        if (!color) throw new OoxmlError('unsupported-image-effect', 'Invalid CT_Blip replacement color');
        for (let i = 0; i < d.length; i += 4) {
          d[i] = color[0]; d[i + 1] = color[1]; d[i + 2] = color[2];
        }
        break;
      }
      case 'unsupported': break; // rejected by preflight above
      case 'duotone':
        if (effect.clr1 && effect.clr2) duotoneImageData(buf, effect.clr1, effect.clr2);
        else if (value.duotone) duotoneImageData(buf, value.duotone.clr1, value.duotone.clr2);
        duotoneApplied = true;
        break;
    }
  }
  if (!duotoneApplied && value.duotone) {
    currentContext.charge(buf.width * buf.height);
    duotoneImageData(buf, value.duotone.clr1, value.duotone.clr2);
  }
  return buf;
}

/** Bound brush construction as well as its output-grid raster/readback.
 * One visit per auxiliary workspace byte conservatively covers pattern/gradient
 * tile painting, mirrored cells, path-shade init/solve/write and readback.
 * Effect fills use a rectangular outline, bounding their edge work too. */
/** Host rotation a rotWithShape="0" brush must undo (§20.1.8.45/.49). The
 * host applies rotation then flips; with an odd number of flips the screen
 * rotation of the mirrored frame is reversed, so the counter-rotation sign
 * flips as well (the same rule as the PPTX shape-fill painter). */
function brushCounterRotation(value: BlipPixelEffects): number {
  const rotation = value.hostRotation ?? 0;
  return (value.hostFlipH ?? false) !== (value.hostFlipV ?? false) ? -rotation : rotation;
}

function fillWork(fill: BlipEffectFill, grid: EffectGrid, value: BlipPixelEffects, runtime?: BlipEffectRuntime): number {
  const pixels = grid.width * grid.height;
  let work = pixels > 0 ? effectFillWork(fill, grid, value.emuPerPixelX ?? 9525, value.emuPerPixelY ?? 9525, runtime,
    brushCounterRotation(value))
    + fillPaintWorkspaceBytes(fill) : 0;
  if (fill.fillType === 'gradient') work += 2 * fill.stops.length * (1 + Math.max(1, Math.ceil(Math.log2(fill.stops.length))));
  return work;
}

function blipAffine(effect: Extract<BlipEffect, { type: 'transform' }>, value: BlipPixelEffects): Affine {
  const ex = value.emuPerPixelX ?? 9525, ey = value.emuPerPixelY ?? 9525;
  return [effect.sx, Math.tan(effect.ky * Math.PI / 180) * ex / ey,
    Math.tan(effect.kx * Math.PI / 180) * ey / ex, effect.sy, effect.tx / ex, effect.ty / ey];
}
function composeEffectBuffers(base: EffectBuffer, overlay: EffectBuffer, mode: BlendMode, context: EffectContext): EffectBuffer {
  const grid = unionEffectGrids(base, overlay);
  context.charge(grid.width * grid.height);
  const bf = base.frame;
  const sameGrid = grid.width === base.width && grid.height === base.height
    && grid.frame.x === (bf?.x ?? 0) && grid.frame.y === (bf?.y ?? 0);
  const data = sameGrid ? base.data : new Uint8ClampedArray(grid.width * grid.height * 4);
  const result = { ...grid, data };
  for (let y = 0; y < grid.height; y++) for (let x = 0; x < grid.width; x++) {
    blendRgba(base.data, sampleEffectIndex(base, result, x, y), overlay.data, sampleEffectIndex(overlay, result, x, y),
      data, (y * grid.width + x) * 4, mode);
  }
  return result;
}

function clonePixels(input: EffectBuffer, context: EffectContext): EffectBuffer {
  context.charge(input.width * input.height);
  // ImageData dimensions are prototype accessors in browsers, so object
  // spread would lose them and turn recursive masks into empty grids.
  return { width: input.width, height: input.height, frame: input.frame,
    data: new Uint8ClampedArray(input.data) };
}
function sampleEffectIndex(source: EffectBuffer, target: EffectBuffer, x: number, y: number): number {
  const sf = source.frame ?? { x: 0, y: 0 }, tf = target.frame ?? { x: 0, y: 0 };
  const sx = x + sf.x - tf.x, sy = y + sf.y - tf.y;
  if (sx < 0 || sy < 0 || sx >= source.width || sy >= source.height) return -1;
  return (sy * source.width + sx) * 4;
}
const TRANSPARENT_PIXEL = [0, 0, 0, 0];
function sampleEffectPixel(source: EffectBuffer, target: EffectBuffer, x: number, y: number): ArrayLike<number> {
  const i = sampleEffectIndex(source, target, x, y);
  return i < 0 ? TRANSPARENT_PIXEL : source.data.subarray(i, i + 4);
}
function evaluateContainer(parent: EffectBuffer, container: BlipEffectContainer,
  value: BlipPixelEffects, runtime: BlipEffectRuntime, context: EffectContext): EffectBuffer {
  // §20.1.10.22: a tree is sequential; siblings independently consume the
  // parent and are composited in authored order. Empty tree = parent;
  // empty sibling container produces no effect pixels.
  if (container.kind === 'tree') return evaluateBlipPixels(clonePixels(parent, context), {
    ...value, duotone: undefined, effects: container.effects,
  }, runtime, context);
  let result: EffectBuffer = { width: parent.width, height: parent.height, frame: parent.frame,
    data: new Uint8ClampedArray(parent.data.length) };
  for (const effect of container.effects) {
    const branch = evaluateBlipPixels(clonePixels(parent, context), { ...value, duotone: undefined, effects: [effect] }, runtime, context);
    result = composeEffectBuffers(result, branch, 'over', context);
  }
  return result;
}

/** Per-channel lookup table for `<a:lum>` (see the evidence above): bright
 *  and contrast are signed fractions in [-1, 1]. */
function luminanceTable(bright: number, contrast: number): Uint8ClampedArray {
  const table = new Uint8ClampedArray(256);
  const b = Math.max(-1, Math.min(1, bright));
  const c = Math.max(-1, Math.min(1, contrast));
  for (let v = 0; v < 256; v++) {
    const shifted = v / 255 + b / 2 - 0.5;
    let out: number;
    if (c >= 1) {
      out = shifted > 0 ? 1 : shifted < 0 ? 0 : 0.5 + b / 2;
    } else {
      const k = c <= 0 ? 1 + c : 1 / (1 - c);
      out = k * shifted + 0.5 + b / 2;
    }
    table[v] = Math.round(Math.max(0, Math.min(1, out)) * 255);
  }
  return table;
}

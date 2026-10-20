import type { BlipEffect, BlipEffectFill, BlipPixelEffects } from './blip-effects.js';
import { OoxmlError } from '../errors/ooxml-error.js';

function invalid(): never { throw new OoxmlError('unsupported-image-effect', 'Malformed image effect parameters'); }
const color = (v: unknown) => typeof v === 'string' && /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(v);
const rgb = (v: unknown) => typeof v === 'string' && /^[0-9a-f]{6}$/i.test(v);
const blend = (v: unknown) => ['over', 'mult', 'screen', 'darken', 'lighten'].includes(v as string);
function numbers(v: object, keys: string[], min = -Infinity, max = Infinity): void {
  for (const key of keys) {
    const n = (v as Record<string, unknown>)[key];
    if (typeof n !== 'number' || !Number.isFinite(n) || n < min || n > max) invalid();
  }
}
function shadow(v: Extract<BlipEffect, { type: 'shadow' }>['shadow']): void {
  if (!v || !rgb(v.color)) invalid();
  numbers(v, ['alpha'], 0, 1); numbers(v, ['blur', 'dist'], 0); numbers(v, ['dir']);
  for (const key of ['sx', 'sy', 'kx', 'ky']) if (key in v) numbers(v, [key]);
  if (v.algn !== undefined && !['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(v.algn)) invalid();
  if (v.rotWithShape !== undefined && typeof v.rotWithShape !== 'boolean') invalid();
}
function fill(v: BlipEffectFill): void {
  if (!v) invalid();
  switch (v.fillType) {
    case 'none': return;
    case 'solid': if (!color(v.color)) invalid(); return;
    case 'pattern': if (!color(v.fg) || !color(v.bg) || !v.preset) invalid(); return;
    case 'gradient':
      if (!Array.isArray(v.stops) || !v.stops.length || !['linear', 'radial'].includes(v.gradType)) invalid();
      numbers(v, ['angle']);
      if (v.path !== undefined && !['rect', 'circle', 'shape'].includes(v.path)) invalid();
      if (v.flip !== undefined && !['none', 'x', 'y', 'xy'].includes(v.flip)) invalid();
      if (v.scaled !== undefined && typeof v.scaled !== 'boolean') invalid();
      if (v.rotWithShape !== undefined && typeof v.rotWithShape !== 'boolean') invalid();
      for (const rect of [v.fillToRect, v.tileRect]) if (rect) {
        for (const key of ['l', 't', 'r', 'b']) if (key in rect) numbers(rect, [key]);
      }
      for (const stop of v.stops) { if (!stop || !color(stop.color)) invalid(); numbers(stop, ['position'], 0, 1); }
      return;
    case 'image':
      if (v.srcRect) numbers(v.srcRect, ['l', 't', 'r', 'b']);
      if (v.fillRect) for (const key of ['l', 't', 'r', 'b']) if (key in v.fillRect) numbers(v.fillRect, [key]);
      if (v.rotWithShape !== undefined && typeof v.rotWithShape !== 'boolean') invalid();
      if (!v.imagePath || !v.mimeType) invalid();
      if ((v.tile != null) === (v.stretch === true)) invalid();
      return;
    default: invalid();
  }
}

/** CT_Blip can also arrive from a model source, bypassing Rust. Validate
 * parameters before planning/painting, so invalid colors, modes and numbers
 * cannot become silent no-ops, NaN pixels or default source-over blending.
 * Structural cycle/depth/count checks are owned by the graph traversal. */
export function validateBlipEffect(effect: BlipEffect, value: BlipPixelEffects): void {
  if (!effect || typeof effect !== 'object') invalid();
  switch (effect.type) {
    case 'grayscale': case 'alphaCeiling': case 'alphaFloor': return;
    case 'biLevel': case 'alphaBiLevel': numbers(effect, ['thresh'], 0, 1); return;
    case 'alphaModFix': numbers(effect, ['amt'], 0); return;
    case 'alphaRepl': numbers(effect, ['alpha'], 0, 1); return;
    case 'alphaInv': if (effect.color !== undefined && !color(effect.color)) invalid(); return;
    case 'colorReplace': if (!rgb(effect.color)) invalid(); return;
    case 'colorChange':
      if (!rgb(effect.from) || !rgb(effect.to) || typeof effect.useAlpha !== 'boolean') invalid();
      numbers(effect, ['fromAlpha', 'toAlpha'], 0, 1); return;
    case 'duotone': {
      const a = effect.clr1 ?? value.duotone?.clr1, b = effect.clr2 ?? value.duotone?.clr2;
      if (!color(a) || !color(b)) invalid(); return;
    }
    case 'luminance': numbers(effect, ['bright', 'contrast'], -1, 1); return;
    case 'hsl': numbers(effect, ['hue']); numbers(effect, ['sat', 'lum'], -1, 1); return;
    case 'tint': numbers(effect, ['hue']); numbers(effect, ['amt'], -1, 1); return;
    case 'blur': numbers(effect, ['radius'], 0); if (typeof effect.grow !== 'boolean') invalid(); return;
    case 'alphaOutset': numbers(effect, ['radius']); return;
    case 'softEdge': numbers(effect, ['radius'], 0); return;
    case 'relativeOffset': numbers(effect, ['tx', 'ty']); return;
    case 'transform': numbers(effect, ['sx', 'sy', 'kx', 'ky', 'tx', 'ty']); return;
    case 'glow':
      if (!effect.glow || !rgb(effect.glow.color)) invalid();
      numbers(effect.glow, ['radius'], 0); numbers(effect.glow, ['alpha'], 0, 1); return;
    case 'shadow':
      shadow(effect.shadow);
      // §A.4.1 gives rotWithShape only to outer shadows.
      if (effect.inner && effect.shadow.rotWithShape === false) invalid();
      return;
    case 'presetShadow':
      if (!Array.isArray(effect.shadows) || effect.shadows.length < 1 || effect.shadows.length > 2) invalid();
      effect.shadows.forEach(shadow); return;
    case 'reflection':
      if (!effect.reflection) invalid();
      if (effect.reflection.rotWithShape !== undefined && typeof effect.reflection.rotWithShape !== 'boolean') invalid();
      if (effect.reflection.algn !== undefined && !['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'].includes(effect.reflection.algn)) invalid();
      numbers(effect.reflection, ['blur', 'dist'], 0);
      numbers(effect.reflection, ['stA', 'stPos', 'endA', 'endPos'], 0, 1);
      numbers(effect.reflection, ['dir', 'sx', 'sy']);
      for (const key of ['kx', 'ky', 'fadeDir']) if (key in effect.reflection) numbers(effect.reflection, [key]);
      return;
    case 'fill': fill(effect.fill); return;
    case 'fillOverlay': fill(effect.fill); if (!blend(effect.blend)) invalid(); return;
    case 'blend': if (!blend(effect.blend)) invalid(); return;
    case 'alphaMod': case 'container': return;
    case 'reference': if (typeof effect.reference !== 'string' || !effect.reference) invalid(); return;
    case 'definitions': numbers(effect, ['work'], 0); if (!Number.isSafeInteger(effect.work)) invalid(); return;
    case 'unsupported': case 'effectCountLimit': return; // typed diagnostics checked by the graph
    default: invalid();
  }
}

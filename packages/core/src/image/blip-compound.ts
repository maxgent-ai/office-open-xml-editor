import type { BlipEffect, BlipPixelEffects } from './blip-effects.js';
import type { EffectBuffer } from './image-effect-frame.js';
import type { Shadow } from '../types/common.js';

/** Lower compound EG_Effect members to the same bounded primitive pipeline.
 * All distances enter as EMUs. No host renderer or parser is consulted here. */
export function expandBlipEffect(effect: BlipEffect, input: Pick<EffectBuffer, 'width' | 'height' | 'frame'>,
  value: BlipPixelEffects): readonly BlipEffect[] | undefined {
  const frame = input.frame ?? { x: 0, y: 0, width: input.width, height: input.height };
  const ex = value.emuPerPixelX ?? 9525, ey = value.emuPerPixelY ?? 9525;
  const transform = (sx: number, sy: number, kx: number, ky: number, tx: number, ty: number): BlipEffect =>
    ({ type: 'transform', sx, sy, kx, ky, tx, ty });
  const aligned = (s: Shadow): BlipEffect => {
    const align = s.algn ?? 'b';
    const ax = align.endsWith('r') || align === 'r' ? frame.width : ['t', 'ctr', 'b'].includes(align) ? frame.width / 2 : 0;
    const ay = align.startsWith('b') ? frame.height : ['l', 'ctr', 'r'].includes(align) ? frame.height / 2 : 0;
    let a = s.sx ?? 1, b = Math.tan((s.ky ?? 0) * Math.PI / 180);
    let c = Math.tan((s.kx ?? 0) * Math.PI / 180), d = s.sy ?? 1;
    let [dx, dy] = effectDirection(s.dir, false, value).map(v => v * s.dist);
    if (s.rotWithShape === false) {
      // §§20.1.8.45/50: page-space transform of the already-oriented
      // silhouette, expressed in the unrotated bitmap: H^-1 A H. Offset
      // direction is likewise pulled back through H. Preserve the host's
      // alignment anchor; counter-rotating only dir loses scale/skew semantics.
      const [co, si] = effectDirection(value.hostRotation ?? 0, false, { effects: [] });
      const fx = value.hostFlipH ? -1 : 1, fy = value.hostFlipV ? -1 : 1;
      const aa = co * (a * co + c * si) + si * (b * co + d * si);
      const bb = fx * fy * (-si * (a * co + c * si) + co * (b * co + d * si));
      const cc = fx * fy * (co * (-a * si + c * co) + si * (-b * si + d * co));
      const dd = -si * (-a * si + c * co) + co * (-b * si + d * co);
      [a, b, c, d] = [aa, bb, cc, dd];
      [dx, dy] = [fx * (co * dx + si * dy), fy * (-si * dx + co * dy)];
    }
    return transform(a, d, Math.atan(c) * 180 / Math.PI, Math.atan(b) * 180 / Math.PI,
      dx + ax * (1 - a) * ex - c * ay * ey,
      dy + ay * (1 - d) * ey - b * ax * ex);
  };
  switch (effect.type) {
    case 'relativeOffset': // §20.1.8.51: percentages of the previous effect's grid.
      return [transform(1, 1, 0, 0, effect.tx * input.width * ex, effect.ty * input.height * ey)];
    case 'glow': // §20.1.8.32: colored, blurred parent coverage.
      return [{ type: 'colorReplace', color: effect.glow.color }, { type: 'alphaModFix', amt: effect.glow.alpha },
        { type: 'blur', radius: effect.glow.radius, grow: true }];
    case 'shadow':
      if (effect.inner) return undefined;
      // §20.1.8.45: align, scale/skew, offset; then blur the colored alpha.
      return [{ type: 'colorReplace', color: effect.shadow.color }, { type: 'alphaModFix', amt: effect.shadow.alpha },
        aligned(effect.shadow), { type: 'blur', radius: effect.shadow.blur, grow: true }];
    case 'presetShadow': // §20.1.10.52 resolves these to ordinary shadows.
      return [{ type: 'container', container: { kind: 'sib', effects:
        effect.shadows.map(shadow => ({ type: 'shadow', shadow: { ...shadow, rotWithShape: false }, inner: false })) } }];
    case 'softEdge':
      // §20.1.8.53 blurs edges while leaving the fill unaffected: radius zero
      // is exact identity, including on translucent pixels. Rasterization
      // policy: feather the nonzero coverage silhouette with our shared blur
      // kernel, then apply that mask once to the original alpha. Blurring the
      // original opacity instead would square it, darkening the entire fill.
      return effect.radius === 0 ? [] : [{ type: 'alphaMod', container: { kind: 'tree', effects: [
        { type: 'alphaCeiling' }, { type: 'blur', radius: effect.radius, grow: false },
      ] } }];
    default: return undefined;
  }
}

/** Operations evaluated on an independent mask copy. Keeping the same
 * expansion in planning and execution binds geometry/work to the actual
 * kernels, including a reflection's transformed and grown grid. */
export function blipMaskOperations(effect: BlipEffect,
  input: Pick<EffectBuffer, 'width' | 'height' | 'frame'>, value: BlipPixelEffects): readonly BlipEffect[] | undefined {
  if (effect.type === 'alphaOutset') return [
    { type: 'alphaCeiling' }, { type: 'blur', radius: Math.abs(effect.radius), grow: effect.radius > 0 },
    { type: effect.radius >= 0 ? 'alphaCeiling' : 'alphaFloor' },
  ];
  if (effect.type === 'shadow' && effect.inner) {
    // CT_InnerShadowEffect has no rotWithShape override (§A.4.1).
    const [dx, dy] = effectDirection(effect.shadow.dir, false, value);
    return [
      { type: 'blur', radius: effect.shadow.blur, grow: true },
      { type: 'transform', sx: 1, sy: 1, kx: 0, ky: 0,
        tx: effect.shadow.dist * dx, ty: effect.shadow.dist * dy },
    ];
  }
  if (effect.type === 'reflection') {
    const r = effect.reflection;
    const expanded = expandBlipEffect({ type: 'shadow', inner: false,
      shadow: { ...r, color: '000000', alpha: 1 } }, input, value) as readonly BlipEffect[];
    return expanded.slice(2);
  }
  return undefined;
}

/** Unit direction in the bitmap's local physical axes. Cardinal directions
 * stay exact, so floor/ceil grid bounds do not gain a spurious sample. */
export function effectDirection(degrees: number, pageFixed: boolean, value: BlipPixelEffects): [number, number] {
  const angle = pageFixed ? degrees - (value.hostRotation ?? 0) : degrees;
  const quarter = ((angle / 90) % 4 + 4) % 4;
  let [x, y] = Number.isInteger(quarter)
    ? ([[1, 0], [0, 1], [-1, 0], [0, -1]][quarter])
    : [Math.cos(angle * Math.PI / 180), Math.sin(angle * Math.PI / 180)];
  if (pageFixed) { if (value.hostFlipH) x = -x; if (value.hostFlipV) y = -y; }
  return [x, y];
}

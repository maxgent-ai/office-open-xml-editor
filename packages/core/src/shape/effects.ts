import type { Shadow, SoftEdge, Reflection } from '../types/common';
import { hexToRgba } from './paint';
import { createAuxCanvas, type AuxCanvas, type AuxContext } from '../canvas/aux-canvas';
import { clampCanvasSize } from '../canvas/clamp';

// createAuxCanvas moved to canvas/aux-canvas.ts so the image/ metafile players
// can share it without an image → shape dependency. Re-exported here to preserve
// the historical import path (`shape/effects` / the core barrel).
export { createAuxCanvas };

/**
 * Canvas 2D rendering of the four DrawingML edge/blur effects that the
 * Canvas shadow primitive (a single offset+blur slot) cannot express:
 *
 *   - outerShdw — ECMA-376 §20.1.8.45 (CT_OuterShadowEffect)
 *   - innerShdw  — ECMA-376 §20.1.8.40 (CT_InnerShadowEffect)
 *   - softEdge   — ECMA-376 §20.1.8.53 (CT_SoftEdgesEffect)
 *   - reflection — ECMA-376 §20.1.8.50 (CT_ReflectionEffect)
 *
 * Each helper takes a `paintShape` callback that renders the shape's opaque
 * silhouette (fill + stroke) into a supplied 2D context, positioned in the
 * SAME device pixel coordinates as the live canvas. The helpers route that
 * silhouette through one or more auxiliary canvases and `globalCompositeOperation`
 * / `filter` passes, then blit the result back onto the live context — so the
 * caller never has to special-case preset vs. custom geometry.
 *
 * Coordinates: `bbox` is the shape bounding box in device pixels (already scaled).
 * Effect radii / distances arrive in EMU and are converted with `scale`.
 */

export type PaintShape = (ctx: AuxContext) => void;

let rasterCoverageWarned = false;
function rasterCoverageUnavailable(): null {
  // Match the scene3d fallback warning: report degradation once per process,
  // without flooding the console for every affected shape or chart mark.
  if (!rasterCoverageWarned && typeof console !== 'undefined' && typeof console.warn === 'function') {
    rasterCoverageWarned = true;
    console.warn('[ooxml] effects: geometric coverage canvas unavailable or exceeds canvas limits; ' +
      'soft edge and inner shadow omitted; the original body is still painted.');
  }
  return null;
}

/** Make a dedicated opaque geometric mask from a complete painted raster.
 * Raster policy: every nonzero-alpha sample belongs to the visible geometry;
 * zero-alpha samples remain holes. Source color and opacity are not coverage.
 * This preserves walls, wire rims, contours and transparent image holes without
 * inventing a second front-face geometry. Subpixel coverage cannot be recovered
 * separately from opacity in an RGBA source, so this is binary raster support,
 * not a claim of exact vector antialiasing or Office mesh fidelity.
 * `bbox` must cover the consumer's entire source raster, including strokes,
 * joins/caps and effect padding, rather than just vector geometry. Use
 * rasterEffectCoverageBounds when the consumer is a cropped edge effect.
 * The bounded mask belongs to this paint call; allocation failure warns once
 * and lets callers omit mask-dependent effects while painting the original body.
 */
export function createRasterCoveragePaint(
  paint: PaintShape,
  bbox: EffectBBox,
): PaintShape | null {
  // bbox is the consumer's complete source domain, including its effect
  // padding and viewport clipping. Round out to cover its raster samples, but
  // do not add another guard outside that domain: those pixels cannot reach
  // the consumer, and would reject sources at the canvas area/axis limit.
  // The size guard is library resource policy, not a DrawingML effect rule.
  const x = Math.floor(bbox.x), y = Math.floor(bbox.y);
  const w = Math.max(1, Math.ceil(bbox.x + bbox.w) - x);
  const h = Math.max(1, Math.ceil(bbox.y + bbox.h) - y);
  if (clampCanvasSize(w, h).clamped) return rasterCoverageUnavailable();
  let canvas: AuxCanvas | null;
  try { canvas = createAuxCanvas(w, h); } catch { return rasterCoverageUnavailable(); }
  const ctx = canvas ? get2d(canvas) : null;
  if (!canvas || !ctx) return rasterCoverageUnavailable();
  paint(offsetPaintCtx(ctx, { x, y, w, h }));
  const image = ctx.getImageData(0, 0, w, h);
  for (let i = 0; i < image.data.length; i += 4) {
    image.data[i] = image.data[i + 1] = image.data[i + 2] = 0;
    image.data[i + 3] = image.data[i + 3] > 0 ? 255 : 0;
  }
  ctx.putImageData(image, 0, 0);
  return target => {
    target.save();
    target.setTransform({ a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } as DOMMatrix);
    target.drawImage(canvas as CanvasImageSource, x, y);
    target.restore();
  };
}

/** Bounding box of the shape in device pixels (post-scale). */
export interface EffectBBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

function shadowAlignmentPoint(bbox: EffectBBox, algn: Shadow['algn']): [number, number] {
  const x = algn === 'tl' || algn === 'l' || algn === 'bl'
    ? bbox.x
    : algn === 'tr' || algn === 'r' || algn === 'br'
      ? bbox.x + bbox.w
      : bbox.x + bbox.w / 2;
  const y = algn === 'tl' || algn === 't' || algn === 'tr'
    ? bbox.y
    : algn === 'l' || algn === 'ctr' || algn === 'r'
      ? bbox.y + bbox.h / 2
      : bbox.y + bbox.h;
  return [x, y];
}

/**
 * EMU → device px. Mirrors `emuToPx` in the pptx renderer, where `scale` is
 * px-per-EMU (it already folds in the EMU→px conversion), so this is a bare
 * multiply. Kept local so the effect helpers have no dependency on the renderer.
 */
function emuToPx(emu: number, scale: number): number {
  return emu * scale;
}

function get2d(canvas: AuxCanvas): AuxContext | null {
  // The union return of getContext('2d') across HTMLCanvasElement /
  // OffscreenCanvas is awkward to type; narrow via unknown.
  return (canvas.getContext('2d') as AuxContext | null) ?? null;
}

// ── bbox-sized auxiliary canvases (perf: A4) ────────────────────────────────
//
// Historically each effect allocated an aux canvas the size of the WHOLE live
// canvas (deviceW × deviceH) and painted the shape into it through the live
// transform (so the silhouette lands at its absolute device-pixel position).
// For a small shape on a large slide that wastes almost the entire allocation and
// makes every fill/blur/composite touch full-canvas pixel counts.
//
// Instead we crop to the shape's device-space bbox grown by the effect's own
// spatial reach (blur radius / feather / offset) plus a 2 px safety pad, clamped
// to the canvas. The shape is painted OFFSET so device (crop.x, crop.y) maps to
// aux-local (0,0); the final blit puts it back at (crop.x, crop.y). Because CSS
// `blur()` treats everything outside the canvas as transparent (alpha 0) — the
// same as the region beyond the old full-canvas silhouette — a margin ≥ the blur
// kernel extent makes the cropped result IDENTICAL, pixel for pixel, to the
// full-canvas version everywhere the shape+effect actually paints.
//
// SCOPE: this applies to innerShadow and softEdge, whose final blit is an
// INTEGER-offset, identity-transform drawImage — a pure pixel copy, byte-exact
// for any source canvas size. applyReflection is deliberately EXEMPT: its final
// blit resamples the aux under a fractional mirror transform, and skia's texture
// sampling (edge behaviour + fixed-point phase) depends on the source's
// size/offset — a cropped source produced platform-dependent 1–7/255 diffs on
// Linux CI (PR #672). See the note inside applyReflection.

/** A sub-rectangle of the device canvas that an effect is confined to. */
interface EffectCrop {
  /** Device-px origin of the crop (top-left) on the live canvas. */
  x: number;
  y: number;
  /** Crop dimensions in device px (already clamped to the canvas). */
  w: number;
  h: number;
}

/**
 * Crop rectangle = the shape's device bbox grown by `margin` px on every side,
 * clamped to `[0, deviceW] × [0, deviceH]`. `margin` must cover the effect's
 * spatial reach (blur kernel extent + any offset), so no painted/blurred pixel
 * the full-canvas version would have produced falls outside the crop.
 */
function computeCrop(
  bbox: EffectBBox,
  margin: number,
  deviceW: number,
  deviceH: number,
): EffectCrop {
  const x0 = Math.max(0, Math.floor(bbox.x - margin));
  const y0 = Math.max(0, Math.floor(bbox.y - margin));
  const x1 = Math.min(deviceW, Math.ceil(bbox.x + bbox.w + margin));
  const y1 = Math.min(deviceH, Math.ceil(bbox.y + bbox.h + margin));
  return { x: x0, y: y0, w: Math.max(1, x1 - x0), h: Math.max(1, y1 - y0) };
}

function softEdgeMargin(softEdge: SoftEdge, scale: number): number {
  return Math.ceil(emuToPx(softEdge.radius, scale)) + 2;
}

function innerShadowMargin(shadow: Shadow, scale: number): number {
  return Math.ceil(3 * emuToPx(shadow.blur, scale) + Math.abs(emuToPx(shadow.dist, scale))) + 2;
}

/**
 * Union of the source domains read by the mask-dependent edge effects.
 * ECMA-376 §§20.1.8.40/.53 apply to the painted object, including its outline.
 * A geometry bbox (even expanded by half the line width) need not contain a
 * miter join, square cap, arrowhead or contour. Capture all raster support in
 * the consumers' source domains instead of imposing a second, tighter clip.
 * This also handles callbacks whose actual line style is resolved while painting.
 *
 * These are the existing compositor crops, not newly inferred vector bounds:
 * pixels outside them were already unavailable to the effects. Keep that crop
 * policy and effect alignment/stretch unchanged; the coverage conversion must
 * preserve every previously available source pixel. Projected callers retain
 * their prepared body/cache bounds, including walls, back rims and padding.
 */
export function rasterEffectCoverageBounds(
  bbox: EffectBBox,
  effects: { softEdge?: SoftEdge | null; innerShadow?: Shadow | null },
  scale: number,
  deviceW: number,
  deviceH: number,
): EffectBBox {
  const margin = Math.max(
    effects.softEdge ? softEdgeMargin(effects.softEdge, scale) : 0,
    effects.innerShadow ? innerShadowMargin(effects.innerShadow, scale) : 0,
  );
  return computeCrop(bbox, margin, deviceW, deviceH);
}

/** The subset of a 2D affine matrix `setTransform(m)` reads (DOMMatrix-shaped). */
interface Matrix2D {
  a: number; b: number; c: number; d: number; e: number; f: number;
}

/**
 * Wrap an aux context so that a `setTransform(m)` from the (opaque) paint callback
 * becomes `setTransform(translate(-crop.x, -crop.y) · m)`. The paint callback sets
 * the live (absolute-device) transform to position the silhouette; prepending the
 * crop offset shifts that absolute placement into the cropped canvas's local space
 * without the callback knowing. Every other property/method forwards unchanged.
 *
 * Composition: the CTM `setTransform(a,b,c,d,e,f)` maps (x,y) → (a·x+c·y+e,
 * b·x+d·y+f). Applying `T(-cx,-cy)` AFTER it just subtracts the crop origin from
 * the translation components, so the composed matrix is `(a,b,c,d,e-cx,f-cy)`.
 * Doing the arithmetic by hand avoids depending on a global `DOMMatrix`, which is
 * absent in some runtimes (plain Node) where core effect code can run.
 *
 * The paint callbacks only ever set the transform via the single-matrix
 * `setTransform(m)` form (they never call `getTransform`, `transform`,
 * `resetTransform`, or the 6-number `setTransform` overload), so this one trap is
 * sufficient; a passthrough covers fills, paths, blurs and save/restore.
 */
function offsetPaintCtx(real: AuxContext, crop: EffectCrop): AuxContext {
  if (crop.x === 0 && crop.y === 0) return real; // no shift needed
  const cx = crop.x;
  const cy = crop.y;
  return new Proxy(real as unknown as Record<string | symbol, unknown>, {
    get(target, prop) {
      if (prop === 'setTransform') {
        return (m: Matrix2D) => {
          (target as unknown as AuxContext).setTransform(
            m.a, m.b, m.c, m.d, m.e - cx, m.f - cy,
          );
        };
      }
      const v = Reflect.get(target, prop);
      return typeof v === 'function' ? v.bind(target) : v;
    },
    set(target, prop, value) {
      // Forward the write to the TARGET (not the Proxy receiver) so a setter
      // runs with `this` bound to the real context — required both for host
      // canvas objects and for class instances with private fields (a mock, or
      // a polyfilled context) that a receiver-bound setter cannot write.
      (target as Record<string | symbol, unknown>)[prop] = value;
      return true;
    },
  }) as unknown as AuxContext;
}

/**
 * outerShdw — ECMA-376 §20.1.8.45. Build the shadow from the alpha of the
 * COMPOSED shape (fill + stroke) rather than putting Canvas's shadow state on
 * each individual fill/stroke call. The latter is observably wrong for a
 * stroked shape: clearing after the fill lets the outline cover the shadow's
 * dense edge, while leaving it enabled makes every path cast another shadow.
 *
 * The cropped aux keeps the cost proportional to the affected shape. Its
 * pixels are recoloured through `source-in`, then the one resulting silhouette
 * is blurred and offset as a single draw operation behind the live shape.
 * Returns false when the host cannot allocate an auxiliary surface so callers
 * may retain a native Canvas-shadow fallback.
 */
export function applyOuterShadow(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  bbox: EffectBBox,
  shadow: Shadow,
  scale: number,
  deviceW: number,
  deviceH: number,
  shapeRotationDegrees = 0,
  alignmentAnchor?: readonly [number, number],
): boolean {
  const blur = Math.max(0, emuToPx(shadow.blur, scale));
  const dist = emuToPx(shadow.dist, scale);
  // rotWithShape defaults true (§20.1.8.45). A false value keeps the effect in
  // page space while the source silhouette itself remains the rotated object.
  const followRotation = shadow.rotWithShape !== false;
  const effectRotation = followRotation ? shapeRotationDegrees : 0;
  const dirRad = ((shadow.dir + effectRotation) * Math.PI) / 180;
  const dx = Math.cos(dirRad) * dist;
  const dy = Math.sin(dirRad) * dist;
  // CSS blur has an unbounded Gaussian tail. Three radii cover the visible
  // contribution while the extra offset/safety band prevents crop-edge cuts.
  const margin = Math.ceil(blur * 3 + Math.max(Math.abs(dx), Math.abs(dy))) + 2;
  const crop = computeCrop(bbox, margin, deviceW, deviceH);
  const aux = createAuxCanvas(crop.w, crop.h);
  if (!aux) return false;
  const c = get2d(aux);
  if (!c) return false;

  paintShape(offsetPaintCtx(c, crop));
  c.save();
  // The paint callback installs the live absolute-device transform (shifted by
  // the crop origin). Recolouring is a crop-local full-surface operation, so it
  // must not inherit that transform. Otherwise the source-in rectangle moves
  // away from silhouettes positioned farther across/down the slide, producing
  // position-dependent partial or entirely missing shadows.
  c.setTransform(1, 0, 0, 1, 0, 0);
  c.globalCompositeOperation = 'source-in';
  c.fillStyle = hexToRgba(shadow.color, shadow.alpha);
  c.fillRect(0, 0, crop.w, crop.h);
  c.restore();

  liveCtx.save();
  if (blur > 0) liveCtx.filter = `blur(${blur}px)`;
  // ECMA-376 defines the order as alignment origin, scale/skew, then offset.
  // The source aux is already in absolute device coordinates; compose the
  // page-space offset separately so it is not itself scaled or skewed.
  // Placement owns the authored shape frame.  In particular, the corner of a
  // rotated or perspective-projected shape is not the corresponding corner of
  // its post-transform AABB.  Callers that know that frame pass the exact
  // device-space anchor; bbox-based placement remains the compatibility
  // fallback for headless/custom callers that only have a silhouette box.
  const [anchorX, anchorY] = alignmentAnchor ??
    shadowAlignmentPoint(bbox, shadow.algn ?? 'b');
  const sx = shadow.sx ?? 1;
  const sy = shadow.sy ?? 1;
  const kx = Math.tan(((shadow.kx ?? 0) * Math.PI) / 180);
  const ky = Math.tan(((shadow.ky ?? 0) * Math.PI) / 180);
  liveCtx.translate(dx, dy);
  liveCtx.translate(anchorX, anchorY);
  if (effectRotation !== 0) liveCtx.rotate((effectRotation * Math.PI) / 180);
  liveCtx.transform(sx, ky, kx, sy, 0, 0);
  if (effectRotation !== 0) liveCtx.rotate((-effectRotation * Math.PI) / 180);
  liveCtx.translate(-anchorX, -anchorY);
  liveCtx.drawImage(aux as CanvasImageSource, crop.x, crop.y);
  liveCtx.restore();
  return true;
}

/**
 * innerShdw — ECMA-376 §20.1.8.40. "A shadow is applied within the edges of the
 * object." The shadow colour shows through where the shape's interior is NOT
 * covered by an offset copy of the shape — i.e. it hugs the inside of the edge
 * opposite the offset direction.
 *
 * Construction (on cropped device-space surfaces with effect padding):
 *   1. Capture the silhouette, then recolour it with the shadow colour.
 *   2. `destination-out` an offset + blurred copy of the silhouette. What
 *      remains is a blurred crescent inside the original edge — the inner shadow.
 *   3. `destination-in` the (un-offset, un-blurred) silhouette so the shadow is
 *      clipped to the shape interior.
 *   4. Blit over the live shape with `source-over` (drawn AFTER the fill).
 *
 * Offset uses dir (degrees clockwise from East) and dist (EMU). The shadow is
 * cast toward `dir`, so the colour collects on the far side — we punch the
 * offset copy toward `dir` and the surviving colour sits opposite it, matching
 * PowerPoint's inset shadow.
 */
export function applyInnerShadow(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  bbox: EffectBBox,
  shadow: Shadow,
  scale: number,
  deviceW: number,
  deviceH: number,
): void {
  const blur = emuToPx(shadow.blur, scale);
  const dist = emuToPx(shadow.dist, scale);
  const dirRad = (shadow.dir * Math.PI) / 180;
  const dx = Math.cos(dirRad) * dist;
  const dy = Math.sin(dirRad) * dist;

  // Margin: the destination-out pass blurs an offset copy of the silhouette. The
  // offset reaches |dist| beyond the bbox and the CSS blur kernel a further ~3σ =
  // 3·blur px (CSS `blur(r)` takes r as the Gaussian std-dev), so the blurred
  // offset silhouette near the bbox edge (which carves the shadow just INSIDE the
  // edge) is fully represented within bbox + (|dist| + 3·blur). +2 px guards the
  // fractional-pixel silhouette edge. The step-3 destination-in additionally clips
  // the result to the interior (⊆ bbox), so nothing outside the crop survives.
  const margin = innerShadowMargin(shadow, scale);
  const crop = computeCrop(bbox, margin, deviceW, deviceH);
  const aux = createAuxCanvas(crop.w, crop.h);
  if (!aux) return;
  const cReal = get2d(aux);
  if (!cReal) return;
  const maskAux = createAuxCanvas(crop.w, crop.h);
  const mc = maskAux ? get2d(maskAux) : null;
  if (!maskAux || !mc) return;
  // Paint geometry once. Callbacks own their fillStyle and absolute transform;
  // setting shadow color or translating before replaying them would be undone.
  // Apply color and offset/blur to the captured raster at identity instead.
  paintShape(offsetPaintCtx(mc, crop));

  // CT_InnerShadowEffect owns its color (ECMA-376 §20.1.8.40), regardless
  // of whether the mask callback fills a path or blits a geometric raster.
  cReal.drawImage(maskAux as CanvasImageSource, 0, 0);
  cReal.globalCompositeOperation = 'source-in';
  cReal.fillStyle = hexToRgba(shadow.color, shadow.alpha);
  cReal.fillRect(0, 0, crop.w, crop.h);

  // Carve with an offset/blurred mask, then clip to the unshifted interior.
  cReal.globalCompositeOperation = 'destination-out';
  cReal.filter = blur > 0 ? `blur(${blur}px)` : 'none';
  cReal.drawImage(maskAux as CanvasImageSource, dx, dy);
  cReal.globalCompositeOperation = 'destination-in';
  cReal.filter = 'none';
  cReal.drawImage(maskAux as CanvasImageSource, 0, 0);

  // 4. Composite over the live shape at the crop origin.
  liveCtx.save();
  liveCtx.drawImage(aux as CanvasImageSource, crop.x, crop.y);
  liveCtx.restore();
}

/**
 * softEdge — ECMA-376 §20.1.8.53. "The edges of the shape are blurred, while the
 * fill is not affected." PowerPoint feathers the shape's alpha SYMMETRICALLY
 * about the geometry by `rad` EMU: edge colours bleed OUTWARD past the rect and
 * fade inward, so the perimeter dissolves smoothly into the background.
 *
 * Masking a hard-clipped image with a feathered silhouette would feather only
 * inward and leave a hard outer step (the image has no pixels outside the rect).
 * Instead this helper extends the colour layer (with its source opacity) past
 * the geometry via an edge-clamp stretch, then clips it to the blurred
 * silhouette — yielding a clean, symmetric Gaussian falloff. The blur std-dev is
 * `rad/3` (the soft-edge `rad` spans ~3σ), matching PowerPoint's rasterised
 * falloff.
 *
 * `paintMask` paints opaque geometric coverage, including any painted edges
 * and holes, independently of source opacity. When omitted, `paintShape` is
 * reused (correct only for an opaque body).
 */
export function applySoftEdge(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  bbox: EffectBBox,
  softEdge: SoftEdge,
  scale: number,
  deviceW: number,
  deviceH: number,
  paintMask?: PaintShape,
): void {
  const radius = emuToPx(softEdge.radius, scale);
  if (radius <= 0) {
    // No feather: paint straight onto the live context.
    paintShape(liveCtx);
    return;
  }

  // Margin: the opaque colour layer is the bbox stretched OUTWARD by `radius`
  // (edge-clamp bleed), so no soft-edge pixel exists past bbox + radius; the
  // radius/3-σ mask blur (≤3σ = radius extent) can only carve within that band.
  // +2 px covers the fractional silhouette edge. So `radius` is the exact reach.
  const margin = softEdgeMargin(softEdge, scale);
  const crop = computeCrop(bbox, margin, deviceW, deviceH);
  // bbox expressed in the cropped canvas's local coordinates.
  const bx = bbox.x - crop.x;
  const by = bbox.y - crop.y;

  const aux = createAuxCanvas(crop.w, crop.h);
  if (!aux) {
    paintShape(liveCtx);
    return;
  }
  const cReal = get2d(aux);
  if (!cReal) {
    paintShape(liveCtx);
    return;
  }
  const c = offsetPaintCtx(cReal, crop);

  const mask = paintMask ?? paintShape;

  // 1. Sharp shape (fill + stroke) in device-pixel space (cropped).
  paintShape(c);

  // PowerPoint's soft edge (ECMA-376 §20.1.8.31) feathers the alpha
  // SYMMETRICALLY about the geometry: edge colours bleed OUTWARD past the rect
  // and fade inward, so the perimeter dissolves smoothly into the background.
  // Masking a hard-clipped image with a feathered silhouette feathers only
  // inward and leaves a hard outer step (the image has no pixels outside the
  // rect). Compose three cropped layers at identity to get the outward bleed:
  const maskAux = createAuxCanvas(crop.w, crop.h);
  const composeAux = createAuxCanvas(crop.w, crop.h);
  const mcReal = maskAux ? get2d(maskAux) : null;
  const cc = composeAux ? get2d(composeAux) : null;
  if (maskAux && mcReal && composeAux && cc) {
    const mc = offsetPaintCtx(mcReal, crop);
    // a. Solid silhouette of the shape (device space; mask applies the live
    //    transform itself). Blurred in step (c) into the alpha ramp.
    mc.fillStyle = '#000';
    mask(mc);
    // b. Extend the colour layer past the geometry: draw the
    //    image's bbox stretched outward by `radius` (edge-clamp → border colours
    //    bleed out), then the sharp image on top to keep the interior crisp.
    //    Source/dest rects are in the cropped canvas's local space (bx,by).
    //    Step (c) multiplies this source opacity by the blurred geometric
    //    coverage, giving the existing symmetric falloff without treating
    //    material opacity as a second geometric mask.
    cc.drawImage(
      aux as CanvasImageSource,
      bx, by, bbox.w, bbox.h,
      bx - radius, by - radius, bbox.w + radius * 2, bbox.h + radius * 2,
    );
    // Replace the stretched layer beneath the actual geometry before drawing
    // the sharp source. Source-over twice would raise a translucent interior's
    // opacity, contrary to §20.1.8.53 (the fill is not affected).
    cc.globalCompositeOperation = 'destination-out';
    cc.drawImage(maskAux as CanvasImageSource, 0, 0);
    cc.globalCompositeOperation = 'source-over';
    cc.drawImage(aux as CanvasImageSource, 0, 0);
    // c. Attenuate source alpha by the blurred coverage → symmetric feather. The
    //    soft-edge `rad` is a blur radius spanning ~3σ (the kernel extent), so
    //    the Gaussian std-dev — which is what CSS blur() takes — is rad/3. This
    //    matches PowerPoint's rasterised falloff (measured σ ≈ rad/3).
    cc.globalCompositeOperation = 'destination-in';
    cc.filter = `blur(${radius / 3}px)`;
    cc.drawImage(maskAux as CanvasImageSource, 0, 0);
    cc.filter = 'none';
    cc.globalCompositeOperation = 'source-over';
    liveCtx.save();
    liveCtx.drawImage(composeAux as CanvasImageSource, crop.x, crop.y);
    liveCtx.restore();
    return;
  }

  // Fallback (e.g. no OffscreenCanvas): paint the shape un-feathered.
  liveCtx.save();
  liveCtx.drawImage(aux as CanvasImageSource, 0, 0);
  liveCtx.restore();
}

/**
 * reflection — ECMA-376 §20.1.8.50. A mirrored copy of the shape rendered below
 * it, faded out by a linear alpha gradient.
 *
 * Steps:
 *   1. Paint the shape onto an aux canvas (its own silhouette + fill/stroke).
 *   2. Apply a vertical alpha gradient via `destination-in`: alpha runs from
 *      `stA` at `stPos` to `endA` at `endPos` along the (post-mirror top→bottom)
 *      ramp. With the default fadeDir=5400000 (90°, downward) the ramp is the
 *      reflection's own vertical axis.
 *   3. Blit the faded copy onto the live context under a transform that mirrors
 *      it (sx/sy), offsets it by `dist` along `dir`, and anchors it to the
 *      shape's bottom edge (algn="b" default).
 *
 * Skew (kx/ky), non-bottom alignment, fadeDir != 90°, and rotWithShape are not
 * carried by the parser model; their spec defaults are assumed (kx=ky=0,
 * algn="b", fadeDir=90°, rotWithShape=true → reflection shares the shape's
 * rotation, which the caller's transform already provides).
 */
export function applyReflection(
  liveCtx: AuxContext,
  paintShape: PaintShape,
  bbox: EffectBBox,
  reflection: Reflection,
  scale: number,
  deviceW: number,
  deviceH: number,
): void {
  // NOT cropped — the A4 bbox-sizing (see innerShadow/softEdge above) is
  // deliberately not applied here. Those two effects blit their aux back with an
  // INTEGER-offset, identity-transform drawImage — a pure pixel copy that is
  // byte-exact for any source canvas size/offset. The reflection's final blit is
  // different in kind: the mirror transform carries a FRACTIONAL translation
  // (`dist` in device px, and `bottom` itself is fractional in general), so
  // drawImage bilinear-RESAMPLES the aux as a texture. Skia's sampling near a
  // texture edge and its fixed-point sample phase depend on the source image's
  // size and offset — with a cropped aux, the crop's bottom edge sits right on
  // the stroke's bbox-overflow (the bbox excludes the stroke half-width), and
  // Linux CI produced 1–7/255 diffs on the flipped edge row that macOS did not
  // (PR #672). Feeding the mirror blit the SAME full-canvas source is the only
  // way byte-exactness holds across skia builds, so the full-size aux stays.
  const aux = createAuxCanvas(deviceW, deviceH);
  if (!aux) return;
  const c = get2d(aux);
  if (!c) return;

  const blur = emuToPx(reflection.blur, scale);

  // 1. Paint the shape onto the aux canvas, optionally blurred.
  c.save();
  if (blur > 0) c.filter = `blur(${blur}px)`;
  paintShape(c);
  c.restore();

  // 2. Fade with a vertical alpha gradient over the shape's bbox band.
  //    ECMA-376 §20.1.8.50: stA/stPos is the alpha at the START of the
  //    reflection (the edge touching the shape) and endA/endPos the END (the
  //    far edge). The reflection is mirrored about the bottom edge (step 3), so
  //    its START maps to the shape's BOTTOM row in this un-mirrored aux and its
  //    END maps upward toward the top. Build the ramp from the bottom up so the
  //    opaque stA band sits at `bottom` (→ reflection top after the flip) and
  //    fades to endA toward `top` — otherwise the visible band lands far below
  //    the shape (off-canvas for tall pictures) and the reflection disappears.
  c.save();
  c.globalCompositeOperation = 'destination-in';
  const top = bbox.y;
  const bottom = bbox.y + bbox.h;
  const grad = c.createLinearGradient(0, bottom, 0, top);
  const stPos = clamp01(reflection.stPos);
  const endPos = clamp01(reflection.endPos);
  // Offsets run 0→1 from `bottom` to `top` (the createLinearGradient axis).
  // stA holds from the bottom edge up to stPos; endA holds from endPos onward.
  grad.addColorStop(0, `rgba(0,0,0,${reflection.stA})`);
  if (stPos > 0) grad.addColorStop(stPos, `rgba(0,0,0,${reflection.stA})`);
  if (endPos < 1 && endPos > stPos) grad.addColorStop(endPos, `rgba(0,0,0,${reflection.endA})`);
  grad.addColorStop(1, `rgba(0,0,0,${reflection.endA})`);
  c.fillStyle = grad;
  c.fillRect(0, 0, deviceW, deviceH);
  c.restore();

  // 3. Blit the faded mirror under the shape. Mirror about the shape's bottom
  //    edge (algn="b"): reflect across y = bottom, then push down by `dist`.
  const dist = emuToPx(reflection.dist, scale);
  const dirRad = (reflection.dir * Math.PI) / 180;
  const offX = Math.cos(dirRad) * dist;
  const offY = Math.sin(dirRad) * dist;

  liveCtx.save();
  // Translate to the bottom edge, apply sx/sy scale (sy<0 = mirror), then
  // translate back. Anchoring at the bottom edge keeps the reflection's top
  // touching the shape's bottom before `dist` separation.
  liveCtx.translate(bbox.x + offX, bottom + offY);
  liveCtx.scale(reflection.sx, reflection.sy);
  liveCtx.translate(-(bbox.x), -bottom);
  liveCtx.drawImage(aux as CanvasImageSource, 0, 0);
  liveCtx.restore();
}

function clamp01(v: number): number {
  return v < 0 ? 0 : v > 1 ? 1 : v;
}

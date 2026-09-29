import type { Fill, GradientFill, PatternFill, Stroke } from '../types/common';
import { buildPatternBitmap } from './pattern-bitmaps';
import { drawingmlLineDashArray, shapeStrokeDashArray } from '../draw/dash';
import { createAuxCanvasForContext } from '../canvas/aux-canvas';

const MAX_GRADIENT_TILE_EDGE = 512;

function tiledGradient(
  fill: GradientFill,
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  shapeRotationDeg: number,
): CanvasPattern | null {
  const tile = fill.tileRect;
  if (!tile) return null;
  // An explicitly authored all-zero CT_RelativeRect is the complete fill
  // region. It has exactly one tile, so allocating two auxiliary canvases and
  // a repeating pattern would change no pixels and only add per-shape work.
  if ((tile.l ?? 0) === 0 && (tile.t ?? 0) === 0
      && (tile.r ?? 0) === 0 && (tile.b ?? 0) === 0) return null;
  const tileX = x + w * (tile.l ?? 0);
  const tileY = y + h * (tile.t ?? 0);
  const tileW = w * (1 - (tile.l ?? 0) - (tile.r ?? 0));
  const tileH = h * (1 - (tile.t ?? 0) - (tile.b ?? 0));
  if (!Number.isFinite(tileW) || !Number.isFinite(tileH)
      || Math.abs(tileW) < 1e-9 || Math.abs(tileH) < 1e-9) return null;

  const scale = Math.min(
    1,
    MAX_GRADIENT_TILE_EDGE / Math.abs(tileW),
    MAX_GRADIENT_TILE_EDGE / Math.abs(tileH),
  );
  const baseW = Math.max(1, Math.ceil(Math.abs(tileW) * scale));
  const baseH = Math.max(1, Math.ceil(Math.abs(tileH) * scale));
  const base = createAuxCanvasForContext(ctx, baseW, baseH);
  const baseCtx = base?.getContext('2d');
  if (!base || !baseCtx) return null;
  const basePaint = resolveFill(
    { ...fill, tileRect: undefined, flip: undefined },
    baseCtx as CanvasRenderingContext2D,
    0,
    0,
    baseW,
    baseH,
    shapeRotationDeg,
  );
  if (!basePaint) return null;
  baseCtx.fillStyle = basePaint;
  baseCtx.fillRect(0, 0, baseW, baseH);

  const flipX = fill.flip === 'x' || fill.flip === 'xy';
  const flipY = fill.flip === 'y' || fill.flip === 'xy';
  let patternSource = base;
  if (flipX || flipY) {
    const repeatW = baseW * (flipX ? 2 : 1);
    const repeatH = baseH * (flipY ? 2 : 1);
    const repeat = createAuxCanvasForContext(ctx, repeatW, repeatH);
    const repeatCtx = repeat?.getContext('2d');
    if (!repeat || !repeatCtx) return null;
    for (let row = 0; row < (flipY ? 2 : 1); row += 1) {
      for (let col = 0; col < (flipX ? 2 : 1); col += 1) {
        repeatCtx.save();
        repeatCtx.translate(col * baseW, row * baseH);
        repeatCtx.scale(col === 1 ? -1 : 1, row === 1 ? -1 : 1);
        repeatCtx.drawImage(base, col === 1 ? -baseW : 0, row === 1 ? -baseH : 0);
        repeatCtx.restore();
      }
    }
    patternSource = repeat;
  }
  const pattern = ctx.createPattern(patternSource, 'repeat');
  if (!pattern || typeof pattern.setTransform !== 'function') return null;
  pattern.setTransform({
    a: tileW / baseW,
    b: 0,
    c: 0,
    d: tileH / baseH,
    e: tileX,
    f: tileY,
  });
  return pattern;
}

/**
 * Convert a 6- or 8-char hex colour to a CSS `rgba()` string.
 * 8-char hex encodes alpha in the last two chars (RRGGBBAA).
 * `alpha` applies to 6-char hex; ignored for 8-char.
 * A leading `#` is tolerated (`#RRGGBB` and `RRGGBB` both work).
 */
export function hexToRgba(hex: string, alpha = 1): string {
  const h = hex.charCodeAt(0) === 35 /* '#' */ ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  const a = h.length >= 8 ? parseInt(h.slice(6, 8), 16) / 255 : alpha;
  return `rgba(${r},${g},${b},${a})`;
}

function colorCanProduceVisiblePixels(color: string): boolean {
  const normalized = color.startsWith('#') ? color.slice(1) : color;
  if (normalized.length < 8) return true;
  const alpha = Number.parseInt(normalized.slice(6, 8), 16);
  // Parsed OOXML colours are well formed. Treat malformed hand-authored
  // public-model input conservatively as visible rather than silently
  // suppressing paint or changing a chart-style role.
  return !Number.isFinite(alpha) || alpha !== 0;
}

/**
 * Static visibility bound for one DrawingML fill recipe. `true` means the
 * recipe may produce visible pixels; image contents themselves remain opaque
 * to this layer. A finite non-positive `a:alphaModFix` is nevertheless known
 * to make every image pixel transparent, so it must not trigger image decode,
 * resource charging, or callout-style selection.
 */
export function fillCanProduceVisiblePixels(fill: Fill | null | undefined): boolean {
  if (!fill || fill.fillType === 'none') return false;
  if (fill.fillType === 'solid') return colorCanProduceVisiblePixels(fill.color);
  if (fill.fillType === 'gradient') {
    return fill.stops.some(stop => colorCanProduceVisiblePixels(stop.color));
  }
  if (fill.fillType === 'pattern') {
    return colorCanProduceVisiblePixels(fill.fg) || colorCanProduceVisiblePixels(fill.bg);
  }
  return fill.alpha == null || !Number.isFinite(fill.alpha) || fill.alpha > 0;
}

/**
 * Rec.601 perceptual luma (`0.299·R + 0.587·G + 0.114·B`) of a colour, on the
 * 0–255 scale. Accepts a 6- or 8-char hex; a leading `#` is tolerated and the
 * alpha byte (if present) is ignored, matching {@link hexToRgba}'s hex
 * normalisation.
 */
export function relativeLuma(hex: string): number {
  const h = hex.charCodeAt(0) === 35 /* '#' */ ? hex.slice(1) : hex;
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

/**
 * Pick black or white text for legibility against a background colour. The
 * mid-gray threshold (128) on the Rec.601 luma splits light vs dark: a dark
 * background ⇒ white text, otherwise black. `bgHex=null` (no background ⇒ page
 * white) ⇒ black text. The black/white pick is implementation-defined — no
 * normative algorithm exists (ECMA-376 §17.3.2.6 `w:color="auto"` only says the
 * consumer chooses "an appropriate color based on the background").
 */
export function autoContrastColor(bgHex: string | null): '#000000' | '#FFFFFF' {
  if (!bgHex) return '#000000';
  return relativeLuma(bgHex) < 128 ? '#FFFFFF' : '#000000';
}

/**
 * Resolve a Fill to a CanvasRenderingContext2D-compatible paint.
 * Gradients require pixel bounds (x, y, w, h) to construct the CanvasGradient.
 * Returns null for noFill.
 */
export function resolveFill(
  fill: Fill | null,
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  shapeRotationDeg = 0,
  patternPtToUserUnits?: number,
  patternCoordinateTransform?: DOMMatrix2DInit,
): string | CanvasGradient | CanvasPattern | null {
  if (!fill || fill.fillType === 'none') return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'pattern') {
    const slideRoot = activePatternCoordinateRoot.get(ctx);
    return resolvePatternFill(
      fill, ctx, patternPtToUserUnits ?? activePatternPointScale.get(ctx) ?? 4 / 3,
      patternCoordinateTransform ?? (slideRoot
        ? patternTransformToRoot(ctx.getTransform(), slideRoot)
        : undefined),
    );
  }
  if (fill.fillType === 'gradient') {
    const stops = fill.stops;
    if (stops.length === 0) return null;
    if (stops.length === 1) return hexToRgba(stops[0].color);

    const repeated = tiledGradient(fill, ctx, x, y, w, h, shapeRotationDeg);
    if (repeated) return repeated;

    let gradient: CanvasGradient;
    const tile = fill.tileRect;
    const tileX = x + w * (tile?.l ?? 0);
    const tileY = y + h * (tile?.t ?? 0);
    const tileW = w * (1 - (tile?.l ?? 0) - (tile?.r ?? 0));
    const tileH = h * (1 - (tile?.t ?? 0) - (tile?.b ?? 0));
    if (fill.gradType === 'radial') {
      // §20.1.8.31: fillToRect is the center-shade (focus) rectangle inside
      // the gradient tile. Canvas has a point focus rather than a rectangular
      // focus, so use its authored centre and retain the full rectangle on the
      // public model for richer hosts.
      const focus = fill.fillToRect;
      const focusX = tileX + tileW * (focus?.l ?? 0);
      const focusY = tileY + tileH * (focus?.t ?? 0);
      const focusW = tileW * (1 - (focus?.l ?? 0) - (focus?.r ?? 0));
      const focusH = tileH * (1 - (focus?.t ?? 0) - (focus?.b ?? 0));
      const cx = focusX + focusW / 2;
      const cy = focusY + focusH / 2;
      const rx = Math.max(Math.abs(cx - tileX), Math.abs(tileX + tileW - cx));
      const ry = Math.max(Math.abs(cy - tileY), Math.abs(tileY + tileH - cy));
      const r = fill.path === 'rect'
        ? Math.max(rx, ry)
        : Math.sqrt(rx * rx + ry * ry);
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.max(r, 1e-9));
    } else {
      const authoredAngle = fill.rotWithShape === false
        ? fill.angle - shapeRotationDeg
        : fill.angle;
      const rad = (authoredAngle * Math.PI) / 180;
      let dx = Math.cos(rad);
      let dy = Math.sin(rad);
      if (fill.scaled === true) {
        // §20.1.8.41: (cos x, sin x) becomes (w cos x, h sin x)
        // before normalization.
        dx *= tileW;
        dy *= tileH;
        const magnitude = Math.hypot(dx, dy);
        if (magnitude > 0) {
          dx /= magnitude;
          dy /= magnitude;
        }
      }
      const cx = tileX + tileW / 2;
      const cy = tileY + tileH / 2;
      const gradLen = (Math.abs(dx) * tileW + Math.abs(dy) * tileH) / 2;
      gradient = ctx.createLinearGradient(
        cx - dx * gradLen, cy - dy * gradLen,
        cx + dx * gradLen, cy + dy * gradLen,
      );
    }
    for (const stop of stops) {
      gradient.addColorStop(Math.min(1, Math.max(0, stop.position)), hexToRgba(stop.color));
    }
    return gradient;
  }
  return null;
}

// Chart families resolve fills through many shared painters. A chart installs
// its host's point scale for the duration of the render, so chart-space,
// series, legend, labels, markers and optional 3-D paths use the same units.
// The scope changes no canvas geometry, so a chart without patterns paints
// exactly as before. The scale is restored for nested charts and later shapes.
const activePatternPointScale = new WeakMap<CanvasRenderingContext2D, number>();
type PatternMatrix = { a: number; b: number; c: number; d: number; e: number; f: number };
const activePatternCoordinateRoot = new WeakMap<CanvasRenderingContext2D, PatternMatrix>();

function patternTransformToRoot(current: DOMMatrix2DInit, root: PatternMatrix): PatternMatrix | undefined {
  const a = current.a ?? 1; const b = current.b ?? 0;
  const c = current.c ?? 0; const d = current.d ?? 1;
  const e = current.e ?? 0; const f = current.f ?? 0;
  const determinant = a * d - b * c;
  if (Math.abs(determinant) < 1e-12) return undefined;
  const ia = d / determinant; const ib = -b / determinant;
  const ic = -c / determinant; const id = a / determinant;
  const ie = -(ia * e + ic * f); const iff = -(ib * e + id * f);
  return {
    a: ia * root.a + ic * root.b,
    b: ib * root.a + id * root.b,
    c: ia * root.c + ic * root.d,
    d: ib * root.c + id * root.d,
    e: ia * root.e + ic * root.f + ie,
    f: ib * root.e + id * root.f + iff,
  };
}

/** PowerPoint PDF paints pattern cells on slide axes even when a shape is
 * rotated, reflected or inside a scaled/rotated group. The 64px tile stays
 * 8pt wide for shapes, table cells, chart fills and text on 4:3, 16:9 and
 * custom slides. Preserve the slide-to-device frame while a host changes its
 * CTM; the pattern resolver cancels only the host-local transform. */
export function withPatternCoordinateSpace<T>(
  ctx: CanvasRenderingContext2D,
  slideToDevice: DOMMatrix2DInit,
  paint: () => T,
): T {
  const previous = activePatternCoordinateRoot.get(ctx);
  activePatternCoordinateRoot.set(ctx, {
    a: slideToDevice.a ?? 1, b: slideToDevice.b ?? 0,
    c: slideToDevice.c ?? 0, d: slideToDevice.d ?? 1,
    e: slideToDevice.e ?? 0, f: slideToDevice.f ?? 0,
  });
  try {
    return paint();
  } finally {
    if (previous) activePatternCoordinateRoot.set(ctx, previous);
    else activePatternCoordinateRoot.delete(ctx);
  }
}

export function withPatternPointScale<T>(
  ctx: CanvasRenderingContext2D,
  ptToUserUnits: number,
  paint: () => T,
): T {
  const previous = activePatternPointScale.get(ctx);
  activePatternPointScale.set(ctx, ptToUserUnits);
  try {
    return paint();
  } finally {
    if (previous === undefined) activePatternPointScale.delete(ctx);
    else activePatternPointScale.set(ctx, previous);
  }
}

/** Preserve a chart's pattern units when effect compositing repaints its body
 * into a temporary canvas. The auxiliary context receives the original chart
 * CTM separately, so its pattern uses the same user-space point grid. */
export function withInheritedPatternScope<T>(
  source: CanvasRenderingContext2D,
  target: CanvasRenderingContext2D,
  paint: () => T,
  deviceOffset?: { x: number; y: number },
  sourceDeviceToTargetDevice?: PatternMatrix,
): T {
  const scale = activePatternPointScale.get(source);
  const root = activePatternCoordinateRoot.get(source);
  // An effect canvas may crop the source, while a bevel canvas additionally
  // changes its axes to the shape's local frame. Transfer the complete affine
  // slide frame in that case; a translation alone would rotate the tile with
  // the bevel when the shape has an authored transform.
  const inherited = root && sourceDeviceToTargetDevice
    ? {
        a: sourceDeviceToTargetDevice.a * root.a + sourceDeviceToTargetDevice.c * root.b,
        b: sourceDeviceToTargetDevice.b * root.a + sourceDeviceToTargetDevice.d * root.b,
        c: sourceDeviceToTargetDevice.a * root.c + sourceDeviceToTargetDevice.c * root.d,
        d: sourceDeviceToTargetDevice.b * root.c + sourceDeviceToTargetDevice.d * root.d,
        e: sourceDeviceToTargetDevice.a * root.e + sourceDeviceToTargetDevice.c * root.f + sourceDeviceToTargetDevice.e,
        f: sourceDeviceToTargetDevice.b * root.e + sourceDeviceToTargetDevice.d * root.f + sourceDeviceToTargetDevice.f,
      }
    : root && deviceOffset
      ? { ...root, e: root.e - deviceOffset.x, f: root.f - deviceOffset.y }
      : root;
  const run = () => inherited
    ? withPatternCoordinateSpace(target, inherited, paint)
    : paint();
  return scale === undefined
    ? run()
    : withPatternPointScale(target, scale, run);
}

/**
 * Build a tiling CanvasPattern for an OOXML preset pattern fill.
 * Falls back to the foreground colour string when the preset name is unknown
 * or the OffscreenCanvas / Canvas environment cannot create a pattern.
 *
 * Cached per (preset, fg, bg, coordinate-unit scale) tuple. The per-context
 * bound limits retained tile resources when a document uses many distinct
 * colours.
 */
const patternCache = new WeakMap<CanvasRenderingContext2D, Map<string, CanvasPattern>>();
// 256 64×64 RGBA tiles retain at most 4 MiB of pixel data per context,
// aside from CanvasPattern/Map bookkeeping; evict oldest colours beyond that.
const MAX_PATTERN_CACHE_ENTRIES = 256;

function resolvePatternFill(
  fill: PatternFill,
  ctx: CanvasRenderingContext2D,
  ptToUserUnits: number,
  coordinateTransform?: DOMMatrix2DInit,
): CanvasPattern | string {
  const tx = coordinateTransform;
  const key = `${fill.preset}|${fill.fg}|${fill.bg}|${ptToUserUnits}|${tx?.a ?? 1},${tx?.b ?? 0},${tx?.c ?? 0},${tx?.d ?? 1},${tx?.e ?? 0},${tx?.f ?? 0}`;
  let perCtx = patternCache.get(ctx);
  if (!perCtx) {
    perCtx = new Map();
    patternCache.set(ctx, perCtx);
  }
  const cached = perCtx.get(key);
  if (cached) return cached;

  const bitmap = buildPatternBitmap(fill.preset, fill.fg, fill.bg);
  if (!bitmap) return hexToRgba(fill.fg);
  const pat = ctx.createPattern(bitmap, 'repeat');
  if (!pat) return hexToRgba(fill.fg);
  // PowerPoint PDF uses one point per 8×8-pixel bitmap cell. In CSS-pixel renderers,
  // 1 pt = 4/3 px; DOCX supplies 1 because its paint coordinates are points.
  // Zero translation anchors the tile to the caller's coordinate origin.
  // PPTX and DOCX paint in slide/page coordinates. XLSX translates into a
  // shape-local frame; its print PDF uses a different phase (see its painter).
  if (typeof pat.setTransform === 'function') {
    const sampleToUser = ptToUserUnits / 8;
    pat.setTransform({
      a: (tx?.a ?? 1) * sampleToUser,
      b: (tx?.b ?? 0) * sampleToUser,
      c: (tx?.c ?? 0) * sampleToUser,
      d: (tx?.d ?? 1) * sampleToUser,
      e: tx?.e ?? 0,
      f: tx?.f ?? 0,
    });
  }
  if (perCtx.size >= MAX_PATTERN_CACHE_ENTRIES) {
    const oldest = perCtx.keys().next().value;
    if (oldest !== undefined) perCtx.delete(oldest);
  }
  perCtx.set(key, pat);
  return pat;
}

/**
 * Apply a Stroke to ctx. `emuPerPx` converts stroke width from EMU to px
 * (e.g. scale factor from pptx's emuToPx).
 *
 * Parsers normalize symbolic dash vocabularies to DrawingML preset names; the
 * shared resolver also accepts VML's numeric relative grammar. Both scale by
 * the pixel line width and return `[]` for solid / unknown styles.
 */
export function applyStroke(
  ctx: CanvasRenderingContext2D,
  stroke: Stroke | null,
  emuPerPx: number,
): void {
  if (!stroke) {
    ctx.strokeStyle = 'transparent';
    ctx.lineWidth = 0;
    ctx.setLineDash([]);
    ctx.lineCap = 'butt';
    ctx.lineJoin = 'miter';
    ctx.miterLimit = 10;
    return;
  }
  ctx.strokeStyle = hexToRgba(stroke.color);
  const lw = Math.max(0.5, stroke.width * emuPerPx);
  ctx.lineWidth = lw;
  const dash = stroke.customDash != null
    ? drawingmlLineDashArray(stroke.customDash, null, lw)
    : stroke.dashStyle
      ? shapeStrokeDashArray(stroke.dashStyle, lw)
      : [];
  const requestedCap = stroke.lineCap ?? 'butt';
  // ECMA-376 Part 4 §19.1.2.21: a zero in VML's numeric dash grammar is a
  // visible fourfold-symmetric dot, even though VML's default endcap is flat.
  // Canvas drops a zero-length dash under its equivalent `butt` cap. A square
  // cap is the exact centered, fourfold-symmetric fallback for that one case;
  // explicit round/square caps keep their authored shape.
  const hasZeroDash = dash.some((length, index) => index % 2 === 0 && length === 0);
  ctx.lineCap = requestedCap === 'butt' && hasZeroDash ? 'square' : requestedCap;
  ctx.lineJoin = stroke.lineJoin ?? 'miter';
  ctx.miterLimit = stroke.miterLimit ?? 10;
  ctx.setLineDash(dash);
}

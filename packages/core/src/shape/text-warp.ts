/**
 * WordArt text-warp envelopes — ECMA-376 §20.1.9.19 (`prstTxWarp`,
 * `ST_TextShapeType`).
 *
 * A text warp is authored as `<a:bodyPr><a:prstTxWarp prst="…"><a:avLst>`. The
 * `prst` names one of the 40 envelopes in the spec's
 * `presetTextWarpDefinitions.xml` (shipped here as `text-warp-presets.json`).
 * Each envelope's `<pathLst>` is NOT a closed silhouette but a pair of OPEN
 * boundary curves — a TOP edge and a BOTTOM edge — spanning the shape's width.
 * The warp maps flat text into the region between the two edges: a glyph at
 * horizontal fraction `u ∈ [0,1]` is placed against the top/bottom points
 * `T(u)` / `B(u)`, rotated to the local slope, and scaled vertically to the
 * edge-to-edge gap. This is the standard per-glyph WordArt approximation on a
 * Canvas 2D surface, which cannot deform glyph outlines directly.
 *
 * The guide-formula grammar is identical to a preset SHAPE's, so this module
 * reuses {@link createEvaluator} / {@link compileFormula} verbatim — the same
 * postfix evaluator that drives `renderPresetShape`. The only new work is
 * FLATTENING each `<path>` into a polyline (arcs / béziers sampled) rather than
 * stroking it, so the envelope can be resampled by arc length. The module is
 * pure (no canvas), so the geometry is unit-testable.
 */

import warpPresetsJson from './preset-geometry/text-warp-presets.json';
import { hasOwnKey, ownValue } from '../internal/own-key';
import {
  createEvaluator,
  compileFormula,
  type Evaluator,
  type CompiledFormula,
} from './preset-geometry/evaluator';

// 60 000-ths of a degree per full revolution (2π radians) — the OOXML angle
// unit shared by <arcTo> stAng/swAng and the evaluator's cd/cd2/… built-ins.
const DEG60K_TO_RAD = (Math.PI * 2) / 21600000;

interface WarpPathDef {
  w: number | null;
  h: number | null;
  fill: string | null;
  stroke: boolean;
  extrusionOk: boolean;
  cmds: Array<[string, ...string[]]>;
}
interface WarpDef {
  adj: [string, string][];
  gd: [string, string][];
  paths: WarpPathDef[];
}
interface CompiledWarpDef {
  adj: [string, CompiledFormula][];
  gd: [string, CompiledFormula][];
  paths: WarpPathDef[];
}

const WARP_PRESETS = warpPresetsJson as unknown as Record<string, WarpDef>;
const COMPILED = new Map<string, CompiledWarpDef>();

/** True when `preset` (case-insensitive) is a known text-warp envelope. */
export function hasTextWarp(preset: string): boolean {
  return hasOwnKey(WARP_PRESETS, preset.toLowerCase());
}

function compiledFor(key: string): CompiledWarpDef | null {
  let c = COMPILED.get(key);
  if (c) return c;
  const def = ownValue(WARP_PRESETS, key);
  if (!def) return null;
  c = {
    adj: def.adj.map(([n, f]) => [n, compileFormula(f)] as [string, CompiledFormula]),
    gd: def.gd.map(([n, f]) => [n, compileFormula(f)] as [string, CompiledFormula]),
    paths: def.paths,
  };
  COMPILED.set(key, c);
  return c;
}

/** A flattened polyline in shape-local pixel space (already scaled to w×h). */
export type Polyline = Array<{ x: number; y: number }>;

/** Number of straight segments a bézier / arc is flattened into. */
const FLATTEN_STEPS = 48;

/**
 * Flatten one warp `<path>` into a polyline in shape-local pixel space. Curves
 * (cubic/quad bézier, elliptical arc) are subdivided into {@link FLATTEN_STEPS}
 * chords. Path-local coords (`path.w`/`path.h`) are scaled to `[0,w]×[0,h]`, the
 * same mapping the path executor applies for shapes.
 */
function flattenPath(path: WarpPathDef, evaluator: Evaluator, w: number, h: number): Polyline {
  const sx = path.w != null ? w / path.w : 1;
  const sy = path.h != null ? h / path.h : 1;
  const X = (v: number) => v * sx;
  const Y = (v: number) => v * sy;
  const pts: Polyline = [];
  let penX = 0;
  let penY = 0;

  for (const cmd of path.cmds) {
    switch (cmd[0]) {
      case 'm': {
        penX = X(evaluator.resolve(cmd[1]));
        penY = Y(evaluator.resolve(cmd[2]));
        pts.push({ x: penX, y: penY });
        break;
      }
      case 'l': {
        penX = X(evaluator.resolve(cmd[1]));
        penY = Y(evaluator.resolve(cmd[2]));
        pts.push({ x: penX, y: penY });
        break;
      }
      case 'C': {
        const c1x = X(evaluator.resolve(cmd[1]));
        const c1y = Y(evaluator.resolve(cmd[2]));
        const c2x = X(evaluator.resolve(cmd[3]));
        const c2y = Y(evaluator.resolve(cmd[4]));
        const ex = X(evaluator.resolve(cmd[5]));
        const ey = Y(evaluator.resolve(cmd[6]));
        for (let i = 1; i <= FLATTEN_STEPS; i++) {
          const t = i / FLATTEN_STEPS;
          const mt = 1 - t;
          const bx =
            mt * mt * mt * penX +
            3 * mt * mt * t * c1x +
            3 * mt * t * t * c2x +
            t * t * t * ex;
          const by =
            mt * mt * mt * penY +
            3 * mt * mt * t * c1y +
            3 * mt * t * t * c2y +
            t * t * t * ey;
          pts.push({ x: bx, y: by });
        }
        penX = ex;
        penY = ey;
        break;
      }
      case 'Q': {
        const c1x = X(evaluator.resolve(cmd[1]));
        const c1y = Y(evaluator.resolve(cmd[2]));
        const ex = X(evaluator.resolve(cmd[3]));
        const ey = Y(evaluator.resolve(cmd[4]));
        for (let i = 1; i <= FLATTEN_STEPS; i++) {
          const t = i / FLATTEN_STEPS;
          const mt = 1 - t;
          const bx = mt * mt * penX + 2 * mt * t * c1x + t * t * ex;
          const by = mt * mt * penY + 2 * mt * t * c1y + t * t * ey;
          pts.push({ x: bx, y: by });
        }
        penX = ex;
        penY = ey;
        break;
      }
      case 'a': {
        // Elliptical arc. Visual (not parametric) start/sweep angles, matching
        // the path-executor's ooxmlArcTo. Radii are in path-local space, then
        // scaled per axis; center is back-solved from the pen.
        const wRlocal = evaluator.resolve(cmd[1]);
        const hRlocal = evaluator.resolve(cmd[2]);
        const wR = wRlocal * sx;
        const hR = hRlocal * sy;
        const stDeg = evaluator.resolve(cmd[3]) * DEG60K_TO_RAD;
        const swDeg = evaluator.resolve(cmd[4]) * DEG60K_TO_RAD;
        const visualToParam = (v: number) =>
          Math.atan2(wRlocal * Math.sin(v), hRlocal * Math.cos(v));
        const TWO_PI = Math.PI * 2;
        const stP = visualToParam(stDeg);
        const fullRevs = Math.trunc(swDeg / TWO_PI);
        const remainder = swDeg - fullRevs * TWO_PI;
        let delta = visualToParam(stDeg + remainder) - stP;
        if (remainder > 0 && delta < 0) delta += TWO_PI;
        else if (remainder < 0 && delta > 0) delta -= TWO_PI;
        const sweepP = delta + fullRevs * TWO_PI;
        const cx = penX - wR * Math.cos(stP);
        const cy = penY - hR * Math.sin(stP);
        // More steps for arcs: they carry the whole textArch/textCircle glyph run.
        const arcSteps = Math.max(FLATTEN_STEPS, Math.ceil((Math.abs(sweepP) / TWO_PI) * 96));
        for (let i = 1; i <= arcSteps; i++) {
          const p = stP + (sweepP * i) / arcSteps;
          pts.push({ x: cx + wR * Math.cos(p), y: cy + hR * Math.sin(p) });
        }
        penX = cx + wR * Math.cos(stP + sweepP);
        penY = cy + hR * Math.sin(stP + sweepP);
        break;
      }
      case 'c':
        // Close: warp envelopes are open edges, so a stray close is a no-op for
        // the polyline (there is nothing to seam back to for a boundary curve).
        break;
    }
  }
  return pts;
}

/** The two boundary edges of a warp, flattened to shape-local pixels. */
export interface WarpEnvelope {
  top: Polyline;
  bottom: Polyline;
  /** Cumulative arc-length lookup, precomputed for even u-sampling. */
  topLen: number[];
  bottomLen: number[];
  /** True for arch/circle presets whose single path is the glyph BASELINE
   *  (top === bottom). The mapper then keeps glyph height fixed and offsets the
   *  em-box symmetrically about the curve, instead of scaling to a gap. */
  singleEdge: boolean;
}

function cumulativeLengths(poly: Polyline): number[] {
  const acc = [0];
  for (let i = 1; i < poly.length; i++) {
    const dx = poly[i].x - poly[i - 1].x;
    const dy = poly[i].y - poly[i - 1].y;
    acc.push(acc[i - 1] + Math.hypot(dx, dy));
  }
  return acc;
}

/**
 * Build a warp's two boundary edges for a shape of `w × h` pixels.
 *
 * `adj` are the `<a:avLst>` adjust values (thousandths of a percent) in
 * adj1/adj2/… order; an empty array uses the preset defaults. Returns `null`
 * when the preset name is unknown (caller then renders flat text).
 *
 * The envelope is assembled from the preset's paths: the FIRST path is the top
 * edge, the SECOND the bottom edge. Single-path presets (textArchUp/Down,
 * textCircle) have no separate bottom edge — the one path is the baseline arc
 * the glyphs sit ON; {@link warpBaselinePoint} handles that by offsetting the
 * glyph box symmetrically about the arc. Presets with more than two paths
 * (textButton, textDeflateInflate, …) use the first and last paths as the outer
 * top/bottom edges, a faithful-enough envelope for per-glyph mapping.
 */
export function buildWarpEnvelope(
  preset: string,
  adj: number[],
  w: number,
  h: number,
): WarpEnvelope | null {
  const key = preset.toLowerCase();
  const def = compiledFor(key);
  if (!def || def.paths.length === 0) return null;
  const evaluator = createEvaluator({ w, h, adj }, def.adj, def.gd);

  const singleEdge = def.paths.length === 1;
  const top = flattenPath(def.paths[0], evaluator, w, h);
  const bottom = singleEdge
    ? top
    : flattenPath(def.paths[def.paths.length - 1], evaluator, w, h);
  return {
    top,
    bottom,
    topLen: cumulativeLengths(top),
    bottomLen: cumulativeLengths(bottom),
    singleEdge,
  };
}

/**
 * Sample a polyline at fraction `u ∈ [0,1]` of its ARC LENGTH. Returns the point
 * and the local unit tangent (direction of travel along the curve). Even
 * arc-length sampling — not raw index — keeps glyph advance uniform where the
 * envelope's control-point spacing is not.
 */
export function samplePolyline(
  poly: Polyline,
  cum: number[],
  u: number,
): { x: number; y: number; tx: number; ty: number } {
  const total = cum[cum.length - 1];
  if (poly.length === 1 || total === 0) {
    return { x: poly[0].x, y: poly[0].y, tx: 1, ty: 0 };
  }
  const target = Math.max(0, Math.min(1, u)) * total;
  // Binary search for the segment containing `target`.
  let lo = 0;
  let hi = cum.length - 1;
  while (lo < hi - 1) {
    const mid = (lo + hi) >> 1;
    if (cum[mid] <= target) lo = mid;
    else hi = mid;
  }
  const segLen = cum[hi] - cum[lo] || 1;
  const f = (target - cum[lo]) / segLen;
  const a = poly[lo];
  const b = poly[hi];
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const norm = Math.hypot(dx, dy) || 1;
  return {
    x: a.x + dx * f,
    y: a.y + dy * f,
    tx: dx / norm,
    ty: dy / norm,
  };
}

/**
 * The baseline arc length (px) of a single-edge warp — the total flattened
 * length of the one boundary curve the glyphs sit on. For paired-edge presets
 * (which have no single baseline) this returns the top edge's length, but the
 * value is only meaningful for {@link WarpEnvelope.singleEdge} presets, where it
 * is the path length used by Follow Path distribution.
 */
export function warpArcLength(env: WarpEnvelope): number {
  return env.topLen[env.topLen.length - 1] ?? 0;
}

/**
 * Follow Path distribution factor for single-edge (arch/circle) warps
 * (ECMA-376 §20.1.9.19). PowerPoint lays WordArt out along an arch/circle
 * baseline using "Follow Path" semantics: the text keeps its NATURAL width and
 * is placed from the path start (stAng) for only `naturalWidth` of arc length —
 * it is NOT stretched to span the whole path. So a glyph at horizontal fraction
 * `u ∈ [0,1]` of the flat text must map to arc parameter `u · (naturalWidth /
 * arcLength)`, i.e. this factor scales the `[0,1]` glyph range down to the arc
 * span the text actually fills.
 *
 * The factor is clamped to `[0,1]`: Follow Path only ever SHRINKS the span (text
 * shorter than the path occupies a leading segment); text longer than the path
 * is clamped to the full path rather than stretched or wrapped. Paired-edge
 * presets return `1` (identity) — they deliberately stretch the flat ink box to
 * fill the envelope width, so their glyph `u` already spans `[0,1]`.
 */
export function followPathUScale(env: WarpEnvelope, naturalWidth: number): number {
  if (!env.singleEdge) return 1;
  const arc = warpArcLength(env);
  if (arc <= 0) return 1;
  return Math.max(0, Math.min(1, naturalWidth / arc));
}

/**
 * A per-glyph placement in warped space. The renderer draws the glyph with
 * `ctx.translate(x, y); ctx.rotate(angle); ctx.transform(1, 0, shear, 1, 0, 0);
 * ctx.scale(hScale, vScale)` and then paints at the glyph's local baseline
 * (`fillText(g, 0, 0)`), so the flat glyph em-box is bent onto the envelope.
 *
 * The four linear terms together reproduce the LOCAL JACOBIAN of the envelope
 * map `P(u, v) = (1−v)·T(u) + v·B(u)` at the glyph centre — see
 * {@link warpGlyphTransform}. `angle` + `hScale` orient/scale the horizontal
 * (advance) axis to the mean edge tangent; `shear` + `vScale` bend the vertical
 * (em-height) axis onto the true gap vector `B(u)−T(u)`. On a sloped envelope
 * edge (a wave crest, a taper) these two axes are NOT orthogonal, so `shear`
 * is what tilts the glyph into a parallelogram — vertical strokes stay aligned
 * with the gap while the baseline follows the slope, matching PowerPoint's
 * WordArt (the glyph "leans" along the wave rather than rigidly rotating).
 */
export interface WarpGlyphTransform {
  /** Canvas-space baseline origin (shape-local; caller adds the shape offset). */
  x: number;
  y: number;
  /** Rotation of the local text (advance) axis, radians (mean edge slope). */
  angle: number;
  /** Vertical scale of the glyph em-box along the (post-shear) up-axis. */
  vScale: number;
  /**
   * Horizontal skew of the em-box in the rotated frame (0 = no shear). Applied
   * as `ctx.transform(1, 0, shear, 1, 0, 0)` AFTER `rotate(angle)` and BEFORE
   * `scale(hScale, vScale)`: it slides the top of the glyph forward/back along
   * the advance axis so the vertical column of the transform lands on the true
   * envelope gap vector `B(u)−T(u)`. Zero for single-edge (arch/circle) presets
   * and for flat (unsloped) envelope points; nonzero on sloped paired edges.
   */
  shear: number;
}

/**
 * Map a glyph anchored at horizontal fraction `u ∈ [0,1]` of the text width into
 * warped space.
 *
 * `boxHeight` is the flat line's height in px (the space the glyph occupies
 * vertically before warping). `baselineFrac` is where the baseline sits within
 * that box, top→bottom (0 = top edge, 1 = bottom edge) — typically ascent/height.
 *
 * Two regimes:
 * - **Paired-edge** (waves, inflate/deflate, cascade, …): the glyph em-box is
 *   fitted between `T(u)` and `B(u)` by the LOCAL JACOBIAN of the envelope map
 *   `P(u, v) = (1−v)·T(u) + v·B(u)`. Its two columns are `∂P/∂u ≈ mean edge
 *   tangent` (the advance axis → `angle`) and `∂P/∂v = B(u)−T(u)` (the em-height
 *   axis → the gap vector). Because these columns are generally NOT orthogonal on
 *   a sloped edge, the transform is a rotate + shear + non-uniform scale, not a
 *   plain rotate+scale: `angle` orients the advance axis, `shear`+`vScale` bend
 *   the vertical column onto the true gap vector so vertical strokes track the
 *   gap while the baseline follows the slope (the glyph LEANS, matching
 *   PowerPoint WordArt — cf. ECMA-376 §20.1.9.19). The baseline lands at
 *   `baselineFrac` of the top→bottom span.
 * - **Single-edge** (arch/circle): the one curve IS the baseline. The glyph sits
 *   ON the curve, its up-axis is the curve NORMAL (so letters stand perpendicular
 *   to the arc), `vScale` stays 1 (glyphs keep their height), `shear` is 0, and
 *   the baseline is nudged along the normal by the box's descent so the arc
 *   passes through the glyph baseline rather than its top. Follow Path presets
 *   deliberately keep a rigid per-glyph rotation (no shear).
 */
export function warpGlyphTransform(
  env: WarpEnvelope,
  u: number,
  boxHeight: number,
  baselineFrac: number,
): WarpGlyphTransform {
  if (env.singleEdge) {
    const c = samplePolyline(env.top, env.topLen, u);
    const angle = Math.atan2(c.ty, c.tx);
    // The inward normal (rotate tangent −90°) points toward the box interior.
    // For an "up" arch the curve is the top of the text, so the box hangs BELOW
    // it (in the direction of increasing box y). We place the baseline a descent
    // below the curve along +normal so the glyph's cap sits near the arc.
    const nx = c.ty; // (tx,ty) rotated +90° → (−ty, tx); inward = (ty, −tx)…
    const ny = -c.tx;
    const descent = boxHeight * (1 - baselineFrac);
    return {
      x: c.x - nx * descent,
      y: c.y - ny * descent,
      angle,
      vScale: 1,
      shear: 0,
    };
  }

  const t = samplePolyline(env.top, env.topLen, u);
  const b = samplePolyline(env.bottom, env.bottomLen, u);
  const gapx = b.x - t.x;
  const gapy = b.y - t.y;
  const bx = t.x + gapx * baselineFrac;
  const by = t.y + gapy * baselineFrac;
  // Advance axis: mean of the two edges' unit tangents (∂P/∂u direction).
  const sumTx = t.tx + b.tx;
  const sumTy = t.ty + b.ty;
  const angle = Math.atan2(sumTy, sumTx);

  // Em-height axis: the TRUE gap vector ∂P/∂v = B(u)−T(u), per unit box height.
  // Decompose it in the rotated (advance-axis) frame. Its component ALONG the
  // advance axis is the shear (skews the glyph top forward/back on a slope); its
  // component PERPENDICULAR is the vertical scale. When the edge is flat this
  // gap vector is already perpendicular to the tangent → shear≈0, vScale=gap/box,
  // recovering the old rotate+scale behaviour exactly.
  const c = Math.cos(angle);
  const s = Math.sin(angle);
  const gapAlong = (c * gapx + s * gapy) / (boxHeight > 0 ? boxHeight : 1);
  const gapPerp = (-s * gapx + c * gapy) / (boxHeight > 0 ? boxHeight : 1);
  const vScale = gapPerp !== 0 ? gapPerp : boxHeight > 0 ? Math.hypot(gapx, gapy) / boxHeight : 1;
  const shear = gapPerp !== 0 ? gapAlong / gapPerp : 0;

  return {
    x: bx,
    y: by,
    angle,
    vScale,
    shear,
  };
}

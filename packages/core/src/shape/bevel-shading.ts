/** DrawingML bevels: ECMA-376 §§20.1.5.3–4/.9/.11/.12, 20.1.10.9/.50;
 * MS-OI29500 §§2.1.1274/.1316/.1324/.1334 define profiles, rigs and materials.
 * Shared Canvas implementation: region-local O(pixels) EDT and normal field,
 * fixed-size profile lookup, at most four lights, no per-pixel allocations.
 * This reference shader replaces the old fitted multiply/screen appearance.
 * See bevel-lighting.ts for the remaining Office photometric fidelity boundary.
 * Control audit: 6,051 Windows PowerPoint front-view rectangle/ellipse controls,
 * all 12 profiles, 27 rigs and 15 material names (wireframe uses raster geometry),
 * eight directions, five fill colors and explicit single/combined light rotations.
 * Office residuals remain, particularly with Fresnel materials. This reference
 * model is not a complete Office pixel match. The profile axis assignment is
 * settled against these controls (see bevel-profiles.ts); Fresnel/compositing
 * remain unresolved; large residuals are not raster AA. The light rotation convention
 * is checked separately against single-axis and combined Office controls.
 * Arbitrary backgrounds, camera-rotated lighting and back/mesh visibility are
 * not established by these controls. No per-case fit or scalar gain is used.
 */
import { profileHeight, profileSlope } from './bevel-profiles';
import { scene3dSurfaceFrame, type CameraInput } from './scene3d-camera';
import { rasterBevelWireframe } from './bevel-wireframe';
import {
  presetMaterial,
  prepareLighting,
  rigVector,
  shadeSurface,
  type PresetMaterial,
} from './bevel-lighting';

export interface Vec3 {
  x: number;
  y: number;
  z: number;
}
export type BevelMaterial = PresetMaterial;
/** ECMA §20.1.5.11 sphere coordinates, in degrees. */
export interface LightRigRot {
  lat: number;
  lon: number;
  rev: number;
}
/** Retained scalar helper contract; full-color shading uses the rig/material tables. */
export interface BevelShadeParams {
  light: Vec3;
  material: BevelMaterial;
  ambient: number;
  diffuse: number;
  specular: number;
  shininess: number;
  fillLight?: Vec3;
  fillDiffuse?: number;
}
/** Visible front-view height from the documented parametric profile. Overhangs
 * use the nearest surface along the viewing ray. Unknown presets use circle,
 * the CT_Bevel schema default. Width zero has no lip.
 */
export function bevelHeightProfile(
  prst: string,
  w: number,
): (d: number) => number {
  return w > 0 ? (d) => profileHeight(prst, d / w) : () => 1;
}

/**
 * 1-D squared Euclidean distance transform (Felzenszwalb & Huttenlocher 2012,
 * "Distance Transforms of Sampled Functions"). Input `f[i]` is the cost at i
 * (0 at a seed, +∞ elsewhere); output is min_j (f[j] + (i-j)²). Running it once
 * per row then once per column gives the exact 2-D squared EDT. O(n) per line.
 */
export function edt1d(f: ArrayLike<number>): Float64Array {
  const n = f.length;
  const d = new Float64Array(n);
  if (n === 0) return d;
  const v = new Int32Array(n); // locations of parabolas in the lower envelope
  const z = new Float64Array(n + 1); // boundaries between parabolas
  let k = 0;
  v[0] = 0;
  z[0] = -Infinity;
  z[1] = Infinity;
  for (let q = 1; q < n; q++) {
    let s = (f[q] + q * q - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
    while (s <= z[k]) {
      k--;
      s = (f[q] + q * q - (f[v[k]] + v[k] * v[k])) / (2 * q - 2 * v[k]);
    }
    k++;
    v[k] = q;
    z[k] = s;
    z[k + 1] = Infinity;
  }
  k = 0;
  for (let q = 0; q < n; q++) {
    while (z[k + 1] < q) k++;
    const dq = q - v[k];
    d[q] = dq * dq + f[v[k]];
  }
  return d;
}

/**
 * Box sizes whose 3-fold convolution approximates a Gaussian of std-dev `sigma`
 * (Wells 1986, "Efficient synthesis of Gaussian filters by cascaded uniform
 * filters"). Three equal-radius box blurs already give a good Gaussian; mixing two
 * adjacent odd widths (`wl` and `wl+2`) matches the target variance more exactly.
 * Returns `n` odd box widths. Used by `gaussianBlur` below.
 */
function boxSizesForGaussian(sigma: number, n = 3): number[] {
  if (sigma <= 0) return new Array(n).fill(1);
  const wIdeal = Math.sqrt((12 * sigma * sigma) / n + 1);
  let wl = Math.floor(wIdeal);
  if (wl % 2 === 0) wl--;
  const wu = wl + 2;
  const mIdeal =
    (12 * sigma * sigma - n * wl * wl - 4 * n * wl - 3 * n) / (-4 * wl - 4);
  const m = Math.round(mIdeal);
  const sizes: number[] = [];
  for (let i = 0; i < n; i++) sizes.push(i < m ? wl : wu);
  return sizes;
}

/**
 * One separable box-blur pass with ZERO padding (out-of-bounds treated as 0).
 * `horizontal` selects the axis. Uses a running sum so cost is O(N) independent
 * of the radius.
 *
 * Zero padding (NOT clamp-to-edge) is the correct boundary condition for a COVERAGE
 * field: a silhouette flush with the canvas edge has transparent (coverage 0) just
 * outside the bounds, so the blurred coverage falls off across that edge and ∇C
 * points outward there — exactly as `distanceToEdge` folds in the out-of-bounds
 * transparent region for the EDT. Clamp-to-edge would instead make the field flat
 * at the border, giving a zero gradient and no lip along a clipped edge.
 */
function boxBlurPass(
  src: Float64Array,
  dst: Float64Array,
  w: number,
  h: number,
  r: number,
  horizontal: boolean,
): void {
  const norm = 1 / (2 * r + 1);
  if (horizontal) {
    for (let y = 0; y < h; y++) {
      const row = y * w;
      // Seed the window [−r, r]: indices < 0 contribute 0 (zero padding).
      let acc = 0;
      for (let k = 0; k <= r; k++) if (k < w) acc += src[row + k];
      for (let x = 0; x < w; x++) {
        dst[row + x] = acc * norm;
        const ai = x + r + 1;
        const si = x - r;
        if (ai < w) acc += src[row + ai];
        if (si >= 0) acc -= src[row + si];
      }
    }
  } else {
    for (let x = 0; x < w; x++) {
      let acc = 0;
      for (let k = 0; k <= r; k++) if (k < h) acc += src[k * w + x];
      for (let y = 0; y < h; y++) {
        dst[y * w + x] = acc * norm;
        const ai = y + r + 1;
        const si = y - r;
        if (ai < h) acc += src[ai * w + x];
        if (si >= 0) acc -= src[si * w + x];
      }
    }
  }
}

/**
 * Approximate Gaussian blur of std-dev `sigma` over a scalar field, via three
 * cascaded separable box blurs (zero-padded). O(N) total, independent of sigma.
 *
 * This is the SAME code path in every environment (browser, OffscreenCanvas,
 * node/skia headless), so the bevel azimuth never depends on `ctx.filter='blur'`
 * — whose support and exact kernel differ across canvas backends. The blur is run
 * on the in-memory coverage field, not on a canvas, so it is fully unit-testable.
 */
export function gaussianBlur(
  src: ArrayLike<number>,
  w: number,
  h: number,
  sigma: number,
): Float64Array {
  const a = Float64Array.from(src as ArrayLike<number> & Iterable<number>);
  if (sigma <= 0 || w <= 0 || h <= 0) return a;
  const scratch = new Float64Array(w * h);
  // Each box does a horizontal pass into `scratch`, then a vertical pass back into
  // `a`, so the result always lands in `a` (no buffer swap to track).
  for (const size of boxSizesForGaussian(sigma, 3)) {
    const r = Math.max(1, (size - 1) / 2);
    boxBlurPass(a, scratch, w, h, r, true);
    boxBlurPass(scratch, a, w, h, r, false);
  }
  return a;
}

/**
 * Exact 2-D Euclidean distance (in px) from every pixel to the nearest pixel
 * whose alpha is below `threshold` (i.e. the boundary of the silhouette). The
 * region OUTSIDE the image bounds is also treated as below-threshold, so a
 * silhouette that runs to the canvas edge still gets a finite edge distance
 * (its bevel lip is drawn along that clipped edge). Pixels below threshold
 * get distance 0. Returns a W·H Float64Array.
 */
export function distanceToEdge(
  alpha: ArrayLike<number>,
  w: number,
  h: number,
  threshold = 128,
): Float64Array {
  const INF = 1e20;
  const f = new Float64Array(w * h);
  for (let i = 0; i < w * h; i++) {
    f[i] = (alpha[i] ?? 0) >= threshold ? INF : 0;
  }
  // The implicit boundary just outside the image: a pixel at row 0 is one px
  // from the (transparent) row −1, so its column-pass cost is capped at that
  // distance. Seed the borders by treating index −1 / N as zero-cost via the
  // edt1d padding below — done by capping the first/last interior value.
  // (edt1d alone can't see outside the line, so fold the out-of-bounds seed in
  //  as min(current, distance-to-border²) after each separable pass.)
  // Transform along columns, then rows (separable).
  const col = new Float64Array(h);
  for (let x = 0; x < w; x++) {
    for (let y = 0; y < h; y++) col[y] = f[y * w + x];
    const dc = edt1d(col);
    for (let y = 0; y < h; y++) f[y * w + x] = dc[y];
  }
  const row = new Float64Array(w);
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) row[x] = f[y * w + x];
    const dr = edt1d(row);
    for (let x = 0; x < w; x++) f[y * w + x] = dr[x];
  }
  // Fold in the implicit transparent region beyond the image edges: row −1 and
  // row h, column −1 and column w are all "outside", so a pixel at (x,y) is at
  // most (y+1) px from the top edge, (h−y) from the bottom, (x+1) from the left,
  // (w−x) from the right. Take the min of the interior EDT and these border
  // distances (squared) so a silhouette flush with the canvas edge still beveled.
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (f[i] === 0) continue; // already a boundary pixel
      const bt = (y + 1) * (y + 1);
      const bb = (h - y) * (h - y);
      const bl = (x + 1) * (x + 1);
      const br = (w - x) * (w - x);
      const border = Math.min(bt, bb, bl, br);
      if (border < f[i]) f[i] = border;
    }
  }
  // f now holds squared distances; sqrt to px.
  for (let i = 0; i < w * h; i++) f[i] = Math.sqrt(f[i]);
  return f;
}

/** Front-view bevel normals. The documented curves have real creases and
 * overhangs; do not feather them into an unchanged face. Their visible envelope
 * supplies the tilt, while the EDT's blurred gradient supplies azimuth. The
 * quarter-band Gaussian is the existing raster regularization policy, not an
 * Office fillet or a fitted lighting parameter. Subtract half a pixel from the
 * center-to-transparent-center EDT to approximate the silhouette boundary.
 * Full camera-dependent tessellation, self-shadowing and back-face occlusion
 * are outside this bitmap surface model.
 */
export function computeBevelNormals(
  alpha: ArrayLike<number>,
  w: number,
  h: number,
  bandPx: number,
  prst: string,
  heightPx: number,
): {
  normals: Float32Array;
  bandMask: Uint8Array;
  bandWeight: Float32Array;
} {
  const normals = new Float32Array(w * h * 3),
    bandMask = new Uint8Array(w * h),
    bandWeight = new Float32Array(w * h);
  if (w <= 0 || h <= 0) return { normals, bandMask, bandWeight };
  // RGBA silhouettes encode opacity as well as raster coverage. Geometry is
  // their nonzero support, not a 50% material-opacity contour.
  const dist = distanceToEdge(alpha, w, h, 1);
  const C =
    bandPx > 0 ? gaussianBlur(dist, w, h, Math.max(1, bandPx * 0.25)) : dist;
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      normals[i * 3 + 2] = 1;
      if ((alpha[i] ?? 0) === 0 || bandPx <= 0 || heightPx <= 0) continue;
      const d = Math.max(0, dist[i] - 0.5);
      if (d >= bandPx) continue;
      bandMask[i] = 1;
      bandWeight[i] = 1;
      const xl = Math.max(0, x - 1),
        xr = Math.min(w - 1, x + 1),
        yu = Math.max(0, y - 1),
        yd = Math.min(h - 1, y + 1);
      const gx = (C[y * w + xr] - C[y * w + xl]) / (xr - xl || 1),
        gy = (C[yd * w + x] - C[yu * w + x]) / (yd - yu || 1);
      const gm = Math.hypot(gx, gy);
      const slope = (profileSlope(prst, d / bandPx) * heightPx) / bandPx;
      const nx = gm > 1e-9 ? (-slope * gx) / gm : 0,
        ny = gm > 1e-9 ? (-slope * gy) / gm : 0;
      const m = Math.hypot(nx, ny, 1);
      normals[i * 3] = nx / m;
      normals[i * 3 + 1] = ny / m;
      normals[i * 3 + 2] = 1 / m;
    }
  return { normals, bandMask, bandWeight };
}
export function lightDirFromRig(
  rig: string,
  dir: string,
  rot?: LightRigRot,
): Vec3 {
  return rigVector(rig, dir, rot);
}
/** Preserve every material name through the historical adapter signature. */
export function materialClass(name: string | undefined): BevelMaterial {
  return presetMaterial(name);
}
/** Scalar compatibility helpers for callers supplying explicit scalar weights.
 * They do not govern the full-color production compositor below.
 */
export function shadeParamsFor(
  material: BevelMaterial,
  light: Vec3,
  fill = true,
): BevelShadeParams {
  const p = prepareLighting(material, light),
    m = p.material;
  return {
    material,
    light,
    ambient: p.ambient,
    diffuse: m.diffuseShape,
    specular: m.specularShape + m.specularWhite,
    shininess: m.power,
    fillLight: fill && p.lights[1] ? p.lights[1] : undefined,
    fillDiffuse: fill && p.lights[1]?.diffuse ? p.lights[1].rgb[0] : 0,
  };
}
export function shadePixel(n: Vec3, p: BevelShadeParams): number {
  let value =
    p.ambient +
    p.diffuse *
      Math.max(0, n.x * p.light.x + n.y * p.light.y + n.z * p.light.z);
  if (p.fillLight && p.fillDiffuse)
    value +=
      p.fillDiffuse *
      Math.max(
        0,
        n.x * p.fillLight.x + n.y * p.fillLight.y + n.z * p.fillLight.z,
      );
  const hm = Math.hypot(p.light.x, p.light.y, p.light.z + 1) || 1;
  if (p.specular > 0)
    value +=
      p.specular *
      Math.pow(
        Math.max(
          0,
          (n.x * p.light.x + n.y * p.light.y + n.z * (p.light.z + 1)) / hm,
        ),
        p.shininess,
      );
  return Math.max(0, value);
}

/** A minimal 2D context surface the bevel compositor needs. */
export interface BevelCtx {
  canvas: { width: number; height: number };
  getImageData(sx: number, sy: number, sw: number, sh: number): ImageData;
  putImageData(data: ImageData, dx: number, dy: number): void;
}

/**
 * A sub-rectangle of the body bitmap that the shading actually touches (perf: A3).
 * The distance transform + normal loop are O(w·h); when the painted silhouette
 * occupies only part of a large offscreen, restricting the `getImageData` window
 * and the loop to `bbox ⊕ effect-reach` (clamped to the canvas) avoids running
 * the whole thing over transparent margin.
 *
 * CORRECTNESS: `distanceToEdge` treats the area OUTSIDE the passed alpha plane as
 * transparent (below-threshold), so a silhouette flush with the plane edge still
 * gets a finite edge distance. Passing a sub-window therefore gives IDENTICAL
 * results to the full canvas **provided the window edge lies in transparent
 * territory** — i.e. the window must contain the silhouette grown by the effect's
 * spatial reach (the bevel band, or the extrusion offset). Callers size the
 * region that way; when omitted the whole canvas is used (unchanged behaviour).
 */
export interface BevelRegion {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Clamp a requested region to the canvas bounds; returns null if it collapses. */
function clampRegion(
  region: BevelRegion | undefined,
  canvasW: number,
  canvasH: number,
): { x: number; y: number; w: number; h: number } {
  if (!region) return { x: 0, y: 0, w: canvasW, h: canvasH };
  const x0 = Math.max(0, Math.floor(region.x));
  const y0 = Math.max(0, Math.floor(region.y));
  const x1 = Math.min(canvasW, Math.ceil(region.x + region.w));
  const y1 = Math.min(canvasH, Math.ceil(region.y + region.h));
  return { x: x0, y: y0, w: Math.max(0, x1 - x0), h: Math.max(0, y1 - y0) };
}

export interface BevelInput {
  widthPx: number;
  heightPx: number;
  prst: string;
  material: BevelMaterial;
  light: Vec3;
  /** Back-face profile (§20.1.5.3). Hidden unless the camera exposes it. */
  bottom?: boolean;
  /** Distance to the back surface, in the same device units as the band. */
  depthPx?: number;
  camera?: CameraInput;
  /** Raster wire width and parallel back-rim offset; no surface lighting. */
  wireframeLinePx?: number;
  wireframeDepth?: { x: number; y: number };
}
/** Shade the entire painted surface, including the flat face. Region callers
 * must include the whole silhouette plus the spatial reach (see BevelRegion).
 * Direct calls with zero/subpixel bevels still shade the flat face. The PPTX
 * adapter currently omits those inputs, so scenes without a nonzero bevel need
 * a separate flat-surface pass in that adapter; this API cannot create a call.
 * ECMA §§20.1.5.3–4 assign bevelB/T to back/front faces. Cull the hidden
 * planar face and shade the visible back with the same profile/material and
 * camera-rotated outward normals. The painter positions it at the extrusion
 * base (MS-OI29500 §2.1.1268). This planar bitmap model does not tessellate
 * overhanging side lips or resolve their self-occlusion; front-view controls
 * establish neither those surfaces nor complete solid/Office fidelity.
 * Wireframe renders rims and connecting edges instead of a filled material.
 */
export function applyBevelShading(
  ctx: BevelCtx,
  input: BevelInput,
  region?: BevelRegion,
): void {
  if (ctx.canvas.width <= 0 || ctx.canvas.height <= 0) return;
  const frame = input.camera ? scene3dSurfaceFrame(input.camera) : undefined;
  const backFacing = frame?.backFacing ?? false;
  if (Boolean(input.bottom) !== backFacing) return;
  const {
    x: rx,
    y: ry,
    w,
    h,
  } = clampRegion(region, ctx.canvas.width, ctx.canvas.height);
  if (w <= 0 || h <= 0) return;
  const img = ctx.getImageData(rx, ry, w, h),
    px = img.data;
  const coverage = new Uint8ClampedArray(w * h);
  for (let i = 0; i < w * h; i++) coverage[i] = px[i * 4 + 3] > 0 ? 255 : 0;
  // Keep geometric support independent of source opacity. Both the distance
  // field and wire crossing predicates use nonzero support; the painter alone
  // multiplies edge coverage by the original source alpha.
  if (input.material === 'legacyWireframe') {
    rasterBevelWireframe(px, coverage, distanceToEdge(coverage, w, h), w, h,
      input.widthPx, input.wireframeLinePx ?? 1, input.wireframeDepth);
    ctx.putImageData(img, rx, ry);
    return;
  }
  const { normals } = computeBevelNormals(
    coverage, w, h, input.widthPx, input.prst, input.heightPx,
  );
  const params = prepareLighting(input.material, input.light),
    out = new Float64Array(4);
  for (let i = 0; i < w * h; i++) {
    const o = i * 4;
    if (px[o + 3] === 0) continue;
    let nx = normals[i * 3], ny = normals[i * 3 + 1], nz = normals[i * 3 + 2];
    if (input.bottom && frame) {
      // Mirroring the profile through its base reverses its axial normal. The
      // camera rotation then maps that outward normal into the lighting frame.
      const R = frame.rotation;
      const x = R[0] * nx + R[1] * ny - R[2] * nz;
      const y = R[3] * nx + R[4] * ny - R[5] * nz;
      nz = R[6] * nx + R[7] * ny - R[8] * nz;
      nx = x; ny = y;
    }
    shadeSurface(
      out,
      px[o],
      px[o + 1],
      px[o + 2],
      px[o + 3],
      nx,
      ny,
      nz,
      params,
    );
    px[o] = out[0];
    px[o + 1] = out[1];
    px[o + 2] = out[2];
    px[o + 3] = out[3];
  }
  ctx.putImageData(img, rx, ry);
}

export interface ExtrusionInput {
  /** Screen-space depth offset of the back face (device px), from the camera. */
  offsetX: number;
  offsetY: number;
  /** Side-wall colour as [r,g,b] (extrusionClr, or a darkened body colour). */
  rgb: [number, number, number];
}

/**
 * Bake an extrusion side-wall band into a painted body bitmap (§20.1.5.12
 * `extrusionH`). The front face sits at z=0 and the back face is pushed to
 * −depth; the screen projection of that push is `(offsetX, offsetY)` device px
 * (from `computeDepthOffset`). The visible side wall is the region swept between
 * the front silhouette and the offset back silhouette.
 *
 * APPROXIMATION (Phase B, documented): we sweep the front silhouette by the
 * single centre depth-offset vector and fill the newly-uncovered pixels (those
 * transparent in the body but covered by an offset copy of the silhouette) with
 * the side-wall colour, UNDER the front face. This is exact for a parallel sweep
 * and a good approximation for small extrusions / gentle tilts; it does NOT
 * model per-silhouette-point frustum divergence or self-occlusion of a deep
 * extrusion under a strong perspective. When the offset is sub-pixel (face-on
 * camera) there is no visible wall and this is a no-op.
 */
export function applyExtrusion(
  ctx: BevelCtx,
  input: ExtrusionInput,
  region?: BevelRegion,
): void {
  const canvasW = ctx.canvas.width;
  const canvasH = ctx.canvas.height;
  if (canvasW <= 0 || canvasH <= 0) return;
  const dx = input.offsetX;
  const dy = input.offsetY;
  const len = Math.hypot(dx, dy);
  if (len < 0.75) return; // sub-pixel — no visible side wall.

  // Restrict to the region the wall can occupy (perf: A3). A wall pixel is a
  // transparent pixel within `len` px of the front silhouette along the depth
  // vector, so it lies inside the silhouette bbox ⊕ |offset|. When the region
  // contains that (callers size it so), the result is identical to the full
  // canvas: the walk-back always moves TOWARD the silhouette (into the region),
  // so no in-region wall pixel needs an out-of-region sample. See BevelRegion.
  const { x: rx, y: ry, w, h } = clampRegion(region, canvasW, canvasH);
  if (w <= 0 || h <= 0) return;

  const img = ctx.getImageData(rx, ry, w, h);
  const px = img.data;
  // Source alpha snapshot, region-local. Nonzero support defines the front
  // silhouette even below 50% opacity; source translucency is not a wall hole.
  const srcA = new Uint8ClampedArray(w * h);
  for (let i = 0; i < w * h; i++) srcA[i] = px[i * 4 + 3];

  const steps = Math.max(1, Math.ceil(len));
  const [r, g, b] = input.rgb;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const i = y * w + x;
      if (srcA[i] > 0) continue; // covered by the front face — wall hidden.
      // Walk back toward the front face along the depth vector; if any sample
      // along the sweep lands on the front silhouette, this pixel is side wall.
      // Coords are region-local (the walk stays in-region — see the note above).
      let onWall = false;
      for (let s = 1; s <= steps; s++) {
        const t = s / steps;
        const sx = Math.round(x - dx * t);
        const sy = Math.round(y - dy * t);
        if (sx < 0 || sy < 0 || sx >= w || sy >= h) continue;
        if (srcA[sy * w + sx] > 0) {
          onWall = true;
          break;
        }
      }
      if (!onWall) continue;
      const o = i * 4;
      px[o] = r;
      px[o + 1] = g;
      px[o + 2] = b;
      px[o + 3] = 255;
    }
  }
  ctx.putImageData(img, rx, ry);
}

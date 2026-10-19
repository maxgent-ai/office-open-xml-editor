/** Office profile control points: MS-OI29500 §2.1.1316 (ECMA §20.1.10.9).
 * Segments contain one, two or three points for line/quadratic/cubic; the
 * coordinates are the documented pseudocode values, unchanged.
 *
 * Axis assignment (observed Office behavior, not the §2.1.1316 prose). The
 * prose calls the first coordinate the height and the second the inward
 * distance. Read that way, `circle` is a concave cove that is flat at the
 * silhouette and vertical at the inner band edge, which paints a dark crease
 * inside the rim and leaves the rim itself unlit. Windows PowerPoint front-view
 * controls (all 12 presets on rectangles, bevel w = h from 3 to 23.4 pt, matte
 * threePt, all four edges) instead place every documented feature at the FIRST
 * coordinate as inward distance and the SECOND as height: circle darkens
 * monotonically toward the silhouette (convex quarter round); cross creases at
 * 0.5 of the band width and its face starts at 0.6 (the first coordinate is not
 * stretched); artDeco creases at 1/3; slope's rising and falling flanks sit at
 * 0–0.125 and 0.75–0.875; softRound rises within the first tenth and descends
 * smoothly to its 1/3-height face. Each of those is mirrored or displaced under
 * the prose reading. So each point is (inward distance, height): the inward
 * axis is scaled 0..1 to 0..bevel width, and the height axis is normalized to
 * its maximum (the prose's "stretched to fully fill the bevel height"). A
 * profile ending at inward < 1 leaves the rest of the band as its flat face at
 * the profile's end height. No coordinate is fitted, reversed or rescaled.
 */
const CURVES: Readonly<Record<string, readonly (readonly number[])[]>> = {
  angle: [[1.0, 1.0]],
  artDeco: [
    [0.0, 0.184095, 0.149238, 0.333333, 0.333333, 0.333333],
    [0.333333, 0.701523, 0.63181, 1.0, 1.0, 1.0],
  ],
  circle: [[0.0, 0.55627, 0.44373, 1.0, 1.0, 1.0]],
  convex: [
    [0.0, 0.07082, 0.0297458, 0.1, 0.101416, 0.1],
    [0.501416, 0.1, 0.9, 0.7, 0.901416, 0.899999],
    [0.9, 0.97167, 0.93343, 1.0, 1.0, 1.0],
  ],
  coolSlant: [
    [0.0, 0.138122, 0.0, 0.2, 0.271356, 0.775535],
    [0.377654, 1.0, 0.519455, 1.0, 0.583589, 1.0],
  ],
  cross: [
    [0.0, 0.05563, 0.04437, 0.1, 0.1, 0.1],
    [0.4, 0.1],
    [0.45563, 0.1, 0.5, 0.14437, 0.5, 0.2],
    [0.5, 0.9],
    [0.5, 0.95563, 0.54437, 1.0, 0.6, 1.0],
  ],
  divot: [
    [0.0, 0.236604, 0.119276, 0.607024, 0.263046, 0.760235],
    [0.361098, 0.864726, 0.457934, 0.909567, 0.537806, 0.925082],
    [0.559245, 0.929246, 0.567625, 0.89793, 0.542066, 0.845567],
    [0.484073, 0.726757, 0.477103, 0.393693, 0.551651, 0.393693],
    [0.899894, 0.393693],
    [0.99148, 0.393693, 0.958466, 0.746659, 0.907348, 0.779629],
    [0.879394, 0.797658],
    [0.845545, 0.819489, 0.848775, 0.954172, 0.874334, 0.971627],
    [0.915883, 1.0, 1.0, 1.0],
  ],
  hardEdge: [
    [0.0, 0.092437, 0.042353, 0.305322],
    [0.170124, 0.947558],
    [0.177203, 0.983142, 0.2, 1.0, 0.268235, 0.998599],
    [0.614118, 0.998599],
    [0.647059, 0.998599, 0.656471, 0.987395, 0.663529, 0.969188],
    [0.687006, 0.908633, 0.802353, 0.822129, 1.0, 0.822129],
  ],
  relaxedInset: [
    [0.0, 0.367, 0.124605, 0.82, 0.507899, 1.0],
    [0.724115, 0.737, 0.790455, 0.64, 1.0, 0.64],
  ],
  riblet: [
    [0.0, 0.238519, 0.132047, 0.500741, 0.357567, 0.731852],
    [0.513167, 0.891311, 0.563798, 0.912593, 0.735905, 0.912593],
    [0.873887, 0.912593],
    [0.956973, 0.912593, 0.878338, 1.0, 1.0, 1.0],
  ],
  slope: [
    [0.0, 0.125, 0.025, 0.25],
    [0.125, 0.75],
    [0.175, 1.0, 0.25, 1.0, 0.375, 1.0],
    [0.5, 1.0],
    [0.625, 1.0, 0.7, 1.0, 0.75, 0.75],
    [0.875, 0.125],
    [0.9, 0.01, 0.98, 0.01, 1.0, 0.01],
  ],
  softRound: [
    [0.0, 0.4775, 0.0968736, 1.0, 0.156301, 1.0],
    [0.264179, 1.0, 0.376919, 0.33333, 1.0, 0.33333],
  ],
};

// A fixed 1024-interval visible-surface lookup bounds work independently of
// canvas size, bevel dimensions and document input. Cubics are subdivided into
// 256 line segments. Where the curve folds back over an inward distance (an
// overhang), select the greatest height along the front viewing ray. Beyond
// the final inward distance the band is the flat face at the end height.
// This is a front-view surface, not a full camera-dependent bevel tessellation.
const RESOLUTION = 1024;
interface Profile {
  height: Float64Array;
  slope: Float64Array;
}
function compile(segments: readonly (readonly number[])[]): Profile {
  const height = new Float64Array(RESOLUTION + 1).fill(-Infinity);
  const slope = new Float64Array(RESOLUTION + 1);
  // x: inward distance (band fraction); y: height. See the CURVES note.
  const xs = [0],
    ys = [0];
  let x0 = 0,
    y0 = 0;
  for (const segment of segments) {
    const degree = segment.length / 2;
    for (let step = 1; step <= 256; step++) {
      const t = step / 256,
        u = 1 - t;
      if (degree === 1) {
        xs.push(u * x0 + t * segment[0]);
        ys.push(u * y0 + t * segment[1]);
      } else if (degree === 2) {
        xs.push(u * u * x0 + 2 * u * t * segment[0] + t * t * segment[2]);
        ys.push(u * u * y0 + 2 * u * t * segment[1] + t * t * segment[3]);
      } else {
        xs.push(
          u * u * u * x0 +
            3 * u * u * t * segment[0] +
            3 * u * t * t * segment[2] +
            t * t * t * segment[4],
        );
        ys.push(
          u * u * u * y0 +
            3 * u * u * t * segment[1] +
            3 * u * t * t * segment[3] +
            t * t * t * segment[5],
        );
      }
    }
    x0 = segment[segment.length - 2];
    y0 = segment[segment.length - 1];
  }
  // Normalize to the curve's own peak height (not a control point's).
  const maxY = Math.max(...ys);
  for (let k = 1; k < xs.length; k++) {
    const px = xs[k - 1],
      x = xs[k],
      py = ys[k - 1] / maxY,
      dx = x - px,
      dy = ys[k] / maxY - py;
    // A vertical piece (dx = 0) is a crease between neighbouring samples.
    if (dx === 0) continue;
    const lo = Math.max(0, Math.ceil(Math.min(px, x) * RESOLUTION));
    const hi = Math.min(RESOLUTION, Math.floor(Math.max(px, x) * RESOLUTION));
    for (let i = lo; i <= hi; i++) {
      const z = py + ((i / RESOLUTION - px) * dy) / dx;
      if (z > height[i]) {
        height[i] = z;
        slope[i] = dy / dx;
      }
    }
  }
  // Beyond the final inward distance the band is the flat face.
  for (let i = 0; i <= RESOLUTION; i++) {
    if (i / RESOLUTION >= x0) {
      height[i] = y0 / maxY;
      slope[i] = 0;
    }
  }
  return { height, slope };
}
// Compiled on first use per preset. Compilation is deterministic, so a lazy
// table returns the same samples as an eager one; most documents use no bevel
// or one preset, and every realm (main thread and each render worker) would
// otherwise compile all presets while the module evaluates.
const PROFILES = new Map<string, Profile>();
function profileFor(name: string): Profile {
  const key = Object.hasOwn(CURVES, name) ? name : 'circle';
  let profile = PROFILES.get(key);
  if (profile === undefined) {
    profile = compile(CURVES[key]);
    PROFILES.set(key, profile);
  }
  return profile;
}
function sample(values: Float64Array, t: number): number {
  const u = Math.max(0, Math.min(1, t)) * RESOLUTION;
  const i = Math.min(RESOLUTION - 1, Math.floor(u));
  return values[i] + (values[i + 1] - values[i]) * (u - i);
}
export function profileHeight(name: string, t: number): number {
  return sample(profileFor(name).height, t);
}
export function profileSlope(name: string, t: number): number {
  return sample(profileFor(name).slope, t);
}

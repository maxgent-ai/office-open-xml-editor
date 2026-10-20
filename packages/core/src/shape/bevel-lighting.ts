import { rigDefinition } from './bevel-light-rigs';
import type { LightRigRot, Vec3 } from './bevel-shading';

export type PresetMaterial =
  | 'legacyMatte'
  | 'legacyPlastic'
  | 'legacyMetal'
  | 'legacyWireframe'
  | 'matte'
  | 'plastic'
  | 'metal'
  | 'warmMatte'
  | 'translucentPowder'
  | 'powder'
  | 'dkEdge'
  | 'softEdge'
  | 'clear'
  | 'flat'
  | 'softmetal';

interface Material {
  specularShape: number;
  specularWhite: number;
  diffuseShape: number;
  diffuseWhite: number;
  diffuseAlpha: number;
  ambientAlpha: number;
  emissive: number;
  power: number;
  blinn: boolean;
  diffuseFresnel: number;
  alphaFresnel: number;
}
function material(
  ss: number,
  sw: number,
  ds: number,
  dw: number,
  power: number,
  blinn = false,
  df = 0,
  af = 0,
  da = 1,
  aa = da,
  emissive = 0,
): Material {
  return {
    specularShape: ss,
    specularWhite: sw,
    diffuseShape: ds,
    diffuseWhite: dw,
    power,
    blinn,
    diffuseFresnel: df,
    alphaFresnel: af,
    diffuseAlpha: da,
    ambientAlpha: aa,
    emissive,
  };
}
/** MS-OI29500 §2.1.1334, including legacyMatte's Office Fresnel correction.
 * Color expressions are affine in the shape color; Brighten(...,1.4) is kept
 * unclamped, exactly as the documented function. Metal's colored specular is
 * already expressed by the color columns (no invented extra metal gain).
 * legacyWireframe has no surface properties (ECMA §20.1.10.50); its geometry
 * is rasterized by the bevel compositor instead of receiving surface lighting.
 */
const MATERIALS: Readonly<Record<PresetMaterial, Material>> = {
  legacyMatte: material(0, 0, 1, 0, 0, true),
  legacyPlastic: material(0, 1, 1, 0, 32, true),
  legacyMetal: material(1, 0, 2 / 3, 0, 32, true),
  legacyWireframe: material(0, 0, 0, 0, 0),
  matte: material(0, 0, 1, 0, 0),
  plastic: material(0, 0.6, 1, 0, 12),
  metal: material(-0.32, 1.32, 1, 0, 12, false, 4),
  warmMatte: material(0, 0.3, 1, 0, 8),
  translucentPowder: material(0, 0.3, 1, 0, 10, true, 2, -1, 0.7),
  powder: material(0, 0.3, 0.5, 0.5, 10, true, 2),
  dkEdge: material(0, 1, 1, 0, 35, false, -2),
  softEdge: material(0, 1, 1, 0, 35, false, 4, -10),
  clear: material(0, 0.6, 1, 0, 20, false, -8, 1, 0.1, 1),
  flat: material(0, 0.8, 0, 0, 50, false, -4, 0, 1, 1, 1),
  softmetal: material(0.5, 0.5, 1, 0, 8),
};
export function presetMaterial(name: string | undefined): PresetMaterial {
  return name && Object.hasOwn(MATERIALS, name)
    ? (name as PresetMaterial)
    : 'matte';
}
interface Light {
  x: number;
  y: number;
  z: number;
  hx: number;
  hy: number;
  hz: number;
  rgb: readonly number[];
  scale: number;
  offset: number;
  specular: boolean;
  diffuse: boolean;
}
export interface RigVector extends Vec3 {
  rigLights?: readonly Light[];
  rigAmbient?: number;
}
const DIRECTIONS: Readonly<Record<string, number>> = {
  t: 0,
  tr: 45,
  r: 90,
  br: 135,
  b: 180,
  bl: 225,
  l: 270,
  tl: 315,
};
/** MS-OI29500 §2.1.1274: M=[0,-1,0;-1,0,0;0,0,1] maps
 * the tabulated light-travel vectors into scene space. Negate for surface-to-
 * light vectors in the bitmap convention (+Y down, +Z toward the viewer).
 * ECMA §20.1.5.11 describes revolution about the latitude/longitude axis,
 * without a numeric matrix. In this basis our rig-to-scene transform applies
 * local Rz(rev), then Rx(lat), then Ry(lon) to the entire mapped rig.
 * This is the inverse of the scene camera's world-to-view rotation, rather
 * than the same transform with only its revolution sign reversed.
 * Office evidence: front-view matte threePt rectangle controls at single-axis
 * 0/90/180/270 degrees, rev=90, (30,60,45), (90,90,90), and dir=t/r/br support
 * the signs/order (face and rim). This does not establish pixel fidelity for
 * other materials, profiles, cameras or arbitrary rotation combinations.
 * MS-OI29500 §2.1.1274(b): explicit rot replaces dir, including all-zero rot.
 * Omitted rot with dir=t is identity too; explicit zero differs from omitted
 * rot for non-top directions, not from the top direction's identity.
 * Keep the whole rig on the returned structural Vec3 so existing callers of
 * lightDirFromRig → BevelInput preserve light identity without a signature change.
 */
export function rigVector(
  name: string,
  dir: string,
  rot?: LightRigRot,
): RigVector {
  const rig = rigDefinition(name);
  const rad = Math.PI / 180;
  const lat = (rot?.lat ?? 0) * rad,
    lon = (rot?.lon ?? 0) * rad;
  const rev =
    (rot ? rot.rev : Object.hasOwn(DIRECTIONS, dir) ? DIRECTIONS[dir] : 0) *
    rad;
  const cx = Math.cos(lat),
    sx = Math.sin(lat),
    cy = Math.cos(lon),
    sy = Math.sin(lon),
    cz = Math.cos(rev),
    sz = Math.sin(rev);
  const lights = rig.lights.map((l) => {
    const [dx, dy, dz] = l.direction;
    const m = Math.hypot(dx, dy, dz) || 1;
    // Negated documented M: surface-to-light, with +Z toward the viewer.
    const ax = dy / m,
      ay = dx / m,
      az = -dz / m;
    const bx = cz * ax - sz * ay,
      by = sz * ax + cz * ay;
    const y = cx * by - sx * az,
      bz = sx * by + cx * az;
    const x = cy * bx + sy * bz,
      z = -sy * bx + cy * bz;
    const hm = Math.hypot(x, y, z + 1) || 1;
    return {
      x,
      y,
      z,
      hx: x / hm,
      hy: y / hm,
      hz: (z + 1) / hm,
      rgb: l.rgb,
      scale: l.scale,
      offset: l.offset,
      specular: l.specular,
      diffuse: l.diffuse,
    };
  });
  return {
    x: lights[0].x,
    y: lights[0].y,
    z: lights[0].z,
    rigLights: lights,
    rigAmbient: rig.ambient,
  };
}
export interface Lighting {
  material: Material;
  lights: readonly Light[];
  ambient: number;
  /** Per-invocation scratch and precomputed flat-face terms, never global. */
  terms: Float64Array;
  flatTerms: Float64Array;
}
export function prepareLighting(name: PresetMaterial, vector: Vec3): Lighting {
  const rig = vector as RigVector;
  const hm = Math.hypot(vector.x, vector.y, vector.z + 1) || 1;
  const lighting: Lighting = {
    material: MATERIALS[name],
    ambient: rig.rigAmbient ?? 0,
    lights: rig.rigLights ?? [
      {
        ...vector,
        hx: vector.x / hm,
        hy: vector.y / hm,
        hz: (vector.z + 1) / hm,
        rgb: [1, 1, 1],
        scale: 1,
        offset: 0,
        specular: true,
        diffuse: true,
      },
    ],
    terms: new Float64Array(6),
    flatTerms: new Float64Array(6),
  };
  lightTerms(lighting.flatTerms, 0, 0, 1, lighting);
  return lighting;
}
/** All color channels share the same angular terms. Compute them once per
 * normal, and once per shape for the flat face (the common interior case).
 */
function lightTerms(
  out: Float64Array,
  nx: number,
  ny: number,
  nz: number,
  p: Lighting,
): void {
  const m = p.material;
  let dr = 0,
    dg = 0,
    db = 0,
    sr = 0,
    sg = 0,
    sb = 0;
  for (const l of p.lights) {
    const ndotl = nx * l.x + ny * l.y + nz * l.z;
    if (l.diffuse) {
      const diffuse = Math.max(0, ndotl * l.scale + l.offset);
      dr += l.rgb[0] * diffuse;
      dg += l.rgb[1] * diffuse;
      db += l.rgb[2] * diffuse;
    }
    if (
      l.specular &&
      ndotl > 0 &&
      (m.specularShape !== 0 || m.specularWhite !== 0)
    ) {
      const alignment = m.blinn
        ? nx * l.hx + ny * l.hy + nz * l.hz
        : 2 * ndotl * nz - l.z;
      const specular = Math.pow(Math.max(0, alignment), m.power);
      sr += l.rgb[0] * specular;
      sg += l.rgb[1] * specular;
      sb += l.rgb[2] * specular;
    }
  }
  out[0] = dr;
  out[1] = dg;
  out[2] = db;
  out[3] = sr;
  out[4] = sg;
  out[5] = sb;
}
/** ECMA §20.1.10.50 describes signed glancing-angle Fresnel effects but does
 * not prescribe their equation or transfer curve. Library reference shader:
 * signed power of (1-|N·V|), Lambert plus Phong/Blinn as selected by the table,
 * evaluated in the Canvas RGB domain. This is NOT a claim to reproduce Office's
 * undocumented tone mapping/Fresnel compositor. No measured gains or per-case
 * adjustments are used. Preserve signed light sums until the final clamp.
 * Output is straight RGBA; caller applies source silhouette coverage once.
 * Work is bounded by four lights and three channels; no per-pixel allocation.
 */
export function shadeSurface(
  out: Float64Array,
  r: number,
  g: number,
  b: number,
  a: number,
  nx: number,
  ny: number,
  nz: number,
  p: Lighting,
): void {
  const m = p.material,
    glance = 1 - Math.min(1, Math.abs(nz));
  const df =
    m.diffuseFresnel === 0 ? 0 : Math.pow(glance, Math.abs(m.diffuseFresnel));
  const af =
    m.alphaFresnel === 0 ? 0 : Math.pow(glance, Math.abs(m.alphaFresnel));
  let alpha = m.diffuseAlpha;
  if (m.alphaFresnel > 0) alpha += (1 - alpha) * af;
  else if (m.alphaFresnel < 0) alpha *= 1 - af;
  const terms = nx === 0 && ny === 0 && nz === 1 ? p.flatTerms : p.terms;
  if (terms === p.terms) lightTerms(terms, nx, ny, nz, p);
  for (let c = 0; c < 3; c++) {
    const shape = (c === 0 ? r : c === 1 ? g : b) / 255;
    let diffuse = shape * m.diffuseShape + m.diffuseWhite;
    diffuse =
      m.diffuseFresnel > 0 ? diffuse + (1 - diffuse) * df : diffuse * (1 - df);
    const specular = shape * m.specularShape + m.specularWhite;
    const value =
      shape * m.emissive +
      (shape * m.diffuseShape + m.diffuseWhite) * p.ambient +
      diffuse * terms[c] +
      specular * terms[c + 3];
    out[c] = Math.max(0, Math.min(255, value * 255));
  }
  // A Canvas bitmap has one opacity channel. Reference surface opacity uses
  // diffuse alpha; the distinct ambient alpha retained in the table (clear) is
  // not an independently composited layer. Office's layer combination remains
  // unresolved; do not infer it from white-background controls alone.
  out[3] = a * Math.max(0, Math.min(1, alpha));
}

import { ownValue } from '../internal/own-key';
/**
 * DrawingML 3D camera — perspective projection of a planar shape (Phase A).
 *
 * ECMA-376 references:
 *   - §20.1.5.5  `camera` (CT_Camera): prst / fov / zoom / rot.
 *   - §20.1.5.11 `rot`    (CT_SphereCoords): lat / lon / rev rotation.
 *   - §20.1.10.47 ST_PresetCameraType: the 62 preset cameras.
 *
 * Fidelity boundary: ECMA-376 names presets and authored FOV/zoom/rotation,
 * but does not supply a numerical preset table or a complete projection matrix.
 * MS-OI29500 §2.1.1270 says non-perspective cameras ignore FOV. This module's
 * modern preset table and projection are bounded Office observations, not
 * normative definitions. Legacy presets retain their previous parallel fallback.
 * See MODERN_SCENE_INCHES for the evidence and scene-scale limitation.
 *
 * Phase A scope: we model the shape as a flat z=0 rectangle (the picture / 2D
 * drawing) and compute the 2D homography that maps the un-projected rectangle
 * to the projected quad. Bevel / extrusion / lightRig shading is Phase B.
 */

/**
 * Geometry-only camera inputs the projection math needs. Defined locally so the
 * core helper has no dependency on the pptx (or any format) type layer; the
 * pptx `Camera3d` / `Rot3d` are structurally assignable to these. Angles are in
 * degrees (the parser already converted from 60000ths).
 */
export interface CameraInput {
  /**
   * Projection coordinate units per document inch; renderer context, not an
   * OOXML camera attribute. Defaults to 96 for CSS pixels. Use 72 for points,
   * or multiply by the raster scale for device-space extrusion. Shape/group
   * resizing changes geometry, not this document-to-render conversion.
   */
  unitsPerInch?: number;
  /**
   * Extent of the rendered solid along object −Z, in projection units. Bind
   * once for all front/back/effect calls; scale with unitsPerInch for device
   * coordinates. This is renderer context, not an OOXML camera attribute.
   */
  sceneDepth?: number;
  /** Preset camera name (`ST_PresetCameraType`). */
  prst: string;
  /** Field-of-view override in degrees. Omitted = preset default. */
  fov?: number;
  /** Zoom factor as a unit ratio (1.0 = 100%). Omitted = 1.0. */
  zoom?: number;
  /** Camera rotation override. Omitted = preset base orientation. */
  rot?: RotInput;
}

/** Sphere-coordinate rotation in degrees — ECMA-376 §20.1.5.11. */
export interface RotInput {
  lat: number;
  lon: number;
  rev: number;
}

/** A 2D point in the projected (screen) plane, pixels relative to bbox origin. */
export interface Vec2 {
  x: number;
  y: number;
}

/** Result of projecting a planar shape through the camera. */
export interface Scene3dQuad {
  /**
   * The four projected corners in CSS pixels, relative to the element's
   * bounding-box top-left, in order: top-left, top-right, bottom-right,
   * bottom-left (matching the source rectangle (0,0)-(w,0)-(w,h)-(0,h)).
   *
   * Projection is about the shape center and can extend beyond its original
   * box; zoom is retained. Callers must allow the projected bitmap to overhang.
   */
  corners: [Vec2, Vec2, Vec2, Vec2];
  /**
   * True when the projection is (numerically) an affine map — i.e. the four
   * corners still form a parallelogram. Orthographic / isometric presets and a
   * pure in-plane `rev` are affine; perspective presets with lat/lon tilt are
   * not. Callers can take a cheaper single-`setTransform` path when affine.
   */
  isAffine: boolean;
  /**
   * True when the projection is (numerically) the identity within the bbox —
   * the renderer should skip projection entirely and draw normally. Happens for
   * `orthographicFront` with no rot override (or rot all-zero).
   */
  isIdentity: boolean;
}

/** Preset camera kinds that determine the projection math. */
type ProjectionKind = 'perspective' | 'orthographic';

/**
 * Per-preset definition: the base camera orientation (replaced by an authored
 * `rot` override) and the fallback field of view in degrees for perspective
 * presets. Orthographic / isometric presets use parallel projection (fov is
 * ignored). Modern oblique presets additionally shear the rotated depth.
 *
 * Base orientation is expressed as the same lat/lon/rev sphere coordinates the
 * file uses, in degrees; authored rot replaces it.
 */
interface PresetDef {
  kind: ProjectionKind;
  /** Base latitude (X-axis tilt), degrees. */
  baseLat: number;
  /** Base longitude (Y-axis turn), degrees. */
  baseLon: number;
  /** Base revolution (Z-axis spin), degrees. */
  baseRev: number;
  /** Default field of view for perspective presets, degrees. */
  fovDeg: number;
  /** Physical scene framing for modern perspective; absent = old fallback. */
  sceneInches?: number;
  /** Parallel depth shear in screen coordinates, after rotation. */
  shearX: number;
  shearY: number;
}

// Observed Office behavior (Windows PowerPoint 16.0, synthetic flat rectangles):
// all 44 modern ST_PresetCameraType values (20 isometric, 8 oblique, 15
// perspective, orthographicFront) were checked with omitted/zero/(25,35,17)
// rotations. All 15 perspective names also cover FOV 5/15/26/45/75/115 degrees,
// rotations (0,30,0)/(30,0,0)/(25,35,17), aspects 0.7/1.5/2.5; zoom controls
// cover 0.5/1/1.5. Max modern corner residuals by preset/lens/parallel-FOV/zoom
// stratum: 0.688/0.556/0.496/0.503pt (0.75pt PDF extraction uncertainty).
// The table's orientation hypothesis was independently checked against those
// Office controls; the numeric angles are observations, not ECMA defaults.
//
// Windows PowerPoint 16.0 controls additionally vary width 47.88..383.04pt,
// aspect 0.7/2.5, shape-image-fill/picture, 2pt top bevel, 6pt extrusion,
// xfrm rotation 0/30 and flipH, all 23 modern perspective/oblique presets,
// wrapped lat/lon/rev, and identity/uniform/nonuniform/nested reflected groups.
// Together with the lens packet above they support a FIXED physical numerator:
// d = 6.28 * unitsPerInch / tan(FOV/2), independent of shape size/aspect/depth.
// Microsoft Open Specifications support question 769795 (2022-03-23) gives
// the 6.28 scene-unit formula; the inch mapping is an Office observation from
// these varied-size controls, not an ECMA definition or a universal scene rule.
// https://learn.microsoft.com/answers/questions/769795/
// Group ext/chExt scale the leaf dimensions before projection (Annex L.4.3);
// composed flips act before the camera and composed 2-D rotation after it.
// Do not scale the physical distance with group/shape width or raster bounds.
// All 1,005 original and 77,648 added modern controls fit a 2pt maximum corner
// tolerance. Thin compressed PDF fiducials require higher-resolution edge
// extraction; this tolerance is for measured front-plane correspondences.
// Interior fiducials measure the z=0 texture plane, not bevel/wall silhouettes.
// Omitted lenses are 45 degrees (80 for both heroic-extreme presets). Near-plane
// nudging, larger depths and other group/angle combinations remain unvalidated
// Office fidelity; cameraDistance below supplies a non-crossing library fallback.
const MODERN_SCENE_INCHES = 6.28;

// Legacy ST_PresetCameraType names intentionally keep the previous parallel
// fallback, including their authored rotations/FOV handling; no legacy framing
// candidate is adopted. In the same 1,572-control investigation, an off-axis
// perspective candidate still missed by 41.425pt; 73 cases exceeded 2pt, all
// legacy perspective. The 3.037pt preset and 4.232pt zoom maxima were legacy,
// not modern failures. A free scale per case hid the error but supplied no rule.
// Audit C camera-round2-residuals.json (local investigation evidence) retains
// every control and the per-case/preset/stratum residual table. Legacy oblique
// shear and legacy perspective framing remain explicit fidelity limitations.
// Retained legacy production maxima by preset/lens/zoom: 24.023/65.689/35.969pt;
// these are unresolved errors, not accepted tolerances.
// Keep the old 26-degree lens only for the unrecognized perspective* fallback;
// no measured modern preset uses it. These fallbacks are library policy.
const LEGACY_PERSP_FOV = 26;

function parallel(lat = 0, lon = 0, rev = 0, shearX = 0, shearY = 0): PresetDef {
  return { kind: 'orthographic', baseLat: lat, baseLon: lon, baseRev: rev, fovDeg: 0, shearX, shearY };
}

function perspective(lat: number, lon: number, rev: number, fovDeg = 45): PresetDef {
  return {
    kind: 'perspective', baseLat: lat, baseLon: lon, baseRev: rev, fovDeg,
    sceneInches: MODERN_SCENE_INCHES, shearX: 0, shearY: 0,
  };
}

/**
 * ECMA-376 §20.1.10.47 names these modern presets; numerical definitions below
 * are the bounded Office observations above. +X right, +Y down, +Z toward viewer.
 * Explicit rot replaces these angles, retaining the preset's projection/shear.
 */
const PRESETS: Record<string, PresetDef> = {
  isometricBottomDown: parallel(35.4, 314.7, 299.8),
  isometricBottomUp: parallel(35.4, 45.3, 60.2),
  isometricLeftDown: parallel(35, 45, 0),
  isometricLeftUp: parallel(325, 45, 0),
  isometricOffAxis1Left: parallel(18, 64, 0),
  isometricOffAxis1Right: parallel(18, 334, 0),
  isometricOffAxis1Top: parallel(301.3, 306.5, 57.6),
  isometricOffAxis2Left: parallel(18, 26, 0),
  isometricOffAxis2Right: parallel(18, 296, 0),
  isometricOffAxis2Top: parallel(301.3, 53.5, 302.4),
  isometricOffAxis3Bottom: parallel(58.7, 306.5, 302.4),
  isometricOffAxis3Left: parallel(342, 64, 0),
  isometricOffAxis3Right: parallel(342, 334, 0),
  isometricOffAxis4Bottom: parallel(58.7, 53.5, 57.6),
  isometricOffAxis4Left: parallel(342, 26, 0),
  isometricOffAxis4Right: parallel(342, 296, 0),
  isometricRightDown: parallel(325, 315, 0),
  isometricRightUp: parallel(35, 315, 0),
  isometricTopDown: parallel(324.6, 45.3, 299.8),
  isometricTopUp: parallel(324.6, 314.7, 60.2),
  obliqueBottom: parallel(0, 0, 0, 0, -0.3),
  obliqueBottomLeft: parallel(0, 0, 0, 0.3 / Math.sqrt(2), -0.3 / Math.sqrt(2)),
  obliqueBottomRight: parallel(0, 0, 0, -0.3 / Math.sqrt(2), -0.3 / Math.sqrt(2)),
  obliqueLeft: parallel(0, 0, 0, 0.3, 0),
  obliqueRight: parallel(0, 0, 0, -0.3, 0),
  obliqueTop: parallel(0, 0, 0, 0, 0.3),
  obliqueTopLeft: parallel(0, 0, 0, 0.3 / Math.sqrt(2), 0.3 / Math.sqrt(2)),
  obliqueTopRight: parallel(0, 0, 0, -0.3 / Math.sqrt(2), 0.3 / Math.sqrt(2)),
  orthographicFront: parallel(0, 0, 0),
  perspectiveAbove: perspective(340, 0, 0),
  perspectiveAboveLeftFacing: perspective(39.3, 14.3, 341.1),
  perspectiveAboveRightFacing: perspective(39.3, 345.7, 18.9),
  perspectiveBelow: perspective(20, 0, 0),
  perspectiveContrastingLeftFacing: perspective(10.4, 43.9, 356.4),
  perspectiveContrastingRightFacing: perspective(10.4, 316.1, 3.6),
  perspectiveFront: perspective(0, 0, 0),
  perspectiveHeroicExtremeLeftFacing: perspective(8.1, 34.5, 357.1, 80),
  perspectiveHeroicExtremeRightFacing: perspective(8.1, 325.5, 2.9, 80),
  perspectiveHeroicLeftFacing: perspective(349, 14.3, 2.6),
  perspectiveHeroicRightFacing: perspective(349, 345.7, 357.4),
  perspectiveLeft: perspective(0, 20, 0),
  perspectiveRelaxed: perspective(309.6, 0, 0),
  perspectiveRelaxedModerately: perspective(324.8, 0, 0),
  perspectiveRight: perspective(0, 340, 0),
};

const UNKNOWN_PERSPECTIVE: PresetDef = {
  ...perspective(0, 0, 0, LEGACY_PERSP_FOV), sceneInches: undefined,
};

/** 3×3 matrix in row-major order. */
type Mat3 = [number, number, number, number, number, number, number, number, number];

function mul3(a: Mat3, b: Mat3): Mat3 {
  const r = new Array(9).fill(0) as number[];
  for (let i = 0; i < 3; i++) {
    for (let j = 0; j < 3; j++) {
      let s = 0;
      for (let k = 0; k < 3; k++) s += a[i * 3 + k] * b[k * 3 + j];
      r[i * 3 + j] = s;
    }
  }
  return r as Mat3;
}

function rotX(deg: number): Mat3 {
  const t = (deg * Math.PI) / 180;
  const c = Math.cos(t);
  const s = Math.sin(t);
  // X-axis rotation (latitude tilt).
  return [1, 0, 0, 0, c, -s, 0, s, c];
}

function rotY(deg: number): Mat3 {
  const t = (deg * Math.PI) / 180;
  const c = Math.cos(t);
  const s = Math.sin(t);
  // Y-axis rotation (longitude turn).
  return [c, 0, s, 0, 1, 0, -s, 0, c];
}

function rotZ(deg: number): Mat3 {
  const t = (deg * Math.PI) / 180;
  const c = Math.cos(t);
  const s = Math.sin(t);
  // Z-axis rotation (revolution / in-plane spin).
  return [c, -s, 0, s, c, 0, 0, 0, 1];
}

function applyMat3(m: Mat3, x: number, y: number, z: number): [number, number, number] {
  return [
    m[0] * x + m[1] * y + m[2] * z,
    m[3] * x + m[4] * y + m[5] * z,
    m[6] * x + m[7] * y + m[8] * z,
  ];
}

/**
 * Object rotation Rz(-rev) · Rx(-lat) · Ry(-lon), with explicit rot replacing
 * the modeled preset orientation. ECMA-376 §§20.1.5.5/.11 describe overrides;
 * the modern-camera controls above support this matrix for their tested angles.
 * It is not a complete legacy camera model (legacy perspective ignores rev).
 */
function buildRotation(preset: PresetDef, rot: RotInput | undefined): Mat3 {
  const lat = rot ? rot.lat : preset.baseLat;
  const lon = rot ? rot.lon : preset.baseLon;
  const rev = rot ? rot.rev : preset.baseRev;
  return mul3(rotZ(-rev), mul3(rotX(-lat), rotY(-lon)));
}

/** Look up a preset while retaining the projection family encoded by its
 * standard name. Unknown perspective names retain the old neutral lens/framing;
 * other unknown and legacy names retain the old parallel fallback. */
function presetDef(prst: string): PresetDef {
  const known = ownValue(PRESETS, prst);
  if (known) return known;
  if (prst.startsWith('perspective')) return UNKNOWN_PERSPECTIVE;
  return PRESETS.orthographicFront;
}

/** Shared by the face and extrusion so their coordinate scales agree. */
function cameraDistance(
  camera: CameraInput, def: PresetDef, rotation: Mat3,
  w: number, h: number, surfaceDepth: number,
): number {
  const fovDeg = camera.fov ?? def.fovDeg;
  // Retain the finite-lens guard; endpoint/near-plane fidelity is unvalidated.
  const fov = (Math.max(1, Math.min(179, fovDeg)) * Math.PI) / 180;
  const numerator = def.sceneInches === undefined
    ? Math.max(w, h) / 2 : def.sceneInches * (camera.unitsPerInch ?? 96);
  const distance = numerator / Math.tan(fov / 2);
  // ECMA-376 §20.1.5.5 does not define camera near-plane handling. Microsoft
  // Open Specifications support question 769795 describes moving the camera
  // behind the scene with a gap of 2% of its projection-space z-volume.
  // Windows PowerPoint 16.0 front-plane markers support the unchanged lens
  // for flat 4in shape/picture and a 12in shape with 8in receding extrusion
  // (lon60/FOV90, no bevel); raised-bevel probes remain unresolved.
  // A 20in crossing shape and a shared-camera group FAIL Office PDF export;
  // they establish no numerical nudge, framing or scene-ownership rule.
  //
  // Outside that measured domain this is LIBRARY POLICY, not an Office fit:
  // use the rendered leaf solid's rotated bounds, move only along view +Z,
  // and use the moved distance in the numerator too (retain z=0 framing).
  // Do not infer shared-group scene bounds from flattened leaf geometry.
  // Bevel shading is baked onto these planes, not independent raised geometry.
  const depth = Math.max(0, camera.sceneDepth ?? 0, surfaceDepth);
  const halfRange = Math.abs(rotation[6]) * w / 2 + Math.abs(rotation[7]) * h / 2;
  const rearZ = -rotation[8] * depth;
  const nearest = halfRange + Math.max(0, rearZ);
  const volume = 2 * halfRange + Math.abs(rearZ);
  // Mathematical guard: the solid contains z=0, hence volume >= nearest >= 0.
  // If nearest is positive the 2% gap dominates roundoff; if it is zero the
  // positive lens distance already clears every point. Thus 1-rz/d >= 1/51
  // throughout the solid (perspective magnification <= 51 before zoom).
  // One scene distance preserves a convex projective map instead of clamping
  // corners separately.
  return Math.max(distance, nearest + .02 * volume);
}

/**
 * Project a planar shape (the w×h rectangle) through the camera and return the
 * four projected corners in shape-local coordinates. A tilted camera can
 * legitimately project outside the original 2-D box; forcing it back into
 * that box would cancel camera zoom and shrink the authored 3-D scene.
 *
 * @param camera  parsed <a:camera> (prst / fov / zoom / rot, angles in degrees).
 * @param w       shape width in CSS pixels.
 * @param h       shape height in CSS pixels.
 * @param surfaceDepth back surface distance along object −Z, in the same units.
 *                MS-OI29500 §2.1.1268 places bevelB at the extrusion base.
 */
export function computeScene3dQuad(camera: CameraInput, w: number, h: number, surfaceDepth = 0): Scene3dQuad {
  const def = presetDef(camera.prst);
  const R = buildRotation(def, camera.rot);

  // Degenerate sizes can't be projected meaningfully.
  if (w <= 0 || h <= 0) {
    return {
      corners: [
        { x: 0, y: 0 },
        { x: w, y: 0 },
        { x: w, y: h },
        { x: 0, y: h },
      ],
      isAffine: true,
      isIdentity: true,
    };
  }

  // Source corners centred at the origin, in the caller's coordinate units.
  const hw = w / 2;
  const hh = h / 2;
  const srcCorners: Array<[number, number]> = [
    [-hw, -hh], // top-left
    [hw, -hh], // top-right
    [hw, hh], // bottom-right
    [-hw, hh], // bottom-left
  ];

  const zoom = camera.zoom ?? 1;

  let projected: Array<[number, number]>;

  if (def.kind === 'perspective') {
    const d = cameraDistance(camera, def, R, w, h, surfaceDepth);
    projected = srcCorners.map(([sx, sy]) => {
      const [rx, ry, rz] = applyMat3(R, sx, sy, -surfaceDepth);
      // Camera at +Z = d looking toward -Z (origin). A point at object-z rz is
      // at camera-relative depth (d - rz); perspective divide by that depth.
      const f = d / (d - rz);
      return [rx * f, ry * f] as [number, number];
    });
  } else {
    // Parallel projection, including the oblique preset's directional depth shear.
    projected = srcCorners.map(([sx, sy]) => {
      const [rx, ry, rz] = applyMat3(R, sx, sy, -surfaceDepth);
      return [rx + rz * def.shearX, ry + rz * def.shearY] as [number, number];
    });
  }

  // Apply zoom about the centre.
  projected = projected.map(([px, py]) => [px * zoom, py * zoom] as [number, number]);

  const corners = projected.map(([px, py]) => ({
    x: w / 2 + px,
    y: h / 2 + py,
  })) as [Vec2, Vec2, Vec2, Vec2];

  // Affine test: a quad is a parallelogram iff (C0 + C2) == (C1 + C3) (the
  // diagonals bisect). Use a relative epsilon vs. the bbox size.
  const eps = 1e-3 * Math.max(w, h);
  const sumDiagX = corners[0].x + corners[2].x - (corners[1].x + corners[3].x);
  const sumDiagY = corners[0].y + corners[2].y - (corners[1].y + corners[3].y);
  const isAffine = Math.abs(sumDiagX) < eps && Math.abs(sumDiagY) < eps;

  // Identity test: corners within epsilon of the un-projected rectangle.
  const ref: Array<[number, number]> = [
    [0, 0],
    [w, 0],
    [w, h],
    [0, h],
  ];
  let isIdentity = true;
  for (let i = 0; i < 4; i++) {
    if (Math.abs(corners[i].x - ref[i][0]) > eps || Math.abs(corners[i].y - ref[i][1]) > eps) {
      isIdentity = false;
      break;
    }
  }

  return { corners, isAffine, isIdentity };
}

/**
 * True when a scene3d camera will visibly change the shape, i.e. the renderer
 * should take the projection path. A camera whose preset is unknown/front-facing
 * with no rot (or an all-zero rot) can be skipped when its depth geometry adds
 * no visible footprint. `depth` uses the same coordinate units as w and h.
 */
export function isScene3dNonIdentity(camera: CameraInput, w?: number, h?: number, depth = 0): boolean {
  // Perspective framing depends on aspect, so a square probe can miss a visible
  // projection on a tall shape. Renderers supply their actual dimensions.
  if (w !== undefined && h !== undefined) {
    if (w <= 0 || h <= 0) return false;
    // The front plane alone is insufficient: default oblique cameras preserve
    // it but shear the extrusion base (§20.1.5.12). Test both ends of the solid
    // with the same geometry tolerance. A modern face-on perspective camera
    // contracts the rear plane inside the front: that hidden surface does not
    // expose a wall and must not introduce another front-face resampling.
    // The planar compositor does not draw independent translucent rear faces.
    camera = { ...camera, sceneDepth: Math.max(camera.sceneDepth ?? 0, depth) };
    if (!computeScene3dQuad(camera, w, h).isIdentity) return true;
    if (depth <= 0) return false;
    const back = computeScene3dQuad(camera, w, h, depth);
    if (presetDef(camera.prst).sceneInches === undefined) return !back.isIdentity;
    const eps = 1e-3 * Math.max(w, h);
    return back.corners.some(p => p.x < -eps || p.y < -eps || p.x > w + eps || p.y > h + eps);
  }
  // Geometry-free callers use a conservative mathematical identity test instead
  // of borrowing a shape's numerical pixel tolerance.
  const def = presetDef(camera.prst), R = buildRotation(def, camera.rot);
  const terms = def.kind === 'perspective'
    ? [R[0] - 1, R[1], R[3], R[4] - 1, R[6], R[7]]
    : [R[0] + def.shearX * R[6] - 1, R[1] + def.shearX * R[7],
       R[3] + def.shearY * R[6], R[4] + def.shearY * R[7] - 1];
  if (depth > 0) {
    if (def.kind === 'perspective') return true;
    terms.push(R[2] + def.shearX * R[8], R[5] + def.shearY * R[8]);
  }
  return (camera.zoom ?? 1) !== 1 || terms.some(value => Math.abs(value) > 1e-12);
}

/** Internal surface frame for the bitmap bevel compositor. Back-face selection
 * uses the camera's viewing direction (including parallel oblique shear), not
 * a rotation-angle threshold. ECMA §§20.1.5.3–4 assign separate back/front faces.
 */
export function scene3dSurfaceFrame(camera: CameraInput) {
  const def = presetDef(camera.prst), rotation = buildRotation(def, camera.rot);
  const facing = rotation[8] - def.shearX * rotation[2] - def.shearY * rotation[5];
  return { rotation, backFacing: facing < 0 };
}

/**
 * Screen-space displacement (in shape-local px, +X right / +Y down) of pushing a
 * point at the shape centre by `depthPx` along the object's −Z axis (i.e. AWAY
 * from the viewer — the direction an extrusion's back face sits). Used by the
 * extrusion renderer to offset the swept side-wall band (§20.1.5.12 extrusionH).
 *
 * For a face-on camera (orthographicFront / perspectiveFront with no rot) the
 * −Z axis projects straight back into the screen, so the screen displacement is
 * ~0 and the side walls are (correctly) invisible. A tilted/rotated camera turns
 * −Z partly into the screen plane, revealing the side wall as this offset.
 *
 * This is the LINEARISED offset of the box centre; for a perspective camera the
 * exact side-wall outline is a frustum sweep that varies per silhouette point.
 * Phase B uses this single centre offset as the documented approximation (good
 * for small extrusions / gentle tilts; see the renderer's extrusion note).
 */
export function computeDepthOffset(camera: CameraInput, w: number, h: number, depthPx: number): Vec2 {
  const def = presetDef(camera.prst);
  const R = buildRotation(def, camera.rot);
  if (w <= 0 || h <= 0 || depthPx === 0) return { x: 0, y: 0 };

  const zoom = camera.zoom ?? 1;
  const distance = def.kind === 'perspective'
    ? cameraDistance(camera, def, R, w, h, depthPx) : 0;

  // Project the centre at object-z 0 and at object-z −depthPx, take the screen
  // delta. Use the same lens, distance and parallel shear as the front face.
  // unitsPerInch must use the same render scale as w/h/depth, keeping the
  // device-space extrusion consistent with the CSS-space face. This remains
  // a centre approximation to the wall sweep.
  const project = (oz: number): [number, number] => {
    const [rx, ry, rz] = applyMat3(R, 0, 0, oz);
    if (def.kind === 'perspective') {
      const f = distance / (distance - rz);
      return [rx * f * zoom, ry * f * zoom];
    }
    return [(rx + rz * def.shearX) * zoom, (ry + rz * def.shearY) * zoom];
  };
  const [x0, y0] = project(0);
  const [x1, y1] = project(-depthPx);
  return { x: x1 - x0, y: y1 - y0 };
}

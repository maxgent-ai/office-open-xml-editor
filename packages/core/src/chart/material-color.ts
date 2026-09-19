import type { ChartThreeD } from '../types/chart';

/** Effective omitted-view camera isolated by the S1-S5 Office boundary set. */
export function isObservedAutomaticSurfaceCamera(view: ChartThreeD): boolean {
  return view.rotationX === 15
    && view.rotationY === 20
    && view.rightAngleAxes === false
    && view.perspective === 30;
}

/**
 * Office's omitted Surface view uses a wider perspective response than the
 * DrawingML angle alone produces in the shared camera. That compatibility
 * gain is observed only for the effective S1-S5 camera; authored cameras keep
 * the specification-derived projection until their own boundaries exist.
 */
export function surfacePerspectiveTangentGain(view: ChartThreeD): number {
  return isObservedAutomaticSurfaceCamera(view) ? 2 : 1;
}

/** Multiply an sRGB chart material by a bounded diffuse-light factor. */
export function scaleHexColor(color: string, factor: number): string {
  const value = color.replace(/^#/, '');
  if (!/^[0-9a-f]{6}$/i.test(value) || !Number.isFinite(factor)) return color;
  const bounded = Math.max(0, factor);
  const channel = (offset: number) => Math.max(0, Math.min(255,
    Math.round(Number.parseInt(value.slice(offset, offset + 2), 16) * bounded),
  )).toString(16).padStart(2, '0');
  return `#${channel(0)}${channel(2)}${channel(4)}`.toUpperCase();
}

const srgbToLinear = (channel: number): number => channel <= 0.04045
  ? channel / 12.92
  : ((channel + 0.055) / 1.055) ** 2.4;

const linearToSrgb = (channel: number): number => channel <= 0.0031308
  ? channel * 12.92
  : 1.055 * channel ** (1 / 2.4) - 0.055;

/** Apply the generated DrawingML shade/tint used by classic Excel chart
 * palettes. XLSX uses the linear-sRGB transform also used by DrawingML theme
 * colours; this mirrors ooxml-common's node-free generated-colour path. */
export function applyLinearTintOrShade(color: string, amount: number): string {
  const value = color.replace(/^#/, '');
  if (!/^[0-9a-f]{6}$/i.test(value) || !Number.isFinite(amount)) return color;
  const bounded = Math.max(-1, Math.min(1, amount));
  const channel = (offset: number): string => {
    const encoded = Number.parseInt(value.slice(offset, offset + 2), 16) / 255;
    const linear = srgbToLinear(encoded);
    const transformed = bounded < 0
      ? linear * (1 + bounded)
      : linear * (1 - bounded) + bounded;
    return Math.round(Math.max(0, Math.min(1, linearToSrgb(transformed))) * 255)
      .toString(16).padStart(2, '0');
  };
  return `#${channel(0)}${channel(2)}${channel(4)}`.toUpperCase();
}

/** Resolve one generated colour from legacy six-accent Pattern 2.
 *
 * ECMA-376 Part 1 §21.2.3.46 Tables 5-6 define the accent cycle and require
 * repeated sets to vary their tint/shade. The application-defined endpoints
 * are the already-registered 1..48-point Office boundary rule used by the
 * shared parser; surface bands need the same rule at paint time because their
 * count depends on the final automatic value-axis plan. */
export function legacyPattern2Color(
  accents: readonly string[],
  objectIndex: number,
  objectCount: number,
  chartStyle: number | null | undefined,
): string | null {
  const accentCount = 6;
  if (accents.length < accentCount || objectIndex < 0 || objectCount <= 0) return null;
  const base = accents[objectIndex % accentCount];
  if (!base) return null;
  if (![2, 10, 18, 26, 34, 42].includes(chartStyle ?? -1)) return `#${base}`.toUpperCase();
  const setIndex = Math.floor(objectIndex / accentCount);
  // Excel's complete 48-object boundary matrix fixes eight transforms by set;
  // 6/7/12/13-object controls establish that an existing object's transform is
  // independent of the final object count. Index 48 starts an unobserved ninth
  // set and therefore falls through instead of extrapolating a transform.
  const amount = [0, -0.4, 0.2, -0.2, 0.4, -0.5, 0.3, -0.3][setIndex];
  if (amount == null) return null;
  return applyLinearTintOrShade(base, amount);
}

/** Surface-only automatic material response.
 *
 * ECMA-376 carries the view and source mesh but leaves the automatic material
 * to the application. The S1-S5/saddle/90° contour boundary set isolates a
 * single camera-space directional response: opposing source-grid triangles
 * may darken or brighten, while authored band paint remains the base colour.
 * Keep that compatibility rule here (and out of general fills/3-D solids).
 */
export function surfaceMaterialFactor(normal: { x: number; y: number; z: number } | null): number {
  if (!normal) return 1;
  const facing = normal.z < 0
    ? { x: -normal.x, y: -normal.y, z: -normal.z }
    : normal;
  // The observed Office surface material is lit from screen upper-right.
  // Shared camera projection maps +x rightward and +y upward on screen.
  const light = { x: 0.24, y: 0.42, z: 0.88 };
  const length = Math.hypot(light.x, light.y, light.z);
  const lambert = Math.max(0,
    (facing.x * light.x + facing.y * light.y + facing.z * light.z) / length,
  );
  return Math.max(0.48, Math.min(1.22, 0.48 + 0.78 * lambert));
}

import { assertChartImageWork, rejectChartImageGeometry } from './image-fill-errors.js';
import type { ImageFill } from '../types/common.js';
import type { ChartThreeDSurface } from '../types/chart.js';
import { MAX_CHART_IMAGE_FILL_TILES } from './resource-limits.js';

export type ChartThreeDSurfaceKind = 'floor' | 'sideWall' | 'backWall';
export interface SurfacePicturePoint { x: number; y: number }
export type SurfacePictureQuad = [
  SurfacePicturePoint,
  SurfacePicturePoint,
  SurfacePicturePoint,
  SurfacePicturePoint,
];

export interface SurfacePicturePlan {
  mode: 'stretch' | 'stack' | 'stackScale' | 'tile';
  repetitions: number;
  stackUnit?: number;
  slabFaces?: {
    front: boolean;
    sides: boolean;
    end: boolean;
  };
}

/** CT_Surface slab faces are emitted as inner, outer, then the four joining
 * faces. ECMA-376 §21.2.2.1-.3 names the visible picture targets front, sides,
 * and end. The observed 25% wall/floor boundary maps the inner face to front,
 * alternating joining faces to end/sides, and leaves the hidden outer face
 * unpainted. */
export function surfacePictureFaceIsEnabled(
  plan: SurfacePicturePlan,
  faceIndex: number,
): boolean {
  if (!Number.isSafeInteger(faceIndex) || faceIndex < 0) return false;
  if (!plan.slabFaces) return faceIndex === 0;
  if (faceIndex === 0) return plan.slabFaces.front;
  if (faceIndex === 1 || faceIndex >= 6) return false;
  return faceIndex % 2 === 0 ? plan.slabFaces.end : plan.slabFaces.sides;
}

/** Whether a repeated picture mode advances across this face. Office applies
 * both plain stack and stackScale across planar targets and thick front/side
 * faces. A slab end face has no corresponding repetition extent, so it maps
 * one complete source image. */
export function surfacePictureFaceUsesStackedMapping(
  plan: SurfacePicturePlan,
  faceIndex: number,
): boolean {
  if ((plan.mode !== 'stack' && plan.mode !== 'stackScale')
    || !surfacePictureFaceIsEnabled(plan, faceIndex)) return false;
  return !plan.slabFaces || faceIndex === 0 || faceIndex % 2 === 1;
}

/** Number of source images mapped along one visible face. The two slab end
 * faces have no value-axis extent, so stackScale maps one complete source;
 * front/side faces retain the value-unit repetition count. */
export function surfacePictureFaceRepetitions(
  plan: SurfacePicturePlan,
  faceIndex: number,
): number {
  if (!surfacePictureFaceIsEnabled(plan, faceIndex)) return 0;
  return surfacePictureFaceUsesValueAxis(plan, faceIndex) ? plan.repetitions : 1;
}

/** Whether stackScale maps value units along this face. A repetition count of
 * one is not enough to answer this: when stackUnit exceeds the visible range,
 * a value-axis face still maps the larger unit and clips it at the face edge. */
export function surfacePictureFaceUsesValueAxis(
  plan: SurfacePicturePlan,
  faceIndex: number,
): boolean {
  return plan.mode === 'stackScale' && surfacePictureFaceUsesStackedMapping(plan, faceIndex);
}

function boundedSurfacePicturePlan(plan: SurfacePicturePlan): SurfacePicturePlan | null {
  let work = 0;
  for (let faceIndex = 0; faceIndex < 6; faceIndex++) {
    work += surfacePictureFaceRepetitions(plan, faceIndex);
    assertChartImageWork(work, MAX_CHART_IMAGE_FILL_TILES);
  }
  return plan;
}

function rectIsIdentity(rect: ImageFill['srcRect'] | ImageFill['fillRect']): boolean {
  return rect == null || [rect.l, rect.t, rect.r, rect.b].every(value => (value ?? 0) === 0);
}

function relativeRectIsSupported(rect: ImageFill['srcRect'] | ImageFill['fillRect']): boolean {
  if (!rect) return true;
  const l = rect.l ?? 0;
  const t = rect.t ?? 0;
  const r = rect.r ?? 0;
  const b = rect.b ?? 0;
  const values = [l, t, r, b];
  const right = 1 - r;
  const bottom = 1 - b;
  if (!values.every(Number.isFinite) || !Number.isFinite(right - l) || !Number.isFinite(bottom - t)) rejectChartImageGeometry();
  return right > l && bottom > t
    && Math.min(1, right) > Math.max(0, l)
    && Math.min(1, bottom) > Math.max(0, t);
}

/** The Office-observed, bounded subset of CT_Surface pictureOptions.
 *
 * ECMA-376 defines the flags and formats but not wall texture projection.
 * Excel/PDF observations establish full-face stretch and value-axis
 * stackScale on planar and positive-thickness back/side walls; floor ignores
 * pictureStackUnit. Planar plain stack uses one projected plot-face reference
 * aspect across floor and walls. Positive-thickness front/sides/end targets are
 * independently authored and map to the bounded six-face slab. Signed source
 * and destination rectangles retain DrawingML's mapping, including observed
 * outsets, on each face. */
export function planChartThreeDSurfacePicture(
  fill: ImageFill,
  surface: ChartThreeDSurface | null | undefined,
  kind: ChartThreeDSurfaceKind,
  valueSpan?: number,
): SurfacePicturePlan | null {
  if ((fill.tile != null) === (fill.stretch === true)
    || !relativeRectIsSupported(fill.srcRect) || !relativeRectIsSupported(fill.fillRect)
    || fill.rotWithShape === false
    || (fill.alpha != null && (!Number.isFinite(fill.alpha) || fill.alpha < 0 || fill.alpha > 1))) {
    return null;
  }
  const options = surface?.pictureOptions;
  if (options?.pictureFormatAuthored === true && options.pictureFormat == null) return null;
  if (options?.pictureStackUnitAuthored === true && options.pictureStackUnit == null) return null;
  const format = options?.pictureFormat ?? 'stretch';
  if ((!rectIsIdentity(fill.srcRect) || !rectIsIdentity(fill.fillRect)) && format !== 'stretch') {
    return null;
  }
  const thickness = surface?.thicknessPercent ?? 0;
  if (!Number.isFinite(thickness) || thickness < 0) rejectChartImageGeometry();
  const slabFaces = thickness === 0
    ? undefined
    : {
      front: options?.applyToFront !== false,
      sides: options?.applyToSides !== false,
      end: options?.applyToEnd !== false,
    };
  if (slabFaces && !Object.values(slabFaces).some(Boolean)) return null;
  if (!slabFaces) {
    if (kind === 'backWall' && options?.applyToFront === false) return null;
    if ((kind === 'floor' || kind === 'sideWall') && options?.applyToSides === false) return null;
  }
  if ((options?.pictureStackUnitAuthored === true || options?.pictureStackUnit != null)
    && format !== 'stackScale') return null;
  if (fill.tile) {
    if (format !== 'stretch' || !rectIsIdentity(fill.fillRect)) {
      return null;
    }
    // Physical tile size and the final aggregate work depend on the decoded
    // image and each projected face. The painter resolves both before drawing,
    // then applies the authored source rectangle inside every bounded tile.
    return boundedSurfacePicturePlan({ mode: 'tile', repetitions: 1, slabFaces });
  }
  if (format === 'stretch') {
    return boundedSurfacePicturePlan({ mode: 'stretch', repetitions: 1, slabFaces });
  }
  if (format === 'stack') {
    // The final repetition count depends on the decoded image aspect and the
    // face geometry. The painter calculates and bounds that count atomically;
    // positive-thickness end faces map one complete source instead.
    return boundedSurfacePicturePlan({ mode: 'stack', repetitions: 1, slabFaces });
  }
  if (format !== 'stackScale') return null;
  const stackUnit = options?.pictureStackUnit;
  if (stackUnit == null) return null;
  if (!(Number.isFinite(stackUnit) && stackUnit > 0)) rejectChartImageGeometry();
  // MS-OE376 §2.1.1543(c): Excel ignores pictureStackUnit on floor.
  if (kind === 'floor') {
    return boundedSurfacePicturePlan({ mode: 'stretch', repetitions: 1, slabFaces });
  }
  if (valueSpan == null) {
    return boundedSurfacePicturePlan({ mode: 'stackScale', repetitions: 1, stackUnit, slabFaces });
  }
  if (!(Number.isFinite(valueSpan) && valueSpan > 0)) rejectChartImageGeometry();
  const repetitions = Math.ceil(valueSpan / stackUnit);
  assertChartImageWork(repetitions, MAX_CHART_IMAGE_FILL_TILES);
  return boundedSurfacePicturePlan({ mode: 'stackScale', repetitions, stackUnit, slabFaces });
}

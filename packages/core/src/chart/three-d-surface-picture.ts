import { assertChartImageWork, rejectChartImageGeometry } from './image-fill-errors.js';
import type { ImageFill } from '../types/common.js';
import type { ChartThreeDSurface } from '../types/chart.js';
import { createAuxCanvasForContext, type AuxCanvas, type AuxContext } from '../canvas/aux-canvas.js';
import { MAX_CANVAS_AREA, MAX_CANVAS_DIMENSION } from '../canvas/clamp.js';
import { cropSourceMapping, drawImageCropped, imageNaturalSize, checkedImageCropProjection } from '../image/crop.js';
import { imageEffectFrame, imageEffectOutsetPeriods } from '../image/image-effect-frame.js';
import { drawProjected, projectQuadPoint } from '../shape/scene3d-draw.js';
import { chartImageTileMetrics, chartImageTileOrigin, isChartFillPreparationSource } from './image-fill.js';
import {
  planChartThreeDSurfacePicture,
  surfacePictureFaceIsEnabled,
  surfacePictureFaceRepetitions,
  surfacePictureFaceUsesStackedMapping,
  surfacePictureFaceUsesValueAxis,
  type ChartThreeDSurfaceKind,
  type SurfacePicturePoint,
  type SurfacePictureQuad,
} from './three-d-surface-picture-plan.js';
import type { ChartThreeDSurfaceGeometry, ThreeDScenePoint } from './three-d.js';
import { MAX_CHART_IMAGE_FILL_TILES } from './resource-limits.js';

function relativeRectQuad(
  quad: SurfacePictureQuad,
  rect: ImageFill['srcRect'] | ImageFill['fillRect'],
): SurfacePictureQuad | null {
  if (!rect || ![rect.l, rect.t, rect.r, rect.b].some(value => (value ?? 0) !== 0)) {
    return quad;
  }
  const left = rect.l ?? 0;
  const top = rect.t ?? 0;
  const right = 1 - (rect.r ?? 0);
  const bottom = 1 - (rect.b ?? 0);
  const points = [
    projectQuadPoint(quad, left, top),
    projectQuadPoint(quad, right, top),
    projectQuadPoint(quad, right, bottom),
    projectQuadPoint(quad, left, bottom),
  ];
  return points.every((point): point is SurfacePicturePoint => point != null)
    ? points as SurfacePictureQuad
    : rejectChartImageGeometry();
}

function screenAlignedFace(
  face: readonly ThreeDScenePoint[],
  project: (point: ThreeDScenePoint) => SurfacePicturePoint,
): [ThreeDScenePoint, ThreeDScenePoint, ThreeDScenePoint, ThreeDScenePoint] | null {
  if (face.length !== 4) return null;
  const byY = face.map((scenePoint, index) => ({
    scenePoint,
    projected: project(scenePoint),
    index,
  })).sort((left, right) => left.projected.y - right.projected.y
    || left.projected.x - right.projected.x);
  if (byY.some(({ projected }) => ![projected.x, projected.y].every(Number.isFinite))) rejectChartImageGeometry();
  const top = byY.slice(0, 2).sort((left, right) => left.projected.x - right.projected.x);
  const bottom = byY.slice(2).sort((left, right) => left.projected.x - right.projected.x);
  if (new Set([...top, ...bottom].map(item => item.index)).size !== 4) return null;
  return [top[0].scenePoint, top[1].scenePoint, bottom[1].scenePoint, bottom[0].scenePoint];
}

export function chartThreeDSurfacePictureSceneFace(
  geometry: ChartThreeDSurfaceGeometry,
  kind: ChartThreeDSurfaceKind,
  faceIndex: number,
  project: (point: ThreeDScenePoint) => SurfacePicturePoint,
): [ThreeDScenePoint, ThreeDScenePoint, ThreeDScenePoint, ThreeDScenePoint] | null {
  if (geometry.thickness === 0 && faceIndex === 0) {
    return [geometry.inner[3], geometry.inner[2], geometry.inner[1], geometry.inner[0]];
  }
  const aligned = screenAlignedFace(geometry.faces[faceIndex] ?? [], project);
  if (!aligned || kind !== 'backWall' || faceIndex !== 4) return aligned;
  // Office maps the back-wall slab's top joining face from its left-back
  // scene corner, whereas the generic screen-upright ordering starts at the
  // following corner and rotates the texture by one quadrant. Other back-wall
  // faces and every floor/side-wall face retain the generic ordering.
  return [aligned[1], aligned[2], aligned[3], aligned[0]];
}

/** Office plain-stack observation: derive one repetition height from the
 * projected plot-face aspect and source aspect, then share it across the
 * selected floor/wall target. Repetitions start at the target's lower edge.
 * Authored DPI does not affect this chart pictureOptions mode. */
function plainStackFraction(
  referenceAspect: number | null | undefined,
  sourceWidth: number,
  sourceHeight: number,
): number | null {
  if (!(referenceAspect != null && Number.isFinite(referenceAspect) && referenceAspect > 0)
    || !(sourceWidth > 0) || !(sourceHeight > 0)) return null;
  const fraction = referenceAspect * sourceHeight / sourceWidth;
  return Number.isFinite(fraction) && fraction > 0 ? fraction : null;
}

function sceneMetricDistance(
  first: ThreeDScenePoint,
  second: ThreeDScenePoint,
  modelDepth: number,
): number {
  return Math.hypot(
    first.x - second.x,
    first.y - second.y,
    (first.depth - second.depth) * modelDepth,
  );
}

/** Rendered size (points) of a wall/floor picture: the scene-metric extent of
 * the surface's inner face, onto which a stretched picture is mapped. It is
 * the occurrence size for per-occurrence CT_Blip chains (chartFillRenderedSize). */
export function chartThreeDSurfacePictureSizePt(
  geometry: ChartThreeDSurfaceGeometry, ptToPx: number,
): { widthPt: number; heightPt: number } | null {
  if (geometry.inner.length !== 4 || !(ptToPx > 0)) return null;
  const width = sceneMetricDistance(geometry.inner[0], geometry.inner[1], geometry.modelDepth) / ptToPx;
  const height = sceneMetricDistance(geometry.inner[1], geometry.inner[2], geometry.modelDepth) / ptToPx;
  return [width, height].every(v => Number.isFinite(v) && v > 0) ? { widthPt: width, heightPt: height } : null;
}

export function paintChartThreeDSurfacePicture(
  ctx: CanvasRenderingContext2D,
  fill: ImageFill,
  image: CanvasImageSource,
  surface: ChartThreeDSurface | null | undefined,
  kind: ChartThreeDSurfaceKind,
  geometry: ChartThreeDSurfaceGeometry,
  visibleFaceIndices: readonly number[],
  project: (point: ThreeDScenePoint) => SurfacePicturePoint,
  valueSpan: number,
  /** Per-mapping source for size-dependent CT_Blip chains: called with the
   * physical size (pt) of each face or repetition the picture is mapped onto,
   * so EMU-measured effects match that rendered size. */
  sourceFor?: (widthPt: number, heightPt: number) => CanvasImageSource | null,
  ptToPx = 1,
): boolean {
  const plan = planChartThreeDSurfacePicture(fill, surface, kind, valueSpan);
  // Scene-metric size of a mapped face span (value axis runs face[3]→face[0]).
  const mapped = (face: readonly ThreeDScenePoint[], lower = 0, upper = 1): CanvasImageSource | null => {
    if (!sourceFor) return image;
    const widthPt = sceneMetricDistance(face[3], face[2], geometry.modelDepth) / ptToPx;
    const heightPt = sceneMetricDistance(face[3], face[0], geometry.modelDepth) * (upper - lower) / ptToPx;
    if (![widthPt, heightPt].every(v => Number.isFinite(v) && v > 0)) return null;
    return sourceFor(widthPt, heightPt);
  };
  if (!plan) return false;
  if (geometry.inner.length !== 4) rejectChartImageGeometry();
  let projection: ReturnType<typeof checkedImageCropProjection>;
  try { projection = checkedImageCropProjection(fill.srcRect); } catch { rejectChartImageGeometry(); }
  // A spatial CT_Blip effect can grow the decoded grid around the authored
  // picture frame. The frame (not the grown grid) is the picture that maps
  // onto the face (§20.1.8.14/.55); its outset lies outside the clipped face.
  // Each per-mapping source has its own grid and outset, so these metrics are
  // derived from every resolved source before it is projected.
  const sourceMetrics = (source: CanvasImageSource) => {
    const effectFrame = imageEffectFrame(source);
    const grid = imageNaturalSize(source);
    const natural = effectFrame ? { w: effectFrame.width, h: effectFrame.height } : grid;
    if (![natural.w, natural.h].every(v => Number.isFinite(v) && v > 0)) rejectChartImageGeometry();
    const frameGrid = effectFrame
      ? { width: effectFrame.width, height: effectFrame.height } as unknown as CanvasImageSource
      : source;
    const crop = cropSourceMapping(frameGrid, fill.srcRect, projection);
    const fx = effectFrame?.x ?? 0, fy = effectFrame?.y ?? 0;
    const sourceRect = crop
      ? { x0: fx + crop.sx, y0: fy + crop.sy, x1: fx + crop.sx + crop.sw, y1: fy + crop.sy + crop.sh }
      : effectFrame ? { x0: fx, y0: fy, x1: fx + natural.w, y1: fy + natural.h } : undefined;
    return { grid, natural, crop, sourceRect };
  };
  const { natural } = sourceMetrics(image);
  // A plain-stack repetition height follows the decoded source aspect. A
  // preparation placeholder has no source frame, so it records no repetition
  // size; the host's next round resolves the real source first (see
  // collectChartFillRenderedSizes).
  if (plan.mode === 'stack' && isChartFillPreparationSource(image)) return false;
  const drawSource = (source: CanvasImageSource, destination: SurfacePictureQuad) => {
    const metrics = sourceMetrics(source);
    drawProjected(source, ctx, metrics.grid.w, metrics.grid.h, destination, 0.5, metrics.sourceRect);
  };
  // Preserve the established planar mapping exactly. Positive-thickness
  // joining faces need their own screen-upright ordering because each face has
  // a different scene-space axis pair.
  const fullQuad: SurfacePictureQuad = [
    project(geometry.inner[3]), project(geometry.inner[2]),
    project(geometry.inner[1]), project(geometry.inner[0]),
  ];
  const interpolate = (
    lower: ThreeDScenePoint,
    upper: ThreeDScenePoint,
    fraction: number,
  ): ThreeDScenePoint => ({
    x: lower.x + (upper.x - lower.x) * fraction,
    y: lower.y + (upper.y - lower.y) * fraction,
    depth: lower.depth + (upper.depth - lower.depth) * fraction,
  });
  const stackQuad = (
    face: readonly ThreeDScenePoint[],
    lower: number,
    upper: number,
  ): SurfacePictureQuad => [
    project(interpolate(face[3], face[0], upper)),
    project(interpolate(face[2], face[1], upper)),
    project(interpolate(face[2], face[1], lower)),
    project(interpolate(face[3], face[0], lower)),
  ];
  const stackFraction = plan.mode === 'stack'
    ? plainStackFraction(geometry.pictureStackAspect, natural.w, natural.h)
    : null;
  const tileMetrics = plan.mode === 'tile'
    ? chartImageTileMetrics(fill, image)
    : null;
  const tileFaces: Array<{
    faceIndex: number;
    width: number;
    height: number;
    origin: { x: number; y: number };
    firstColumn: number;
    firstRow: number;
    lastColumn: number;
    lastRow: number;
    canvas: AuxCanvas;
    context: AuxContext;
  }> = [];

  try {
  if (plan.mode === 'stack') {
    if (stackFraction == null) rejectChartImageGeometry();
    let work = 0;
    let hasFace = false;
    for (const faceIndex of visibleFaceIndices) {
      if (!surfacePictureFaceIsEnabled(plan, faceIndex)) continue;
      const face = chartThreeDSurfacePictureSceneFace(geometry, kind, faceIndex, project);
      if (!face) continue;
      const repetitions = surfacePictureFaceUsesStackedMapping(plan, faceIndex)
        ? Math.ceil(1 / stackFraction)
        : 1;
      if (repetitions < 1) rejectChartImageGeometry();
      assertChartImageWork(repetitions, MAX_CHART_IMAGE_FILL_TILES);
      work += repetitions;
      assertChartImageWork(work, MAX_CHART_IMAGE_FILL_TILES);
      hasFace = true;
    }
    if (!hasFace) return false;
  }

  if (plan.mode === 'tile') {
    if (!tileMetrics) rejectChartImageGeometry();
    let work = 0;
    let canvasArea = 0;
    let hasFace = false;
    for (const faceIndex of visibleFaceIndices) {
      if (!surfacePictureFaceIsEnabled(plan, faceIndex)) continue;
      const face = chartThreeDSurfacePictureSceneFace(geometry, kind, faceIndex, project);
      if (!face) continue;
      // ECMA-376 §20.1.8.58 defines the tile grid but not its ordering relative
      // to a 3-D chart-surface projection. Desktop Excel output shows that the
      // grid is foreshortened with the wall; keeping back-wall tiles in device
      // space makes them too large when the authored projection is below 1×.
      const width = sceneMetricDistance(face[0], face[1], geometry.modelDepth);
      const height = sceneMetricDistance(face[0], face[3], geometry.modelDepth);
      if (![width, height].every(Number.isFinite)) rejectChartImageGeometry();
      if (!(width > 0) || !(height > 0)) continue;
      const origin = chartImageTileOrigin(tileMetrics, width, height);
      // Neighbouring tiles whose CT_Blip effect outset reaches the face are
      // drawn and charged as well; the authored period is unchanged.
      const reach = imageEffectOutsetPeriods(image, fill.srcRect != null);
      const firstColumn = Math.floor(-origin.x / tileMetrics.tileW) - reach.x;
      const firstRow = Math.floor(-origin.y / tileMetrics.tileH) - reach.y;
      const lastColumn = Math.ceil((width - origin.x) / tileMetrics.tileW) + reach.x;
      const lastRow = Math.ceil((height - origin.y) / tileMetrics.tileH) + reach.y;
      // Resource policy: safe work counts alone do not prove that shifted
      // indices advance or retain alternate-tile parity in IEEE-754.
      if (![firstColumn, firstRow, lastColumn, lastRow].every(Number.isSafeInteger)) {
        assertChartImageWork(Number.MAX_SAFE_INTEGER, MAX_CHART_IMAGE_FILL_TILES);
      }
      const columns = Math.max(0, lastColumn - firstColumn);
      const rows = Math.max(0, lastRow - firstRow);
      // A positive face always intersects at least one tile. Cancellation
      // against a large phase must reject, rather than erase that face.
      if (columns === 0 || rows === 0) rejectChartImageGeometry();
      const repetitions = columns * rows;
      assertChartImageWork(repetitions, MAX_CHART_IMAGE_FILL_TILES);
      work += repetitions;
      assertChartImageWork(work, MAX_CHART_IMAGE_FILL_TILES);
      const canvasWidth = Math.ceil(width);
      const canvasHeight = Math.ceil(height);
      if (!(canvasWidth > 0 && canvasWidth <= MAX_CANVAS_DIMENSION)
        || !(canvasHeight > 0 && canvasHeight <= MAX_CANVAS_DIMENSION)
        || canvasWidth > Math.floor((MAX_CANVAS_AREA - canvasArea) / canvasHeight)) {
        assertChartImageWork(Number.MAX_SAFE_INTEGER, MAX_CANVAS_AREA);
      }
      canvasArea += canvasWidth * canvasHeight;
      const canvas = createAuxCanvasForContext(ctx, canvasWidth, canvasHeight);
      const context = canvas?.getContext('2d');
      if (!canvas || !context) {
        if (canvas) canvas.width = 0;
        rejectChartImageGeometry();
      }
      tileFaces.push({
        faceIndex,
        width,
        height,
        origin,
        firstColumn,
        firstRow,
        lastColumn,
        lastRow,
        canvas,
        context,
      });
      hasFace = true;
    }
    if (!hasFace) return false;
  }

  ctx.save();
  if (fill.alpha != null) ctx.globalAlpha *= fill.alpha;
  if (plan.mode === 'stretch') {
    for (const faceIndex of visibleFaceIndices) {
      if (!surfacePictureFaceIsEnabled(plan, faceIndex)) continue;
      const sceneFace = chartThreeDSurfacePictureSceneFace(geometry, kind, faceIndex, project);
      const quad = geometry.thickness === 0 && faceIndex === 0
        ? fullQuad
        : sceneFace?.map(project) as SurfacePictureQuad | undefined;
      if (!quad) continue;
      // sceneFace shares fullQuad's corner order on the planar face, so its
      // width runs face[3]→face[2] and its height face[3]→face[0].
      const faceImage = sceneFace ? mapped(sceneFace) : null;
      if (!faceImage) continue;
      const { grid, crop, sourceRect } = sourceMetrics(faceImage);
      const fillDestination = relativeRectQuad(quad, fill.fillRect);
      if (!fillDestination) continue;
      const destination = crop
        ? relativeRectQuad(fillDestination, {
          l: crop.dxFraction,
          t: crop.dyFraction,
          r: 1 - crop.dxFraction - crop.dwFraction,
          b: 1 - crop.dyFraction - crop.dhFraction,
        })
        : fillDestination;
      if (!destination) continue;
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(quad[0].x, quad[0].y);
      for (let index = 1; index < quad.length; index++) ctx.lineTo(quad[index].x, quad[index].y);
      ctx.closePath();
      ctx.clip();
      drawProjected(faceImage, ctx, grid.w, grid.h, destination, 0.5, sourceRect);
      ctx.restore();
    }
  } else {
    for (const faceIndex of visibleFaceIndices) {
      if (!surfacePictureFaceIsEnabled(plan, faceIndex)) continue;
      const face = chartThreeDSurfacePictureSceneFace(geometry, kind, faceIndex, project);
      if (!face) continue;
      const quad = face.map(project) as SurfacePictureQuad;
      ctx.save();
      ctx.beginPath();
      ctx.moveTo(quad[0].x, quad[0].y);
      for (let index = 1; index < quad.length; index++) ctx.lineTo(quad[index].x, quad[index].y);
      ctx.closePath();
      ctx.clip();
      // Repetition follows the value axis across the slab's front and side
      // faces. End faces have no value-axis extent, so Office maps one whole
      // source there instead of compressing every repetition into thickness.
      const repetitions = surfacePictureFaceRepetitions(plan, faceIndex);
      if (plan.mode === 'tile') {
        if (!tileMetrics) continue;
        const tileFace = tileFaces.find(candidate => candidate.faceIndex === faceIndex);
        if (!tileFace) continue;
        const scaleX = tileFace.canvas.width / tileFace.width;
        const scaleY = tileFace.canvas.height / tileFace.height;
        tileFace.context.save();
        tileFace.context.scale(scaleX, scaleY);
        for (let ri = 0; ri < tileFace.lastRow - tileFace.firstRow; ri++) {
          const row = tileFace.firstRow + ri;
          for (let ci = 0; ci < tileFace.lastColumn - tileFace.firstColumn; ci++) {
            const column = tileFace.firstColumn + ci;
            const dx = tileFace.origin.x + column * tileMetrics.tileW;
            const dy = tileFace.origin.y + row * tileMetrics.tileH;
            const mirrorX = tileMetrics.flipX && Math.abs(column) % 2 === 1;
            const mirrorY = tileMetrics.flipY && Math.abs(row) % 2 === 1;
            tileFace.context.save();
            tileFace.context.translate(
              dx + (mirrorX ? tileMetrics.tileW : 0),
              dy + (mirrorY ? tileMetrics.tileH : 0),
            );
            tileFace.context.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1);
            drawImageCropped(
              tileFace.context,
              image,
              fill.srcRect,
              0,
              0,
              tileMetrics.tileW,
              tileMetrics.tileH,
            );
            tileFace.context.restore();
          }
        }
        tileFace.context.restore();
        drawProjected(
          tileFace.canvas,
          ctx,
          tileFace.canvas.width,
          tileFace.canvas.height,
          quad,
          0.5,
        );
      } else if (plan.mode === 'stack') {
        if (stackFraction == null) continue;
        if (surfacePictureFaceUsesStackedMapping(plan, faceIndex)) {
          for (let index = 0; index < Math.ceil(1 / stackFraction); index++) {
            const unit = mapped(face, index * stackFraction, (index + 1) * stackFraction);
            if (unit) drawSource(unit, stackQuad(face, index * stackFraction, (index + 1) * stackFraction));
          }
        } else {
          const whole = mapped(face);
          if (whole) drawSource(whole, quad);
        }
      } else if (plan.stackUnit != null && surfacePictureFaceUsesValueAxis(plan, faceIndex)) {
        for (let index = 0; index < repetitions; index++) {
          const lower = index * plan.stackUnit / valueSpan;
          const upper = (index + 1) * plan.stackUnit / valueSpan;
          const unit = mapped(face, lower, upper);
          if (unit) drawSource(unit, stackQuad(face, lower, upper));
        }
      } else if (plan.stackUnit != null) {
        const whole = mapped(face);
        if (whole) drawSource(whole, quad);
      }
      ctx.restore();
    }
  }
  ctx.restore();
  return true;
  } finally {
    // A later face can reject after an earlier allocation. Scratch ownership
    // ends on success and failure; the projected pixels have already been
    // copied synchronously into the destination when paint succeeds.
    for (const face of tileFaces) face.canvas.width = 0;
  }
}

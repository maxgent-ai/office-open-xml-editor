import type { Matrix2DData, PointPt } from './types.js';

/** Canvas-order composition: the returned transform applies `inner`, then `outer`. */
export function composeAffine(outer: Matrix2DData, inner: Matrix2DData): Matrix2DData {
  return Object.freeze({
    a: outer.a * inner.a + outer.c * inner.b,
    b: outer.b * inner.a + outer.d * inner.b,
    c: outer.a * inner.c + outer.c * inner.d,
    d: outer.b * inner.c + outer.d * inner.d,
    e: outer.a * inner.e + outer.c * inner.f + outer.e,
    f: outer.b * inner.e + outer.d * inner.f + outer.f,
  });
}

export function scaleAffine(scale: number): Matrix2DData {
  return Object.freeze({ a: scale, b: 0, c: 0, d: scale, e: 0, f: 0 });
}

export function translationAffine(x: number, y: number): Matrix2DData {
  return Object.freeze({ a: 1, b: 0, c: 0, d: 1, e: x, f: y });
}

/** Exact quarter-turn data avoids floating-point axis drift from Math.cos(π/2). */
export function quarterTurnAffine(direction: 1 | -1): Matrix2DData {
  return direction === 1
    ? Object.freeze({ a: 0, b: 1, c: -1, d: 0, e: 0, f: 0 })
    : Object.freeze({ a: 0, b: -1, c: 1, d: 0, e: 0, f: 0 });
}

export function mapAffinePoint(matrix: Matrix2DData, point: PointPt): PointPt {
  return {
    xPt: matrix.a * point.xPt + matrix.c * point.yPt + matrix.e,
    yPt: matrix.b * point.xPt + matrix.d * point.yPt + matrix.f,
  };
}

export function inverseMapAffinePoint(matrix: Matrix2DData, point: PointPt): PointPt | null {
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  if (!Number.isFinite(determinant) || determinant === 0) return null;
  const x = point.xPt - matrix.e;
  const y = point.yPt - matrix.f;
  const result = {
    xPt: (matrix.d * x - matrix.c * y) / determinant,
    yPt: (-matrix.b * x + matrix.a * y) / determinant,
  };
  return Number.isFinite(result.xPt) && Number.isFinite(result.yPt) ? result : null;
}

/**
 * The rectangle whose image under `matrix` is `rect`, when `matrix` maps
 * rectangles to rectangles: a translation or scale, optionally with an exact
 * quarter turn or reflection. Any other rotation or skew maps no rectangle
 * to `rect`, and the result is null.
 */
export function inverseMapAxisAlignedRect(
  matrix: Matrix2DData,
  rect: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
): Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }> | null {
  const axisAligned = (matrix.b === 0 && matrix.c === 0) || (matrix.a === 0 && matrix.d === 0);
  if (!axisAligned) return null;
  const corners = [
    { xPt: rect.xPt, yPt: rect.yPt },
    { xPt: rect.xPt + rect.widthPt, yPt: rect.yPt + rect.heightPt },
  ].map((corner) => inverseMapAffinePoint(matrix, corner));
  if (corners.some((corner) => corner === null)) return null;
  const [first, second] = corners as PointPt[];
  const xPt = Math.min(first!.xPt, second!.xPt);
  const yPt = Math.min(first!.yPt, second!.yPt);
  return Object.freeze({
    xPt,
    yPt,
    widthPt: Math.max(first!.xPt, second!.xPt) - xPt,
    heightPt: Math.max(first!.yPt, second!.yPt) - yPt,
  });
}

export function inverseMapAffineVector(
  matrix: Matrix2DData,
  vector: PointPt,
): PointPt | null {
  const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
  if (!Number.isFinite(determinant) || determinant === 0) return null;
  const result = {
    xPt: (matrix.d * vector.xPt - matrix.c * vector.yPt) / determinant,
    yPt: (-matrix.b * vector.xPt + matrix.a * vector.yPt) / determinant,
  };
  return Number.isFinite(result.xPt) && Number.isFinite(result.yPt) ? result : null;
}

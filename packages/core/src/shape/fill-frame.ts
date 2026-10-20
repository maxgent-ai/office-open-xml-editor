/** ECMA-376 §20.1.8.14/33: a non-rotating fill retains its box centre
 * while its complete paint frame counter-rotates. Scaling/cropping happens
 * within that frame first; changing only a gradient angle is insufficient. */
export function counterRotatedFillTransform(
  matrix: DOMMatrix2DInit, rotation: number, cx: number, cy: number,
): DOMMatrix2DInit {
  if (rotation === 0) return matrix;
  const angle = -rotation * Math.PI / 180;
  const c = Math.cos(angle); const s = Math.sin(angle);
  const a = matrix.a ?? 1; const b = matrix.b ?? 0;
  const cc = matrix.c ?? 0; const d = matrix.d ?? 1;
  const x = (matrix.e ?? 0) - cx; const y = (matrix.f ?? 0) - cy;
  return { a: c * a - s * b, b: s * a + c * b,
    c: c * cc - s * d, d: s * cc + c * d,
    e: cx + c * x - s * y, f: cy + s * x + c * y };
}

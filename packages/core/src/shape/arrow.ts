import type { ArrowEnd, Stroke } from '../types/common';
import { hexToRgba, resolveFill, usesPathShade } from './paint';

/** A 2D point in canvas pixels. */
export interface Point {
  x: number;
  y: number;
}

/**
 * Resolve a line-end decoration's pixel geometry. The width/length steps
 * (sm/med/lg) are *relative* in the spec (§20.1.10.31–.32); the multipliers of
 * line width below are calibrated against PowerPoint. `lw` is the line width in
 * device px, `halfW` the half-span across the line, `len` the extent along it.
 */
function arrowGeom(
  arrowEnd: ArrowEnd,
  stroke: Stroke,
  scale: number,
): { lw: number; halfW: number; len: number } {
  const lw = Math.max(0.5, stroke.width * scale);
  const wMul = arrowEnd.w === 'sm' ? 4 : arrowEnd.w === 'lg' ? 8 : 6;
  const lMul = arrowEnd.len === 'sm' ? 4 : arrowEnd.len === 'lg' ? 8 : 6;
  return { lw, halfW: (lw * wMul) / 2, len: lw * lMul };
}

/** Decorations whose filled body covers the tip→`-len` span, so the leader line
 *  must stop at `-len` (retract) for its cap to hide inside the shape. The open
 *  `arrow` (a stroked V) and `none` keep the line running all the way to the tip. */
const RETRACTING_ENDS = new Set(['triangle', 'stealth', 'diamond', 'oval']);

/**
 * How far (device px) the leader line should be pulled back from the tip so a
 * line-end decoration's filled body hides the line's end cap. Zero for `arrow`
 * and `none`. Matches the `len` used by {@link drawArrowHead} so the line stops
 * exactly at the decoration's base.
 */
export function lineEndRetract(arrowEnd: ArrowEnd, stroke: Stroke, scale: number): number {
  if (!RETRACTING_ENDS.has(arrowEnd.type)) return 0;
  return arrowGeom(arrowEnd, stroke, scale).len;
}

/** Conservative radius around a line-end tip occupied by its painted pixels. */
export function lineEndPaintExtent(arrowEnd: ArrowEnd, stroke: Stroke, scale: number): number {
  if (arrowEnd.type === 'none') return 0;
  const { lw, halfW, len } = arrowGeom(arrowEnd, stroke, scale);
  return Math.max(len, halfW) + lw / 2;
}

/**
 * Pull `p` toward its neighbour `toward` by `amount` px, clamped so it never
 * passes the neighbour. Used to retract a polyline's terminal vertex before
 * stroking, so a decorated end stops at the decoration's base.
 */
export function retractLineEndpoint(p: Point, toward: Point, amount: number): Point {
  if (amount <= 0) return { x: p.x, y: p.y };
  const dx = toward.x - p.x;
  const dy = toward.y - p.y;
  const d = Math.hypot(dx, dy);
  if (d < 1e-9) return { x: p.x, y: p.y };
  const t = Math.min(amount, d) / d;
  return { x: p.x + dx * t, y: p.y + dy * t };
}

/**
 * Draw a DrawingML line-end decoration (arrow head) at `(tipX, tipY)`,
 * oriented along `angle` radians (0 = pointing right, +x axis).
 *
 * ECMA-376 §20.1.8.3 (CT_LineEndProperties) / §20.1.10.33 (ST_LineEndType:
 * none / triangle / stealth / diamond / oval / arrow) / §20.1.10.31–.32
 * (ST_LineEndWidth / ST_LineEndLength: sm / med / lg). The spec only names
 * the w/len steps as *relative* sizes, not exact ratios — the multiples of
 * line width below are calibrated against PowerPoint's rendering and shared
 * between the pptx and docx renderers so connector arrows look identical.
 *
 * `scale` is the EMU → device-px factor (same convention as core's
 * `applyStroke`, where stroke width in px is `stroke.width * scale`).
 */
export function drawArrowHead(
  ctx: CanvasRenderingContext2D,
  tipX: number,
  tipY: number,
  angle: number,
  arrowEnd: ArrowEnd,
  stroke: Stroke,
  scale: number,
  effectivePaint?: string | CanvasGradient | CanvasPattern,
): void {
  if (arrowEnd.type === 'none') return;
  const { lw, halfW, len } = arrowGeom(arrowEnd, stroke, scale);
  const paint = effectivePaint ?? hexToRgba(stroke.color);

  // Only the new untiled rect/shape raster is authored in the host frame.
  // Ordinary patterns and tiled gradients retain main's decoration-local CTM;
  // native solid/linear/circle paints also keep their existing frames. A
  // CanvasPattern alone cannot distinguish those brushes from a path raster.
  const hostTransform = usesPathShade(stroke.fill)
    && typeof ctx.getTransform === 'function' ? ctx.getTransform() : undefined;
  const anchorPaint = () => {
    // Canvas patterns follow the paint-time CTM; native gradients retain their
    // creation frame. The path has already captured the decoration transform.
    // Paint a host-box pattern at the host CTM so rotating/translating the
    // arrow geometry does not translate/rotate its gradient a second time.
    if (hostTransform && typeof paint === 'object' && 'setTransform' in paint) ctx.setTransform(hostTransform);
  };
  ctx.save();
  ctx.translate(tipX, tipY);
  ctx.rotate(angle);
  ctx.fillStyle = paint;
  ctx.strokeStyle = paint;
  ctx.lineWidth = lw;
  ctx.setLineDash([]);
  ctx.beginPath();
  switch (arrowEnd.type) {
    case 'triangle':
    case 'stealth':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len, -halfW);
      ctx.lineTo(-len, halfW);
      ctx.closePath();
      anchorPaint(); ctx.fill();
      break;
    case 'arrow':
      // An open DrawingML arrow is one continuous chevron. Separate subpaths
      // leave two flat caps stacked at the tip, producing a visibly broken,
      // jagged point. PowerPoint joins the two arms and rounds the exposed ends.
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.moveTo(-len, -halfW);
      ctx.lineTo(0, 0);
      ctx.lineTo(-len, halfW);
      anchorPaint(); ctx.stroke();
      break;
    case 'diamond':
      ctx.moveTo(0, 0);
      ctx.lineTo(-len / 2, -halfW);
      ctx.lineTo(-len, 0);
      ctx.lineTo(-len / 2, halfW);
      ctx.closePath();
      anchorPaint(); ctx.fill();
      break;
    case 'oval':
      ctx.ellipse(-len / 2, 0, len / 2, halfW, 0, 0, Math.PI * 2);
      anchorPaint(); ctx.fill();
      break;
  }
  ctx.restore();
}

/** Resolve decoration paint from the very geometry drawn by drawArrowHead.
 * The filled end is contained by [-len,0] × [-halfW,halfW]; oval uses its
 * exact ellipse support. Open arrows add the transformed round stroke pen.
 * Keep the gradient's host frame separate from this device-space coverage. */
export function resolveArrowPaint(
  ctx: CanvasRenderingContext2D, stroke: Stroke, scale: number,
  tipX: number, tipY: number, angle: number, end: ArrowEnd,
  frame: { x: number; y: number; w: number; h: number }, rotation: number, ptUnits: number,
): string | CanvasGradient | CanvasPattern | undefined {
  if (!stroke.fill || end.type === 'none') return undefined;
  if (!usesPathShade(stroke.fill)
    || typeof ctx.getTransform !== 'function') {
    return resolveFill(stroke.fill, ctx, frame.x, frame.y, frame.w, frame.h, rotation, ptUnits) ?? undefined;
  }
  const { lw, halfW, len } = arrowGeom(end, stroke, scale);
  const m = ctx.getTransform(); const c = Math.cos(angle); const s = Math.sin(angle);
  const a = m.a * c + m.c * s; const b = m.b * c + m.d * s;
  const cc = -m.a * s + m.c * c; const d = -m.b * s + m.d * c;
  const cx = m.a * tipX + m.c * tipY + m.e - a * len / 2;
  const cy = m.b * tipX + m.d * tipY + m.f - b * len / 2;
  const hx = end.type === 'oval' ? Math.hypot(a * len / 2, cc * halfW) : Math.abs(a) * len / 2 + Math.abs(cc) * halfW;
  const hy = end.type === 'oval' ? Math.hypot(b * len / 2, d * halfW) : Math.abs(b) * len / 2 + Math.abs(d) * halfW;
  const px = end.type === 'arrow' ? lw / 2 * Math.hypot(a, cc) : 0;
  const py = end.type === 'arrow' ? lw / 2 * Math.hypot(b, d) : 0;
  return resolveFill(stroke.fill, ctx, frame.x, frame.y, frame.w, frame.h, rotation, ptUnits,
    undefined, undefined, { x: cx - hx - px, y: cy - hy - py, w: 2 * (hx + px), h: 2 * (hy + py) }) ?? undefined;
}

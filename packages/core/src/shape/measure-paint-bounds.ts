import { createAuxCanvasForContext } from '../canvas/aux-canvas';
import { appendPathStateCommand, strokePathStateBounds, type GeometryBox, type PathState } from './path-enclosure';

/** Conservative, unclipped bounds of Canvas text/image/path paint. Run the
 * actual painter so layout, insets, vertical glyph transforms, decorations and
 * embedded objects cannot diverge from a second layout implementation. Only a
 * 1px state/measurement canvas and the current path are retained. Clipping and
 * transparent image margins may overestimate bounds; neither may remove ink.
 * Observes current-path operations, not opaque Path2D objects. The supplied
 * painter must not use readback of its destination as input. */
export function measurePaintBounds(
  source: CanvasRenderingContext2D,
  paint: (ctx: CanvasRenderingContext2D) => void,
): GeometryBox | undefined {
  const canvas = createAuxCanvasForContext(source, 1, 1);
  const ctx = canvas?.getContext('2d') as CanvasRenderingContext2D | null;
  if (!ctx) return undefined;
  const path: PathState = { paths: [] };
  let left = Infinity, top = Infinity, right = -Infinity, bottom = -Infinity;
  const add = (box: GeometryBox | undefined) => {
    if (!box || ![box.x, box.y, box.w, box.h].every(Number.isFinite)) return;
    left = Math.min(left, box.x); top = Math.min(top, box.y);
    right = Math.max(right, box.x + box.w); bottom = Math.max(bottom, box.y + box.h);
    // Canvas shadows use device-space offsets/blur, independent of the CTM.
    // Retain applyOuterShadow's conservative 3× blur crop policy; this is a
    // raster allocation bound, not an Office shadow-transfer equation.
    if (ctx.shadowColor !== 'rgba(0, 0, 0, 0)' && ctx.shadowColor !== 'transparent') {
      const blur = 3 * ctx.shadowBlur;
      left = Math.min(left, box.x + ctx.shadowOffsetX - blur);
      top = Math.min(top, box.y + ctx.shadowOffsetY - blur);
      right = Math.max(right, box.x + box.w + ctx.shadowOffsetX + blur);
      bottom = Math.max(bottom, box.y + box.h + ctx.shadowOffsetY + blur);
    }
  };
  const rect = (x: number, y: number, w: number, h: number, stroke = 0) => {
    const m = ctx.getTransform();
    const points = [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
    const xs = points.map(([px, py]) => m.a * px + m.c * py + m.e);
    const ys = points.map(([px, py]) => m.b * px + m.d * py + m.f);
    const rx = stroke * Math.hypot(m.a, m.c), ry = stroke * Math.hypot(m.b, m.d);
    add({ x: Math.min(...xs) - rx, y: Math.min(...ys) - ry,
      w: Math.max(...xs) - Math.min(...xs) + 2 * rx,
      h: Math.max(...ys) - Math.min(...ys) + 2 * ry });
  };
  const proxy = new Proxy(ctx, {
    get(target, property) {
      // Effect helpers need the real viewport dimensions, not the scratch size.
      if (property === 'canvas') return source.canvas;
      const value = Reflect.get(target, property, target);
      if (typeof value !== 'function') return value;
      return (...args: unknown[]) => {
        const name = String(property);
        const m = ctx.getTransform();
        appendPathStateCommand(path, { a: m.a, b: m.b, c: m.c, d: m.d, e: m.e, f: m.f }, name, args as number[]);
        if (name === 'fillText' || name === 'strokeText') {
          const [text, x, y, maxWidth] = args as [string, number, number, number?];
          const metrics = ctx.measureText(text);
          const stretch = maxWidth !== undefined && metrics.width > maxWidth ? maxWidth / metrics.width : 1;
          rect(x - metrics.actualBoundingBoxLeft * stretch, y - metrics.actualBoundingBoxAscent,
            (metrics.actualBoundingBoxLeft + metrics.actualBoundingBoxRight) * stretch,
            metrics.actualBoundingBoxAscent + metrics.actualBoundingBoxDescent,
            name === 'strokeText' ? ctx.lineWidth / 2 * (ctx.lineJoin === 'miter' ? ctx.miterLimit : 1) : 0);
          return;
        }
        if (name === 'fill' || name === 'stroke') {
          if (m.a * m.d - m.b * m.c === 0) return;
          add(strokePathStateBounds(path, m, { lineWidth: name === 'stroke' ? ctx.lineWidth : 0,
            lineCap: ctx.lineCap, lineJoin: ctx.lineJoin, miterLimit: ctx.miterLimit, dash: ctx.getLineDash() }));
          return;
        }
        if (name === 'fillRect' || name === 'strokeRect') {
          const [x, y, w, h] = args as number[];
          rect(x, y, w, h, name === 'strokeRect' ? ctx.lineWidth / 2 * Math.max(1, ctx.miterLimit) : 0);
          return;
        }
        if (name === 'drawImage') {
          const image = args[0] as { width: number; height: number };
          const [x, y, w, h] = args.length === 9 ? args.slice(5) as number[]
            : args.length === 5 ? args.slice(1) as number[] : [args[1] as number, args[2] as number, image.width, image.height];
          rect(x, y, w, h);
          return;
        }
        if (name === 'clearRect') return;
        return Reflect.apply(value, target, args);
      };
    },
    set(target, property, value) { return Reflect.set(target, property, value, target); },
  });
  paint(proxy);
  return right > left && bottom > top ? { x: left, y: top, w: right - left, h: bottom - top } : undefined;
}

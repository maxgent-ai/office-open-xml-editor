/** ECMA-376 §20.1.10.50: legacyWireframe has no surface material; the
 * illustrated geometry consists of silhouette/profile rims and connecting
 * edges, with an empty face. The standard does not prescribe tessellation.
 * Library raster policy: caller-scaled rims and one-pixel mesh edges,
 * 32 azimuth sectors for curved bevel strips, and the same parallel depth
 * sweep as the extrusion compositor.
 * These are mesh/raster choices, not measured Office material parameters.
 */
export function rasterBevelWireframe(
  pixels: Uint8ClampedArray,
  coverage: Uint8ClampedArray,
  distance: Float64Array,
  w: number,
  h: number,
  band: number,
  lineWidth: number,
  offset?: { x: number; y: number },
): void {
  // Geometry and source opacity are separate planes: coverage controls mesh
  // topology; the original RGBA alpha only attenuates the painted edges.
  const source = pixels.slice();
  const sectors = new Int8Array(w * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = y * w + x;
    const gx = (x + 1 < w ? distance[i + 1] : 0) - (x > 0 ? distance[i - 1] : 0);
    const gy = (y + 1 < h ? distance[i + w] : 0) - (y > 0 ? distance[i - w] : 0);
    sectors[i] = Math.floor((Math.atan2(gy, gx) + Math.PI) * 16 / Math.PI) % 32;
    pixels[i * 4 + 3] = 0;
  }
  const paint = (x: number, y: number, sourceIndex: number, coverage: number) => {
    if (x < 0 || x >= w || y < 0 || y >= h) return;
    const dest = (y * w + x) * 4, src = sourceIndex * 4;
    const a = source[src + 3] * Math.min(1, Math.max(0, coverage));
    if (a <= pixels[dest + 3]) return;
    pixels[dest] = source[src];
    pixels[dest + 1] = source[src + 1];
    pixels[dest + 2] = source[src + 2];
    pixels[dest + 3] = a;
  };
  const dx = offset && Number.isFinite(offset.x) ? offset.x : 0;
  const dy = offset && Number.isFinite(offset.y) ? offset.y : 0;
  const seeds = offset ? new Uint8Array(w * h) : undefined;
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = y * w + x;
    if (coverage[i] === 0) continue;
    const d = Math.max(0, distance[i] - .5);
    const outer = Math.max(0, Math.min(1, lineWidth + .5 - d));
    const inner = band > lineWidth ? Math.max(0, lineWidth / 2 + .5 - Math.abs(d - band)) : 0;
    const crossing = (x + 1 < w && coverage[i + 1] > 0 && sectors[i] !== sectors[i + 1]) ||
      (y + 1 < h && coverage[i + w] > 0 && sectors[i] !== sectors[i + w]);
    const spoke = band > 0 && d <= band && crossing ? 1 : 0;
    const edgeCoverage = Math.max(outer, inner, spoke);
    if (edgeCoverage > 0) {
      paint(x, y, i, edgeCoverage);
      if (offset) paint(Math.round(x + dx), Math.round(y + dy), i, edgeCoverage);
    }
    if (seeds && outer > 0 && crossing) seeds[i] = 1;
  }
  // All depth edges are parallel. Propagate their union along sheared raster
  // scanlines in O(pixels), rather than drawing a depth-length line for every
  // boundary pixel (which is quadratic on a noisy picture silhouette). Each
  // scanline retains only its last seed and segment end; even enormous authored
  // depth cannot add iterations or allocations beyond this local raster.
  if (!seeds || (dx === 0 && dy === 0)) return;
  const majorX = Math.abs(dx) >= Math.abs(dy);
  const majorSize = majorX ? w : h, minorSize = majorX ? h : w;
  const delta = majorX ? dx : dy, slope = (majorX ? dy : dx) / delta;
  const direction = delta > 0 ? 1 : -1;
  const sources = new Int32Array(minorSize + 2 * majorSize + 1).fill(-1);
  const ends = new Float64Array(sources.length);
  for (let major = direction > 0 ? 0 : majorSize - 1;
    major >= 0 && major < majorSize; major += direction) {
    const shift = Math.round(slope * major);
    for (let minor = 0; minor < minorSize; minor++) {
      const x = majorX ? major : minor, y = majorX ? minor : major;
      const i = y * w + x, line = minor - shift + majorSize;
      if (seeds[i]) {
        sources[line] = i;
        ends[line] = direction * major + Math.abs(delta);
      }
      if (sources[line] >= 0 && direction * major <= ends[line])
        paint(x, y, sources[line], 1);
    }
  }
}

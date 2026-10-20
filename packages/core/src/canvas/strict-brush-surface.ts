import { OoxmlError } from '../errors/ooxml-error.js';

// Ordinary shape paint keeps its established policy: when an auxiliary
// brush surface (tiled gradient cell, path-shade raster, repeat pattern) is
// unavailable, the painter degrades to a native gradient. Inside a CT_Blip
// effect chain that substitution would silently alter authored pixels, so the
// effect evaluator runs the shared painter in a strict scope where the same
// allocation failure is the typed effect error instead. Synchronous only.
let strictDepth = 0;

export function withStrictBrushSurfaces<T>(paint: () => T): T {
  strictDepth++;
  try {
    return paint();
  } finally {
    strictDepth--;
  }
}

/** Return null (the ordinary degrade signal) or, in a strict scope, throw. */
export function brushSurfaceUnavailable(): null {
  if (strictDepth > 0) throw new OoxmlError('unsupported-image-effect', 'Brush surface unavailable');
  return null;
}

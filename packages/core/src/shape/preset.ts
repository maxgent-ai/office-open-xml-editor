import type { GeometryPathSink } from './path-data';
import { ownValue } from '../internal/own-key';
import { buildPresetGeometryPath } from './preset-geometry/geometry';

// Fallback path builder for OOXML `<a:prstGeom>` names (ECMA-376 §20.1.9.18;
// ST_ShapeType §20.1.10.56).
//
// Every ST_ShapeType value except `rect` is defined in
// presetShapeDefinitions.xml (ECMA-376 Part 1, DrawingML Geometries) and is
// rendered by the spec engine (shape/preset-geometry, presets.json). The live
// callers — the DrawingML geometry resolver (docx/pptx/xlsx shape bodies and
// clip/shade silhouettes) and the pptx renderer's fill outline / body
// fallback — consult the spec engine FIRST (`resolvePresetGeometryPaths`,
// `buildPresetGeometryFillPath`, `hasPreset`) and reach this function only for
// a name the engine does not define: `rect`, whose geometry the spec leaves
// implicit (the full shape box), or a schema-invalid value.
//
// The former hand-written per-preset switch (approximate, uncited geometry
// such as a homePlate pointing down and fixed-ratio arrow heads) was therefore
// unreachable from every caller and has been deleted (#1856). Spec-defined
// names passed here still resolve through the spec engine so the function is
// correct for any caller.

/**
 * Non-ST_ShapeType spellings the former hand-written switch accepted, mapped to
 * the ECMA-376 preset whose body they shared. A schema-invalid `prst` is not
 * produced by any of the repository's parsers; this map is library policy that
 * keeps such input on the same shape class it historically rendered as, rather
 * than collapsing it to a rectangle.
 */
const LEGACY_SPEC_ALIASES: Record<string, string> = {
  oval: 'ellipse',
  rtriangle: 'rttriangle',
  roundrectangle: 'roundrect',
  star: 'star5',
  nosmokingsign: 'nosmoking',
  document: 'foldedcorner',
  flowchartmanuallabel: 'flowchartonlinestorage',
  flowchartpuncheddisk: 'flowchartonlinestorage',
  // Historic misspelling of flowChartSummingJunction.
  flowchartsumingjunction: 'flowchartsummingjunction',
};

/** Build the canvas path for an OOXML preset geometry (`<a:prstGeom>`).
 *
 * The caller is responsible for `ctx.beginPath()` / `ctx.fill()` /
 * `ctx.stroke()` — this function only emits path commands.
 *
 * Spec-defined names (and the legacy aliases above) emit every `<path>` of the
 * ECMA-376 definition via the spec engine. `rect` — and any other value the
 * engine does not define — emits the shape box, which is `rect`'s implicit
 * geometry and the library's fallback for schema-invalid names.
 */
export function buildShapePath(
  ctx: GeometryPathSink,
  geom: string,
  x: number,
  y: number,
  w: number,
  h: number,
  adj: number | null = null,
  adj2: number | null = null,
  adj3: number | null = null,
  adj4: number | null = null,
) {
  const raw = geom.toLowerCase();
  const key = ownValue(LEGACY_SPEC_ALIASES, raw) ?? raw;
  if (buildPresetGeometryPath(ctx, key, x, y, w, h, [adj, adj2, adj3, adj4])) return;
  ctx.rect(x, y, w, h);
}

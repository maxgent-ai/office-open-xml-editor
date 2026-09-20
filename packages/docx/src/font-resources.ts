import {
  embeddedFontBytesAreWithinLimit,
  normalizeFontMetricFamily,
  parseOpenTypeResourceMetrics,
  registerEmbeddedFonts,
  type EmbeddedFontFace,
  type ResolvedFontMetric,
} from '@silurus/ooxml-core';
import {
  wordOpenTypeDesignLineRatios,
  wordOpenTypeEastAsianSingleLineRatio,
} from './layout/line-compatibility.js';

/** One concrete static sfnt font face supplied by the application.
 *
 * The bytes, rather than a family-name metric table, are the authority for
 * pagination. WOFF/WOFF2 containers are intentionally not accepted here: the
 * OpenType tables must be directly inspectable before the face can influence
 * layout. Variable faces (`fvar`) are unsupported until instance metrics can
 * be resolved. A failed or unsupported resource retains the existing fallback path;
 * it does not opt native or embedded fonts into a new metric policy.
 */
export interface DocxFontResource {
  readonly family: string;
  readonly bytes: Uint8Array;
  readonly weight?: 400 | 700;
  readonly style?: 'normal' | 'italic';
}

export interface LoadedDocxFontResources {
  readonly faces: FontFace[];
  readonly metrics: Readonly<Record<string, ResolvedFontMetric>>;
}

// Application-provided resources cross the worker boundary by structured
// clone. Bound both fan-out and retained bytes before cloning/registration can
// become an accidental process-wide font cache. These are hard safety ceilings,
// not rendering heuristics; documents remain renderable through normal fallback.
const MAX_PROVIDED_FONT_FACES = 64;
const MAX_PROVIDED_FONT_BYTES = 128 * 1024 * 1024;

async function resourceAlias(bytes: Uint8Array): Promise<string | null> {
  if (!globalThis.crypto?.subtle) return null;
  const source = bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
  const digest = new Uint8Array(await globalThis.crypto.subtle.digest('SHA-256', source));
  return `__ooxml_provided_${[...digest]
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')}`;
}

/** Validate and bound resources before a render-worker structured clone. */
export function snapshotDocxFontResources(
  resources: readonly DocxFontResource[] | undefined,
): readonly DocxFontResource[] {
  if (!resources?.length) return [];
  const snapshot: DocxFontResource[] = [];
  const seen = new Set<string>();
  let bytes = 0;
  for (const resource of resources) {
    if (snapshot.length >= MAX_PROVIDED_FONT_FACES) break;
    const family = resource.family.trim();
    const weight = resource.weight ?? 400;
    const style = resource.style ?? 'normal';
    if (!family || !(resource.bytes instanceof Uint8Array)
      || !embeddedFontBytesAreWithinLimit(resource.bytes)) continue;
    const key = tupleKey(family, weight, style);
    if (seen.has(key) || bytes > MAX_PROVIDED_FONT_BYTES - resource.bytes.byteLength) continue;
    seen.add(key);
    bytes += resource.bytes.byteLength;
    snapshot.push(Object.freeze({
      family,
      // Own the admitted bytes. Callers may reuse or mutate their source after
      // load begins without racing OpenType parsing or FontFace construction.
      bytes: resource.bytes.slice(),
      weight,
      style,
    }));
  }
  return Object.freeze(snapshot);
}

function tupleKey(
  family: string,
  weight: 400 | 700,
  style: 'normal' | 'italic',
): string {
  return `${normalizeFontMetricFamily(family)}:${weight}:${style}`;
}

/** Register caller-owned font resources and derive immutable layout metrics.
 * Duplicate family/style tuples are first-wins, matching CSS face selection
 * and preventing metrics from one byte resource describing another face.
 */
export async function loadDocxFontResources(
  resources: readonly DocxFontResource[] | undefined,
): Promise<LoadedDocxFontResources> {
  if (!resources?.length) return { faces: [], metrics: {} };

  const entries: Array<{
    face: EmbeddedFontFace;
    metric: ResolvedFontMetric;
    key: string;
    loadedKey: string;
  }> = [];
  const seen = new Set<string>();
  let acceptedBytes = 0;
  for (const resource of resources) {
    if (entries.length >= MAX_PROVIDED_FONT_FACES) break;
    const family = resource.family.trim();
    const weight = resource.weight ?? 400;
    const style = resource.style ?? 'normal';
    if (!family || !(resource.bytes instanceof Uint8Array)
      || !embeddedFontBytesAreWithinLimit(resource.bytes)) continue;
    const key = tupleKey(family, weight, style);
    if (seen.has(key)) continue;
    seen.add(key);
    if (acceptedBytes > MAX_PROVIDED_FONT_BYTES - resource.bytes.byteLength) continue;

    const openType = parseOpenTypeResourceMetrics(resource.bytes);
    const design = openType ? wordOpenTypeDesignLineRatios(openType) : null;
    if (!openType?.unicodeRanges?.length || !design) continue;
    acceptedBytes += resource.bytes.byteLength;
    // A content-derived alias makes CSS selection unambiguous and keeps the
    // same resource identity deterministic across main/worker contexts.
    const alias = await resourceAlias(resource.bytes);
    if (!alias) continue;
    const loadedKey = tupleKey(alias, weight, style);
    const eastAsianLineHeightRatio = openType.hasEastAsianCmap
      ? wordOpenTypeEastAsianSingleLineRatio(openType)
      : 0;
    entries.push({
      key,
      loadedKey,
      face: {
        family: alias,
        bytes: resource.bytes,
        odttf: false,
        weight: weight === 700 ? 'bold' : 'normal',
        style,
      },
      metric: Object.freeze({
        family: alias,
        requestedFamily: family,
        weight,
        style,
        sourceIdentity: `provided-sfnt:${alias}`,
        synthesized: false,
        unicodeRanges: openType.unicodeRanges,
        ...design,
        ...(eastAsianLineHeightRatio > 0 ? { eastAsianLineHeightRatio } : {}),
      }),
    });
  }
  if (entries.length === 0) return { faces: [], metrics: {} };

  const faces = await registerEmbeddedFonts(entries.map((entry) => entry.face));
  const loaded = new Set(faces.map((face) => tupleKey(
    face.family.trim().replace(/^(['"])(.*)\1$/, '$2'),
    face.weight === 'bold' ? 700 : 400,
    face.style === 'italic' ? 'italic' : 'normal',
  )));
  const metrics: Record<string, ResolvedFontMetric> = {};
  for (const entry of entries) {
    if (!loaded.has(entry.loadedKey)) continue;
    const requestedFamily = normalizeFontMetricFamily(entry.metric.requestedFamily ?? '');
    metrics[entry.face.weight === 'normal' && entry.face.style === 'normal'
      ? requestedFamily
      : entry.key] = entry.metric;
  }
  return { faces, metrics: Object.freeze(metrics) };
}

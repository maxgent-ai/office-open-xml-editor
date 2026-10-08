import { cjkLangFromLanguage, type CjkLang } from '@silurus/ooxml-core';
import type { LayoutDiagnostic } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import {
  createCanvasFontRoute,
  type CanvasFontRoute,
  type FontSubstituteScript,
} from '@silurus/ooxml-core';

export type FontResolutionSource = 'embedded' | 'local' | 'css' | 'google' | 'substitute' | 'native' | 'generic';
export type FontStyle = 'normal' | 'italic';

export interface FontRequest {
  readonly cjkFallback?: CjkLang;
  readonly language?: string;
  readonly requestedFamily?: string | null;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly weight?: number;
  readonly style?: FontStyle;
  /** Script that the requested text belongs to, when it belongs to one that a
   * scoped visual substitute covers. A scoped inventory face answers only such
   * a request; any other request resolves as if that face did not exist. */
  readonly script?: FontSubstituteScript;
}

export interface FontResolution {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly route: CanvasFontRoute;
  readonly source: FontResolutionSource;
  /** Identity of the registered resource that supplied this face, when known. */
  readonly resourceIdentity?: string;
  readonly weight: number;
  readonly style: FontStyle;
  readonly diagnostics: readonly LayoutDiagnostic[];
  readonly genericFamily: 'serif' | 'sans-serif' | 'monospace';
}

export interface FontResolver {
  readonly fingerprint: string;
  resolve(request: Readonly<FontRequest>): FontResolution;
  /** Registry scope used only to delimit the shared run-context proof rule.
   * It does not authorize changing a scalar's slot or selected face. */
  configuredSubstituteScript?(requestedFamily: string | null | undefined): FontSubstituteScript | undefined;
  /** Script of the scoped substitute that actually wins this resource tuple.
   * Authored embedded/local/installed faces retain §17.3.2.26 scalar slots. */
  scopedSubstituteScript?(requestedFamily: string | null | undefined, weight?: number, style?: FontStyle): FontSubstituteScript | undefined;
}

export interface FontInventoryFace {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly source: Exclude<FontResolutionSource, 'generic' | 'native'>;
  readonly resourceIdentity?: string;
  readonly weight?: number;
  readonly style?: FontStyle;
  /** Visual substitute limited to one script (core `substitute-script.ts`). */
  readonly script?: FontSubstituteScript;
}

export interface FontResolverOptions {
  /** Immutable regional routes; their concrete outputs participate in cache identity. */
  readonly regionalFamilyLists?: Partial<Record<CjkLang, Readonly<Record<string, string>>>>;
  /** Stable DOCX fallback routes derived from document metadata and rendered faces. */
  readonly nativeFamilyLists?: Readonly<Record<string, string>>;
  /**
   * Authored families (normalized) whose web substitute is script-scoped (core
   * substitute-script.ts), with the substitute families of that script. For a
   * request not in that script, those families are removed from EVERY route
   * of the family (explicit, native and regional CSS lists), so Canvas can
   * never select them for other characters, even as a CSS fallback.
   */
  readonly scriptScopedFamilies?: Readonly<Record<string, Readonly<{
    script: FontSubstituteScript;
    substituteFamilies: readonly string[];
  }>>>;
}

function normalizeFamily(value: string): string {
  return value.trim().toLocaleLowerCase('en-US');
}

function normalizedWeight(value: number | undefined): number {
  if (value == null || !Number.isFinite(value)) return 400;
  return Math.min(900, Math.max(100, Math.round(value / 100) * 100));
}

function freezeResolution(value: FontResolution): FontResolution {
  return Object.freeze({ ...value, diagnostics: Object.freeze([...value.diagnostics]) });
}

function quoteCssFamily(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`;
}

function cssFamilyList(family: string, generic: FontResolution['genericFamily']): string {
  return `${quoteCssFamily(family)}, ${generic}`;
}

/**
 * Snapshot the font inventory used by one document. ECMA-376 §17.8.2 leaves
 * the substitution algorithm implementation-defined, so a substituted or
 * generic result is carried as an explicit diagnostic instead of being hidden
 * in paragraph geometry.
 */
export function createFontResolver(
  inventory: readonly FontInventoryFace[],
  options: Readonly<FontResolverOptions> = {},
): FontResolver {
  const sourcePriority: Readonly<Record<FontInventoryFace['source'], number>> = {
    embedded: 0,
    local: 1,
    css: 2,
    google: 3,
    substitute: 4,
  };
  const faces = inventory
    .filter((face) => face.requestedFamily.trim() && face.resolvedFamily.trim())
    .map((face) => Object.freeze({
      ...face,
      weight: normalizedWeight(face.weight),
      style: face.style ?? 'normal',
    }))
    .sort((a, b) => {
      const family = normalizeFamily(a.requestedFamily).localeCompare(normalizeFamily(b.requestedFamily));
      return family || sourcePriority[a.source] - sourcePriority[b.source]
        || a.resolvedFamily.localeCompare(b.resolvedFamily)
        || a.weight - b.weight
        || a.style.localeCompare(b.style);
    });
  const byFamily = new Map<string, (typeof faces)[number][]>();
  for (const face of faces) {
    const key = normalizeFamily(face.requestedFamily);
    byFamily.set(key, [...(byFamily.get(key) ?? []), face]);
  }
  const nativeFamilyLists = Object.freeze(Object.fromEntries(
    Object.entries(options.nativeFamilyLists ?? {})
      .filter(([family, familyList]) => family.trim() && familyList.trim())
      .map(([family, familyList]) => [normalizeFamily(family), familyList] as const)
      .sort(([a], [b]) => a.localeCompare(b)),
  ));
  const regionalFamilyLists = Object.freeze(Object.fromEntries(
    Object.entries(options.regionalFamilyLists ?? {}).map(([region, lists]) => [
      region,
      Object.freeze(Object.fromEntries(Object.entries(lists)
        .map(([family, list]) => [normalizeFamily(family), list])
        .sort(([a], [b]) => a.localeCompare(b)))),
    ]).sort(([a], [b]) => String(a).localeCompare(String(b))),
  )) as Readonly<Partial<Record<CjkLang, Readonly<Record<string, string>>>>>;
  const familyListFor = (
    family: string,
    language: string | undefined,
    fallback: CjkLang | undefined,
  ): string | undefined => {
    const region = cjkLangFromLanguage(language) ?? fallback;
    return (region ? regionalFamilyLists[region]?.[normalizeFamily(family)] : undefined)
      ?? nativeFamilyLists[normalizeFamily(family)];
  };
  const scriptScopedFamilies = Object.freeze(Object.fromEntries(
    Object.entries(options.scriptScopedFamilies ?? {})
      .map(([family, scope]) => [normalizeFamily(family), Object.freeze({
        script: scope.script,
        substituteFamilies: Object.freeze([...scope.substituteFamilies]
          .map((name) => normalizeFamily(name)).sort()),
      })] as const)
      .sort(([a], [b]) => a.localeCompare(b)),
  ));
  // Remove a scoped substitute's families from a CSS family list. Names never
  // contain commas in these generated lists; each entry is a quoted family or
  // a CSS generic keyword.
  const withoutScopedSubstitutes = (
    familyList: string,
    requestedFamily: string,
    script: FontSubstituteScript | undefined,
  ): string => {
    const scope = scriptScopedFamilies[normalizeFamily(requestedFamily)];
    if (!scope || scope.script === script) return familyList;
    return familyList.split(',').map((entry) => entry.trim())
      .filter((entry) => !scope.substituteFamilies.includes(
        normalizeFamily(entry.replace(/^"(.*)"$/u, '$1').replaceAll('\\"', '"').replaceAll('\\\\', '\\')),
      ))
      .join(', ');
  };
  const fingerprint = stableFingerprint('fonts', {
    faces, nativeFamilyLists, regionalFamilyLists, scriptScopedFamilies,
  });

  // Every shaped script span, line segment and east-Asian line floor asks for a
  // resolution, and each fresh answer carried its own copy of the complete CSS
  // family list plus a percent-encoded route fingerprint (together ~1.7 KB for
  // a registered local face). A long document retained hundreds of thousands
  // of value-identical copies across the text caches and the retained layout.
  // `resolve` is a pure function of the request fields below and this
  // resolver's immutable snapshot, and every resolution is deep-frozen, so
  // equal requests share one answer. The raw fields are the key (not their
  // normalized forms), so each entry is exactly what that request computed.
  // The bound only caps pathological documents; a miss recomputes.
  const resolutionMemoLimit = 4096;
  const resolutions = new Map<string, FontResolution>();
  // One resource selector serves resolution and the scope gate. Inventory
  // priority, exact tuple matching and script admission cannot diverge.
  const selectedFace = (family: string, weight: number, style: FontStyle, script?: FontSubstituteScript) =>
    (byFamily.get(normalizeFamily(family)) ?? []).find((candidate) =>
      candidate.weight === weight && candidate.style === style
      && (candidate.script === undefined || candidate.script === script));
  const resolveUncached = (request: Readonly<FontRequest>): FontResolution => {
    const requestedFamily = request.requestedFamily?.trim() || request.genericFamily || 'sans-serif';
    const weight = normalizedWeight(request.weight);
    const style = request.style ?? 'normal';
    const face = selectedFace(requestedFamily, weight, style, request.script);
    if (face) {
      const diagnostics: LayoutDiagnostic[] = face.source === 'substitute'
        ? [{
            code: 'UNSUPPORTED_FEATURE',
            severity: 'warning',
            message: `ECMA-376 §17.8.2 implementation-dependent font substitution: ${requestedFamily} resolved to ${face.resolvedFamily}`,
          }]
        : [];
      const fallbackList = familyListFor(requestedFamily, request.language, request.cjkFallback);
      const familyList = fallbackList
        ? `${quoteCssFamily(face.resolvedFamily)}, ${fallbackList}`
        : cssFamilyList(face.resolvedFamily, request.genericFamily ?? 'sans-serif');
      return freezeResolution({
        requestedFamily,
        resolvedFamily: face.resolvedFamily,
        route: createCanvasFontRoute(
          withoutScopedSubstitutes(familyList, requestedFamily, request.script),
          'registered',
        ),
        source: face.source,
        ...(face.resourceIdentity === undefined ? {} : { resourceIdentity: face.resourceIdentity }),
        weight,
        style,
        diagnostics,
        genericFamily: request.genericFamily ?? 'sans-serif',
      });
    }

    const generic = request.genericFamily ?? 'sans-serif';
    const authored = request.requestedFamily?.trim();
    if (authored) {
      const familyList = familyListFor(authored, request.language, request.cjkFallback)
        ?? cssFamilyList(authored, generic);
      return freezeResolution({
        requestedFamily,
        resolvedFamily: authored,
        route: createCanvasFontRoute(
          withoutScopedSubstitutes(familyList, authored, request.script),
          'native',
        ),
        source: 'native',
        weight,
        style,
        diagnostics: [],
        genericFamily: generic,
      });
    }
    return freezeResolution({
      requestedFamily,
      resolvedFamily: generic,
      route: createCanvasFontRoute(
        familyListFor(generic, request.language, request.cjkFallback) ?? generic,
        'generic',
      ),
      source: 'generic',
      weight,
      style,
      diagnostics: [],
      genericFamily: generic,
    });
  };

  return Object.freeze({
    fingerprint,
    configuredSubstituteScript(requestedFamily: string | null | undefined): FontSubstituteScript | undefined {
      return requestedFamily ? scriptScopedFamilies[normalizeFamily(requestedFamily)]?.script : undefined;
    },
    scopedSubstituteScript(requestedFamily: string | null | undefined, weight?: number, style?: FontStyle): FontSubstituteScript | undefined {
      const family = requestedFamily?.trim();
      if (!family) return undefined;
      const script = scriptScopedFamilies[normalizeFamily(family)]?.script;
      if (!script) return undefined;
      // Same priority and tuple selection as resolveUncached: a configured
      // substitute is insufficient when an authored resource outranks it.
      const face = selectedFace(family, normalizedWeight(weight), style ?? 'normal', script);
      return face?.source === 'substitute' && face.script === script ? script : undefined;
    },
    resolve(request: Readonly<FontRequest>): FontResolution {
      const key = JSON.stringify([
        request.requestedFamily ?? null,
        request.genericFamily ?? null,
        request.weight ?? null,
        request.style ?? null,
        request.language ?? null,
        request.cjkFallback ?? null,
        request.script ?? null,
      ]);
      const retained = resolutions.get(key);
      if (retained) return retained;
      const resolution = resolveUncached(request);
      resolutions.set(key, resolution);
      if (resolutions.size > resolutionMemoLimit) {
        resolutions.delete(resolutions.keys().next().value as string);
      }
      return resolution;
    },
  });
}

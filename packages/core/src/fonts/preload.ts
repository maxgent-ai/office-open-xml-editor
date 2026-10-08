/**
 * Google Fonts preload utility shared by docx / pptx / xlsx viewers.
 *
 * The contract is intentionally narrow: callers pass the set of font-family
 * names they want available, plus a static map from a lower-cased key to a
 * Google Fonts CSS URL (and optionally an alternate FontFaceSet family name
 * for Office substitutes such as Calibri → Carlito). Names without a map
 * entry are skipped (the renderer falls back to the system font).
 *
 * Rather than inject a `<link rel="stylesheet">` and read `document.fonts`
 * (which is impossible inside a Web Worker), this fetches the Google Fonts CSS
 * directly, parses its `@font-face` rules, and registers `FontFace` objects
 * into whichever FontFaceSet exists in the current JS context —
 * `document.fonts` on the main thread, `self.fonts` in a worker. This keeps the
 * loader FontFaceSet-agnostic so both the main-thread and worker rendering
 * modes share one code path.
 *
 * Font load is forced via `face.load()` rather than `FontFaceSet.load()`
 * because canvas-only rendering does not put glyphs into the DOM, so the
 * unicode-range gating in modern Google Fonts CSS would otherwise leave the
 * `FontFace` entries in the `unloaded` state — the first paint would then
 * use a system fallback and shift once a later interaction re-rasterized
 * the canvas after the font landed.
 */
import { retainFace, releaseFaces } from './font-registry.js';
import type { FontSubstituteScript } from './substitute-script.js';

export interface FontPreloadEntry {
  /** Google Fonts CSS URL — `display=swap` recommended. */
  url: string;
  /**
   * Family name to drive {@link FontFaceSet} loading when the substitute
   * differs from the requested face (e.g. Calibri → Carlito). Defaults to
   * the requested name when omitted.
   */
  loadFamily?: string;
  /**
   * Script scope of a visual substitute. Such a substitute may paint and
   * measure only characters of this script. See `substitute-script.ts`.
   */
  script?: FontSubstituteScript;
}

/**
 * Hard ceiling so a wedged network (a stylesheet fetch that never settles, or a
 * `FontFace.load()` that never resolves) cannot hang the caller forever. This is
 * a SAFETY NET, not the normal exit: on a reachable network every awaited
 * promise settles well within it, so first paint is deterministic. It is
 * intentionally generous — the previous 3 s timeout RACED the font loads and
 * could resolve while faces were still downloading, which is exactly the
 * cold-cache flicker this function must prevent.
 */
const HARD_CEILING_MS = 15000;

/** Race a font-load promise against a generous hard ceiling so a wedged network
 *  or a `FontFace.load()` that never settles cannot hang the caller forever.
 *  Shared by the Google-Fonts and embedded-font loaders (same first-paint
 *  determinism contract). */
export function withFontCeiling<T>(p: Promise<T>): Promise<T | void> {
  return Promise.race([
    p,
    new Promise<void>((resolve) => setTimeout(resolve, HARD_CEILING_MS)),
  ]);
}

/** In-flight (or completed) stylesheet FETCH promise per CSS url, keyed by url.
 *  Resolves with the PARSED `@font-face` rules of that stylesheet (empty on a
 *  failed fetch). This dedups only the NETWORK fetch: a concurrent caller with
 *  the same url JOINs the first fetch instead of re-downloading, and the join
 *  cannot resolve before the fetch settles (first-paint determinism). The actual
 *  `FontFace` objects are dedup + refcounted separately per call in the shared
 *  {@link ./font-registry.ts} (so holders in one FontFaceSet share one face;
 *  another document set receives its own face). The stored
 *  promise ALWAYS resolves: a failed fetch records the failed families + deletes
 *  the entry (so a later call retries) inside the producing call, so a cache-hit
 *  awaiter never throws. */
const cssFetches = new Map<string, Promise<ParsedFontFace[]>>();

/** Test hook — clears the per-context CSS fetch cache. */
export function _resetCssCacheForTests(): void {
  cssFetches.clear();
}

export interface ParsedFontFace {
  family: string;
  src: string;
  descriptors: FontFaceDescriptors;
}

/** Extract @font-face rules from a Google Fonts stylesheet. Deliberately
 *  minimal: Google's CSS is machine-generated (one declaration per line, no
 *  nesting), so a brace-block regex is sufficient and avoids a CSS parser. */
export function parseFontFaceRules(css: string, stylesheetUrl?: string): ParsedFontFace[] {
  const faces: ParsedFontFace[] = [];
  const blockRe = /@font-face\s*\{([^}]*)\}/g;
  let m: RegExpExecArray | null;
  while ((m = blockRe.exec(css))) {
    const body = m[1];
    const prop = (name: string): string | undefined =>
      body.match(new RegExp(`(?:^|;|\\n)\\s*${name}\\s*:\\s*([^;]+)`, 'i'))?.[1].trim();
    const familyRaw = prop('font-family');
    const rawSrc = prop('src');
    if (!familyRaw || !rawSrc) continue;
    const src = stylesheetUrl
      ? rawSrc.replace(
          /url\(\s*(?:(['"])(.*?)\1|([^)]*?))\s*\)/gi,
          (original, _quote: string | undefined, quoted: string | undefined, unquoted: string | undefined) => {
            const value = (quoted ?? unquoted ?? '').trim();
            if (!value) return original;
            try {
              return `url("${new URL(value, stylesheetUrl).href}")`;
            } catch {
              return original;
            }
          },
        )
      : rawSrc;
    const descriptors: FontFaceDescriptors = {};
    const style = prop('font-style');
    if (style) descriptors.style = style;
    const weight = prop('font-weight');
    if (weight) descriptors.weight = weight;
    const stretch = prop('font-stretch');
    if (stretch) descriptors.stretch = stretch;
    const unicodeRange = prop('unicode-range');
    if (unicodeRange) descriptors.unicodeRange = unicodeRange;
    faces.push({ family: familyRaw.replace(/^['"]|['"]$/g, ''), src, descriptors });
  }
  return faces;
}

/** The FontFaceSet of the current context: `document.fonts` on the main
 *  thread, `self.fonts` in a worker, null elsewhere (Node without a shim).
 *  Exported so the embedded-font loader shares one FontFaceSet-resolution rule
 *  with the Google-Fonts loader (both must register into the SAME set). */
export function activeFontSet(): FontFaceSet | null {
  if (typeof document !== 'undefined' && document && document.fonts) return document.fonts;
  if (typeof self !== 'undefined' && self && 'fonts' in self) {
    return (self as unknown as { fonts: FontFaceSet }).fonts;
  }
  return null;
}

/** Stable signature for one Google-Fonts `@font-face`, keyed to the CSS url it
 *  came from + its identity (family + all CSS descriptors). `gfonts:` namespaces
 *  it away from embedded-font signatures in the shared registry. Two documents
 *  requesting the same web font compute the SAME signature and, within one
 *  FontFaceSet, share one refcounted FontFace; distinct subsets
 *  key distinctly. */
function googleFaceSignature(url: string, f: ParsedFontFace): string {
  const d = f.descriptors;
  return [
    'gfonts',
    url,
    f.family.toLowerCase(),
    d.style ?? '',
    d.weight ?? '',
    d.stretch ?? '',
    d.unicodeRange ?? '',
    // `src` last: two rules identical in every descriptor but pointing at
    // different files (theoretically) still key apart.
    f.src,
  ].join('|');
}

/** A proven complete scalar overapproximation, or conservative all-rules demand. */
export type FontPreloadDemand = 'all' | readonly number[];

/** CSS Fonts unicode-range is an exclusion hint, never a support certificate.
 * Reject the whole descriptor unless every component is recognized. */
function unicodeRanges(value?: string): readonly [number, number][] | null {
  if (!value?.trim()) return null;
  const ranges: [number, number][] = [];
  for (const component of value.split(',')) {
    const part = component.trim();
    let lo: number, hi: number;
    const match = /^U\+([0-9A-F]{1,6})(?:-([0-9A-F]{1,6}))?$/i.exec(part);
    const wildcard = /^U\+([0-9A-F]{0,5})(\?{1,6})$/i.exec(part);
    if (match) { lo = parseInt(match[1], 16); hi = parseInt(match[2] ?? match[1], 16); }
    else if (wildcard && wildcard[1].length + wildcard[2].length <= 6) {
      lo = parseInt(wildcard[1] + '0'.repeat(wildcard[2].length), 16);
      hi = parseInt(wildcard[1] + 'F'.repeat(wildcard[2].length), 16);
    } else return null;
    if (lo > hi || hi > 0x10ffff) return null;
    ranges.push([lo, hi]);
  }
  return ranges;
}

function rangeIntersects(ranges: readonly [number, number][] | null, points: readonly number[]): boolean {
  if (ranges === null) return true;
  // Sorted once per ensure, not once per face. Lower bounds avoid a product of
  // retained document scalars and every CSS interval without changing matches.
  for (const [lo, hi] of ranges) {
    let left = 0, right = points.length;
    while (left < right) {
      const mid = (left + right) >>> 1;
      if (points[mid] < lo) left = mid + 1;
      else right = mid;
    }
    if (left < points.length && points[left] <= hi) return true;
  }
  return false;
}

interface OwnedGoogleFace {
  face: FontFace;
  ranges: readonly [number, number][] | null;
  readiness?: Promise<void>;
}

/** Presentation-owned monotone lease. Every CSS rule is registered in original
 * order, including excluded rules: lazy insertion can reverse overlapping
 * unicode-range precedence when later slides introduce another scalar. Only
 * explicit binary loads are filtered. Omitted/unknown demand preserves legacy
 * DOCX/XLSX behavior. No range match changes font selection or proves cmap/GSUB
 * support. CSS Font Loading §2.2 load() joins a shared face's readiness. */
export class GoogleFontPreloadLease {
  private closed = false;
  private all = false;
  private readonly points = new Set<number>();
  private readonly urls = new Set<string>();
  private readonly owned = new Map<string, OwnedGoogleFace>();
  private queue: Promise<unknown> = Promise.resolve();

  constructor(
    private readonly map: Record<string, FontPreloadEntry>,
    private readonly fonts: FontFaceSet | null = activeFontSet(),
  ) {}

  get faces(): FontFace[] { return [...this.owned.values()].map(entry => entry.face); }

  ensure(names: Iterable<string | null | undefined>, demand: FontPreloadDemand = 'all'): Promise<void> {
    // Snapshot only the caller's delta; serialized ownership avoids duplicate
    // retains while stylesheet fetches overlap. The caller budgets this delta.
    const requested = [...names];
    const delta = demand === 'all' ? demand : [...demand];
    const task = this.queue.then(() => this.ensureNow(requested, delta));
    this.queue = task.catch(() => undefined);
    return task;
  }

  private async ensureNow(names: (string | null | undefined)[], demand: FontPreloadDemand): Promise<void> {
    if (this.closed || !this.fonts || typeof FontFace === 'undefined' || typeof fetch === 'undefined') return;
    const delta = new Set<number>();
    const becameAll = demand === 'all' && !this.all;
    if (demand === 'all') { this.all = true; this.points.clear(); }
    else if (!this.all) for (const cp of demand) if (!this.points.has(cp)) { this.points.add(cp); delta.add(cp); }
    const fresh = new Set<OwnedGoogleFace>();
    const urlTargets = new Map<string, Set<string>>();
    for (const name of names) {
      const key = name?.trim().toLowerCase();
      const entry = key && this.map[key];
      if (!entry) continue;
      const targets = urlTargets.get(entry.url) ?? new Set<string>();
      targets.add((entry.loadFamily ?? (name as string).trim()).toLowerCase());
      urlTargets.set(entry.url, targets);
    }
    const failed = new Set<string>();
    try {
      const groups = await withFontCeiling(Promise.all([...urlTargets.keys()].filter(url => !this.urls.has(url)).map(async url => {
        let fetching = cssFetches.get(url);
        if (!fetching) {
          fetching = (async () => {
            try {
              const res = await fetch(url);
              if (!res.ok) throw new Error(`HTTP ${res.status}`);
              return parseFontFaceRules(await res.text(), res.url || url);
            } catch { cssFetches.delete(url); return []; }
          })();
          cssFetches.set(url, fetching);
        }
        return { url, rules: await fetching };
      })));
      if (this.closed) return;
      for (const group of Array.isArray(groups) ? groups : []) {
        if (!group.rules.length) { for (const family of urlTargets.get(group.url) ?? []) failed.add(family); continue; }
        for (const rule of group.rules) {
          const signature = googleFaceSignature(group.url, rule);
          if (this.owned.has(signature)) continue;
          const { face } = retainFace(signature, this.fonts, () => {
            const created = new FontFace(rule.family, rule.src, rule.descriptors);
            try { this.fonts!.add(created); }
            catch (error) { this.fonts!.delete(created); throw error; }
            return created;
          });
          const owned = { face, ranges: unicodeRanges(rule.descriptors.unicodeRange) };
          this.owned.set(signature, owned);
          fresh.add(owned);
        }
        this.urls.add(group.url);
      }
      const deltaIndex = [...delta].sort((a, b) => a - b);
      const cumulativeIndex = fresh.size ? [...this.points].sort((a, b) => a - b) : deltaIndex;
      const waits: Promise<void>[] = [];
      for (const entry of this.owned.values()) {
        if (entry.readiness) { waits.push(entry.readiness); continue; }
        const inspected = fresh.has(entry) ? cumulativeIndex : deltaIndex;
        const eligible = (this.all && (becameAll || fresh.has(entry))) || entry.ranges === null || rangeIntersects(entry.ranges, inspected);
        if (!eligible) continue;
        // Cache our readiness, including failures; another owner still calls
        // idempotent load() to join the same browser font-status promise.
        entry.readiness ??= Promise.resolve().then(() => this.closed ? undefined : entry.face.load()).then(() => undefined, () => {
          failed.add(entry.face.family.replace(/['"]/g, '').toLowerCase());
        });
        waits.push(entry.readiness);
      }
      await withFontCeiling(Promise.all(waits));
      if (!this.closed && failed.size) console.warn(
        `[ooxml] failed to preload web font(s): ${[...failed].join(', ')}; falling back to system fonts (text may shift or differ).`,
      );
    } catch (error) { this.release(); throw error; }
  }

  release(): void {
    if (this.closed) return;
    this.closed = true;
    releaseFaces(this.faces);
    this.owned.clear(); this.points.clear(); this.urls.clear();
  }
}

export async function preloadGoogleFonts(
  fontNames: Iterable<string | null | undefined>,
  map: Record<string, FontPreloadEntry>,
  targetFontSet: FontFaceSet | null = activeFontSet(),
): Promise<FontFace[]> {
  const lease = new GoogleFontPreloadLease(map, targetFontSet);
  await lease.ensure(fontNames);
  return lease.faces;
}

/**
 * Release the Google-Fonts `FontFace` objects a document/presentation/workbook
 * preloaded (the array returned by {@link preloadGoogleFonts}). Refcounted +
 * dedup-safe via the shared {@link ./font-registry.ts}: each face is removed
 * from its FontFaceSet only when the LAST holder releases it, so a web font used
 * by two holders in the same document survives until both are destroyed. Double-release safe
 * (a face passed twice, or a re-release after full release, is a no-op) and safe
 * in a context without a FontFaceSet. Twin of `unregisterEmbeddedFonts`; called
 * from each viewer's `destroy()` to fix the SPA leak where every opened document
 * left its Google FontFace objects in `document.fonts` forever.
 */
export function unloadGoogleFonts(faces: Iterable<FontFace>): void {
  releaseFaces(faces);
}

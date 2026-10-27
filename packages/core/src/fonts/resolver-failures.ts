/**
 * One document's resolver outcomes that the realm cache does not keep.
 * Library resource policy, not OOXML behaviour.
 *
 * The realm resolution cache never keeps a failure (a later document retries
 * the same source), but every requester of one document must report one
 * outcome per family: an XLSX workbook's viewer-realm sessions measure with
 * the faces its worker paints. Every terminal non-success of a family within
 * the document (a failed answer, a malformed answer, a ceiling breach, the
 * font ceiling, a bridge that cannot answer) is recorded here, through
 * {@link recordFontResolverFailures}, keyed by source identity and scoped
 * family; every requester consults it before and after resolving.
 *
 * Light on purpose: the main-thread bridge host records a bootstrap failure
 * without loading the responder.
 *
 * The record lives on main only. A bridged call's deadline is main's (the
 * worker sends the time left of its ceiling with each request), so main sees
 * every terminal outcome of the worker's resolver calls and records it here.
 */
import type { BridgedFamilyAnswer, FontResolveOutcome } from './resolver-bridge.js';
import type { FamilyResolution, SourceCache } from './source-cache.js';
import type { FontResolverSource, FontSource } from './source.js';
import type { FontSubstituteScript } from './substitute-script.js';
import { fontFamilyKey, scopedFamilyKey, type FontSubstituteScripts } from './substitutions.js';

/** Failure reason by {@link fontResolverFailureKey}. */
export type FontResolverFailures = Map<string, string>;

/** Key of one scoped family of one source identity (`fontSourceId`). */
export function fontResolverFailureKey(sourceId: string, name: string, scripts: FontSubstituteScripts): string {
  return `${sourceId}\u0000${scopedFamilyKey(name, scripts)}`;
}

/** Record `reason` for `families` of `sourceId`. Private: every write goes
 * through {@link concludeFamilies}, which writes only undecided families. */
function recordFontResolverFailures(
  failures: FontResolverFailures | undefined,
  sourceId: string,
  families: readonly { readonly name: string; readonly scripts: FontSubstituteScripts }[],
  reason: string,
): void {
  if (!failures) return;
  for (const { name, scripts } of families) {
    const key = fontResolverFailureKey(sourceId, name, scripts);
    if (!failures.has(key)) failures.set(key, reason);
  }
}

/**
 * The outcome this document already reached for one scoped family of
 * `cache`'s source: a recorded failure, or a resolution (claimed or absent)
 * the realm cache keeps, which `holder` then holds (see
 * `SharedTaskMap.holdKept`). Undefined while undecided.
 */
export function decidedFamilyOutcome(
  failures: FontResolverFailures | undefined,
  cache: SourceCache,
  name: string,
  scripts: FontSubstituteScripts,
  holder: object,
): FamilyResolution | undefined {
  const reason = failures?.get(fontResolverFailureKey(cache.id, name, scripts));
  if (reason !== undefined) return { kind: 'failed', reason };
  return cache.families.holdKept(scopedFamilyKey(name, scripts), holder);
}

/**
 * The single writer of a document's resolver outcomes: conclude `families`
 * of `cache`'s source that reached no resolution of their own (a timeout, a
 * ceiling breach, an error, a failed answer, a bridge that cannot answer).
 * A family the document already decided keeps that outcome (held for
 * `holder` when it comes from the cache); only an undecided family fails
 * with `reason`, which is recorded for every requester of the document.
 * Returns lower-cased family → outcome.
 */
export function concludeFamilies(
  failures: FontResolverFailures | undefined,
  cache: SourceCache,
  families: readonly { readonly name: string; readonly scripts: FontSubstituteScripts }[],
  reason: string,
  holder: object,
): Map<string, FamilyResolution> {
  const outcomes = new Map<string, FamilyResolution>();
  for (const { name, scripts } of families) {
    const decided = decidedFamilyOutcome(failures, cache, name, scripts, holder);
    if (decided) { outcomes.set(fontFamilyKey(name), decided); continue; }
    recordFontResolverFailures(failures, cache.id, [{ name, scripts }], reason);
    outcomes.set(fontFamilyKey(name), { kind: 'failed', reason });
  }
  return outcomes;
}

/**
 * The scoped families a bridged request names, read leniently (the request
 * may be malformed): authored and fallback names with the substitution scope
 * the demand proves, as the responder derives them. Empty when nothing
 * readable is named.
 */
export function bridgedRequestFamilies(
  demand: unknown,
): { readonly name: string; readonly scripts: FontSubstituteScripts }[] {
  if (demand === null || typeof demand !== 'object') return [];
  const fields = demand as { families?: unknown; fallbackFamilies?: unknown; substituteScripts?: unknown };
  const names = [fields.families, fields.fallbackFamilies]
    .flatMap((list) => (Array.isArray(list) ? list.filter((name): name is string => typeof name === 'string') : []));
  const record = fields.substituteScripts;
  const scoped = record !== null && typeof record === 'object' && !Array.isArray(record)
    ? new Map(Object.entries(record).map(([family, list]) =>
      [fontFamilyKey(family), Array.isArray(list)
        ? [...new Set(list.filter((script): script is FontSubstituteScript => typeof script === 'string'))].sort()
        : []]))
    : undefined;
  return names.map((name) => ({
    name,
    scripts: scoped ? Object.freeze(scoped.get(fontFamilyKey(name)) ?? []) : 'all',
  }));
}

/** A bridged request's families split by whether this document already
 * decided them. */
export interface SettledBridgedRequest {
  /** Families with a terminal outcome in this document: a recorded failure,
   * or a resolution (claimed or absent) this realm's cache holds for a
   * requester of the document (the XLSX viewer session). */
  readonly settled: readonly BridgedFamilyAnswer[];
  /** Families no requester has decided yet. */
  readonly unsettled: readonly { readonly name: string; readonly scripts: FontSubstituteScripts }[];
  /** The document's handle on the source's cache (holder of `settled`). */
  readonly cache: SourceCache;
}

/** One family's outcome as the bridge sends it. */
export function bridgedAnswer(family: string, outcome: FamilyResolution): BridgedFamilyAnswer {
  return outcome.kind === 'claimed' ? { family, kind: 'claimed', declarations: outcome.declarations }
    : outcome.kind === 'failed' ? { family, kind: 'failed', reason: outcome.reason }
      : { family, kind: 'absent' };
}

/** The requesting document's handle on one resolver source's realm cache
 * (acquired once per source, released when the document is disposed); it is
 * also the holder of every entry the document receives an answer from. */
export type DocumentSourceCache = (sourceIndex: number, source: FontResolverSource) => SourceCache;

/**
 * Split a bridged request by the document's settled outcomes, or undefined
 * when it does not name one of the document's resolver sources. A family's
 * outcome is decided once, by whichever requester reaches it first, success
 * included; the bridge answers settled families from here whatever the state
 * of its responder, so the worker adopts the viewer realm's outcome. A
 * resolution taken from the cache is held by the requesting document (its
 * cache handle) until it is disposed, as a resolution it ran itself would be,
 * so a later requester of the document (its viewer realm) still finds it
 * after the requester that produced it released it.
 */
export function settledBridgedRequest(
  failures: FontResolverFailures,
  sources: readonly FontSource[],
  request: { readonly sourceIndex?: unknown; readonly resolverId?: unknown; readonly demand?: unknown },
  cacheFor: DocumentSourceCache,
): SettledBridgedRequest | undefined {
  const { sourceIndex, resolverId } = request;
  if (typeof sourceIndex !== 'number' || !Number.isSafeInteger(sourceIndex) || sourceIndex < 0) return undefined;
  const source = sources[sourceIndex];
  if (source?.kind !== 'font-resolver' || source.id !== resolverId) return undefined;
  const cache = cacheFor(sourceIndex, source);
  const settled: BridgedFamilyAnswer[] = [];
  const unsettled: { name: string; scripts: FontSubstituteScripts }[] = [];
  for (const { name, scripts } of bridgedRequestFamilies(request.demand)) {
    const decided = decidedFamilyOutcome(failures, cache, name, scripts, cache);
    if (decided) settled.push(bridgedAnswer(fontFamilyKey(name), decided));
    else unsettled.push({ name, scripts });
  }
  return { settled, unsettled, cache };
}

/**
 * The single way a bridged request ends without a resolution of its own
 * (its deadline, an unknown source, a responder that cannot load or answer):
 * families the document already decided adopt that outcome (held for the
 * document, see {@link settledBridgedRequest}), the rest fail with `reason`,
 * recorded for the whole document. A request that names none of the
 * document's sources fails as a whole and records nothing.
 */
export function concludeBridgedRequest(
  failures: FontResolverFailures,
  sources: readonly FontSource[],
  request: { readonly sourceIndex?: unknown; readonly resolverId?: unknown; readonly demand?: unknown },
  reason: string,
  cacheFor: DocumentSourceCache,
): FontResolveOutcome {
  const { sourceIndex, resolverId } = request;
  if (typeof sourceIndex !== 'number' || !Number.isSafeInteger(sourceIndex) || sourceIndex < 0) return { status: 'failed', reason };
  const source = sources[sourceIndex];
  if (source?.kind !== 'font-resolver' || source.id !== resolverId) return { status: 'failed', reason };
  const cache = cacheFor(sourceIndex, source);
  const outcomes = concludeFamilies(failures, cache, bridgedRequestFamilies(request.demand), reason, cache);
  return { status: 'families', families: [...outcomes].map(([family, outcome]) => bridgedAnswer(family, outcome)) };
}

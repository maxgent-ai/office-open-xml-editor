/**
 * `fontResolver` resolution through a realm's shared resolution cache
 * (design §4.4). Library resource policy, not OOXML behaviour.
 *
 * One path serves every requester in a realm: a painting session
 * (`source-session.ts`) and, on main, the worker bridge
 * (`resolver-bridge-responder.ts`). Each demanded family joins the cache
 * entry keyed by (source identity, family, substitution scope); families no
 * live entry covers are resolved by one shared application call. Whoever asks
 * first calls the function, later requesters (another document, the viewer
 * realm of an XLSX worker-mode workbook, the bridge for its worker) join the
 * in-flight or kept answer, so a resolver runs once per family per realm
 * while any requester holds the entry.
 */
import { createAbortError } from '../internal/abort.js';
import { nullRecord } from '../internal/own-key.js';
import { claimFromIndex, indexFontDeclarations } from './files-source.js';
import {
  concludeFamilies,
  decidedFamilyOutcome,
  fontResolverFailureKey,
  type FontResolverFailures,
} from './resolver-failures.js';
import { admitFontResolverResult, isBridgedFontSourceLimitError } from './resolver-admission.js';
import { SharedTask, type FamilyResolution, type SourceCache } from './source-cache.js';
import {
  fontResolverFunction,
  type FontDemand,
  type FontResolveContext,
  type FontResolverSource,
} from './source.js';
import { FONT_SUBSTITUTE_SCRIPTS, type FontSubstituteScript } from './substitute-script.js';
import { fontFamilyKey, scopedFamilyKey, type FontSubstituteScripts } from './substitutions.js';

/**
 * Per-family answers another realm's resolution produced (the worker bridge:
 * main resolved through its own cache). A runner returns this instead of a
 * declaration list; each family's declarations are admitted and claimed
 * again here, so the painting realm trusts nothing but the bytes.
 */
export class BridgedFamilyResolutions {
  constructor(readonly families: ReadonlyMap<string,
    | { readonly kind: 'claimed'; readonly declarations: readonly unknown[] }
    | { readonly kind: 'failed'; readonly reason: string }
    | { readonly kind: 'absent' }>) {}
}

/** What a runner learns about the call besides the demand. */
export interface ResolverCallInfo {
  /** Position of the source in the requesting document's `fonts`. */
  readonly sourceIndex: number;
  /** Milliseconds left of the requester's font ceiling for this ensure
   * (main mode's deadline is cumulative over the source walk; a bridged
   * call must expire at the same moment). */
  readonly remainingMs: number;
}

/** Runs one `font-resolver` source. The worker bridge supplies its own
 * implementation; the default calls this realm's function directly.
 * `sourceIndex` is the source's position in the requesting document's
 * `fonts`, which names the exact descriptor (two configurations of one
 * function share an id but not their substitutes). */
export type FontResolverRunner = (
  source: FontResolverSource,
  demand: FontDemand,
  context: FontResolveContext,
  call: ResolverCallInfo,
) => readonly unknown[] | BridgedFamilyResolutions | Promise<readonly unknown[] | BridgedFamilyResolutions>;

/** One demanded family with the substitution scope its text proves. */
export interface ResolverDemandedName {
  readonly name: string;
  readonly fallback: boolean;
  readonly scripts: FontSubstituteScripts;
}

const KNOWN_SUBSTITUTE_SCRIPTS: ReadonlySet<string> = new Set<FontSubstituteScript>(FONT_SUBSTITUTE_SCRIPTS);

/** Scope lookup for a demand's `substituteScripts` (undefined: no restriction;
 * a family absent from the record proves none). */
export function substituteScriptsOf(value: unknown): (name: string) => FontSubstituteScripts {
  if (value === undefined) return () => 'all';
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new TypeError('font demand substituteScripts must be an object');
  }
  const scripts = new Map<string, readonly FontSubstituteScript[]>();
  for (const [family, list] of Object.entries(value)) {
    if (!Array.isArray(list) || !list.every((script) => typeof script === 'string' && KNOWN_SUBSTITUTE_SCRIPTS.has(script))) {
      throw new TypeError(`font demand substituteScripts[${JSON.stringify(family)}] must list substitute scripts`);
    }
    scripts.set(fontFamilyKey(family), Object.freeze([...new Set(list as FontSubstituteScript[])].sort()));
  }
  return (name) => scripts.get(fontFamilyKey(name)) ?? [];
}

/** Main-realm resolver execution: the application function behind a
 * descriptor this realm's `fontResolver()` issued. */
export const runResolverInRealm: FontResolverRunner = (source, demand, context) => {
  const resolve = fontResolverFunction(source);
  if (!resolve) throw new Error('font resolver function is not available in this realm');
  return resolve(demand, context);
};

export type { FontResolverFailures };

/**
 * Resolve `families` through `source`'s shared cache on behalf of `holder`
 * (which keeps the entries it joined until `cache.families.release(holder)`).
 * One application call covers the families no live entry has; the batch is
 * handed out per family (each family takes and removes its own entry) and only
 * per-family results are cached, so unclaimed or rejected declarations are
 * unreachable once the call settles. Returns lower-cased family → answer.
 */
export async function resolveFamiliesThroughCache(options: {
  readonly source: FontResolverSource;
  readonly cache: SourceCache;
  readonly runner: FontResolverRunner;
  readonly families: readonly ResolverDemandedName[];
  readonly tuples: FontDemand['tuples'];
  readonly scalars: 'all' | readonly number[];
  readonly signal: AbortSignal;
  readonly holder: object;
  /** Position of `source` in the requesting document's `fonts`. */
  readonly sourceIndex: number;
  /** Milliseconds left of the requester's ceiling (see ResolverCallInfo). */
  readonly remainingMs: () => number;
  /** The requesting document's shared failure record, if any. */
  readonly failures?: FontResolverFailures;
}): Promise<Map<string, FamilyResolution>> {
  const { source, cache, runner, tuples, scalars, signal, holder, sourceIndex, remainingMs, failures } = options;
  const settled = new Map<string, FamilyResolution>();
  // A family the document already decided keeps that outcome.
  const families = options.families.filter(({ name, scripts }) => {
    const decided = decidedFamilyOutcome(failures, cache, name, scripts, holder);
    if (!decided) return true;
    settled.set(fontFamilyKey(name), decided);
    return false;
  });
  const needed = families.filter(({ name, scripts }) => !cache.families.has(scopedFamilyKey(name, scripts)));
  let call: SharedTask<Map<string, FamilyResolution>> | undefined;
  if (needed.length) {
    const neededKeys = new Set(needed.map(({ name }) => fontFamilyKey(name)));
    const demand: FontDemand = Object.freeze({
      families: Object.freeze(needed.filter((entry) => !entry.fallback).map(({ name }) => name)),
      fallbackFamilies: Object.freeze(needed.filter((entry) => entry.fallback).map(({ name }) => name)),
      tuples: Object.freeze(tuples.filter((tuple) => neededKeys.has(fontFamilyKey(tuple.family)))),
      scalars: scalars === 'all' ? 'all' : Object.freeze([...scalars]),
      ...(needed.some(({ scripts }) => scripts !== 'all') ? {
        // Keyed by document family names: no prototype, in either realm.
        substituteScripts: Object.freeze(Object.assign(nullRecord<readonly FontSubstituteScript[]>(),
          Object.fromEntries(needed.flatMap(({ name, scripts }) =>
            scripts === 'all' ? [] : [[fontFamilyKey(name), scripts] as const])))),
      } : {}),
    });
    call = new SharedTask((callSignal) => callResolver(runner, source, demand, callSignal, { sourceIndex, remainingMs: remainingMs() }), () => undefined);
  }
  const entries = await Promise.all(families.map(async ({ name, scripts }) => {
    const key = fontFamilyKey(name);
    const answer = await cache.families.join(scopedFamilyKey(name, scripts), async (taskSignal) => {
      if (!call) return { kind: 'absent' } satisfies FamilyResolution;
      const batch = await call.join(taskSignal);
      const own = batch.get(key) ?? { kind: 'absent' };
      batch.delete(key);
      return own;
    }, signal, holder);
    // A rejection (abort, timeout, ceiling breach) concludes the whole batch
    // at the requester, which knows why (`concludeFamilies`, per family).
    // A failed answer, or a failure another requester of this document
    // recorded while this answer was in flight, concludes this family.
    const reason = answer.kind === 'failed' ? answer.reason : failures?.get(fontResolverFailureKey(cache.id, name, scripts));
    if (reason === undefined) return [key, answer] as const;
    return [key, concludeFamilies(failures, cache, [{ name, scripts }], reason, holder).get(key) ?? answer] as const;
  }));
  return new Map([...settled, ...entries]);
}

/**
 * Run one shared resolver call. Module-level and given its runner explicitly:
 * the call is shared by every requester waiting on the same families, so it
 * must not read any requester's state. The runner of the requester that
 * created the call is kept by the call until it settles, so releasing that
 * requester never breaks the other waiters.
 */
async function callResolver(
  runResolver: FontResolverRunner,
  source: FontResolverSource,
  demand: FontDemand,
  signal: AbortSignal,
  info: ResolverCallInfo,
): Promise<Map<string, FamilyResolution>> {
  const names = [...demand.families, ...demand.fallbackFamilies];
  const failAll = (reason: string) => new Map(
    names.map((name) => [fontFamilyKey(name), { kind: 'failed', reason } satisfies FamilyResolution]));
  // The requester's deadline by wall clock as well as by timer: a resolver
  // that blocks the thread synchronously keeps every timer from firing, so
  // an answer that settles after the deadline is a timeout, never a success
  // (and, being a failure, is not cached). The realm that runs the function
  // judges this; a bridged answer (or failure) is main's verdict and is taken
  // as it is.
  const deadline = Date.now() + info.remainingMs;
  const late = () => runResolver === runResolverInRealm && Date.now() > deadline;
  let raw: readonly unknown[] | BridgedFamilyResolutions;
  try {
    raw = await runResolver(source, demand, Object.freeze({ signal }), info);
  } catch (error) {
    // A bridged call reports the main-side admission's ceiling breach, which
    // rejects the load exactly as this realm's own admission would.
    if (signal.aborted || isBridgedFontSourceLimitError(error)) throw error;
    if (late()) return failAll('timed out');
    return failAll(error instanceof Error ? error.message : String(error));
  }
  if (signal.aborted) throw createAbortError('Font resolver abandoned');
  if (late()) return failAll('timed out');
  const scriptsOf = substituteScriptsOf(demand.substituteScripts);
  const restricted = demand.substituteScripts !== undefined;
  const claim = (declarations: ReturnType<typeof indexFontDeclarations>, name: string) =>
    claimFromIndex(declarations, name, source.substitutes, restricted ? scriptsOf(name) : 'all');
  if (raw instanceof BridgedFamilyResolutions) {
    const bridged = raw.families;
    return new Map(names.map((name): [string, FamilyResolution] => {
      const key = fontFamilyKey(name);
      const answer = bridged.get(key);
      if (!answer || answer.kind === 'absent') return [key, { kind: 'absent' }];
      if (answer.kind === 'failed') return [key, { kind: 'failed', reason: answer.reason }];
      // The other realm's claim is re-derived from its declarations here.
      const admitted = admitFontResolverResult(answer.declarations);
      if (!admitted.ok) return [key, { kind: 'failed', reason: admitted.reason }];
      return [key, claim(indexFontDeclarations(admitted.declarations), name)];
    }));
  }
  const admitted = admitFontResolverResult(raw);
  if (!admitted.ok) return failAll(admitted.reason);
  const index = indexFontDeclarations(admitted.declarations);
  return new Map(names.map((name) => [fontFamilyKey(name), claim(index, name)]));
}

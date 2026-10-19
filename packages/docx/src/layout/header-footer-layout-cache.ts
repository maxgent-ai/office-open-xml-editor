import type { StoryLayoutCache } from './story-layout-cache.js';
import type { StoryLayout } from './types.js';

export const MAX_HEADER_FOOTER_CACHE_ENTRIES = 4096;
export const MAX_HEADER_FOOTER_CACHE_UNITS = 2_000_000;

/** Retention accounting charges every property, primitive and UTF-16 code unit.
 * It bounds retained graph size, not exact JS heap bytes. Shared objects within
 * one layout are counted once; sharing between entries is conservatively charged
 * again. Oversized layouts remain usable but are never admitted to this cache. */
function retainedUnits(value: unknown, seen = new WeakSet<object>()): number {
  if (typeof value === 'string') return 1 + value.length;
  if (!value || typeof value !== 'object') return 1;
  if (seen.has(value)) return 0;
  seen.add(value);
  let units = 1;
  for (const key in value) {
    if (!Object.hasOwn(value, key)) continue;
    units += key.length + retainedUnits((value as Record<string, unknown>)[key], seen);
    if (units > MAX_HEADER_FOOTER_CACHE_UNITS) break;
  }
  return units;
}

/** One pagination owns this LRU; no pass or DocumentLayout owns its handle.
 * Two placements per exact occurrence and aggregate entry/graph-size ceilings
 * prevent retaining whole histories of reserve, field or anchor trials. Eviction
 * only costs reacquisition; it never changes exact convergence acceptance. */
export function createHeaderFooterLayoutCache(chargeTrial: () => void): StoryLayoutCache {
  const entries = new Map<string, { occurrence: string; story: StoryLayout; units: number }>();
  const occurrences = new Map<string, Set<string>>();
  let units = 0;
  const remove = (key: string) => {
    const entry = entries.get(key)!;
    entries.delete(key);
    units -= entry.units;
    const keys = occurrences.get(entry.occurrence)!;
    keys.delete(key);
    if (keys.size === 0) occurrences.delete(entry.occurrence);
  };
  return { layout(request, acquire) {
    const key = JSON.stringify([request.occurrence, request.placement]);
    const known = entries.get(key);
    if (known) {
      entries.delete(key);
      entries.set(key, known);
      const keys = occurrences.get(request.occurrence)!;
      keys.delete(key);
      keys.add(key);
      return known.story;
    }
    if (request.trial) chargeTrial();
    const story = acquire();
    const weight = key.length + request.occurrence.length + retainedUnits(story);
    if (weight > MAX_HEADER_FOOTER_CACHE_UNITS) return story;
    let keys = occurrences.get(request.occurrence);
    if (!keys) occurrences.set(request.occurrence, keys = new Set());
    keys.add(key);
    entries.set(key, { occurrence: request.occurrence, story, units: weight });
    units += weight;
    while (keys.size > 2) remove(keys.values().next().value!);
    while (entries.size > MAX_HEADER_FOOTER_CACHE_ENTRIES || units > MAX_HEADER_FOOTER_CACHE_UNITS) {
      remove(entries.keys().next().value!);
    }
    return story;
  } };
}

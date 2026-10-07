import type { StoryLayout } from './types.js';

/**
 * How long one acquired story layout is retained ({@link createStoryLayoutCache}):
 * - `session`: for the session under its exact placement (a trial placement
 *   keeps only its occurrence's most recent placements instead);
 * - `latest`: only the occurrence's most recent placement, so a shared
 *   definition (a reserved note separator root) replaces its previous context
 *   instead of accumulating one per context;
 * - `none`: never, so destination-dependent full acquisitions (continued
 *   footnotes that are not proved page-independent) are not retained once per
 *   destination page.
 */
export type StoryLayoutRetention = 'session' | 'latest' | 'none';

/** Story layouts retained by one body layout session ({@link createStoryLayoutCache}). */
export interface StoryLayoutCache {
  /** The layout of `request`, acquired by `acquire` on a miss. `occurrence`
   * is the story source on its page (or the source alone for a proved
   * page-independent story); `placement` everything else the layout depends
   * on; `trial` marks a placement a convergence is solving for; `retention`
   * (default `session`) is the explicit admission of a miss. */
  layout(
    request: Readonly<{
      occurrence: string;
      placement: string;
      trial: boolean;
      retention?: StoryLayoutRetention;
    }>,
    acquire: () => StoryLayout,
  ): StoryLayout;
}

/** Trial layouts retained per story occurrence: the two most recently used,
 * as a paragraph keeps its two most recent placements (runtime-state.ts). */
export const RETAINED_TRIALS_PER_STORY = 2;

/**
 * The story layouts of one body layout session. A story occurrence laid out
 * without a trial placement keeps each layout for the session, as it always
 * did. A trial placement — a band translation or page frames that a
 * band-dependent story's convergence is solving for (body-paginator.ts
 * layoutBandStory and the footnote band plan; paragraph.ts text box solves,
 * nested in one another) — is speculative work: each miss
 * is charged to the session's acquisition budget (`chargeTrial`:
 * runtime-state.ts noteMiss, which fails closed), and only the most recently
 * used trials of an occurrence are retained. A solver's rejected candidates
 * are released with it while its accepted (last evaluated) one stays
 * reachable; a miss only costs re-acquisition, and acceptance stays exact.
 * The trial charge applies to every retention, including `none`, whose
 * requests are always misses. An acquisition that throws (an exhausted
 * budget, a failed solve) retains nothing: entries are stored only once
 * acquired, and a `latest` entry is replaced only by a completed one.
 */
export function createStoryLayoutCache(chargeTrial: () => void): StoryLayoutCache {
  const stable = new Map<string, StoryLayout>();
  const trials = new Map<string, Map<string, StoryLayout>>();
  const latest = new Map<string, Readonly<{ placement: string; story: StoryLayout }>>();
  return {
    layout(request, acquire) {
      const retention = request.retention ?? 'session';
      if (retention === 'none') {
        if (request.trial) chargeTrial();
        return acquire();
      }
      if (retention === 'latest') {
        const known = latest.get(request.occurrence);
        if (known?.placement === request.placement) return known.story;
        if (request.trial) chargeTrial();
        const acquired = acquire();
        latest.set(request.occurrence, Object.freeze({ placement: request.placement, story: acquired }));
        return acquired;
      }
      if (!request.trial) {
        const key = JSON.stringify([request.occurrence, request.placement]);
        const known = stable.get(key);
        if (known) return known;
        const acquired = acquire();
        stable.set(key, acquired);
        return acquired;
      }
      const known = trials.get(request.occurrence)?.get(request.placement);
      if (known) {
        const byPlacement = trials.get(request.occurrence)!;
        byPlacement.delete(request.placement);
        byPlacement.set(request.placement, known);
        return known;
      }
      chargeTrial();
      const acquired = acquire();
      // Looked up after acquisition, which may itself have retained trials.
      let byPlacement = trials.get(request.occurrence);
      if (!byPlacement) {
        byPlacement = new Map();
        trials.set(request.occurrence, byPlacement);
      }
      byPlacement.delete(request.placement);
      byPlacement.set(request.placement, acquired);
      while (byPlacement.size > RETAINED_TRIALS_PER_STORY) {
        byPlacement.delete(byPlacement.keys().next().value!);
      }
      return acquired;
    },
  };
}

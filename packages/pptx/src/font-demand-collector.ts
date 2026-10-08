import type { FontDemandCollector } from './presentation-preflight.js';

/** Analysis/data is optional: ordinary parser/viewer graphs do not execute or
 * fetch it. The builder's ACK transaction stays synchronous after injection. */
export async function loadFontDemandCollector(enabled: boolean): Promise<FontDemandCollector | undefined> {
  return enabled ? (await import('./font-preload-demand.js')).collectSlideFontDemand : undefined;
}

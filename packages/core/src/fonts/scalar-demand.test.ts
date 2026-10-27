import { describe, expect, it } from 'vitest';
import { HARD_MAX_FONT_DEMAND_SCALARS } from '../worker/resource-policy.generated.js';
import { FontDemandScalarCollector, collectFontDemandScalars } from './scalar-demand.js';

const SENTINELS = [0x4d, 0x69, 0x20]; // M, i, space
const BUDGET = 1 << 20;

describe('collectFontDemandScalars', () => {
  it('returns first-seen scalars after the measurement sentinels, deduplicated', () => {
    expect(collectFontDemandScalars(['Ab', 'b漢𠀋'], { budget: BUDGET }))
      .toEqual([...SENTINELS, 0x41, 0x62, 0x6f22, 0x2000b]);
  });

  it('answers all for every text whose painted scalars its spelling cannot prove', () => {
    for (const text of ['é', 'é', '\ud800', '͸', '', 'שלום', '‮(']) {
      expect(collectFontDemandScalars(['A', text], { budget: BUDGET }), JSON.stringify(text)).toBe('all');
    }
  });

  it('charges only uncommitted scalars and degrades to all beyond the JSON budget', () => {
    const committed = new Set([...SENTINELS, 0x41]);
    expect(collectFontDemandScalars(['AB'], { budget: BUDGET, committed })).toEqual([0x42]);
    // "[]" alone needs two bytes.
    expect(collectFontDemandScalars([], { budget: 1, committed })).toBe('all');
    expect(collectFontDemandScalars([], { budget: 2, committed })).toEqual([]);
    // "[66]" is four bytes; "[66,67]" seven.
    expect(collectFontDemandScalars(['B'], { budget: 5, committed })).toEqual([0x42]);
    expect(collectFontDemandScalars(['BC'], { budget: 6, committed })).toBe('all');
  });

  it('degrades to all beyond HARD_MAX_FONT_DEMAND_SCALARS, counting committed scalars', () => {
    const committedOf = (size: number) => new Set([...SENTINELS,
      ...Array.from({ length: size - SENTINELS.length }, (_, i) => 0x4e00 + i)]);
    expect(collectFontDemandScalars(['A'], { budget: BUDGET, committed: committedOf(HARD_MAX_FONT_DEMAND_SCALARS - 1) }))
      .toEqual([0x41]);
    expect(collectFontDemandScalars(['A'], { budget: BUDGET, committed: committedOf(HARD_MAX_FONT_DEMAND_SCALARS) }))
      .toBe('all');
  });

  it('stays all once declined, so later texts cannot resurrect a partial list', () => {
    const collector = new FontDemandScalarCollector({ budget: BUDGET });
    expect(collector.add('A')).toBe(true);
    expect(collector.decline()).toBe(false);
    expect(collector.add('B')).toBe(false);
    expect(collector.result()).toBe('all');
  });
});

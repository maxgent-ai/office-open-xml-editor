import { describe, expect, it } from 'vitest';
import { LayoutInvariantError } from './diagnostics.js';
import { RETAINED_TRIALS_PER_STORY, createStoryLayoutCache } from './story-layout-cache.js';
import type { StoryLayout } from './types.js';

const story = (label: string) => ({ story: 'textbox', label } as unknown as StoryLayout);

describe('createStoryLayoutCache', () => {
  it('keeps stable layouts for the session and charges no budget for them', () => {
    let charged = 0;
    let acquired = 0;
    const cache = createStoryLayoutCache(() => { charged += 1; });
    const acquire = (label: string) => () => { acquired += 1; return story(label); };
    for (let page = 0; page < 50; page += 1) {
      cache.layout({ occurrence: `header:${page}`, placement: 'p', trial: false }, acquire(`h${page}`));
    }
    for (let page = 0; page < 50; page += 1) {
      expect(cache.layout({ occurrence: `header:${page}`, placement: 'p', trial: false }, acquire('x')))
        .toMatchObject({ label: `h${page}` });
    }
    expect({ charged, acquired }).toEqual({ charged: 0, acquired: 50 });
  });

  it('charges every distinct trial of a nested solve and keeps only the most recent per occurrence', () => {
    let charged = 0;
    const cache = createStoryLayoutCache(() => { charged += 1; });
    // A nested solve: every outer candidate re-solves the inner story over
    // its own candidates; each distinct trial is acquired once and charged.
    for (let outer = 0; outer < 16; outer += 1) {
      cache.layout({ occurrence: 'outer', placement: `o${outer}`, trial: true }, () => {
        for (let inner = 0; inner < 16; inner += 1) {
          cache.layout({ occurrence: 'inner', placement: `${outer}:${inner}`, trial: true }, () => story('inner'));
        }
        return story(`outer${outer}`);
      });
    }
    expect(charged).toBe(16 + 16 * 16);
    // The most recent candidates (the accepted one last) stay reachable,
    // uncharged; an older one was released and is re-acquired, charged.
    const before = charged;
    for (let recent = 0; recent < RETAINED_TRIALS_PER_STORY; recent += 1) {
      const outer = 15 - recent;
      expect(cache.layout({ occurrence: 'outer', placement: `o${outer}`, trial: true }, () => story('again')))
        .toMatchObject({ label: `outer${outer}` });
    }
    expect(charged).toBe(before);
    const released = 15 - RETAINED_TRIALS_PER_STORY;
    expect(cache.layout({ occurrence: 'outer', placement: `o${released}`, trial: true }, () => story('re')))
      .toMatchObject({ label: 're' });
    expect(cache.layout({ occurrence: 'inner', placement: '0:0', trial: true }, () => story('inner re')))
      .toMatchObject({ label: 'inner re' });
    expect(charged).toBe(before + 2);
  });

  it('fails closed before acquiring a trial once the session budget is spent', () => {
    let remaining = 3;
    let acquired = 0;
    const cache = createStoryLayoutCache(() => {
      remaining -= 1;
      if (remaining < 0) throw new LayoutInvariantError('NON_CONVERGENCE', 'budget spent');
    });
    for (let index = 0; index < 3; index += 1) {
      cache.layout({ occurrence: 's', placement: `${index}`, trial: true }, () => { acquired += 1; return story('s'); });
    }
    expect(() => cache.layout({ occurrence: 's', placement: '3', trial: true }, () => {
      acquired += 1;
      return story('s');
    })).toThrow(/NON_CONVERGENCE/);
    expect(acquired).toBe(3);
  });

  it('replaces a latest-retained context and never retains an uncached one', () => {
    let charged = 0;
    const cache = createStoryLayoutCache(() => { charged += 1; });
    // A shared separator root: one active context, replaced (not accumulated)
    // when its placement changes, and free while it is unchanged.
    const separator = (placement: string, label: string) => cache.layout(
      { occurrence: 'separator', placement, trial: false, retention: 'latest' }, () => story(label));
    expect(separator('w1', 'a')).toMatchObject({ label: 'a' });
    expect(separator('w1', 'x')).toMatchObject({ label: 'a' });
    expect(separator('w2', 'b')).toMatchObject({ label: 'b' });
    expect(separator('w1', 'c')).toMatchObject({ label: 'c' });
    // A destination-dependent continued note: every request is acquired, and
    // a trial placement is charged before each acquisition.
    let acquired = 0;
    const note = (trial: boolean) => cache.layout(
      { occurrence: 'note:page:1', placement: 'band', trial, retention: 'none' },
      () => { acquired += 1; return story(`n${acquired}`); });
    expect(note(true)).toMatchObject({ label: 'n1' });
    expect(note(true)).toMatchObject({ label: 'n2' });
    expect(note(false)).toMatchObject({ label: 'n3' });
    expect(charged).toBe(2);
  });

  it('retains nothing from an acquisition that throws', () => {
    const cache = createStoryLayoutCache(() => {});
    for (const trial of [true, false]) {
      const request = { occurrence: `s:${trial}`, placement: 'p', trial };
      expect(() => cache.layout(request, () => {
        // A trial nested in the failing one is its own completed acquisition.
        cache.layout({ occurrence: 'nested', placement: `${trial}`, trial: true }, () => story('nested'));
        throw new LayoutInvariantError('NON_CONVERGENCE', 'solve failed');
      })).toThrow(/NON_CONVERGENCE/);
      expect(cache.layout(request, () => story('retry'))).toMatchObject({ label: 'retry' });
      expect(cache.layout({ occurrence: 'nested', placement: `${trial}`, trial: true }, () => story('x')))
        .toMatchObject({ label: 'nested' });
    }
  });
});

import { describe, it, expect } from 'vitest';
import {
  buildSlidePartIndex,
  resolveSlidePartTarget,
  resolveInternalSlideTarget,
} from './slide-nav';

// A three-slide deck's part names in sldIdLst order, as the parser stamps them.
const PART_NAMES = ['ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml'];

describe('buildSlidePartIndex', () => {
  it('maps each part name to its sldIdLst-order index', () => {
    const idx = buildSlidePartIndex(PART_NAMES);
    expect(idx.get('ppt/slides/slide1.xml')).toBe(0);
    expect(idx.get('ppt/slides/slide2.xml')).toBe(1);
    expect(idx.get('ppt/slides/slide3.xml')).toBe(2);
  });

  it('skips undefined / empty entries and keeps the first on a duplicate', () => {
    const idx = buildSlidePartIndex([
      'ppt/slides/slide1.xml',
      undefined,
      '',
      'ppt/slides/slide1.xml', // duplicate — first (index 0) wins
    ]);
    expect(idx.size).toBe(1);
    expect(idx.get('ppt/slides/slide1.xml')).toBe(0);
  });
});

describe('resolveSlidePartTarget', () => {
  const idx = buildSlidePartIndex(PART_NAMES);

  it('resolves a slide-rel target (../slides/slideN.xml) to its index', () => {
    // The exact hlinksldjump rel Target: authored relative to ppt/slides.
    expect(resolveSlidePartTarget('../slides/slide3.xml', idx)).toBe(2);
    expect(resolveSlidePartTarget('../slides/slide1.xml', idx)).toBe(0);
  });

  it('resolves a root-absolute slide target', () => {
    expect(resolveSlidePartTarget('/ppt/slides/slide2.xml', idx)).toBe(1);
  });

  it('resolves a canonical target inherited from a master or layout part', () => {
    expect(resolveSlidePartTarget('ppt/slides/slide2.xml', idx)).toBe(1);
  });

  it('returns undefined for a target that names no known slide', () => {
    expect(resolveSlidePartTarget('../slides/slide9.xml', idx)).toBeUndefined();
    expect(resolveSlidePartTarget('https://example.com/', idx)).toBeUndefined();
    expect(resolveSlidePartTarget('', idx)).toBeUndefined();
  });
});

describe('resolveInternalSlideTarget', () => {
  const idx = buildSlidePartIndex(PART_NAMES);

  it('resolves a specific slide-part jump', () => {
    expect(resolveInternalSlideTarget('../slides/slide3.xml', idx, 0)).toBe(2);
  });

  it('resolves relative show-jump verbs against the current slide', () => {
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=firstslide', idx, 1)).toBe(0);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=lastslide', idx, 1)).toBe(2);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=nextslide', idx, 1)).toBe(2);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=previousslide', idx, 1)).toBe(
      0,
    );
  });

  it('clamps nextslide/previousslide at the deck ends', () => {
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=nextslide', idx, 2)).toBe(2);
    expect(resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=previousslide', idx, 0)).toBe(
      0,
    );
  });

  it('returns undefined for an unresolvable internal ref', () => {
    // A show-action verb we don't navigate on, and no part match either.
    expect(
      resolveInternalSlideTarget('ppaction://hlinkshowjump?jump=endshow', idx, 0),
    ).toBeUndefined();
  });
});

// M2: main mode reads each slide's `partName` off the parsed model; worker mode
// SERIALIZES the same per-slide `partName` array into `PresentationMeta.partNames`
// (the array that rides through `postMessage`) and the main-thread proxy rebuilds
// the part-index map from it. Both feed the SAME `buildSlidePartIndex`, so a
// serialization drop or re-order is the only way an internal slide jump could
// diverge across modes. Pin that the two paths build an identical map + resolve
// every part name to the same index.
describe('slide part-index — main/worker serialization equivalence (M2)', () => {
  it('resolves identically whether built from the model array or the worker-meta array', () => {
    // main: presentation.ts does `(slides ?? []).map((s) => s.partName)`.
    const fromModel = PART_NAMES.map((p) => p); // simulate s.partName per slide
    // worker: render-worker.ts posts `partNames: pres.slides.map((s) => s.partName)`,
    // which arrives as the same array after structured-clone.
    const fromMeta: (string | undefined)[] = JSON.parse(JSON.stringify(fromModel));

    const mainIdx = buildSlidePartIndex(fromModel);
    const workerIdx = buildSlidePartIndex(fromMeta);

    expect(workerIdx.size).toBe(mainIdx.size);
    expect(mainIdx.size).toBeGreaterThan(0); // non-degenerate
    for (const name of PART_NAMES) {
      expect(resolveSlidePartTarget(`../slides/${name.split('/').pop()}`, workerIdx)).toBe(
        resolveSlidePartTarget(`../slides/${name.split('/').pop()}`, mainIdx),
      );
      expect(workerIdx.get(name)).toBe(mainIdx.get(name));
    }
  });
});

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
let { loadOfficeFontFallbacks, unloadOfficeFontFallbacks, probeFontFamilyPresence } = await import('./office-fallback.js');
import referenceData from './reference-font-metrics-data.json';

beforeEach(async () => {
  vi.resetModules();
  ({ loadOfficeFontFallbacks, unloadOfficeFontFallbacks, probeFontFamilyPresence } = await import('./office-fallback.js'));
});

const globals = globalThis as unknown as Record<string, unknown>;
const originals = Object.fromEntries(
  ['FontFace', 'fetch', 'document', 'self'].map((name) => [name, globals[name]]),
);

type Face = { family: string; source: string | ArrayBuffer; status: FontFaceLoadStatus };

function fontSet(
  installed: readonly string[],
  delayMs = 0,
  declared: readonly string[] = [],
  stalled: readonly string[] = [],
) {
  const added: Face[] = [];
  const deleted: Face[] = [];
  let active = 0;
  let peak = 0;
  const set = {
    add(face: Face) { added.push(face); },
    delete(face: Face) { deleted.push(face); return true; },
    *[Symbol.iterator]() {
      for (const family of declared) yield {
        family, source: '', status: 'loaded', weight: '400', style: 'normal',
      };
      for (const face of added) if (!deleted.includes(face)) yield face;
    },
  } as unknown as FontFaceSet;
  class FakeFace implements Face {
    status: FontFaceLoadStatus = 'unloaded';
    constructor(readonly family: string, readonly source: string | ArrayBuffer) {}
    async load(): Promise<this> {
      const probeSource = this.source;
      if (typeof probeSource === 'string'
        && stalled.some((name) => probeSource.includes(`local("${name}")`))) {
        return new Promise<this>(() => {});
      }
      active++;
      peak = Math.max(peak, active);
      if (delayMs) await new Promise((resolve) => setTimeout(resolve, delayMs));
      active--;
      const source = this.source;
      if (typeof source !== 'string' ||
          !installed.some((name) => source.includes(`local("${name}")`))) {
        this.status = 'error';
        throw new Error('missing local face');
      }
      this.status = 'loaded';
      return this;
    }
  }
  globals.FontFace = FakeFace;
  return { set, added, deleted, peak: () => peak };
}


afterEach(() => {
  vi.useRealTimers();
  for (const [name, value] of Object.entries(originals)) {
    if (value === undefined) delete globals[name];
    else globals[name] = value;
  }
});

describe('loadOfficeFontFallbacks', () => {
  it('leaves an unavailable Calibri face unresolved without a network request', async () => {
    const { set, added, deleted } = fontSet([]);
    let fetchCount = 0;
    globals.fetch = async () => { fetchCount++; throw new Error('network must remain unused'); };

    const loaded = await loadOfficeFontFallbacks([{ family: 'Calibri', weight: 700 }], set);

    expect(loaded).toEqual({ faces: [], routes: {}, checked: ['calibri:700:normal'] });
    expect(fetchCount).toBe(0);
    expect(added).toHaveLength(1);
    expect(deleted).toEqual(added);
  });

  it('uses an exact local styled face without borrowing a missing style', async () => {
    const { set, deleted } = fontSet(['Calibri-Bold']);
    globals.fetch = async () => { throw new Error('unexpected fetch'); };

    const result = await loadOfficeFontFallbacks([
      { family: 'Calibri', weight: 700 },
      { family: 'Calibri', weight: 700 },
      { family: 'Calibri' },
      { family: 'Calibri Light' },
    ], set);

    expect(result.faces).toHaveLength(1);
    expect(result.routes.calibri).toBeUndefined();
    expect(result.routes['calibri:700:normal']).toMatchObject({
      source: 'local', weight: 700, style: 'normal',
      resourceIdentity: 'office-local:local("Calibri-Bold")',
      metric: { sourceIdentity: 'office-local:local("Calibri-Bold")' },
    });
    expect(result.routes['calibri:700:normal'].metric.lineHeightRatio).toBeUndefined();
    unloadOfficeFontFallbacks(result.faces);
    expect(deleted).toContain(result.faces[0]);
  });

  it('never lets a catalog family name stand in for a named cut of that family', async () => {
    // "Arial Black" resolves to the catalog's "Arial" 900 profile, whose aliases
    // include "Arial". Loading local("Arial") would paint the regular face.
    const { set, added } = fontSet(['Arial', 'Arial Black']);
    globals.fetch = async () => { throw new Error('unexpected fetch'); };
    const result = await loadOfficeFontFallbacks([{ family: 'Arial Black', weight: 900 }], set);
    expect(added.length).toBeGreaterThan(0);
    for (const face of added) expect(String(face.source)).not.toContain('local("Arial")');
    expect(result.routes['arial black:900:normal']?.resourceIdentity).toContain('local("Arial Black")');
    unloadOfficeFontFallbacks(result.faces);
  });

  it('never lets a broader catalog family prove a named cut present', async () => {
    // Only regular Arial / Calibri are installed: their names must not prove
    // the absent Arial Black / Calibri Light cuts present.
    const { set } = fontSet(['Arial', 'Calibri']);
    globals.fetch = async () => { throw new Error('unexpected fetch'); };
    const result = await loadOfficeFontFallbacks([
      { family: 'Arial Black', presenceOnly: true }, { family: 'Calibri Light', presenceOnly: true },
      { family: 'Arial', presenceOnly: true },
    ], set);
    expect(result.installed).toEqual(['arial']);
  });

  it('retains one local face across concurrent holders in the same font set', async () => {
    const { set, added, deleted } = fontSet(['Calibri']);
    const [first, second] = await Promise.all([
      loadOfficeFontFallbacks([{ family: 'Calibri' }], set),
      loadOfficeFontFallbacks([{ family: 'Calibri' }], set),
    ]);
    expect(added).toHaveLength(1);
    expect(first.faces[0]).toBe(second.faces[0]);
    unloadOfficeFontFallbacks(first.faces);
    expect(deleted).toHaveLength(0);
    unloadOfficeFontFallbacks(second.faces);
    expect(deleted).toEqual(added);
  });

  it('registers an exact tuple independently in distinct main and worker font sets', async () => {
    const firstSet = fontSet(['Calibri']);
    const secondSet = fontSet(['Calibri']);

    const [main, worker] = await Promise.all([
      loadOfficeFontFallbacks([{ family: 'Calibri' }], firstSet.set),
      loadOfficeFontFallbacks([{ family: 'Calibri' }], secondSet.set),
    ]);

    expect(main.faces).toHaveLength(1);
    expect(worker.faces).toHaveLength(1);
    expect(main.faces[0]).not.toBe(worker.faces[0]);
    unloadOfficeFontFallbacks(main.faces);
    expect(firstSet.deleted).toContain(main.faces[0]);
    expect(secondSet.deleted).not.toContain(worker.faces[0]);
    unloadOfficeFontFallbacks(worker.faces);
  });

  it('resolves a catalogued authored family through a present exact face and isolates its style', async () => {
    const { set } = fontSet(['TimesNewRomanPSMT', 'TimesNewRomanPS-BoldMT']);
    const result = await loadOfficeFontFallbacks([
      { family: 'Times New Roman' },
      { family: 'Times New Roman', weight: 700 },
      { family: 'Times New Roman', style: 'italic' },
    ], set);
    expect(Object.keys(result.routes).sort()).toEqual([
      'times new roman', 'times new roman:700:normal',
    ]);
    expect(result.routes['times new roman'].resourceIdentity).toContain('TimesNewRomanPSMT');
    expect(result.routes['times new roman:700:normal'].resourceIdentity).toContain('TimesNewRomanPS-BoldMT');
    unloadOfficeFontFallbacks(result.faces);
  });

  it('resolves a translated authored family by its catalogued local name', async () => {
    const { set } = fontSet(['Meiryo']);
    const result = await loadOfficeFontFallbacks([
      { family: 'Meiryo' }, { family: 'メイリオ' },
    ], set);
    expect(result.faces).toHaveLength(1);
    expect(result.routes.meiryo).toBeDefined();
    expect(result.routes['メイリオ']).toMatchObject({
      requestedFamily: 'メイリオ', source: 'local', weight: 400,
    });
    expect(result.routes['メイリオ'].resourceIdentity).toContain('local("Meiryo")');
    unloadOfficeFontFallbacks(result.faces);
  });

  it('resolves the Japanese BIZ UD names through the English-named installed faces', async () => {
    // The macOS asset faces carry only English name records.
    const { set } = fontSet(['BIZUDMincho-Regular', 'BIZUDGothic-Bold']);
    const result = await loadOfficeFontFallbacks([
      { family: 'BIZ UD明朝' }, { family: 'BIZ UDゴシック', weight: 700 },
    ], set);
    expect(result.routes['biz ud明朝']).toMatchObject({
      requestedFamily: 'BIZ UD明朝', source: 'local', weight: 400,
    });
    expect(result.routes['biz ud明朝'].resourceIdentity).toContain('local("BIZUDMincho-Regular")');
    expect(result.routes['biz udゴシック:700:normal'].resourceIdentity)
      .toContain('local("BIZUDGothic-Bold")');
    // A family name shared by regular and bold never proves the bold face.
    expect(result.routes['biz udゴシック:700:normal'].resourceIdentity)
      .not.toContain('local("BIZ UDゴシック")');
    unloadOfficeFontFallbacks(result.faces);
  });

  it('detects an installed catalogued family whatever styles the document uses', async () => {
    // Only a bold tuple is requested, but the regular face is installed.
    const { set } = fontSet(['TimesNewRomanPSMT']);
    const result = await loadOfficeFontFallbacks([
      { family: 'Times New Roman', weight: 700 },
      { family: 'Times New Roman', presenceOnly: true },
    ], set);
    expect(result.routes['times new roman:700:normal']).toBeUndefined();
    expect(result.installed).toEqual(['times new roman']);
  });

  it('returns presence at the deadline, releases completed probes and starts no new probe', async () => {
    vi.useFakeTimers();
    const stalled = ['Stall A', 'Stall B', 'Stall C', 'Stall D'];
    const { set, added, deleted } = fontSet(['Franklin Gothic Book', 'Late Face'], 0, [], stalled);
    const pending = loadOfficeFontFallbacks(
      ['Franklin Gothic Book', ...stalled, 'Late Face'].map((family) => ({ family, presenceOnly: true })),
      set,
    );
    await vi.advanceTimersByTimeAsync(8_000);
    const result = await pending;
    expect(result.installed).toEqual(['franklin gothic book']);
    // Four concurrent probes; the completed one left the set immediately and
    // the queued "Late Face" never started.
    expect(added).toHaveLength(5);
    const completed = added.find((face) => String(face.source).includes('Franklin Gothic Book'));
    expect(deleted).toContain(completed);
    await vi.advanceTimersByTimeAsync(20_000);
    expect(added).toHaveLength(5);
    expect(added.filter((face) => !deleted.includes(face))).toHaveLength(0);
  });

  it('probes every family for font-source demand, past the per-call source cap', async () => {
    // 40 uncatalogued families, the last one installed: one presence call
    // stops at the cap, the batched probe reaches it.
    const families = Array.from({ length: 40 }, (_, index) => `Face ${index}`);
    const { set } = fontSet(['Face 39']);
    const capped = await loadOfficeFontFallbacks(families.map((family) => ({ family, presenceOnly: true })), set);
    expect(capped.installed).toBeUndefined();
    const presence = await probeFontFamilyPresence(families, set);
    expect([...presence.installed]).toEqual(['face 39']);
    expect(presence.checked.size).toBe(40);
  });

  it('reports a family whose probe missed the deadline as unchecked, not missing', async () => {
    vi.useFakeTimers();
    const stalled = ['Stall A', 'Stall B', 'Stall C', 'Stall D'];
    const { set } = fontSet(['Late Face'], 0, [], stalled);
    const pending = probeFontFamilyPresence([...stalled, 'Late Face', 'Missing Face'], set);
    await vi.advanceTimersByTimeAsync(8_000);
    const presence = await pending;
    expect(presence.installed.size).toBe(0);
    expect(presence.checked.size).toBe(0);
  });

  it('probes uncatalogued installed families by exact name and releases the probe faces', async () => {
    const { set, added, deleted } = fontSet(['Franklin Gothic Book']);
    const result = await loadOfficeFontFallbacks([
      { family: 'Franklin Gothic Book', presenceOnly: true },
      { family: 'Missing Face', presenceOnly: true },
    ], set);
    expect(result.installed).toEqual(['franklin gothic book']);
    expect(result.routes).toEqual({});
    // No route is created: every probe face leaves the FontFaceSet again.
    expect(added.filter((face) => !deleted.includes(face))).toHaveLength(0);
  });

  it('bounds independent local probes and cleans up failed faces', async () => {
    const { set, added, deleted, peak } = fontSet([], 3);
    const result = await loadOfficeFontFallbacks([
      'Calibri', 'Times New Roman', 'Meiryo', 'Arial', 'Courier New', 'Georgia', 'Cambria',
    ].map((family) => ({ family })), set);
    expect(result.routes).toEqual({});
    expect(added.length).toBeGreaterThan(4);
    expect(peak()).toBeGreaterThan(1);
    expect(peak()).toBeLessThanOrEqual(4);
    expect(deleted).toEqual(added);
  });

  it('leaves an application-declared authored family to normal CSS resolution', async () => {
    const { set, added } = fontSet(['TimesNewRomanPSMT'], 0, ['"Times New Roman"']);
    const result = await loadOfficeFontFallbacks([{ family: 'Times New Roman' }], set);
    expect(result).toEqual({ faces: [], routes: {}, checked: [] });
    expect(added).toEqual([]);
  });

  it('preflights a bold local face when the application declares only regular', async () => {
    const { set } = fontSet(['Calibri-Bold'], 0, ['Calibri']);
    const result = await loadOfficeFontFallbacks([{ family: 'Calibri', weight: 700 }], set);
    expect(result.routes['calibri:700:normal']).toMatchObject({ source: 'local', weight: 700 });
    unloadOfficeFontFallbacks(result.faces);
  });

  it('keeps italic and weight-range declarations confined to their own tuples', async () => {
    const { set, added } = fontSet(['Calibri']);
    const declaredSet = {
      ...set,
      *[Symbol.iterator]() {
        yield { family: 'Calibri', weight: '400', style: 'italic', status: 'loaded' };
        yield { family: 'Calibri', weight: '600 800', style: 'normal', status: 'loaded' };
        for (const face of set) yield face;
      },
    } as FontFaceSet;
    const result = await loadOfficeFontFallbacks([
      { family: 'Calibri' },
      { family: 'Calibri', weight: 700 },
      { family: 'Calibri', style: 'italic' },
    ], declaredSet);
    expect(result.checked).toEqual(['calibri']);
    expect(Object.keys(result.routes)).toEqual(['calibri']);
    expect(added).toHaveLength(1);
    unloadOfficeFontFallbacks(result.faces);
  });

  it('does not let an application face still loading suppress exact local preflight', async () => {
    const { set, added } = fontSet(['Calibri']);
    const loadingSet = {
      ...set,
      *[Symbol.iterator]() {
        yield { family: 'Calibri', weight: '400', style: 'normal', status: 'loading' };
        for (const face of set) yield face;
      },
    } as FontFaceSet;
    const result = await loadOfficeFontFallbacks([{ family: 'Calibri' }], loadingSet);
    expect(added).toHaveLength(1);
    expect(result.routes.calibri).toMatchObject({ source: 'local' });
    unloadOfficeFontFallbacks(result.faces);
  });

  it('caps the number of local source probes for a font-heavy document', async () => {
    const { set, added, deleted } = fontSet([]);
    const families = [...new Set(referenceData.profiles
      .filter((profile) => profile.weight === 400 && profile.style === 'normal')
      .map((profile) => profile.family))].slice(0, 80);
    await loadOfficeFontFallbacks(families.map((family) => ({ family })), set);
    expect(families.length).toBe(80);
    expect(added.length).toBeGreaterThan(0);
    expect(added.length).toBeLessThanOrEqual(32);
    expect(deleted).toEqual(added);
  });

  it('returns at the document deadline and releases a face that finishes later', async () => {
    vi.useFakeTimers();
    const added: Face[] = [];
    const deleted: Face[] = [];
    let finish!: () => void;
    class DeferredFace implements Face {
      status: FontFaceLoadStatus = 'unloaded';
      constructor(readonly family: string, readonly source: string | ArrayBuffer) {}
      load(): Promise<this> {
        return new Promise((resolve) => {
          finish = () => { this.status = 'loaded'; resolve(this); };
        });
      }
    }
    globals.FontFace = DeferredFace;
    const set = {
      add(face: Face) { added.push(face); },
      delete(face: Face) { deleted.push(face); return true; },
    } as unknown as FontFaceSet;
    const pending = loadOfficeFontFallbacks([{ family: 'Calibri' }], set);
    await vi.advanceTimersByTimeAsync(8_000);
    expect(await pending).toEqual({ faces: [], routes: {}, checked: [] });
    expect(added).toHaveLength(1);
    finish();
    await vi.advanceTimersByTimeAsync(0);
    expect(deleted).toEqual(added);
  });
});
